#!/usr/bin/env python
#-*- coding: utf8 -*-

# -----------------------------------------------------------------------------------------
# This is a Python 3 port of compressed_rtf at https://github.com/delimitry/compressed_rtf,
# which is MIT licensed.
# -----------------------------------------------------------------------------------------

"""
Compressed Rich Text Format (RTF) worker

Based on Rich Text Format (RTF) Compression Algorithm
https://msdn.microsoft.com/en-us/library/cc463890(v=exchg.80).aspx
"""

import struct
from crc32 import crc32
from io import BytesIO

__all__ = ['decompress', 'compress']

INIT_DICT = (
    b'{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\'
    b'fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New '
    b'RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\'
    b'f0\\fs20\\b\\i\\u\\tab\\tx'
)

INIT_DICT_SIZE = 207
MAX_DICT_SIZE = 4096

COMPRESSED = b'LZFu'
UNCOMPRESSED = b'MELA'


def decompress(data):
    """
    Decompress data
    """
    # set init dict
    init_dict = list(INIT_DICT)
    init_dict += ' ' * (MAX_DICT_SIZE - INIT_DICT_SIZE)
    if len(data) < 16:
        raise Exception('Data must be at least 16 bytes long')
    write_offset = INIT_DICT_SIZE
    output_buffer = BytesIO()
    # make stream
    in_stream = BytesIO(data)
    # read compressed RTF header
    comp_size = struct.unpack('<I', in_stream.read(4))[0]
    raw_size = struct.unpack('<I', in_stream.read(4))[0]
    comp_type = in_stream.read(4)
    crc_value = struct.unpack('<I', in_stream.read(4))[0]
    # get only data
    contents = BytesIO(in_stream.read(comp_size - 12))
    if comp_type == COMPRESSED:
        # check CRC
        if crc_value != crc32(contents.read()):
            raise Exception('CRC is invalid! The file is corrupt!')
        contents.seek(0)
        end = False
        while not end:
            val = contents.read(1)
            if not val:
                break
            control = '{0:08b}'.format(ord(val))
            # check bits from LSB to MSB
            for i in range(1, 9):
                if control[-i] == '1':
                    # token is reference (16 bit)
                    val = contents.read(2)
                    if not val:
                        break
                    token = struct.unpack('>H', val)[0]  # big-endian
                    # extract [12 bit offset][4 bit length]
                    offset = (token >> 4) & 0b111111111111
                    length = token & 0b1111
                    # end indicator
                    if write_offset == offset:
                        end = True
                        break
                    actual_length = length + 2
                    for step in range(actual_length):
                        read_offset = (offset + step) % MAX_DICT_SIZE
                        char = init_dict[read_offset]
                        output_buffer.write(bytes([char]))
                        init_dict[write_offset] = char
                        write_offset = (write_offset + 1) % MAX_DICT_SIZE
                else:
                    # token is literal (8 bit)
                    val = contents.read(1)
                    if not val:
                        break
                    output_buffer.write(val)
                    init_dict[write_offset] = val[0]
                    write_offset = (write_offset + 1) % MAX_DICT_SIZE
    elif comp_type == UNCOMPRESSED:
        return contents.read(raw_size)
    else:
        raise Exception('Unknown type of RTF compression!')
    return output_buffer.getvalue()


def compress(data, compressed=True):
    """
    Compress `data` with `compressed` flag
    """
    output_buffer = ''
    # set init dict
    init_dict = list(INIT_DICT + ' ' * (MAX_DICT_SIZE - INIT_DICT_SIZE))
    write_offset = INIT_DICT_SIZE
    # compressed
    if compressed:
        comp_type = COMPRESSED
        # make stream
        in_stream = BytesIO(data)
        # init params
        control_byte = 0
        control_bit = 1
        token_offset = 0
        token_buffer = ''
        match_len = 0
        longest_match = 0
        while True:
            # find longest match
            dict_offset, longest_match, write_offset = \
                _find_longest_match(init_dict, in_stream, write_offset)
            char = in_stream.read(longest_match if longest_match > 1 else 1)
            # EOF input stream
            if not char:
                # update params
                control_byte |= 1 << control_bit - 1
                control_bit += 1
                token_offset += 2
                # add dict reference
                dict_ref = (write_offset & 0xfff) << 4
                token_buffer += struct.pack('>H', dict_ref)
                # add to output
                output_buffer += struct.pack('B', control_byte)
                output_buffer += token_buffer[:token_offset]
                break
            else:
                if longest_match > 1:
                    # update params
                    control_byte |= 1 << control_bit - 1
                    control_bit += 1
                    token_offset += 2
                    # add dict reference
                    dict_ref = (dict_offset & 0xfff) << 4 | (
                        longest_match - 2) & 0xf
                    token_buffer += struct.pack('>H', dict_ref)
                else:
                    # character is not found in dictionary
                    if longest_match == 0:
                        init_dict[write_offset] = char
                        write_offset = (write_offset + 1) % MAX_DICT_SIZE
                    # update params
                    control_byte |= 0 << control_bit - 1
                    control_bit += 1
                    token_offset += 1
                    # add literal
                    token_buffer += char
                longest_match = 0
                if control_bit > 8:
                    # add to output
                    output_buffer += struct.pack('B', control_byte)
                    output_buffer += token_buffer[:token_offset]
                    # reset params
                    control_byte = 0
                    control_bit = 1
                    token_offset = 0
                    token_buffer = ''
    else:
        # if uncompressed - copy data to output
        comp_type = UNCOMPRESSED
        output_buffer = data
    # write compressed RTF header
    comp_size = struct.pack('<I', len(output_buffer) + 12)
    raw_size = struct.pack('<I', len(data))
    crc_value = struct.pack('<I', crc32(output_buffer))
    return comp_size + raw_size + comp_type + crc_value + output_buffer


def _find_longest_match(init_dict, stream, write_offset):
    """
    Find the longest match
    """
    # read the first char
    char = stream.read(1)
    if not char:
        return 0, 0, write_offset
    prev_write_offset = write_offset
    dict_index = 0
    match_len = 0
    longest_match_len = 0
    dict_offset = 0
    # find the first char
    while True:
        if init_dict[dict_index % MAX_DICT_SIZE] == char:
            match_len += 1
            # if found longest match
            if match_len <= 17 and match_len > longest_match_len:
                dict_offset = dict_index - match_len + 1
                # add to dictionary and update longest match
                init_dict[write_offset] = char
                write_offset = (write_offset + 1) % MAX_DICT_SIZE
                longest_match_len = match_len
            # read the next char
            char = stream.read(1)
            if not char:
                stream.seek(stream.tell() - match_len, 0)
                return dict_offset, longest_match_len, write_offset
        else:
            stream.seek(stream.tell() - match_len - 1, 0)
            match_len = 0
            # read the first char
            char = stream.read(1)
            if not char:
                break
        dict_index += 1
        if dict_index >= prev_write_offset + longest_match_len:
            break
    stream.seek(stream.tell() - match_len - 1, 0)
    return dict_offset, longest_match_len, write_offset
