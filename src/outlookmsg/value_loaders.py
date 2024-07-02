from functools import reduce
import logging

logger = logging.getLogger(__name__)

FALLBACK_ENCODING = "cp1252"


class FixedLengthValueLoader(object):
    pass


class NULL(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring with unused content.
        return None


class BOOLEAN(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding a two-byte integer.
        return value[0] == 1


class INTEGER16(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding a two-byte integer.
        return reduce(lambda a, b: (a << 8) + b, reversed(value[0:2]))


class INTEGER32(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding a four-byte integer.
        return reduce(lambda a, b: (a << 8) + b, reversed(value[0:4]))


class INTEGER64(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring holding an eight-byte integer.
        return reduce(lambda a, b: (a << 8) + b, reversed(value))


class INTTIME(FixedLengthValueLoader):
    @staticmethod
    def load(value):
        # value is an eight-byte long bytestring encoding the integer number of
        # 100-nanosecond intervals since January 1, 1601.
        from datetime import datetime, timedelta

        value = reduce(
            lambda a, b: (a << 8) + b, reversed(value)
        )  # bytestring to integer
        try:
            value = datetime(1601, 1, 1) + timedelta(seconds=value / 10000000)
        except OverflowError:
            value = None

        return value


# TODO: The other fixed-length data types:
# "FLOAT", "DOUBLE", "CURRENCY", "APPTIME", "ERROR"


class VariableLengthValueLoader(object):
    pass


class BINARY(VariableLengthValueLoader):
    @staticmethod
    def load(value, **kwargs):
        # value is a bytestring. Just return it.
        return value


class STRING8(VariableLengthValueLoader):
    @staticmethod
    def load(value, encodings, **kwargs):
        # Value is a "bytestring" and encodings is a list of Python
        # codecs to try. If all fail, try the fallback codec with
        # character replacement so that this never fails.
        for encoding in encodings:
            try:
                return value.decode(encoding=encoding, errors="strict")
            except Exception:
                # Try the next one.
                pass
        return value.decode(encoding=FALLBACK_ENCODING, errors="replace")


class UNICODE(VariableLengthValueLoader):
    @staticmethod
    def load(value, **kwargs):
        # value is a bytestring encoded in UTF-16.
        decoded: str = value.decode("utf16")
        # do c-style strings get encoded as utf16?
        # is there an off-by-one error in the variable length math?
        decoded =  decoded.removesuffix('\x00')
        return decoded


# TODO: The other variable-length tag types are "CLSID", "OBJECT".
