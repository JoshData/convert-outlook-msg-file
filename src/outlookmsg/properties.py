import logging

from .types import (
    code_pages,
    property_tags,
    property_types,
)

from .embedded_msg import EMBEDDED_MESSAGE
from .value_loaders import (
    FixedLengthValueLoader,
    VariableLengthValueLoader,
)

logger = logging.getLogger(__name__)


def parse_properties(properties, is_top_level, container, doc):
    # Read a properties stream and return a Python dictionary
    # of the fields and values, using human-readable field names
    # in the mapping at the top of this module.

    # Load stream content.
    with doc.open(properties) as stream:
        stream = stream.read()

    # Skip header.
    i = 32 if is_top_level else 24

    # Read 16-byte entries.
    raw_properties = {}
    while i < len(stream):
        # Read the entry.
        property_type = stream[i + 0 : i + 2]
        property_tag = stream[i + 2 : i + 4]
        value = stream[i + 8 : i + 16]
        i += 16

        # Turn the byte strings into numbers and look up the property type.
        property_type = property_type[0] + (property_type[1] << 8)
        property_tag = property_tag[0] + (property_tag[1] << 8)
        if property_tag not in property_tags:
            continue  # should not happen
        tag_name, _ = property_tags[property_tag]
        tag_type = property_types.get(property_type)

        if tag_name == 'DISPLAY_BCC':
            pass  #debug breakpoint

        # Fixed Length Properties.
        if isinstance(tag_type, FixedLengthValueLoader):
            # The value comes from the stream above.
            pass

        # Variable Length Properties.
        elif isinstance(tag_type, VariableLengthValueLoader):
            # Look up the stream in the document that holds the value.
            streamname = "__substg1.0_{0:0{1}X}{2:0{3}X}".format(
                property_tag, 4, property_type, 4
            )
            try:
                with doc.open(container[streamname]) as innerstream:
                    value = innerstream.read()
            except Exception:
                # Stream isn't present!
                logger.error("stream missing {}".format(streamname))
                continue

        elif isinstance(tag_type, EMBEDDED_MESSAGE):
            # Look up the stream in the document that holds the attachment.
            streamname = "__substg1.0_{0:0{1}X}{2:0{3}X}".format(
                property_tag, 4, property_type, 4
            )
            try:
                value = container[streamname]
            except Exception:
                # Stream isn't present!
                logger.error("stream missing {}".format(streamname))
                continue

        else:
            # unrecognized type
            logger.error("unhandled property type {}".format(hex(property_type)))
            continue

        raw_properties[tag_name] = (tag_type, value)

    # Decode all FixedLengthValueLoader properties so we have codepage
    # properties.
    properties = {}
    for tag_name, (tag_type, value) in raw_properties.items():
        if not isinstance(tag_type, FixedLengthValueLoader):
            continue
        try:
            properties[tag_name] = tag_type.load(value)
        except Exception as e:
            logger.error("Error while reading stream: {}".format(str(e)))

    # String8 strings use code page information stored in other
    # properties, which may not be present. Find the Python
    # encoding to use.

    # The encoding of the "BODY" (and HTML body) properties.
    body_encoding = None
    if (
        "PR_INTERNET_CPID" in properties
        and properties["PR_INTERNET_CPID"] in code_pages
    ):
        body_encoding = code_pages[properties["PR_INTERNET_CPID"]]

    # The encoding of "string properties of the message object".
    properties_encoding = None
    if (
        "PR_MESSAGE_CODEPAGE" in properties
        and properties["PR_MESSAGE_CODEPAGE"] in code_pages
    ):
        properties_encoding = code_pages[properties["PR_MESSAGE_CODEPAGE"]]

    # Decode all of the remaining properties.
    for tag_name, (tag_type, value) in raw_properties.items():
        if isinstance(tag_type, FixedLengthValueLoader):
            continue  # already done, above

        # The codepage properties may be wrong. Fall back to
        # the other property if present.
        encodings = (
            [body_encoding, properties_encoding]
            if tag_name == "BODY"
            else [properties_encoding, body_encoding]
        )

        try:
            properties[tag_name] = tag_type.load(value, encodings=encodings, doc=doc)
        except KeyError as e:
            logger.error("Error while reading stream: {} not found".format(str(e)))
        except Exception as e:
            logger.error("Error while reading stream: {}".format(str(e)))

    return properties
