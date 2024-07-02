import os.path

from .properties import parse_properties


def process_attachment(msg, entry, doc):
    # Load attachment stream.
    props = parse_properties(entry["__properties_version1.0"], False, entry, doc)

    # The attachment content...
    blob = props["ATTACH_DATA_BIN"]

    # Get the filename and MIME type of the attachment.
    filename = (
        props.get("ATTACH_LONG_FILENAME")
        or props.get("ATTACH_FILENAME")
        or props.get("DISPLAY_NAME")
    )
    if isinstance(filename, bytes):
        filename = filename.decode("utf8")

    mime_type = props.get("ATTACH_MIME_TAG", "application/octet-stream")
    if isinstance(mime_type, bytes):
        mime_type = mime_type.decode("utf8")

    filename = os.path.basename(filename)

    # Python 3.6.
    if isinstance(blob, str):
        msg.add_attachment(blob, filename=filename)
    elif isinstance(blob, bytes):
        msg.add_attachment(
            blob,
            maintype=mime_type.split("/", 1)[0],
            subtype=mime_type.split("/", 1)[-1],
            filename=filename,
        )
    else:  # a Message instance
        msg.add_attachment(blob, filename=filename)
