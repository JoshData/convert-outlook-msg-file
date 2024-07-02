class EMBEDDED_MESSAGE(object):
    @staticmethod
    def load(entry, doc, **kwargs):
        from .stream import load_message_stream

        return load_message_stream(entry, False, doc)
