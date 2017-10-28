class NoAttachmentException(Exception):
    def __init__(self, value=None):
        Exception.__init__(self, value)

