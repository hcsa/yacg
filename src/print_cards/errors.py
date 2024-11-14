class CardPrintError(ValueError):
    """
    Raised when something unexpected happened while printing a card
    """

    def __init__(self, message):
        super().__init__(message)


class IllustratorTemplateError(CardPrintError):
    """
    Raised when an unexpected element or setting of an Illustrator template is found
    """

    def __init__(self, message):
        super().__init__(message)
