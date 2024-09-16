class DownloadPdfError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
        
class MiroNotFoundError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
        
class ValidarChaveAcessoError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class VerificQuantDocumentsError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
    
class NoDocuments(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)