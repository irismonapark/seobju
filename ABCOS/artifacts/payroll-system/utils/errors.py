class UserFacingError(Exception):
    """사용자에게 그대로 보여줄 수 있는 변환/입력 오류."""

    def __init__(self, message: str):
        self.message = message
        super().__init__(message)
