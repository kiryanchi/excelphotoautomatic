class ResizeError(Exception):
    def __init__(self):
        super().__init__('사진 크기 변환에 실패했습니다.')