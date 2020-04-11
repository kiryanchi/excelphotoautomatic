from handling.error import ResizeError

'''
    사진 넣기 전 사진 크기 조절하는 함수
'''


def resizing(img):
    fixed_size = (514.0, 335.6)
    try:
        img.width, img.height = fixed_size
    except FileNotFoundError:
        raise ResizeError
