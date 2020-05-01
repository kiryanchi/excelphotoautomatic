# -*- coding:utf-8 -*-
from openpyxl.drawing.image import Image

'''
    엑셀 파일 이름을 넘기면, 이미지를 폴더에서 찾은 다음, 
'''
FIXED_SIZE = (515.527559055, 335.622047244)


def insertinexcel(img_name, column, row, sheet):
    try:
        print(img_name)
        img = Image(img_name)
    except FileNotFoundError:
        return '[에러 3] 사진을 찾을 수 없습니다.'
    else:
        img.width, img.height = FIXED_SIZE
        sheet.add_image(img, column + str(row))
        return '[완료] ' + str(column) + str(row) + '에 사진을 넣었습니다.'