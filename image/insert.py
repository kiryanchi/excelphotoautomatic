# -*- coding:utf-8 -*-
import openpyxl
from openpyxl.drawing.image import Image
from .resize import resizing
from handling.error import ResizeError
import os

'''
    엑셀 파일 이름을 넘기면, 이미지를 폴더에서 찾은 다음, 
'''


def insert(folder_name, column, row, sheet):
    """
    폴더를 골라 사진을 삽입하는 함수
    :param folder_name: string
    :param sheet: sheet
    :param row: list
    :param column: string
    :return: bool
    """
    for r in range(row):
        r += 1
        try:
            img = Image(folder_name + '\\' + str(r) + '.jpg')
            # print(folder_name)
        except FileNotFoundError:
            print("'" + folder_name + '\' 폴더를 찾을 수 없습니다.')
            print(r, '번 째 사진을 찾지 못 했습니다.')
            raise FileNotFoundError
        else:
            try:
                resizing(img)
            except ResizeError:
                print(r, '번 째 사진을 찾지 못 했습니다.')
            else:
                sheet.add_image(img, column + str((r-1) * 19 + 2))
                print(folder_name + '폴더의 ', r, '번 째 사진을 추가했습니다.')


def insert_image(file, BASE_DIR, output='output.xlsx'):
    """
    전, 후 폴더에서 사진을 불러와 리사이즈 후 사진을 삽입한다.
    :param output: output
    :param file: 파일 이름
    :return:
    """
    before_folder = BASE_DIR + '\\전'
    after_folder = BASE_DIR + '\\후'
    row, sheet = 0, 0
    # print(before_folder)
    # print(after_folder)
    while True:
        try:
            wb = openpyxl.load_workbook(file)
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다. 만약 파일이 있다면 파일 이름을 확인해주세용.")
            file = input('파일 이름: ')
            file = os.path.basename(BASE_DIR + '\\' + file + '.' + 'xlsx')
        else:
            print('\n' + file + '파일을 찾았습니다.\n')
            sheet = wb['작업사진']
            row = (sheet.max_row - 1) // 19
            break

        # 사진 추가
    try:
        insert(before_folder, 'A', row, sheet)
    except FileNotFoundError:
        print("\n'전' 폴더 사진 추가를 실패했습니다.\n")
        return
    else:
        try:
            print('\n전 폴더 사진 추가를 완료했습니다.\n')
            insert(after_folder, 'I', row, sheet)
        except FileNotFoundError:
            print("\n'후' 폴더 사진 추가를 실패했습니다.\n")
            return

    wb.save(output)


# def delete_image(file):
#     wb = openpyxl.load_workbook(file)
#     sheet = wb['작업사진']
#     imagelist = sheet._images
#     del imagelist[0]
#     del imagelist[2]
#     del imagelist[3]
#     del imagelist[5]
#     wb.save('test.xlsx')
