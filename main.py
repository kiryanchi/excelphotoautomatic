# -*- coding:utf-8 -*-
from image.insert import insert_image
from menu.printmenu import print_menu
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def run():
    print_menu()
    file_name = input('파일 이름: ')
    file_name = os.path.basename(BASE_DIR + '\\' + file_name + '.' + 'xlsx')
    insert_image(file_name, BASE_DIR)
    print('작업을 완료했습니다.')
    os.system('Pause')


if __name__ == '__main__':
    run()



