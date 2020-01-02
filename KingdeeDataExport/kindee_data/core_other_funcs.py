import os
import re
import configparser
import msvcrt
import datetime


def exit_with_anykey():
    print("按任意键退出！")
    ord(msvcrt.getch())
    os._exit(1)


def get_config(key, value):
    cf = configparser.ConfigParser()
    cf.read('config.txt')
    return cf.get(key, value)


def mkdir(path):

    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)
    print('检查路径...')
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        return path
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(path+' 目录已存在！')
        exit_with_anykey()


def change_dir(path):
    try:
        os.chdir(path)
    except Exception as e:
        print(e)
        exit_with_anykey()


def convert_str(stri):
    li = list(stri)
    li.insert(0, "'")
    li.append("'")
    ii = ''.join(li)
    return ii


def emb_numbers(s):

    re_digits = re.compile(r'(\d+)')
    pieces = re_digits.split(s)
    pieces[1::2] = map(int, pieces[1::2])
    return pieces


def get_current_time():
    return datetime.datetime.now().strftime("%Y%m%d%H%M%S")

def check_file():  # 检查文件是否存在
    if os.path.exists('config.txt'):
        pass
    else:
        print("未找到配置文件！")
        exit_with_anykey()