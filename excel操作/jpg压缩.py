# coding=utf-8
from PIL import Image
import glob, os


# 处理图片大小信息为kb
def get_FileSize(picture):
    f1 = os.path.getsize(picture)
    f2 = f1 / float(1024)
    return f2


# 按比例减小分辨率
def small(SIZE):
    SIZE *= 0.99
    return SIZE


a = float(input("图片大小限制kb："))

# 处理过程
for infile in glob.glob('*.jpg'):
    file, ext = os.path.splitext(infile)
    im = Image.open(infile)

    # 获取图片像素
    b = get_FileSize(infile)
    size = im.size
    size1, size2 = size

    # 循环至达到要求
    while b > a:
        size1, size2 = small(size1), small(size2)
        size = size1, size2
        im.thumbnail(size, Image.ANTIALIAS)
        im.save(file + 'add.jpg', 'jpeg')
        b = get_FileSize(file + 'add.jpg')
