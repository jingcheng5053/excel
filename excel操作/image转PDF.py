from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, PageBreak
from reportlab.lib.pagesizes import A4, A3, A2, A1, legal, landscape
from reportlab.lib.utils import ImageReader
# import PIL.Image, PIL.ExifTags
from PIL import Image, ImageTk
from os import listdir
import os, re
import time
from tkinter import Frame, StringVar, IntVar, Label, Entry, LEFT, \
    Button, Tk, mainloop, Menu, Listbox, TOP, END
from tkinter.filedialog import askopenfilename, askdirectory  # 文件打开对话框
from tkinter import messagebox as msgbox
import copy

from reportlab.lib.units import inch


class ImageToPDF:
    def __init__(self):
        print(A4)
        self.A4 = (794, 1123)
        self.A4 = (int(A4[0]), int(A4[1]))
        self.A4_row = (self.A4[1], self.A4[0])
        self.path = ""
        self.filenames = []
        '''用于生成主界面用于填写'''
        self.top = Tk()
        self.sw = self.top.winfo_screenwidth()
        self.sh = self.top.winfo_screenheight()
        self.topw = 650
        self.toph = 500
        self.top.title('图片转pdf生成器')
        self.top.geometry("%dx%d+%d+%d" % (self.topw, self.toph, (self.sw - self.topw) / 2, (self.sh - self.toph) / 2))

        # <editor-fold desc="---菜单模块---">
        menubar = Menu(self.top)
        # 添加菜单条
        self.top['menu'] = menubar
        # 创建file_menu菜单，它被放入menubar中
        file_menu = Menu(menubar, tearoff=0)
        # 使用add_cascade方法添加file_menu菜单
        menubar.add_cascade(label='文件', menu=file_menu)
        # 创建lang_menu菜单，它被放入menubar中
        lang_menu = Menu(menubar, tearoff=0)
        help_menu = Menu(menubar, tearoff=0)
        # 使用add_cascade方法添加lang_menu菜单
        menubar.add_cascade(label='选择语言', menu=lang_menu)
        menubar.add_cascade(label='帮助', menu=help_menu)

        # 使用add_command方法为file_menu添加菜单项
        file_menu.add_command(label="新建", command=None,
                              image="", compound=LEFT)
        file_menu.add_command(label="打开", command=None,
                              image="", compound=LEFT)
        help_menu.add_command(label="版本", command=None, compound=LEFT)
        # 使用add_command方法为file_menu添加分隔条
        file_menu.add_separator()
        # 为file_menu创建子菜单
        sub_menu = Menu(file_menu, tearoff=0)
        # 使用add_cascade方法添加sub_menu子菜单
        file_menu.add_cascade(label='选择性别', menu=sub_menu)
        self.genderVar = IntVar()
        # 使用循环为sub_menu子菜单添加菜单项
        for i, im in enumerate(['男', '女', '保密']):
            # 使用add_radiobutton方法为sub_menu子菜单添加单选菜单项
            # 绑定同一个变量，说明它们是一组
            sub_menu.add_radiobutton(label=im, command=self.choose_gender,
                                     variable=self.genderVar, value=i)
        self.langVars = [StringVar(), StringVar(), StringVar(), StringVar()]
        # 使用循环为lang_menu菜单添加菜单项
        for i, im in enumerate(('Python', 'Kotlin', 'Swift', 'Java')):
            # 使用add_add_checkbutton方法为lang_menu菜单添加多选菜单项
            lang_menu.add_checkbutton(label=im, command=self.choose_lang,
                                      onvalue=im, variable=self.langVars[i])
        # </editor-fold>

        self._DIRPATH = StringVar(self.top)

        self.emptfmone = Frame(self.top, height=50)
        self.emptfmone.pack()

        self.dirfm = Frame(self.top)
        self.descriptLabel = Label(self.dirfm, width=4, text='路径：')
        self.descriptLabel.pack(side=LEFT)
        self.dirn = Entry(self.dirfm, width=50, textvariable=self._DIRPATH)
        # self.dirn.bind('<Return>', self.setPath)
        self.dirn.pack(side=LEFT)
        self.dirfm.pack()

        self.emptfmtwo = Frame(self.top, height=30)
        self.emptfmtwo.pack()

        self.btnfm = Frame(self.top)
        self.converBtn = Button(self.btnfm, width=10, text='读取', command=self.img_search,
                                activeforeground='white', activebackground='blue')
        self.save_Btn = Button(self.btnfm, width=10, text='保存', command=self.save_images_to_pdf,
                               activeforeground='white', activebackground='blue')
        self.quitBtn = Button(self.btnfm, width=10, text='退出', command=self.top.quit, activeforeground='white',
                              activebackground='blue')
        self.converBtn.pack(side=LEFT, padx=10)
        self.save_Btn.pack(side=LEFT, padx=10)
        self.quitBtn.pack(side=LEFT, padx=10)
        self.btnfm.pack()
        show_frame = Frame(self.top).pack(side=TOP)
        self.list_box_var = StringVar()
        self.list_box = Listbox(show_frame, listvariable=self.list_box_var, width=50)
        self.list_box.pack(side=LEFT)
        # self.list_box.bind('<Double-Button-1>', self.CallOn)
        self.list_box.bind('<ButtonRelease-1>', self.CallOn)
        self.img_label = Label(show_frame, text="ddddd", image="")
        self.img_label.pack(side=LEFT)
        self.show_img(self.new_a4_img())
        self.pop_root1 = None
        pass

    def choose_gender(self):
        msgbox.showinfo(message=('选择的性别为: %s' % self.genderVar.get()))

    def choose_lang(self):
        rt_list = [e.get() for e in self.langVars]
        msgbox.showinfo(message=('选择的语言为: %s' % ','.join(rt_list)))

    def img_search(self):
        mypath = askdirectory()
        mypath = self.converPath(mypath)
        print(mypath)
        """读取指定文件夹下所有的JPEG图片，存入列表"""
        if mypath is None or len(mypath) == 0:
            raise ValueError('dirPath不能为空，该值为存放图片的具体路径文件夹！')
        if os.path.isfile(mypath):
            raise ValueError('dirPath不能为具体文件，该值为存放图片的具体路径文件夹！')
        self.path = mypath
        self._DIRPATH.set(mypath)
        self.filenames = []
        for imageName in os.listdir(mypath):
            if imageName.endswith('.jpg') or imageName.endswith('.jpeg') or imageName.endswith('.png'):
                path = os.path.join(mypath, imageName)
                print(path)
                self.filenames.append(path)
        # self.img_to_pdf()
        # self.cover_to_pdf(self.path, self.filenames)
        self.read_to_arr(self.filenames)

    def new_a4_img(self, row=0) -> Image.Image:
        if row == 0:
            return Image.new('RGB', self.A4, (255, 255, 255))
        else:
            return Image.new('RGB', self.A4_row, (255, 255, 255))

    def show_img(self, a4img: Image.Image):
        imgin4 = a4img.resize((int(a4img.size[0] * 0.25), int(a4img.size[1] * 0.25)), Image.ANTIALIAS)
        imgin4 = ImageTk.PhotoImage(image=imgin4)
        self.img_label.config(image=imgin4)
        self.img_label.image = imgin4  # keep a reference

    def converPath(self, dirPath):
        '''用于转换路径，判断路径后是否为\\，如果有则直接输出，如果没有则添加'''
        if dirPath is None or len(dirPath) == 0:
            raise ValueError('dirPath不能为空！')
        if os.path.isfile(dirPath):
            raise ValueError('dirPath不能为具体文件，该值为文件夹路径！')
        if not str(dirPath).endswith("\\"):
            return dirPath + "\\"
        return dirPath

    def get_resize(self, max_width, max_height, image_width, image_height):
        basewidth = max_width

        wpercent = (basewidth / float(image_width))
        hsize = int((float(image_height) * float(wpercent)))
        if hsize > max_height:
            hsize = max_height
        return basewidth, hsize

    # 按比例缩放至背景大小
    def get_min_resize(self, max_width, max_height, image_width, image_height):
        print("修改前:", max_width, max_height, image_width, image_height)
        if max_width / max_height > image_width / image_height:
            r_height = max_height
            r_width = int(max_height / image_height * image_width)
        else:
            r_width = max_width
            r_height = int(max_width / image_width * image_height)
        print(r_width, r_height)
        return r_width, r_height

    def CallOn(self, event):
        if self.list_box.size() == 0:
            return
        # if self.pop_root1:
        #     self.pop_root1.destroy()
        print(self.list_box.curselection()[0])
        img = self.images[self.list_box.curselection()[0]]
        img = self.img_in_page(img)
        self.show_img(img)
        # self.pop_root1 = Tk()
        # Label(self.pop_root1, text='你的选择是' + self.list_box.get(self.list_box.curselection()) + "语言！").pack()
        # Button(self.pop_root1, text='退出', command=self.pop_root1.destroy).pack()

    def img_in_page(self, img: Image.Image) -> Image.Image:
        if img.size[0] < img.size[1]:
            resize_page_size = self.A4
            page_size = self.A4
        else:
            resize_page_size = self.A4_row
            page_size = self.A4_row
        a4im = Image.new('RGB',
                         page_size,
                         (255, 255, 255))
        print(page_size, img.size)
        image_resize = self.get_min_resize(resize_page_size[0], resize_page_size[1], img.size[0], img.size[1])
        print("需要修改的图片大小:", img.size, "-->", image_resize)
        img = img.resize(image_resize, Image.ANTIALIAS)
        image_site = int(page_size[0] - image_resize[0]) // 2, int(page_size[1] - image_resize[1]) // 2
        a4im.paste(img, image_site)
        print("修改后:页面信息,图片信息:", page_size, img.size)
        return a4im

    def read_to_arr(self, images_path):
        images_path = copy.deepcopy(images_path)
        images = []
        self.list_box_var.set(())
        # self.list_box.delete(0, self.list_box.size())
        for img_path in images_path:
            self.list_box.insert(END, img_path)
            img = Image.open(img_path)
            img = img.convert('RGB')
            images.append(img)
            self.show_img(img)
        self.images = images
        pass

    def save_images_to_pdf(self):
        out_imgs = []
        for img in self.images:
            out_imgs.append(self.img_in_page(img))
        img = out_imgs.pop(0)
        path = self.path + 'out.pdf'
        img.save(path, save_all=True, append_images=out_imgs)


    def cover_to_pdf(self, pdf_path, images):
        images = copy.deepcopy(images)
        resize_page_size = self.A4
        page_size = self.A4
        a4im = Image.new('RGB',
                         page_size,
                         (255, 255, 255))

        first_image = images.pop(0)
        img = Image.open(first_image)
        img = img.convert('RGB')
        image_resize = self.get_resize(resize_page_size[0], resize_page_size[1], img.size[0], img.size[1])
        img = img.resize(image_resize, Image.ANTIALIAS)
        image_site = int(page_size[0] - image_resize[0]) // 2, int(page_size[1] - image_resize[1]) // 2
        a4im.paste(img, image_site)
        other_images = []
        for i in images:
            a4im2 = Image.new('RGB',
                              page_size,
                              (255, 255, 255))
            img_2 = Image.open(i)
            img_2 = img_2.convert('RGB')
            # this_size = img_2.size
            this_size = img_2.size
            print(this_size)
            img_2 = img_2.rotate(90, expand=1)
            this_size = img_2.size
            print(this_size)
            image_resize = self.get_resize(resize_page_size[0], resize_page_size[1], img_2.size[0], img_2.size[1])
            img_2 = img_2.resize(image_resize, Image.ANTIALIAS)
            image_site = int(page_size[0] - image_resize[0]) // 2, int(page_size[1] - image_resize[1]) // 2
            # print("self.image_site", image_site)
            a4im2.paste(img_2, image_site)
            other_images.append(a4im2)
        # Image.Image保存为PDF格式文件
        a4im.save(pdf_path + 'out.pdf', save_all=True, append_images=other_images)
        a4im.save("d:/pdf/www.jpg")
        size = (800, 1200)
        box = (10, 10, 500, 500)
        aaa = a4im.resize(size, Image.ANTIALIAS)
        a4im.save("d:/pdf/www2.jpg", 'JPEG', quality=50)
        aaa.save("d:/pdf/www3.jpg", 'JPEG', quality=50)

    def img_to_pdf(self):
        output_file_name = 'out.pdf'
        output_file_name = self.path + output_file_name
        # save_file_name = 'ex.pdf'
        # doc = SimpleDocTemplate(save_file_name, pagesize=A1,
        #                     rightMargin=72, leftMargin=72,
        #                     topMargin=72, bottomMargin=18)
        imgDoc = canvas.Canvas(output_file_name)  # pagesize=letter
        imgDoc.setPageSize(A4)
        document_width, document_height = A4
        filenames = self.filenames
        for image in filenames:
            try:
                # image_file = PIL.Image.open(image)
                # image_file = rotate_img_to_proper(image_file)
                im = Image.Image
                image_file = Image.open(image)

                # image_file.show()

                image_width, image_height = image_file.size
                print('img size:', image_file.size)
                if not (image_width > 0 and image_height > 0):
                    raise Exception
                image_aspect = image_height / float(image_width)
                # Determins the demensions of the image in the overview
                print_width = document_width
                print_height = document_width * image_aspect
                img = ImageReader(image_file)
                print(type(img), img)
                imgDoc.drawImage(ImageReader(image_file), 0, 0, width=print_width,
                                 height=print_height)
                # imgDoc.drawImage(ImageReader(image_file), document_width - print_width,
                #                  document_height - print_height, width=print_width,
                #                  height=print_height, preserveAspectRatio=True)
                # inform the reportlab we want a new page
                imgDoc.showPage()
            except Exception as e:
                print('error:', e, image)
        imgDoc.save()
        print('Done')


def img_search2(mypath, filenames):
    for lists in os.listdir(mypath):
        path = os.path.join(mypath, lists)
        if os.path.isfile(path):
            expression = r'[\w]+\.(jpg|png|jpeg)$'
            if re.search(expression, path, re.IGNORECASE):
                filenames.append(path)
        elif os.path.isdir(path):
            img_search(path, filenames)


def img_search(mypath, filenames):
    '''读取指定文件夹下所有的JPEG图片，存入列表'''
    if mypath is None or len(mypath) == 0:
        raise ValueError('dirPath不能为空，该值为存放图片的具体路径文件夹！')
    if os.path.isfile(mypath):
        raise ValueError('dirPath不能为具体文件，该值为存放图片的具体路径文件夹！')
    for imageName in os.listdir(mypath):
        if imageName.endswith('.jpg') or imageName.endswith('.jpeg') or imageName.endswith('.png'):
            path = os.path.join(mypath, imageName)
            print(path)
            filenames.append(path)


def img_search1(mypath, filenames):
    for lists in os.listdir(mypath):
        path = os.path.join(mypath, lists)
        if os.path.isfile(path):
            a = path.split('.')
            if a[-1] in ['jpg', 'png', 'JPEG']:
                print(a[-1])
                filenames.append(path)
        elif os.path.isdir(path):
            img_search1(path, filenames)


def rotate_img_to_proper(image):
    try:
        # image = Image.open(filename)
        if hasattr(image, '_getexif'):  # only present in JPEGs
            for orientation in PIL.ExifTags.TAGS.keys():
                if PIL.ExifTags.TAGS[orientation] == 'Orientation':
                    break
            e = image._getexif()  # returns None if no EXIF data
            if e is not None:
                # log.info('EXIF data found: %r', e)
                exif = dict(e.items())
                orientation = exif[orientation]
                # print('found, ',orientation)

                if orientation == 3:
                    image = image.transpose(Image.ROTATE_180)
                elif orientation == 6:
                    image = image.transpose(Image.ROTATE_270)
                elif orientation == 8:
                    image = image.rotate(90, expand=True)
    except:
        pass
    return image


def win_main():
    ImageToPDF()
    mainloop()


if __name__ == '__main__':
    win_main()
    # main(src_folder='C:\\Users\\Administrator\\Desktop\\L1-U5-L3');
    # main(src_folder='D:\\pdf')
