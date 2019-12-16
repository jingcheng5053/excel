from tkinter import ttk
import xlrd
import xlrd.biffh
from tkinter import *
from tkinter.filedialog import askopenfilename  # 文件打开对话框
# from PIL import Image, ImageTk
from tkinter.messagebox import askyesno, showerror, showinfo
import os  # ,io
import xlsxwriter  # pip install xlsxwriter
from tkinter import messagebox as msgbox
from images.icon import Icon
import base64


class MyWin:
    def __init__(self, master):
        # 每一门满分分值--静态值
        self.Max_NAMES = ['语文', '数学', '英语', '物理', '政治', '历史', '地理', '生物', '化学']
        self.Max_TOTAL = [100, 100, 100, 100, 100, 100, 100, 100, 100]
        self.count_student_num = StringVar(value=65)  # '参评人数'
        self.this_show_lesson = ''  # 当前显示的课程名称
        self.add_one = {}
        for name in self.Max_NAMES:
            self.add_one[name] = 0

        self.master = master

        # <editor-fold desc="----------------------菜单模块-----------------------">
        menubar = Menu(self.master)
        # 添加菜单条
        self.master['menu'] = menubar
        # 创建file_menu菜单，它被放入menubar中
        file_menu = Menu(menubar, tearoff=0)
        # 使用add_cascade方法添加file_menu菜单
        menubar.add_cascade(label='文件', menu=file_menu)
        # 创建lang_menu菜单，它被放入menubar中
        lang_menu = Menu(menubar, tearoff=0)
        help_menu = Menu(menubar, tearoff=0)
        # 使用add_cascade方法添加lang_menu菜单
        # menubar.add_cascade(label='选择语言', menu=lang_menu)
        menubar.add_cascade(label='帮助', menu=help_menu)

        # 使用add_command方法为file_menu添加菜单项
        file_menu.add_command(label="读取", command=self.choose_xls,
                              image="", compound=LEFT)
        file_menu.add_command(label="保存", command=self.write_fun,
                              image="", compound=LEFT)
        help_menu.add_command(label="版本", command=self.show_info, compound=LEFT)
        # 使用add_command方法为file_menu添加分隔条
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.exit_root,
                              image="", compound=LEFT)
        # # 为file_menu创建子菜单
        # sub_menu = Menu(file_menu, tearoff=0)
        # # 使用add_cascade方法添加sub_menu子菜单
        # file_menu.add_cascade(label='选择性别', menu=sub_menu)
        # self.genderVar = IntVar()
        # # 使用循环为sub_menu子菜单添加菜单项
        # for i, im in enumerate(['男', '女', '保密']):
        #     # 使用add_radiobutton方法为sub_menu子菜单添加单选菜单项
        #     # 绑定同一个变量，说明它们是一组
        #     sub_menu.add_radiobutton(label=im, command=self.choose_gender,
        #                              variable=self.genderVar, value=i)
        # self.langVars = [StringVar(), StringVar(), StringVar(), StringVar()]
        # # 使用循环为lang_menu菜单添加菜单项
        # for i, im in enumerate(('Python', 'Kotlin', 'Swift', 'Java')):
        #     # 使用add_add_checkbutton方法为lang_menu菜单添加多选菜单项
        #     lang_menu.add_checkbutton(label=im, command=self.choose_lang,
        #                               onvalue=im, variable=self.langVars[i])
        # </editor-fold desc="----------------------菜单模块-----------------------">

        self.path = StringVar(value="先设置好参数在读取数据")
        self.total_name_var = StringVar(value='还未读取数据')
        student_num_frame = ttk.Frame(root)
        student_num_frame.pack(side=TOP, fill=X)
        student_num_label_frame = ttk.LabelFrame(student_num_frame, text='参评人数', labelanchor='w')
        student_num_label_frame.pack(side=LEFT)
        self.student_num_label = ttk.Label(student_num_label_frame, text="").pack(side=LEFT)
        test_cmd = student_num_label_frame.register(self.SacnexydwzVar)  # 注册testCMD绑定函数在Entry中执行
        self.student_num_spinbox = ttk.Spinbox(student_num_label_frame, textvariable=self.count_student_num, width=5,
                                               validate="key",
                                               validatecommand=(test_cmd, '%P', '%V', '%W'),
                                               font=('StSong', 12),
                                               from_=1, to=100).pack(side=LEFT, padx=10)
        add_label_frame = ttk.LabelFrame(student_num_frame, text='功能', labelanchor='w')
        add_label_frame.pack(side=LEFT, padx=10)
        self.add_class_btn = ttk.Button(add_label_frame, text="合并班级评分", command=self.add_class_fun)
        self.add_class_btn.pack(side=LEFT, padx=10)
        self.add_one_btn = ttk.Button(add_label_frame, text="教双班的加1分", command=self.add_one_fun)
        self.add_one_btn.pack(side=LEFT, padx=10)
        self.read_btn = ttk.Button(student_num_frame, text="读取", command=self.choose_xls)
        self.read_btn.pack(side=LEFT, padx=10)
        self.write_btn = ttk.Button(student_num_frame, text="保存", command=self.write_fun)
        self.write_btn.pack(side=LEFT, padx=10)

        read_frame = ttk.Frame(root)
        read_frame.pack(side=TOP)
        font1 = {'family': 'SimHei',
                 'weight': "BOLD",
                 # 'style':'italic',
                 'size': 12,
                 }
        self.label_loop = Label(read_frame, height=1, fg='blue',
                                textvariable=self.total_name_var, font=("黑体", 18, "bold"))  # , bitmap='warning'
        self.label_loop.pack(side=LEFT)

        self.tree_box = ttk.Treeview(root, height=16, show="headings")  # 表格show="headings",隐藏首列
        self.tree_box.pack(side=TOP)
        names = ("班级", '参评人数', '总评分', '总名次', '优秀人数', '优秀百分率', '优秀名次', '优秀评分',
                 '及格人数', '及格百分率', '及格名次', '及格评分', '总分', '平均分', '平均分名次', '成绩评分')

        names_show = ("班级", "参评人数", "总名次", "总评估值", "总分", "平均分", "名次", "评估值",
                      "优秀人数", "优秀率", "名次", "评估值", "及格人数", "及格率", "名次", "评估值")
        self.tree_box["columns"] = names  # 定义列
        for i in range(len(names)):
            self.tree_box.column(names[i], width=60, anchor='center')  # 表示列,不显示
            print(names_show[i])
            self.tree_box.heading(names[i], text=names_show[i])  # #设置显示的表头名

        self.root_lessons = ttk.LabelFrame(root, height=10)
        self.root_lessons.pack(side=TOP, fill=X)
        self.btn_lessons = {}
        name = "班主任"
        self.btn_lessons[name] = ttk.Button(self.root_lessons, text=name,
                                            command=lambda name=name: self.show_classes_total(name))
        self.btn_lessons[name].pack(side=LEFT)
        self.btn_lessons[name]['state'] = 'disabled'
        for name in self.Max_NAMES:
            self.btn_lessons[name] = ttk.Button(self.root_lessons, text=name,  # relief=GROOVE,
                                                command=lambda name=name: self.show_classes_total(name))
            self.btn_lessons[name].pack(side=LEFT)
            self.btn_lessons[name]['state'] = 'disabled'
        self.root_max_totals = ttk.Frame(root, height=10)
        self.root_max_totals.pack(side=TOP, fill=X)
        self.max_label = ttk.Label(self.root_max_totals, text='每一门满分:')
        self.max_label.pack(side=LEFT, padx=10)
        self.testCMD = self.root_max_totals.register(self.SacnexydwzVar)  # 注册testCMD绑定函数在Entry中执行
        self.max_totals_var = {}
        self.max_total_entrys = {}
        for name in self.Max_NAMES:
            self.max_totals_var[name] = StringVar(value='100')
            self.max_total_entrys[name] = ttk.Spinbox(self.root_max_totals,
                                                      textvariable=self.max_totals_var[name], width=5, validate="key",
                                                      validatecommand=(self.testCMD, '%P', '%V', '%W'),
                                                      font=('StSong', 12),
                                                      from_=10, to=150, increment=10).pack(side=LEFT, padx=14)

        self.root_end_totals = ttk.Frame(root, width=360, height=10)
        self.root_end_totals.pack(side=BOTTOM, fill=BOTH)
        self.label_end = Label(self.root_end_totals, textvariable=self.path, width=360, anchor='w', borderwidth=1,
                               relief='groove', fg='#555555', bg='#eeeeee')  # , bitmap='warning',bordercolor="#aaaaaa",
        self.label_end.pack(side=LEFT, fill=BOTH)

        # 当前读入字段,和最大分值
        self.data_names = ['学号', '语文', '数学', '英语', '物理', '政治', '历史', '地理', '生物', '化学']
        self.data_names_total = [100, 100, 100, 100, 100, 100, 100, 100, 100]

        self.lesson_count = 0
        self.data = []  # 读取数据
        self.data_classes = []  # 分班后数据
        self.data_classes_lesson = []  # 构建班级每一门评分数据库
        self.data_classes_class = []  # 构建班级评分数据库
        self.data_all_totals = {'学科': ['每班数据总排名,优秀,及格,平均']}  # 数据库
        self.year = 2019
        self.grade = 7
        self.max_class = 0  # 当前年级一共班级数
        pass

    def exit_root(self):
        self.master.destroy()
        os._exit(0)

    def show_info(self):
        showinfo('软件版本', '          ----------周口市第四初级中学评分软件 v1.1----------'
                         '\n' + '\n'
                                '功能:\n'
                                '    分别读取excel格式各个年级段成绩,并计算及格,优秀,总分等各项分值\n'
                                '计算公式:\n'
                                '    与教导处出具的计算公式完全一致。\n'
                                '导出结果:\n'
                                '    保存在读取文件的同目录下,以各个段的名称命名的excel文件中\n'
                                '\n\n'
                                '软件编写人:  李  军')

    def choose_gender(self):
        msgbox.showinfo(message=('选择的性别为: %s' % self.genderVar.get()))

    def choose_lang(self):
        rt_list = [e.get() for e in self.langVars]
        msgbox.showinfo(message=('选择的语言为: %s' % ','.join(rt_list)))

    # 注册testCMD绑定函数在Entry中执行
    def SacnexydwzVar(self, t, o, p, mini: int = 1, maxi: int = 150, max_len: int = 3):
        # print("this t:", t, o, "该组件的名字:", p, type(mini), maxi, max_len)

        if t == "":
            return True
        if not t.isdecimal() or int(t) < int(mini) or int(t) > int(maxi) or len(t) >= int(max_len):
            return False
        else:
            return True

    # 加一分
    def add_one_fun(self):
        lesson = self.this_show_lesson
        if lesson == '班主任':
            return
        data = self.data_all_totals[lesson]
        if self.add_one[lesson] > 0:
            showerror("重复加分", '当前课程%s已经加分,不需要重复加分' % lesson)
            print('当前课程已经加分,不需要重复加分', lesson)
            return
        for class_data in data:
            if len(class_data[0].split(',')) > 1:
                class_data[12] += class_data[1]
        self.add_one[lesson] = 1
        # 重新计算当前课程分数
        self.count_lesson_data(data)
        self.show_classes_total(lesson)
        pass

    # 保存结果到同目录下的文件
    def write_fun(self):
        path, file = os.path.split(self.path.get())
        if not path:
            showerror('保存错误', '没有读入文件,不能写入')
            print('没有读入文件,不能写入')
            return
        sheet_name = str(self.grade) + '年级评分'
        filename = path + '/' + sheet_name + '.xlsx'
        print(filename)
        workbook = xlsxwriter.Workbook(filename)

        worksheet = workbook.add_worksheet(sheet_name)
        # headings = ["班级", '参评人数', '总评分', '总名次', '优秀人数', '优秀百分率', '优秀名次', '优秀评分',
        #             '及格人数', '及格百分率', '及格名次', '及格评分', '总分', '平均分', '平均分名次', '成绩评分']
        headings = ["任课教师", "参评人数", "总名次", "总评估值", "总分", "平均分", "名次", "评估值",
                    "优秀人数", "优秀率", "名次", "评估值", "及格人数", "及格率", "名次", "评估值"]
        out_num = [0, 1, 3, 2, 12, 13, 14, 15, 4, 5, 6, 7, 8, 9, 10, 11]
        out_list = []
        begin_row = 1
        begin_col = 1
        bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        for key, d in self.data_all_totals.items():
            begin_row += 2
            worksheet.write(begin_row, begin_col - 1, key)
            worksheet.write_row(begin_row, begin_col, headings)
            worksheet.merge_range(begin_row + 1, begin_col - 1, begin_row + len(d), begin_col - 1, key, bold)
            for class_total in d:
                begin_row += 1
                out_list = []
                for i in out_num:
                    out_list.append(class_total[i])
                if key == "班主任":
                    out_list[5] = round(out_list[5] / (self.data_names.__len__() - 1), 2)
                    out_list[9] = round(out_list[9] / (self.data_names.__len__() - 1), 2)
                    out_list[13] = round(out_list[13] / (self.data_names.__len__() - 1), 2)
                if key != '班主任' and self.add_one[key] > 0 and len(out_list[0].split(',')) > 1:
                    out_list[5] = str(out_list[5] - 1) + "+1"
                    out_list[4] -= out_list[1]

                out_list[9] = str(out_list[9]) + "%"
                out_list[13] = str(out_list[13]) + "%"
                worksheet.write_row(begin_row, begin_col, out_list)
        try:
            workbook.close()
        except xlsxwriter.workbook.FileCreateError as e:
            showerror("保存失败", "文件:  '" + sheet_name + ".xlsx'" +
                      "是否已被其他程序打开,请先关闭在保存!")
            raise Exception('保存失败,弹出错误框')
        # workbook.close()
        showinfo("保存成功", "已将评分数据导出为文件:    " + sheet_name + '.xlsx')
        pass

    # 数据清零
    def all_zero(self):
        # 两个班以上的,平均分加1分
        for name in self.Max_NAMES:
            self.add_one[name] = 0

        for name, btn in self.btn_lessons.items():
            if name == '班主任' or name in self.data_names:
                btn['state'] = 'normal'
            else:
                btn['state'] = 'disabled'

    def btn_lesson_fun(self):
        # print(self.)
        pass

    def rank_get_list(self, ls: list, rank="max") -> list:
        """
        :param ls:需要排序的list
        :param rank: max 降序,min 升序
        :return:ls内数据的排序序号
        """
        if len(ls) == 0:
            raise Exception('list长度为0,不能排序')
        rank_bool = True if rank == 'max' else False
        ls_copy = []
        for i in range(len(ls)):
            ls_copy.append([i, ls[i], 0])
        ls_copy.sort(key=lambda asd: asd[1], reverse=rank_bool)  # 排序列表的第二个元素
        for i in range(ls_copy.__len__()):
            ls_copy[i][2] = i + 1
        # ls_copy.sort(key=take_first)
        ls_copy.sort(key=lambda asd: asd[0])  # 排序列表的第一个元素
        print(ls_copy)
        return [ls_copy[k][2] for k in range(len(ls_copy))]

    def delButton(self):
        x = self.tree_box.get_children()
        for item in x:
            self.tree_box.delete(item)

    # 显示班级评分
    def show_classes_total(self, show_name='班主任'):
        # names = tuple(self.data_names)
        if show_name in self.data_all_totals.keys():
            self.this_show_lesson = show_name
            s = str(self.grade) + '年级  ' + show_name + '  评分'
            self.total_name_var.set(s)
        else:
            print("当前课程没有成绩")
            return
        this_data = self.data_all_totals[show_name]
        self.delButton()
        names = ("班级", '参评人数', '总评分', '总名次', '优秀人数', '优秀百分率', '优秀名次', '优秀评分',
                 '及格人数', '及格百分率', '及格名次', '及格评分', '总分', '平均分', '平均分名次', '成绩评分')
        self.tree_box["columns"] = names  # 定义列
        for name in names:
            self.tree_box.column(name, width=60, anchor='center')  # 表示列,不显示
            self.tree_box.heading(name, text=name)  # #设置显示的表头名
        for i in range(len(this_data)):
            self.tree_box.insert('', i, values=(this_data[i]))
        bind = self.tree_box.bind("<Double-1>", self.onDBClick)
        # self.tree_box.unbind_all("<Double-1>")
        # self.tree_box.unbind("<Double-1>", bind)
        # self.tree_box.bind("<ButtonRelease-1>", self.on_click)

        # self.tree_box.insert("", 0, text="line1", values=("卡恩", "18", "180", "65"))  # #给第0行添加数据，索引值可重复

        pass

    # 合并班级,评分合并,并重新计算显示
    def add_class_fun(self):
        if len(self.tree_box.selection()) < 2:
            showerror("没用选中班级", '请选中两个以上的班级进行合并' + '\n' +
                      '提示:按"Ctrl"+鼠标左键进行多选')
            print('请选中两个以上的班级进行合并'
                  '提示:按"Ctrl"+鼠标左键进行多选')
            return
        classes = []
        self.tree_box.selection(items=1)
        print(self.tree_box.get_children())
        for item in list(self.tree_box.selection()):
            print(item, type(item))
            print(list(self.tree_box.get_children()).index(item))
            classes.append(list(self.tree_box.get_children()).index(item))
            # classes.append(self.tree_box.item(item, "values")[0])
        classes.reverse()
        data = self.data_all_totals[self.this_show_lesson]
        for n in range(classes.__len__() - 1):
            data[classes[n + 1]][0] += (',' + data[classes[n]][0])
            for i in [1, 4, 8, 12]:
                data[classes[n + 1]][i] += data[classes[n]][i]
            del data[classes[n]]
        print(self.tree_box.selection())
        print(classes)
        self.count_lesson_data(data)
        self.show_classes_total(self.this_show_lesson)

    def on_click(self, event):
        print(event)
        print(self.tree_box.selection())
        item = self.tree_box.selection()[0]
        print(self.tree_box.get_children()[0])
        # self.tree_box.selection_add(self.tree_box.get_children()[0])
        # self.tree_box.selection_toggle(self.tree_box.get_children()[1])
        # print(self.tree_box.)

    def onDBClick(self, event):
        print(self.tree_box.selection())
        item = self.tree_box.selection()[0]
        print("you clicked on ", self.tree_box.item(item, "values"))
        self.recount_lesson(self.tree_box.item(item, "values")[0])

    # 合并班级,评分合并,重新计算单科成绩,并刷新
    def recount_lesson(self, c: str):
        lesson = self.this_show_lesson
        print(lesson, c)
        data = self.data_all_totals[lesson]
        print(data)
        n = 100
        for i in range(len(data)):
            if c == data[i][0]:
                n = i
        for i in [1, 4, 8, 12]:
            data[n - 1][i] += data[n][i]
        print(n)
        data[n - 1][0] += (',' + c)
        del data[n]
        print(data)
        self.count_lesson_data(data)
        self.show_classes_total(lesson)
        pass

    # 计算成绩
    def count_every_class(self):
        # 构建学科每一门评分数据库
        self.data_classes_lesson = []
        for i in range(self.max_class):
            self.data_classes_lesson.append([])
            for j in range(len(self.data_names) - 1):
                max_total = self.data_names_total[j]
                self.data_classes_lesson[i].append(self.count_lesson_total(i, j, self.data_names_total[j]))
        print('班级每一门成绩评分', self.data_classes_lesson)
        # 构建班级评分数据库
        self.data_classes_class = []
        for i in range(self.max_class):
            self.data_classes_class.append(self.count_class_total(i))
        print('每班成绩评分', self.data_classes_class)
        # 构建评分,平均分,排名数据库,方便打印
        names = self.data_names[1:]
        self.data_all_totals = {}  # 初始化总数据排名,打印,显示数据库
        for j in range(len(names)):
            self.data_all_totals[names[j]] = []
            for i in range(self.max_class):
                self.data_all_totals[names[j]].append(self.count_lesson_total(i, j, self.data_names_total[j]))
        print(self.data_all_totals)
        print('_____以上是所有门分别入数据库______')
        for i, j in self.data_all_totals.items():
            self.data_all_totals[i] = self.put_lesson_to_data(j)
        print('总成绩表:', self.data_all_totals)
        self.data_classes_class = self.put_lesson_to_data(self.data_classes_class)
        self.data_all_totals['班主任'] = self.data_classes_class
        print('班主任:', self.data_classes_class)

    # 计算每一门详细评分,及排名
    def put_lesson_to_data(self, lesson_data: list):
        p = int(self.count_student_num.get())
        return_data = []
        for i in range(len(lesson_data)):
            a, b, x = lesson_data[i]
            if i < 9:
                ii = '0' + str(i + 1)
            else:
                ii = str(i + 1)
            return_data.append([ii, p, 14, 15, a, 0, 0, 0, b, 0, 0, 0, x, 0, 0, 0])
        print(return_data)
        return return_data

    # 直接计算所有门成绩
    def count_all_lesson_data(self):
        for i, d in self.data_all_totals.items():
            self.count_lesson_data(d)
            print(i, d)
        pass

    # 直接计算一门成绩
    def count_lesson_data(self, this_data: list):
        for d in this_data:
            for i in [4, 8, 12]:
                if i == 12:
                    d[i + 1] = round(d[i] / d[1], 2)
                else:
                    d[i + 1] = round(d[i] / d[1] * 100, 2)
        for i in [5, 9, 13]:
            this_data.sort(key=lambda asd: asd[i], reverse=True)
            n = 1
            this_data[0][i + 1] = n
            this_data[0][i + 2] = 60
            if i == 13:
                this_data[0][i + 2] = 80
            for j in range(1, len(this_data)):
                n += 1
                if this_data[j][i] == this_data[j - 1][i]:
                    this_data[j][i + 1] = this_data[j - 1][i + 1]
                else:
                    this_data[j][i + 1] = n
                if i == 13:
                    this_data[j][i + 2] = this_data[0][i + 2] - (this_data[j][i + 1] - 1) * 4
                else:
                    this_data[j][i + 2] = this_data[0][i + 2] - (this_data[j][i + 1] - 1) * 3
            for d in this_data:
                d[2] = d[7] + d[11] + d[15]
            this_data.sort(key=lambda asd: asd[2], reverse=True)
            n = 1
            this_data[0][2 + 1] = n
            for j in range(1, len(this_data)):
                n += 1
                if this_data[j][2] == this_data[j - 1][2]:
                    this_data[j][2 + 1] = this_data[j - 1][2 + 1]
                else:
                    this_data[j][2 + 1] = n
            this_data.sort(key=lambda asd: asd[0])
        pass

    # 计算每一科得分(及格,优秀,总分)
    def count_lesson_total(self, class_num, lesson_num, max_total=100):
        a = b = 0
        max_total = int(self.max_totals_var[self.data_names[lesson_num + 1]].get())
        for n in self.data_classes[class_num][lesson_num]:
            if n >= 0.60 * max_total:
                b += 1  # 及格人数
                if n >= 0.80 * max_total:
                    a += 1  # 优秀人数
        total = sum(self.data_classes[class_num][lesson_num])
        return [a, b, total]

    # 计算每一班得分
    def count_class_total(self, class_num):
        a = b = total = 0
        for lesson in self.data_classes_lesson[class_num]:
            a += lesson[0]
            b += lesson[1]
            total += lesson[2]
        return [a, b, total]

    # 数据分班
    def to_class(self):
        xue = str(int(self.data[0][0]))
        self.year = int(xue[:2])
        self.grade = int(xue[3:4])
        self.class_mate = int(xue[4:6])
        n = 0  # 查找最大班级数
        for xue in self.data[0]:
            xue = str(int(xue))
            if len(xue) > 2 and int(xue[:2]) == self.year:
                if n < int(xue[4:6]):
                    n = int(xue[4:6])
            else:
                pass
        print('当前最大班级数为:', n)
        self.max_class = n
        # 构建数据库
        self.data_classes = []
        for i in range(n):
            self.data_classes.append([])
            for j in range(len(self.data_names) - 1):
                self.data_classes[i].append([])
        # 成绩分别分入班级
        for i in range(len(self.data[0])):
            xue = str(int(self.data[0][i]))
            if len(xue) > 0 and int(xue[:2]) == self.year:
                n = int(xue[4:6]) - 1
                for j in range(len(self.data_names) - 1):
                    try:
                        self.data_classes[n][j].append(self.data[j + 1][i])
                    except IndexError as e:
                        print('错误IndexError:', e)
                        print('n,i,j:', n, i, j)
                        print('self.data_names:', self.data_names)
                        raise Exception('tingzhi')
            else:
                pass
        # 数据截取,保留每班60人
        for i in range(len(self.data_classes)):
            for j in range(len(self.data_classes[i])):
                self.data_classes[i][j] = self.data_classes[i][j][:int(self.count_student_num.get())]

        for d in self.data_classes:
            print(len(d), len(d[0]), len(d[1]), len(d[2]))
            print(d)

    # 读入数据
    def open(self):
        # xlrd.open_workbook(r'/root/excel/chat.xls')
        # xlrd.open_workbook(filename_path)
        try:
            workbook = xlrd.open_workbook(self.path.get())
        except xlrd.biffh.XLRDError as e:
            print("无法识别的excel文件")
            showerror('错误', "无法识别的excel文件!")
            print(e)
            return

        # workbook = xlrd.open_workbook(self.path.get())

        # 获取所有sheet名称
        sheet_names = workbook.sheet_names()
        # 根据sheet索引或者名称获取sheet内容
        sheet = workbook.sheet_by_index(0)
        # sheet = workbook.sheet_by_name('Sheet1')
        # sheet = workbook.sheet_by_name(sheet_names[0])
        # 获取整行和整列的值（数组）
        # rows = sheet.row_values(1)  # 获取第2行内容
        # cols = sheet.col_values(0)  # 获取第1列内容
        self.data = []
        self.data_names = []
        self.data_names_total = []
        if not '学号' in sheet.row_values(1):
            showerror('错误', '没有<学号>字段,无法判断学生具体班级信息!')
            raise Exception('error!!!!!!')
        for name in ['学号', '语文', '数学', '英语', '物理', '政治', '历史', '地理', '生物', '化学']:
            if name in sheet.row_values(1):
                # n = sheet.row_values(1).index(name)
                self.data.append(sheet.col_values(sheet.row_values(1).index(name))[2:])
                self.data_names.append(name)  # 添加当前读入字段名称
                # 添加当前学科最高分
                if name != '学号':
                    self.data_names_total.append(self.Max_TOTAL[self.Max_NAMES.index(name)])
            else:
                print(name, 'no in sheet.')
        print('读入的字段为:', self.data_names)
        print('读入的学科最大分值为:', self.data_names_total)
        print('读入的数据为:', self.data)
        print('------------读入数据完成--------open()---------')
        # sheet的名称，行数，列数
        print(sheet.name, sheet.nrows, sheet.ncols)
        # 获取(第几行, 第几列)单元格内容
        # t = sheet.cell_value(2, 4)
        # print('第sheet.cell_value(2, 4)个数据为:', t)
        # print(type(t))

    # 选择excel文件
    def choose_xls(self):
        path_ = askopenfilename()
        if path_ == '':
            print("没有选择文件,错误返回")
            return
        if not (path_.endswith('xlsx') or path_.endswith('xls')):
            showerror('错误', "打开的不是excel文件!")
            print('打开的不是excel文件!')
            return
        self.path.set(path_)
        self.open()
        self.to_class()  # 数据分班
        self.all_zero()  # 数据清零(双班加1分,按钮disabled)
        self.count_every_class()  # 计算每一门,每一班的评分
        self.count_all_lesson_data()
        self.count_lesson_data(self.data_classes_class)
        print('bzr:', self.data_classes_class)
        self.show_classes_total()


def close_window():
    ans = askyesno(title='Warning', message='是否关闭?')
    if ans:
        root.destroy()
        os._exit(0)
        # print("destroy之后")
    else:
        return


if __name__ == "__main__":
    root = Tk()
    root.title("周口市第四初级中学评分软件v1.1")
    root.geometry('960x550+100+100')
    with open('tmp.ico', 'wb') as tmp:
        tmp.write(base64.b64decode(Icon().img))
    root.iconbitmap('tmp.ico')
    os.remove('tmp.ico')
    # root.iconbitmap('images/config.ico')
    MyWin(root)
    root.protocol('WM_DELETE_WINDOW', close_window)  # 截获win窗口关闭事件，并处理
    root.mainloop()
    # root.main
