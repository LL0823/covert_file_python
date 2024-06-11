#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author: liu
# time: 2024/5/25 16:53
import os
import sys
import tkinter as tk
from tkinter import filedialog as fd
from xmind_to_xlsx import run


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class XmindToExcelTkinter(tk.Tk):  # 继承TK类，去实现界面效果，这样稍微面向对象一些
    def __init__(self):
        # 界面的初始化
        super().__init__()
        self.title('xmind转excel文件')
        self.iconbitmap(get_resource_path('res/yaya.ico'))
        self.geometry('500x300')
        # text控件
        self.console_text = None
        # 可选框checkbutton的选择值
        self.check_button_var = tk.IntVar()
        # 页面元素布局
        self.init_ui()

    def xmind_to_excel(self, xmind_file_path):
        """
        给button按钮使用的方法，转换文件
        :param xmind_file_path:
        :return:
        """
        self.output_to_console('开始将xmind转换成excel了\n')
        self.output_to_console(f'文件路径为：{xmind_file_path}\n')
        # 判断选择的文件是否为xmind格式，如果不是直接提示。
        if xmind_file_path.endswith(".xmind") is False:
            self.output_to_console('选择的文件不是xmind格式，请重新选择\n')
            self.output_to_console('======================================\n')
            return
        excel_path = run(xmind_file_path, self.check_button_var.get())
        self.output_to_console(f'生成的excel文件路径为：{excel_path}\n')
        self.output_to_console('======================================\n')

    def output_to_console(self, new_text):
        """
        在Text控件中打印
        :param new_text: 需要打印的文案
        :return:
        """
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, new_text)
        self.console_text.see(tk.END)
        self.console_text.config(state=tk.DISABLED)

    def init_ui(self):
        """
        初始化页面元素
        :return:
        """
        xmind_file_path = tk.StringVar()
        xmind_file_path.set('')
        f1 = tk.Frame(self)
        f1.pack(padx=20, pady=10)
        tk.Button(f1, text='选择文件', command=lambda: xmind_file_path.set(fd.askopenfilename(title='选择文件'))).pack(
            side=tk.LEFT,
            padx=10)
        tk.Button(f1, text='转换成excel', command=lambda: self.xmind_to_excel(xmind_file_path.get())).pack(side=tk.LEFT,
                                                                                                           padx=10)
        # self.check_button_var = tk.IntVar()
        # 增加一个可选框，选择是否要对excel进行合并单元格
        tk.Checkbutton(self, text='勾选后合并前两列单元格', variable=self.check_button_var).pack(anchor='w', padx=5,
                                                                                                 pady=5)
        label = tk.Label(self, textvariable=xmind_file_path, bg='#ffffff')
        label.pack(expand=True, fill='x', padx=5, pady=5)
        # 加一个不让编辑的文本框，显示转换进度及转换后的文件路径
        self.console_text = tk.Text(self, fg="green", bg="black", state=tk.DISABLED)  # 不让编辑
        self.console_text.pack(padx=5, pady=5)


if __name__ == '__main__':
    app = XmindToExcelTkinter()
    app.mainloop()
