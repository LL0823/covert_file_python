# covert_file_python

# 带GUI的文件转换练习项目
简介：练习python的gui库实现的带界面的文件转换工具，为了把xmind的用例转换成符合格式的excel用例。

###### 环境说明：
Python: 3.11.8、
PyInstaller: 6.7.0

## 1、convert_xmind_to_xlsx_gui
用tkinter界面实现的xmind转换成xlsx文件。

界面效果：
![img.png](img.png)

使用pyinstaller打包，打包遇到的巨多坑都在
https://blog.csdn.net/u010964317/article/details/139510480?spm=1001.2014.3001.5502
https://blog.csdn.net/u010964317/article/details/139524346?spm=1001.2014.3001.5502

文件注释：

├── convert_xmind_to_xlsx_gui.py   //主界面

├── convert_xmind_to_xlsx_gui.spec //pyinstaller打包配置，包括打包多个文件及资源打包

├── xmind_to_xlsx.py               //文件转换的逻辑

├── dist                           //打出的exe包存放位置

## 2、convert_xmind_to_xlsx_gui

