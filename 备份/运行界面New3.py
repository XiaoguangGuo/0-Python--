import tkinter as tk
import os
import subprocess

# 存放程序文件路径的字典
programs = {
    'check_data_source.py': '检查数据源',
    'move_data_to_historical.py': '移动数据到历史文件夹',
    'check_raw_data.py': '检查原始数据',
    'convert_files.py': '转换文件格式',
    'analyze_files.py': '分析文件',
    'clean_files.py': '清理文件',
    'backup_files.py': '备份文件'
}

# 创建主窗口
root = tk.Tk()
root.title('运营程序')

# 创建菜单栏
menubar = tk.Menu(root)

# 创建计划和广告调整菜单
menu_adjustment = tk.Menu(menubar, tearoff=0)
for file_name, description in programs.items():
    menu_adjustment.add_command(label=description, command=lambda file_name=file_name: subprocess.Popen(f'pythonw.exe D:/运营/{file_name}'))
menubar.add_cascade(label='计划和广告调整', menu=menu_adjustment)

# 创建搜索词库菜单
menu_keyword = tk.Menu(menubar, tearoff=0)
for i in range(3):
    file_path = f'D:/0-Python程序/keyword_{i}.py'
    if os.path.isfile(file_path):
        menu_keyword.add_command(label=f'搜索词库{i}', command=lambda file_path=file_path: subprocess.Popen(f'pythonw.exe {file_path}'))
    else:
        menu_keyword.add_command(label=f'搜索词库{i}', state=tk.DISABLED)
menubar.add_cascade(label='搜索词库', menu=menu_keyword)

# 创建基础数据维护菜单
menu_data_maintain = tk.Menu(menubar, tearoff=0)
for i in range(3):
    file_path = f'D:/0-Python程序/data_maintain_{i}.py'
    if os.path.isfile(file_path):
        menu_data_maintain.add_command(label=f'基础数据维护{i}', command=lambda file_path=file_path: subprocess.Popen(f'pythonw.exe {file_path}'))
    else:
        menu_data_maintain.add_command(label=f'基础数据维护{i}', state=tk.DISABLED)
menubar.add_cascade(label='基础数据维护', menu=menu_data_maintain)

# 创建财务菜单
menu_finance = tk.Menu(menubar, tearoff=0)
for i in range(3):
    file_path = f'D:/0-Python程序/finance_{i}.py'
    if os.path.isfile(file_path):
        menu_finance.add_command(label=f'财务{i}', command=lambda file_path=file_path: subprocess.Popen(f'pythonw.exe {file_path}'))
    else:
        menu_finance.add_command(label=f'财务{i}', state=tk.DISABLED)
menubar.add_cascade(label='财务', menu=menu_finance)


