import tkinter as tk
import os
import subprocess

# 存放程序文件路径的字典
programs = {
    'test整理数据20230220-3.py': '数据整理',
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

# 创建一个菜单栏
menubar = tk.Menu(root)

# 创建5个下拉菜单，添加到菜单栏
plan_menu = tk.Menu(menubar, tearoff=0)
keyword_menu = tk.Menu(menubar, tearoff=0)
data_menu = tk.Menu(menubar, tearoff=0)
finance_menu = tk.Menu(menubar, tearoff=0)
inventory_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label='计划和广告调整', menu=plan_menu)
menubar.add_cascade(label='搜索词库', menu=keyword_menu)
menubar.add_cascade(label='基础数据维护', menu=data_menu)
menubar.add_cascade(label='财务', menu=finance_menu)
menubar.add_cascade(label='库存管理', menu=inventory_menu)

# 遍历程序字典，为每个程序创建一个菜单项
for file_name, description in programs.items():
    # 创建一个回调函数
    def run_program(file_name=file_name):
        # 打开IDLE Shell运行程序
        cmd = f'pythonw.exe D:/运营/0-Python程序{file_name}'
        subprocess.Popen(cmd)

    # 将菜单项添加到对应的下拉菜单中
    if file_name in ('test整理数据20230220-3.py', 'move_data_to_historical.py'):
        plan_menu.add_command(label=description, command=run_program)
    elif file_name == 'check_raw_data.py':
        data_menu.add_command(label=description, command=run_program)
    elif file_name in ('convert_files.py', 'analyze_files.py', 'clean_files.py'):
        keyword_menu.add_command(label=description, command=run_program)
    elif file_name == 'backup_files.py':
        inventory_menu.add_command(label=description, command=run_program)

# 设置菜单栏
root.config(menu=menubar)

# 进入主循环
root.mainloop()
