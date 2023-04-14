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
menu_plan_ad = tk.Menu(menubar, tearoff=0)
for file_name, description in programs.items():
    # 创建一个Button运行程序
    def run_program(event, file_name=file_name, button=button):
        # 将按钮置为灰色
        button.configure(state=tk.DISABLED)

        # 打开IDLE Shell运行程序
        cmd = f'pythonw.exe D:/运营/{file_name}'
        subprocess.Popen(cmd)

    # 将回调函数绑定到按钮上
    button = tk.Button(root, text='运行', bg='light blue', command=run_program)
    button.pack(side=tk.TOP, pady=10)

    # 创建一个Label显示程序的描述
    label = tk.Label(root, text=description, font=('Helvetica', 14))
    label.pack(side=tk.TOP, padx=10)

    # 添加按钮到计划和广告调整菜单
    menu_plan_ad.add_command(label=description, command=run_program)

# 添加计划和广告调整菜单到菜单栏
menubar.add_cascade(label='计划和广告调整', menu=menu_plan_ad)

# 创建搜索词库菜单
menu_search = tk.Menu(menubar, tearoff=0)
for i in range(3):
    program_name = f"program_{i+1}.py"
    program_path = os.path.join('D:', '0-Python程序', program_name)
    def run_program(event, program_path=program_path, button=button):
        # 将按钮置为灰色
        button.configure(state=tk.DISABLED)

        # 打开IDLE Shell运行程序
        cmd = f'pythonw.exe {program_path}'
        subprocess.Popen(cmd)

    # 将回调函数绑定到按钮上
    button = tk.Button(root, text='运行', bg='light blue', command=run_program)
    button.pack(side=tk.TOP, pady=10)

    # 创建一个Label显示程序的描述
    label = tk.Label(root, text=f"搜索词库程序 {i+1}", font=('Helvetica', 14))
    label.pack(side=tk.TOP, padx=10)

    # 添加按钮到搜索词库菜单
    menu_search.add_command(label=f"搜索词库程序 {i+1}", command=run_program)

# 添加搜索词库菜单到菜单栏
menubar.add_cascade(label='搜索词库', menu=menu_search)

# 创建基础数据维护菜单
menu_data_maintain = tk.Menu(menubar, tearoff=
