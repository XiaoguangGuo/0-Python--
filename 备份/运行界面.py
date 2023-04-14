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

# 遍历程序字典，为每个程序创建一个按钮
for file_name, description in programs.items():
    # 创建一个Frame作为按钮的容器
    frame = tk.Frame(root)
    frame.pack(side=tk.TOP, pady=10)

    # 创建一个Label显示程序的描述
    label = tk.Label(frame, text=description, font=('Helvetica', 14))
    label.pack(side=tk.LEFT, padx=10)

    # 创建一个Button运行程序
    button = tk.Button(frame, text='运行', bg='light blue')
    button.pack(side=tk.LEFT, padx=10)

    # 定义按钮的回调函数
    def run_program(event, file_name=file_name, button=button):
        # 将按钮置为灰色
        button.configure(state=tk.DISABLED)

        # 打开IDLE Shell运行程序
        cmd = f'pythonw.exe D:/运营/{file_name}'
        subprocess.Popen(cmd)

    # 将回调函数绑定到按钮上
    button.bind('<Button-1>', run_program)

# 进入主循环
root.mainloop()
