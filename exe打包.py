import os
import shutil
import subprocess
import threading
import logging
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Entry, Checkbutton, IntVar, Text
import sys

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def center_window(window):
    """将窗口居中显示"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')
    # 设置窗口置顶
    window.attributes('-topmost', True)
    # 延迟500毫秒取消窗口置顶
    window.after(500, lambda: window.attributes('-topmost', False))
    
    # 设置窗口图标
    root.iconbitmap('icon.ico')


def on_map(event):
    """窗口显示后调用此函数进行居中"""
    root.after(10, lambda: center_window(root))


def select_main_script():
    filepath = filedialog.askopenfilename(filetypes=[("Python files", "*.py")])
    if filepath:
        main_script.set(filepath)
        update_status_label("主脚本已选择，请继续选择其他文件或开始打包。")
        logging.info(f"用户选择了主脚本: {filepath}")


def select_data_folder():
    folderpath = filedialog.askdirectory()
    if folderpath:
        data_folder.set(folderpath)
        update_status_label("数据文件夹已选择，请继续选择其他文件或开始打包。")
        logging.info(f"用户选择了数据文件夹: {folderpath}")


def select_icon_file():
    filepath = filedialog.askopenfilename(filetypes=[("Icon files", "*.ico")])
    if filepath:
        icon_path.set(filepath)
        update_status_label("图标文件已选择，请继续选择其他文件或开始打包。")
        logging.info(f"用户选择了图标文件: {filepath}")


def open_folder(path):
    """打开文件夹，支持不同操作系统"""
    if sys.platform.startswith('win'):
        os.startfile(path)
    elif sys.platform.startswith('darwin'):
        subprocess.run(['open', path])
    elif sys.platform.startswith('linux'):
        subprocess.run(['xdg-open', path])
    logging.info(f"尝试打开文件夹: {path}")


def start_packaging_in_thread():
    threading.Thread(target=pack_application, daemon=True).start()


def generate_spec_file(main_script_path, datas, add_libs, output_dir, icon_path=None, hide_console=False):
    """生成.spec文件"""
    spec_filename = os.path.join(output_dir, os.path.splitext(os.path.basename(main_script_path))[0] + '.spec')
    logging.info(f"开始生成 .spec 文件，目标路径: {spec_filename}")

    main_script_utf8 = repr(main_script_path)[1:-1]
    datas_utf8 = [(repr(os.path.normpath(src).replace('\\', '/'))[1:-1], dest.replace('\\', '/')) for src, dest in datas]

    hidden_imports_utf8 = [repr(lib.split("==")[0].split(">=")[0].split("<=")[0].strip())[1:-1] for lib in add_libs]

    icon_str = f'icon="{icon_path}", ' if icon_path else ''
    console_str = 'console=False' if hide_console else 'console=True'

    try:
        with open(spec_filename, 'w', encoding='utf-8') as f:
            f.write(f"""
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['{main_script_utf8}'],
    pathex=[],
    binaries=[],
    datas={datas_utf8},
    hiddenimports={hidden_imports_utf8},
    hookspath=[],
    runtime_hooks=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=os.path.splitext(os.path.basename('{main_script_utf8}'))[0],
    debug=False,
    strip=False,
    upx=True,
    {console_str},
    {icon_str}
    onefile=True
)
""")
        logging.info(f".spec 文件生成成功: {spec_filename}")
    except FileNotFoundError:
        logging.error(f"生成.spec文件时，找不到目标路径: {spec_filename}")
        messagebox.showerror("错误", f"生成.spec文件时，找不到目标路径: {spec_filename}")
        return None
    except PermissionError:
        logging.error(f"生成.spec文件时，没有写入权限: {spec_filename}")
        messagebox.showerror("错误", f"生成.spec文件时，没有写入权限: {spec_filename}")
        return None
    except Exception as e:
        logging.error(f"生成.spec文件时出现未知错误: {e}，目标路径: {spec_filename}")
        messagebox.showerror("错误", f"生成.spec文件时出现未知错误: {e}")
        return None

    return spec_filename


def get_datas(main_script_dir, main_script_name, data_folder_path):
    """获取要打包的数据文件列表"""
    datas = []
    logging.info(f"开始获取主脚本所在目录（{main_script_dir}）的文件列表")
    # 添加主脚本所在目录的所有文件（不包括子文件夹），同时剔除主脚本本身
    try:
        for item in os.listdir(main_script_dir):
            abs_path = os.path.join(main_script_dir, item)
            if os.path.isfile(abs_path) and item != main_script_name:  # 剔除主脚本本身
                # 将文件直接放入根目录
                datas.append((abs_path, '.'))
        logging.info(f"成功获取主脚本所在目录（{main_script_dir}）的文件列表")
    except FileNotFoundError:
        logging.error(f"主脚本所在目录不存在: {main_script_dir}")
        messagebox.showerror("错误", f"主脚本所在目录不存在: {main_script_dir}")
        return []
    except PermissionError:
        logging.error(f"没有权限访问主脚本所在目录: {main_script_dir}")
        messagebox.showerror("错误", f"没有权限访问主脚本所在目录: {main_script_dir}")
        return []
    except Exception as e:
        logging.error(f"获取主脚本所在目录文件时出现未知错误: {e}，目录路径: {main_script_dir}")
        messagebox.showerror("错误", f"获取主脚本所在目录文件时出现未知错误: {e}")
        return []

    # 添加数据文件夹内容到 datas（包括子目录）
    if data_folder_path:
        base_folder = data_folder_path
        # 动态读取数据文件夹的名称
        target_folder_name = os.path.basename(base_folder)
        logging.info(f"开始获取数据文件夹（{base_folder}）的内容")
        try:
            for root, _, files in os.walk(base_folder):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(root, base_folder)
                    # 目标路径为数据文件夹名称加上相对路径
                    target_path = os.path.join(target_folder_name, rel_path) if rel_path != '.' else target_folder_name
                    datas.append((abs_path, target_path))
            logging.info(f"成功获取数据文件夹（{base_folder}）的内容")
        except FileNotFoundError:
            logging.error(f"数据文件夹不存在: {base_folder}")
            messagebox.showerror("错误", f"数据文件夹不存在: {base_folder}")
        except PermissionError:
            logging.error(f"没有权限访问数据文件夹: {base_folder}")
            messagebox.showerror("错误", f"没有权限访问数据文件夹: {base_folder}")
        except Exception as e:
            logging.error(f"获取数据文件夹内容时出现未知错误: {e}，文件夹路径: {base_folder}")
            messagebox.showerror("错误", f"获取数据文件夹内容时出现未知错误: {e}")

    return datas


def get_add_libs(main_script_dir):
    """获取要添加的库列表"""
    add_libs = []
    # 如果存在 add_libs.txt 文件，则读取其中的库列表
    add_lib_path = os.path.normpath(os.path.join(main_script_dir, 'add_libs.txt'))
    if os.path.exists(add_lib_path):
        logging.info(f"开始读取 add_libs.txt 文件: {add_lib_path}")
        try:
            with open(add_lib_path, encoding='utf-8') as f:
                add_libs = [line.strip() for line in f.readlines()]
            logging.info(f"成功读取 add_libs.txt 文件: {add_lib_path}")
        except FileNotFoundError:
            logging.error(f"找不到 add_libs.txt 文件: {add_lib_path}")
            messagebox.showerror("错误", f"找不到 add_libs.txt 文件: {add_lib_path}")
        except PermissionError:
            logging.error(f"没有权限读取 add_libs.txt 文件: {add_lib_path}")
            messagebox.showerror("错误", f"没有权限读取 add_libs.txt 文件: {add_lib_path}")
        except Exception as e:
            logging.error(f"读取 add_libs.txt 文件时出现未知错误: {e}，文件路径: {add_lib_path}")
            messagebox.showerror("错误", f"读取 add_libs.txt 文件时出现未知错误: {e}")
    return add_libs


def pack_application():
    status_label.config(state='normal')
    status_label.delete('1.0', 'end')
    insert_status("正在选择主脚本...")
    main_script_path = main_script.get()
    if not main_script_path:
        messagebox.showerror("错误", "请选择主脚本")
        return

    # 获取主脚本所在的文件夹和主脚本的名字
    main_script_dir = os.path.dirname(main_script_path)
    main_script_name = os.path.basename(main_script_path)

    insert_status("选择输出文件夹...")
    output_dir = filedialog.askdirectory(title="选择输出文件夹")
    if not output_dir:
        messagebox.showwarning("警告", "未选择输出文件夹，操作已取消。")
        return
    logging.info(f"用户选择的输出文件夹: {output_dir}")

    datas = get_datas(main_script_dir, main_script_name, data_folder.get())
    add_libs = get_add_libs(main_script_dir)

    icon_path_value = icon_path.get()
    hide_console_value = hide_console.get()

    build_dir = os.path.join(main_script_dir, 'build_temp')
    logging.info(f"开始创建目录: {output_dir} 和 {build_dir}")
    try:
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(build_dir, exist_ok=True)
        logging.info(f"成功创建目录: {output_dir} 和 {build_dir}")
    except FileNotFoundError:
        logging.error(f"创建目录时，找不到父目录: {output_dir} 或 {build_dir}")
        messagebox.showerror("错误", f"创建目录时，找不到父目录: {output_dir} 或 {build_dir}")
        return
    except PermissionError:
        logging.error(f"没有权限创建目录: {output_dir} 或 {build_dir}")
        messagebox.showerror("错误", f"没有权限创建目录: {output_dir} 或 {build_dir}")
        return
    except Exception as e:
        logging.error(f"创建目录时出现未知错误: {e}，目录路径: {output_dir} 和 {build_dir}")
        messagebox.showerror("错误", f"创建目录时出现未知错误: {e}")
        return

    insert_status("生成.spec文件...")
    spec_filename = generate_spec_file(main_script_path, datas, add_libs, output_dir, icon_path_value,
                                       hide_console_value)
    if spec_filename is None:
        return

    insert_status("执行PyInstaller命令，过程比较长，请耐心等待...")
    logging.info(f"开始执行 PyInstaller 命令，使用 .spec 文件: {spec_filename}")

    cmd = [
        'pyinstaller',
        '--distpath', output_dir,
        '--workpath', build_dir,
        '--noconfirm',
        '--clean',
        spec_filename
    ]

    startupinfo = subprocess.STARTUPINFO()
    if sys.platform.startswith('win'):
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE

    try:
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                   startupinfo=startupinfo)
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                insert_status(output.strip())
        return_code = process.poll()
        if return_code != 0:
            raise subprocess.CalledProcessError(return_code, cmd)
        logging.info(f"PyInstaller 命令执行成功，使用 .spec 文件: {spec_filename}")
    except FileNotFoundError:
        logging.error("找不到 pyinstaller 命令，请确保已经安装。")
        messagebox.showerror("错误", "找不到 pyinstaller 命令，请确保已经安装。")
        return
    except subprocess.CalledProcessError as e:
        logging.error(f"PyInstaller 执行失败: {e}，使用 .spec 文件: {spec_filename}")
        messagebox.showerror("错误", f"PyInstaller 执行失败: {e}")
        return
    except Exception as e:
        logging.error(f"执行过程中出现未知错误: {e}，使用 .spec 文件: {spec_filename}")
        messagebox.showerror("错误", f"执行过程中出现未知错误: {e}")
        return

    # 删除 __pycache__ 目录
    pycache_dir = os.path.join(main_script_dir, '__pycache__')
    if os.path.exists(pycache_dir):
        logging.info(f"开始删除 __pycache__ 目录: {pycache_dir}")
        try:
            shutil.rmtree(pycache_dir, ignore_errors=True)
            logging.info(f"成功删除 __pycache__ 目录: {pycache_dir}")
        except PermissionError:
            logging.error(f"没有权限删除 __pycache__ 目录: {pycache_dir}")
            messagebox.showerror("错误", f"没有权限删除 __pycache__ 目录: {pycache_dir}")
        except Exception as e:
            logging.error(f"删除 __pycache__ 目录时出现未知错误: {e}，目录路径: {pycache_dir}")
            messagebox.showerror("错误", f"删除 __pycache__ 目录时出现未知错误: {e}")

    logging.info(f"开始删除 build 目录: {build_dir}")
    try:
        shutil.rmtree(build_dir, ignore_errors=True)
        logging.info(f"成功删除 build 目录: {build_dir}")
    except PermissionError:
        logging.error(f"没有权限删除 build 目录: {build_dir}")
        messagebox.showerror("错误", f"没有权限删除 build 目录: {build_dir}")
    except Exception as e:
        logging.error(f"删除 build 目录时出现未知错误: {e}，目录路径: {build_dir}")
        messagebox.showerror("错误", f"删除 build 目录时出现未知错误: {e}")

    logging.info(f"开始删除 .spec 文件: {spec_filename}")
    try:
        os.remove(spec_filename)
        logging.info(f"成功删除 .spec 文件: {spec_filename}")
    except FileNotFoundError:
        logging.error(f"找不到 .spec 文件: {spec_filename}")
    except PermissionError:
        logging.error(f"没有权限删除 .spec 文件: {spec_filename}")
        messagebox.showerror("错误", f"没有权限删除 .spec 文件: {spec_filename}")
    except Exception as e:
        logging.error(f"删除 .spec 文件时出现未知错误: {e}，文件路径: {spec_filename}")
        messagebox.showerror("错误", f"删除 .spec 文件时出现未知错误: {e}")

    open_folder(output_dir)
    insert_status("打包完成，等待新任务...")
    status_label.config(state='disabled')


def insert_status(message):
    status_label.insert('end', message + "\n")
    status_label.yview_moveto(1)  # 自动滚动到最新消息
    logging.info(message)


def update_status_label(text):
    """更新状态标签的内容"""
    status_label.config(state='normal')
    status_label.delete('1.0', 'end')
    status_label.insert('end', text + "\n")
    status_label.yview_moveto(1)  # 自动滚动到最新消息
    status_label.config(state='disabled')


def create_gui():
    global root, main_script, data_folder, icon_path, hide_console, status_label
    root = Tk()
    root.title("Python 应用程序打包工具")

    # 调整窗口大小
    root.geometry('555x350')

    main_script = StringVar()
    data_folder = StringVar()
    icon_path = StringVar()
    hide_console = IntVar()

    initial_instructions = (
        '\u3000\u30001.打包主程序所在目录的全部文件，如果主程序所在目录存在add_libs.txt文件还会自动打包其指定的库，'
        '文件内容一行一个库，如“PyQt5==5.15.9”。'
        '\n\u3000\u30002.添加数据文件夹后需注意调用方法，如：if getattr(sys, "frozen", False):解包后主目录base_path=sys._MEIPASS，'
        'else:开发环境主目录base_path=os.path.dirname(os.path.abspath(__file__))，切换目录os.chdir(base_path)，'
        '生成全路径文件名os.path.join(base_path,"数据文件夹名","文件名")。'
        '\n\u3000\u30003.打包时务必确保“upx.exe”文件与当前脚本在同一目录下！打包完成的可执行文件使用upx压缩，文件相对较小。'
    )

    Label(root, text="脚本(必选):").grid(row=0, column=0, padx=5, pady=(20, 5), sticky='e')
    Entry(root, textvariable=main_script).grid(row=0, column=1, padx=5, pady=(20, 5), sticky='ew')
    Button(root, text="浏览", command=select_main_script, bg="#00da6a", fg="white").grid(row=0, column=2, padx=(5, 25), pady=(20, 5), sticky='ew')
    Label(root, text="图标(可选):").grid(row=1, column=0, padx=5, pady=5, sticky='e')
    Entry(root, textvariable=icon_path).grid(row=1, column=1, padx=5, pady=5, sticky='ew')
    Button(root, text="浏览", command=select_icon_file, bg="#2a5a63", fg="white").grid(row=1, column=2, padx=(5, 25), pady=5, sticky='ew')

    Label(root, text="数据(可选):").grid(row=2, column=0, padx=5, pady=(5, 15), sticky='e')
    Entry(root, textvariable=data_folder).grid(row=2, column=1, padx=5, pady=(5, 15), sticky='ew')
    Button(root, text="浏览", command=select_data_folder, bg="#005bbb", fg="white").grid(row=2, column=2, padx=(5, 25), pady=(5, 15), sticky='ew')

    Checkbutton(root, text="隐藏控制台", variable=hide_console).grid(row=3, column=0, padx=(25, 5), pady=(5, 15), sticky='w')
    Button(root, text="开始打包", command=start_packaging_in_thread, bg="#ff0048", fg="white").grid(row=3, column=1, columnspan=1, padx=5,
                                                                          pady=(5, 15), sticky='ew')

    # 创建一个Text组件来显示状态信息，并允许文本自动换行
    status_label = Text(root, wrap='word', state='disabled', height=9, bg='white', fg='blue', relief='sunken', bd=1)
    status_label.grid(row=4, column=0, columnspan=3, padx=(25, 25), pady=(5, 15), sticky='nsew')

    # 设置初始状态说明
    update_status_label(initial_instructions)

    root.grid_columnconfigure(1, weight=1)
    root.bind("<Map>", on_map)

    root.mainloop()


if __name__ == "__main__":
    create_gui()