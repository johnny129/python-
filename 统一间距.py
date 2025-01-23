import tkinter as tk
from tkinter.filedialog import askopenfilename, askdirectory
import zipfile, os, re, shutil, time

# 选择文件
def selectFile():
    global filepath, targetpath
    filepath = askopenfilename()
    if filepath:
        filePath.set(filepath)
        targetpath = os.path.dirname(filepath) + "/_target_"
        dirPath.set("")

# 选择目录
def selectDirFile():
    global dirpath, targetpath
    dirpath = askdirectory()
    if dirpath:
        dirPath.set(dirpath)
        targetpath = dirpath + "/_target_"
        filePath.set("")

# 保存文件
def fileSave():
    if dirpath or filepath:
        logs.delete(0.0, tk.END)
        root.update()
        replacePPTX(filepath or dirpath)
    else:
        writeLog("文件或目录至少选择一个！")

# 写日志
def writeLog(msg):
    current_time = time.strftime('【%H:%M:%S】', time.localtime(time.time()))
    logmsg_in = str(current_time) + str(msg) + "\n"
    logs.insert(tk.END, logmsg_in)
    logs.yview_moveto(1)
    root.update()

# 获取文件列表
def getFiles(directory):
    if os.path.isdir(directory):
        files = [os.path.join(directory, f) for f in os.listdir(directory)]
    else:
        files = [directory]
    return files

# 压缩文件
def zip_file(fn):
    global targetpath
    if not os.path.exists(targetpath):
        os.mkdir(targetpath)
    with zipfile.ZipFile(os.path.join(targetpath, fn), 'w', zipfile.ZIP_DEFLATED) as z:
        for dirpath, _, filenames in os.walk("temp"):
            fpath = dirpath.replace("temp", '')
            for filename in filenames:
                z.write(os.path.join(dirpath, filename), os.path.join(fpath, filename))
    shutil.rmtree("temp")
    writeLog(f"{fn} 替换成功")

# 处理文件
def modefile(files):
    pptx_files = [f for f in files if f.lower().endswith(".pptx")]
    for fs in pptx_files:
        writeLog("开始处理文件：" + fs)
        z = zipfile.ZipFile(fs, "r")
        z.extractall("temp")
        # 处理幻灯片和母版
        for xml_dir in ["temp/ppt/slides", "temp/ppt/slideMasters"]:
            for xml in getFiles(xml_dir):
                if os.path.isfile(xml):
                    with open(xml, "r", encoding="utf-8") as f:
                        file_data = f.read()
                    # 替换换行符和间距
                    file_data = re.sub("<a:br>.+?</a:br>", "</a:p><a:p>", file_data)
                    file_data = re.sub('spc="-?[\d]+"', " ", file_data)
                    with open(xml, "w", encoding="utf-8") as f:
                        f.write(file_data)
        zip_file(os.path.basename(fs))

# 替换PPTX
def replacePPTX(path):
    files = getFiles(path)
    modefile(files)
    os.startfile(targetpath)

# GUI部分
root = tk.Tk()
root.title("ppt译前处理")
filePath = tk.StringVar()
dirPath = tk.StringVar()
stateLable = tk.StringVar()

w, h = 580, 360
sw, sh = (root.winfo_screenwidth() - w) // 2, (root.winfo_screenheight() - h) // 2
root.geometry(f'{w}x{h}+{sw}+{sh}')

tk.Label(root, text='文件：').grid(row=1, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=filePath).grid(row=1, column=1, columnspan=4, padx=5, pady=5, ipadx=145)
tk.Button(root, width=8, text='选择', command=selectFile).grid(row=1, column=5, padx=5, pady=5)

tk.Label(root, text='目录：').grid(row=2, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=dirPath).grid(row=2, column=1, columnspan=4, padx=5, pady=5, ipadx=145)
tk.Button(root, width=8, text='选择', command=selectDirFile).grid(row=2, column=5, padx=5, pady=5)

tk.Button(root, width=8, text='确定', command=fileSave, height=2).grid(row=5, column=5, padx=5, pady=5, rowspan=2, sticky=tk.N+tk.S)

tk.Label(root, text='* 新文件保存在原目录下的_target_文件夹中', fg="#666").grid(row=5, column=1, sticky=tk.NW, padx=5, pady=0)
state = tk.Label(root, textvariable=stateLable, fg="red").grid(row=6, column=1, padx=5, pady=0, sticky=tk.NW)

logs = tk.Text(root, height=11, width=70)
logs.grid(row=10, column=1, padx=0, pady=5, columnspan=5, sticky=tk.W)

scroll = tk.Scrollbar(root)
scroll.set(0.5, 1)
scroll.grid(row=10, column=5, sticky=tk.N+tk.S+tk.E, padx=10)
scroll.config(command=logs.yview)
logs.config(yscrollcommand=scroll.set)

tk.Label(root, text='替换完成后自动打开目标文件夹').grid(row=14, column=0, columnspan=6, pady=1)

root.mainloop()