#!/usr/bin/env python
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import *
from tkinter.filedialog import *
import zipfile, os, re, time, shutil

global dirpath, filepath, targetpath
filepath = ""
dirpath = ""

def selectFile():
    global filepath, dirpath, targetpath
    filepath = askopenfilename()  # 选择打开什么文件，返回文件名
    if filepath:
        filePath.set(filepath)
        targetpath = os.path.dirname(filepath)+"/_target_"
        
        # 重置目录选择窗口
        dirPath.set("")
        dirpath = ""

def selectDirFile():
    global dirpath, filepath, targetpath
    dirpath = askdirectory()
    if dirpath:
        dirPath.set(dirpath)   # 设置变量dirPath的值
        targetpath = dirpath+"/_target_"
        # 重置文件选择窗口
        filePath.set("")
        filepath = ""

def fileSave():
    global filepath, dirpath
    if dirpath or filepath:
        logs.delete(0.0, END)
        root.update()

        replacePPTX(filepath or dirpath)
    else:
        writeLog("文件或目录至少选择一个！")

def writeLog(msg):
    current_time = time.strftime('【%H:%M:%S】', time.localtime(time.time()))
    logmsg_in = str(current_time) + str(msg) + "\n"      # 换行
    logs.insert(END, logmsg_in)
    logs.yview_moveto(1)
    root.update()

# 生成文件list
def getFiles(directory):
    if os.path.isdir(directory):
        if directory[-1:] == "\\" or directory[-1:] == "/":
            files = [directory+f for f in os.listdir(directory)]
        else:
            files = [directory+"/"+f for f in os.listdir(directory)]
    else:
        files = [directory]
    return files

# 生成pptx
def zip_file(fn):
    global targetpath
    if not os.path.exists(targetpath):
        os.mkdir(targetpath)
    z = zipfile.ZipFile(targetpath+"/"+fn, 'w', zipfile.ZIP_DEFLATED)
    for dirpath, dirnames, filenames in os.walk("temp"):
        fpath = dirpath.replace("temp", '')
        fpath = fpath and fpath + os.sep or ''
        for filename in filenames:
            z.write(os.path.join(dirpath, filename), fpath+filename)
    z.close()
    shutil.rmtree("temp")
    writeLog("%s 替换成功" % fn)

# 替换内容
def replace_content(xmls):
    var_value = var.get()
    for xml in xmls:
        if os.path.isfile(xml):
            file_data = ""
            br = False
            with open(xml, "r", encoding="utf-8") as f:
                for line in f.readlines():
                    if re.search("<a:br>.+?</a:br>", line) is not None or re.search('spc="-?[\d]+"', line):
                        if var_value == 1:
                            line = re.sub("<a:br>.+?</a:br>", "</a:p><a:p>", line)
                            line = re.sub('spc="-?[\d]+"', " ", line)
                        elif var_value == 2:
                            line = re.sub("<a:br>.+?</a:br>", "</a:p><a:p>", line)
                        else:
                            line = re.sub("<a:br>.+?</a:br>", " ", line)
                        br = True
                    file_data += line
                f.close()
            if br:
                with open(xml, "w", encoding="utf-8") as nf:
                    nf.write(file_data)
                    nf.close()

# 解压并替换
def modefile(files):
    nf = []
    for f in files:
        if f.split(".")[-1:][0].lower() == "pptx":
            nf.append(f)
    files = nf
    length = len(files)
    cf = 0
    for fs in files:
        writeLog("-"*20+" * "+"-"*20)
        writeLog("开始处理文件："+fs)
        z = zipfile.ZipFile(fs, "r")
        z.extractall("temp")
        
        # 处理幻灯片
        slide_xmls = getFiles("temp/ppt/slides")
        replace_content(slide_xmls)
        
        # 处理母版
        master_xmls = getFiles("temp/ppt/slideMasters")
        replace_content(master_xmls)
        
        zip_file(fs.split("/")[-1:][0])
        cf += 1
        stateLable.set("* 共%s个pptx文件，已处理%s个，还剩下%s个" % (length, cf, length-cf))

def replacePPTX(path):
    global targetpath
    files = getFiles(path)
    modefile(files)
    os.startfile(targetpath)


root = tk.Tk()
root.title("ppt译前处理")
filePath = tk.StringVar()
dirPath = tk.StringVar()
stateLable = tk.StringVar()

w = 580
h = 360
sw = int((root.winfo_screenwidth()-w)/2)
sh = int((root.winfo_screenheight()-h)/2)

root.geometry('%sx%s+%s+%s' % (w, h, sw, sh))

tk.Label(root, text='文件：').grid(row=1, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=filePath).grid(row=1, column=1, columnspan=4, padx=5, pady=5, ipadx=145,)
tk.Button(root, width=8, text='选择', command=selectFile).grid(row=1, column=5, padx=5, pady=5)

tk.Label(root, text='目录：').grid(row=2, column=0, padx=5, pady=5)
tk.Entry(root, textvariable=dirPath).grid(row=2, column=1, padx=5, columnspan=4, pady=5, ipadx=145,)
tk.Button(root, width=8, text='选择', command=selectDirFile).grid(row=2, column=5, padx=5, pady=5)

var = tk.IntVar()
var.set(1)
tk.Radiobutton(root, text="替换成硬回车+统一间距", variable=var, value=1).grid(row=3, column=1, pady=5, sticky=W)
tk.Radiobutton(root, text="替换成硬回车", variable=var, value=2).grid(row=3, column=2,  pady=5, sticky=W)
tk.Radiobutton(root, text="替换成空格", variable=var, value=3).grid(row=3, column=3,  pady=5, sticky=W)

tk.Button(root, width=8, text='确定', command=fileSave, height=2).grid(row=5, column=5, padx=5, pady=5, rowspan=2, sticky=N+S)

tk.Label(root, text='* 新文件保存在原目录下的_target_文件夹中', fg="#666").grid(row=5, column=1, sticky=NW, padx=5, pady=0)
state = tk.Label(root, textvariable=stateLable, fg="red").grid(row=6, column=1, padx=5, pady=0, sticky=NW,)

logs = tk.Text(root, height=11, width=70)
logs.grid(row=10, column=1, padx=0, pady=5, columnspan=5, sticky=W,)

scroll = tk.Scrollbar(root)
scroll.set(0.5, 1)
scroll.grid(row=10, column=5, sticky=N+S+E, padx=10)

scroll.config(command=logs.yview)
logs.config(yscrollcommand=scroll.set)

tk.Label(root, text='替换完成后自动打开目标文件夹').grid(row=14, column=0, columnspan=6, pady=1)

root.mainloop()
