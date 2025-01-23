import os
import tkinter as tk
from tkinter import filedialog
from pptx import Presentation
from pptx.util import Pt

def set_character_spacing(run, spacing):
    """通过修改 XML 设置字符间距。"""
    rPr = run._r.get_or_add_rPr()
    if spacing != 0:
        rPr.set('kern', str(int(spacing * 100)))  # 设置字符间距，单位是 1/100 磅
    else:
        if 'kern' in rPr.attrib:
            del rPr.attrib['kern']  # 移除字符间距设置

def adjust_text_format(presentation, font_scale, line_spacing, character_spacing, apply_spacing):
    """
    调整文本格式，包括字符间距、字体缩放和行距。
    遍历幻灯片的所有元素，包括表格、母版、文本框和组合。
    """
    def adjust_shape_text(shape):
        """调整形状中的文字，包括普通文本框和表格内文字。"""
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # 调整字体大小
                    if run.font.size is not None:
                        run.font.size = Pt(run.font.size.pt * font_scale)
                    # 设置字符间距
                    if apply_spacing:
                        set_character_spacing(run, character_spacing)
                # 设置段落行距
                paragraph.space_after = 0
                paragraph.space_before = 0
                paragraph.line_spacing = line_spacing

        # 如果形状是表格，遍历单元格
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text_frame:  # 确保单元格具有 text_frame
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                # 调整字体大小
                                if run.font.size is not None:
                                    run.font.size = Pt(run.font.size.pt * font_scale)
                                # 设置字符间距
                                if apply_spacing:
                                    set_character_spacing(run, character_spacing)
                            # 设置段落行距
                            paragraph.space_after = 0
                            paragraph.space_before = 0
                            paragraph.line_spacing = line_spacing

        # 如果形状是组合，递归处理组合中的每个形状
        if shape.shape_type == 6:  # 6 表示组合
            for sub_shape in shape.shapes:
                adjust_shape_text(sub_shape)

    # 遍历幻灯片内容
    for slide in presentation.slides:
        for shape in slide.shapes:
            adjust_shape_text(shape)

    # 遍历所有母版内容
    for slide_master in presentation.slide_masters:
        for layout in slide_master.slide_layouts:
            for shape in layout.shapes:
                adjust_shape_text(shape)

    return presentation

def process_ppt(input_path, output_folder, result_text, font_scale, line_spacing, character_spacing, apply_spacing):
    """处理单个 PPT 文件。"""
    try:
        presentation = Presentation(input_path)
        adjusted_presentation = adjust_text_format(presentation, font_scale, line_spacing, character_spacing, apply_spacing)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        output_ppt = os.path.join(output_folder, os.path.basename(input_path))
        adjusted_presentation.save(output_ppt)

        result_text.config(state=tk.NORMAL)
        result_text.insert(tk.END, f"处理成功：{output_ppt}\n")
        result_text.config(state=tk.DISABLED)
    except Exception as e:
        result_text.config(state=tk.NORMAL)
        result_text.insert(tk.END, f"处理失败：{os.path.basename(input_path)}，错误: {str(e)}\n")
        result_text.config(state=tk.DISABLED)

def browse_folder(entry):
    """选择文件夹并在 Entry 中显示路径。"""
    selected_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, selected_path)

def browse_file(entry):
    """选择单个文件并在 Entry 中显示路径。"""
    selected_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    entry.delete(0, tk.END)
    entry.insert(0, selected_path)

def process():
    """处理按钮的回调函数。"""
    folder_path = entry_folder.get()
    file_path = entry_file.get()
    result_text.config(state=tk.NORMAL)
    result_text.delete(1.0, tk.END)

    font_scale = float(entry_font_scale.get())
    line_spacing = float(entry_line_spacing.get())
    character_spacing = float(entry_character_spacing.get())
    apply_spacing = apply_spacing_var.get()

    if folder_path:
        output_folder = os.path.join(folder_path, "output")
        for filename in os.listdir(folder_path):
            if filename.endswith(".pptx"):
                input_path = os.path.join(folder_path, filename)
                process_ppt(input_path, output_folder, result_text, font_scale, line_spacing, character_spacing, apply_spacing)
    elif file_path:
        output_folder = os.path.join(os.path.dirname(file_path), "output")
        process_ppt(file_path, output_folder, result_text, font_scale, line_spacing, character_spacing, apply_spacing)

    result_text.config(state=tk.DISABLED)

# 创建主窗口
root = tk.Tk()
root.title("PPT 译后预处理工具")

# 创建 UI 组件
label_folder = tk.Label(root, text="选择文件夹:")
entry_folder = tk.Entry(root, width=50)
btn_browse_folder = tk.Button(root, text="选择文件夹", command=lambda: browse_folder(entry_folder))

label_file = tk.Label(root, text="选择文件:")
entry_file = tk.Entry(root, width=50)
btn_browse_file = tk.Button(root, text="选择文件", command=lambda: browse_file(entry_file))

label_font_scale = tk.Label(root, text="字体缩放比例:")
entry_font_scale = tk.Entry(root, width=10)
entry_font_scale.insert(0, "0.6")

label_line_spacing = tk.Label(root, text="行距:")
entry_line_spacing = tk.Entry(root, width=10)
entry_line_spacing.insert(0, "1.0")

label_character_spacing = tk.Label(root, text="字符间距（磅）:")
entry_character_spacing = tk.Entry(root, width=10)
entry_character_spacing.insert(0, "0.0")

apply_spacing_var = tk.BooleanVar()
chk_apply_spacing = tk.Checkbutton(root, text="统一字符间距", variable=apply_spacing_var)

btn_process = tk.Button(root, text="处理", command=process)

result_text = tk.Text(root, wrap=tk.WORD, height=15, width=60, state=tk.DISABLED)

# 布局
label_folder.grid(row=0, column=0, pady=5, sticky="w")
entry_folder.grid(row=0, column=1, pady=5, sticky="ew")
btn_browse_folder.grid(row=0, column=2, padx=5, pady=5)

label_file.grid(row=1, column=0, pady=5, sticky="w")
entry_file.grid(row=1, column=1, pady=5, sticky="ew")
btn_browse_file.grid(row=1, column=2, padx=5, pady=5)

label_font_scale.grid(row=2, column=0, pady=5, sticky="w")
entry_font_scale.grid(row=2, column=1, pady=5, sticky="w")

label_line_spacing.grid(row=3, column=0, pady=5, sticky="w")
entry_line_spacing.grid(row=3, column=1, pady=5, sticky="w")

label_character_spacing.grid(row=4, column=0, pady=5, sticky="w")
entry_character_spacing.grid(row=4, column=1, pady=5, sticky="w")

chk_apply_spacing.grid(row=5, column=0, pady=5, sticky="w")

btn_process.grid(row=6, column=0, columnspan=3, pady=10)

result_text.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# 调整窗口比例
root.columnconfigure(1, weight=1)
root.rowconfigure(7, weight=1)

# 启动主循环
root.mainloop()