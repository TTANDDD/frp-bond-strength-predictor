import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import joblib
import os
import sys
import logging
import pandas as pd
import traceback
import threading
from PIL import Image, ImageTk

# ------------------------- 全局异常处理 -------------------------
def except_hook(exc_type, exc_value, exc_traceback):
    """
    全局异常钩子，用于捕获未处理的异常并记录日志和弹出错误对话框。
    """
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    error_message = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    logging.error(f"未捕获的异常: {error_message}")
    messagebox.showerror("未捕获的异常", error_message)

# 设置全局异常钩子
sys.excepthook = except_hook

# ------------------------- 日志记录配置 -------------------------
logging.basicConfig(
    filename='prediction_app.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ------------------------- 资源路径函数 -------------------------
def resource_path(relative_path):
    """
    获取资源文件的绝对路径，兼容打包后运行。
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ------------------------- 加载模型和缩放器 -------------------------
try:
    model_path = resource_path("model.pkl")  # 模型文件路径
    scaler_path = resource_path("scaler.pkl")  # 缩放器文件路径
    model = joblib.load(model_path)  # 加载预训练模型
    scaler = joblib.load(scaler_path)  # 加载预训练缩放器
    logging.info("模型和缩放器加载成功")
except Exception as e:
    logging.error(f"加载模型或缩放器时出错: {str(e)}")
    messagebox.showerror("错误", f"加载模型或缩放器时出错: {str(e)}")
    sys.exit(1)

# ------------------------- 创建主窗口 -------------------------
root = tk.Tk()
root.title("FRP-混凝土粘结强度预测")
root.geometry("1200x800")  # 增大宽度以容纳批量预测功能
root.configure(bg="#f0f4f7")  # 浅灰蓝色背景

# ------------------------- 设置样式 -------------------------
style = ttk.Style()
style.theme_use('clam')  # 使用 'clam' 主题，可以更好地自定义

# 定义颜色
primary_color = "#4a90e2"  # 蓝色
secondary_color = "#ffffff"  # 白色
accent_color = "#50e3c2"  # 青绿色
text_color = "#333333"  # 深灰色
frame_background = "#f0f4f7"  # 浅灰蓝色

# 设置TFrame背景
style.configure('TFrame', background=frame_background)

# 设置TLabel样式
style.configure('TLabel', background=frame_background, foreground=text_color, font=('Helvetica', 12))

# 设置TEntry样式
style.configure('TEntry', foreground=text_color, font=('Helvetica', 12))
style.map('TEntry',
          foreground=[('readonly', text_color)],
          background=[('readonly', '#e6e6e6')])  # 只读状态下背景变灰

# 设置TButton样式
style.configure('TButton', font=('Helvetica', 12, 'bold'), padding=6)
style.map('TButton',
          background=[('active', primary_color)],
          foreground=[('active', secondary_color)])

# 设置TCombobox样式
style.configure('TCombobox', foreground=text_color, font=('Helvetica', 12), padding=5)
style.map('TCombobox',
          fieldbackground=[('readonly', '#ffffff')],
          background=[('readonly', '#ffffff')])

# 设置TNotebook样式
style.configure('TNotebook', background=frame_background)
style.configure('TNotebook.Tab', font=('Helvetica', 12, 'bold'), padding=[10, 5], foreground=text_color)
style.map('TNotebook.Tab',
          background=[('selected', primary_color)],
          foreground=[('selected', secondary_color)])

# ------------------------- 语言翻译字典 -------------------------
translations = {
    'en': {
        "title": "FRP-Concrete Bond Strength Prediction",
        "tab_individual": "Individual Prediction",
        "tab_batch": "Batch Prediction",
        "tab_instructions": "Home",
        "labels": [
            'T (Temperature)',  # 温度
            'FM (Failure Mode)',  # 破坏模式
            'FT (FRP Type)',  # FRP类型
            'BS (FRP Bar Surface)',  # 表面形式
            'd (FRP Bar Diameter)',  # FRP筋直径
            'la (Anchorage Length)',  # 锚固长度
            'fc (Concrete Compressive Strength)',  # 混凝土抗压强度
            'c/d (Concrete Cover / Bar Diameter)'  # 归一化混凝土保护层厚度
        ],
        "units": ['°C', '', '', '', 'mm', 'mm', 'MPa', ''],
        "predict_button": "Predict",
        "result_prefix": "Predicted τu: ",
        "batch_title": "Batch Prediction",
        "select_file": "Selected Excel File:",
        "browse": "Browse",
        "batch_button": "Run Batch Prediction",
        "batch_status_initial": "",
        "batch_status_select": "File selected. Ready to start batch prediction.",
        "batch_status_start": "Starting batch prediction...",
        "batch_status_reading": "Reading file...",
        "batch_status_scaling": "Scaling input data...",
        "batch_status_predicting": "Performing prediction...",
        "batch_status_saving": "Saving prediction results...",
        "success_save": "Batch prediction completed. Results saved to ",
        "warning_no_file": "No file selected",
        "warning_no_file_msg": "Please select an Excel file for batch prediction.",
        "error_input": "Input Error: Please ensure all fields contain valid numbers.",
        "error_occurred": "An error occurred:",
        "success": "Success",
        "error_title": "Error",
        "FM_values": [("FRP-sand layer debonding", 1), ("FRP bar pullout", 2), ("Concrete shear failure", 3),
                      ("FRP bar rupture", 4), ("Concrete splitting", 5)],
        "FT_values": [("GFRP", 1), ("CFRP", 2), ("BFRP", 3)],
        "BS_values": [("Sand-coated", 1), ("Fiber-wrapped sand-coated", 2), ("Ribbed", 3),
                      ("Fiber-wrapped", 4), ("Sand-coated ribbed", 5)],
        "instructions_text": """
**Inputs:**
1. **T (Temperature)**: Temperature in °C.
2. **FM (Failure Mode)**: Mode of failure.
   - 1: FRP-sand layer debonding
   - 2: FRP bar pullout
   - 3: Concrete shear failure
   - 4: FRP bar rupture
   - 5: Concrete splitting
3. **FT (FRP Type)**: Type of FRP.
   - 1: GFRP
   - 2: CFRP
   - 3: BFRP
4. **BS (FRP Bar Surface)**: Surface form of FRP bar.
   - 1: Sand-coated
   - 2: Fiber-wrapped sand-coated
   - 3: Ribbed
   - 4: Fiber-wrapped
   - 5: Sand-coated ribbed
5. **d (FRP Bar Diameter)**: Diameter of FRP bar in mm.
6. **la (Anchorage Length)**: Anchorage length in mm.
7. **fc (Concrete Compressive Strength)**: Compressive strength of concrete in MPa.
8. **c/d (Concrete Cover / Bar Diameter)**: Normalized concrete cover thickness.

**Output:**
- **τu (Interface Bond Strength)**: Interface bond strength in MPa.

**Batch Prediction Instructions:**
- Ensure that the Excel file has at least 8 columns corresponding to the input features in the following order:
  1. T (Temperature)
  2. FM (Failure Mode)
  3. FT (FRP Type)
  4. BS (FRP Bar Surface)
  5. d (FRP Bar Diameter)
  6. la (Anchorage Length)
  7. fc (Concrete Compressive Strength)
  8. c/d (Concrete Cover / Bar Diameter)
- Discrete features (FM, FT, BS) should be encoded as specified above.
- All feature columns must contain valid numerical values.

""",
        "experiment_diagram_desc": "Figure 1: Experimental Setup Schematic",
        "dt_model_structure_desc": "Figure 2: Decision Tree Model Structure"
    },
    'zh': {
        "title": "FRP-混凝土粘结强度预测",
        "tab_individual": "单个预测",
        "tab_batch": "批量预测",
        "tab_instructions": "主页",
        "labels": [
            'T (温度)',  # 温度
            'FM (失效模式)',  # 破坏模式
            'FT (FRP类型)',  # FRP类型
            'BS (FRP钢筋表面)',  # 表面形式
            'd (FRP钢筋直径)',  # FRP筋直径
            'la (锚固长度)',  # 锚固长度
            'fc (混凝土抗压强度)',  # 混凝土抗压强度
            'c/d (混凝土覆盖/钢筋直径)'  # 归一化混凝土保护层厚度
        ],
        "units": ['°C', '', '', '', 'mm', 'mm', 'MPa', ''],
        "predict_button": "预测",
        "result_prefix": "预测 τu: ",
        "batch_title": "批量预测",
        "select_file": "选择的Excel文件:",
        "browse": "浏览",
        "batch_button": "运行批量预测",
        "batch_status_initial": "",
        "batch_status_select": "文件已选择。准备开始批量预测。",
        "batch_status_start": "开始批量预测...",
        "batch_status_reading": "正在读取文件...",
        "batch_status_scaling": "正在缩放输入数据...",
        "batch_status_predicting": "正在进行预测...",
        "batch_status_saving": "正在保存预测结果...",
        "success_save": "批量预测完成，结果已保存到 ",
        "warning_no_file": "未选择文件",
        "warning_no_file_msg": "请选择一个Excel文件进行批量预测。",
        "error_input": "输入错误: 请确保所有字段包含有效数字。",
        "error_occurred": "发生错误:",
        "success": "成功",
        "error_title": "错误",
        "FM_values": [("FRP沙层脱粘", 1), ("FRP钢筋拔出", 2), ("混凝土剪切破坏", 3),
                      ("FRP钢筋断裂", 4), ("混凝土劈裂", 5)],
        "FT_values": [("GFRP", 1), ("CFRP", 2), ("BFRP", 3)],
        "BS_values": [("砂涂层", 1), ("纤维包裹砂涂层", 2), ("肋纹", 3),
                      ("纤维包裹", 4), ("砂涂层肋纹", 5)],
        "instructions_text": """
**输入特征:**
1. **T (温度)**: 温度，单位为°C。
2. **FM (失效模式)**: 破坏模式。
   - 1: FRP沙层脱粘
   - 2: FRP钢筋拔出
   - 3: 混凝土剪切破坏
   - 4: FRP钢筋断裂
   - 5: 混凝土劈裂
3. **FT (FRP类型)**: FRP类型。
   - 1: GFRP
   - 2: CFRP
   - 3: BFRP
4. **BS (FRP钢筋表面)**: FRP钢筋表面形式。
   - 1: 砂涂层
   - 2: 纤维包裹砂涂层
   - 3: 肋纹
   - 4: 纤维包裹
   - 5: 砂涂层肋纹
5. **d (FRP钢筋直径)**: FRP钢筋直径，单位为mm。
6. **la (锚固长度)**: 锚固长度，单位为mm。
7. **fc (混凝土抗压强度)**: 混凝土抗压强度，单位为MPa。
8. **c/d (混凝土覆盖/钢筋直径)**: 归一化混凝土保护层厚度。

**输出特征:**
- **界面粘结强度 (τu)**: 界面粘结强度，单位为MPa。

**批量预测使用说明:**
- 请确保Excel文件至少包含8列，对应以下输入特征，顺序如下：
  1. T (温度)
  2. FM (失效模式)
  3. FT (FRP类型)
  4. BS (FRP钢筋表面)
  5. d (FRP钢筋直径)
  6. la (锚固长度)
  7. fc (混凝土抗压强度)
  8. c/d (混凝土覆盖/钢筋直径)
- 离散特征（FM, FT, BS）请按照上述编码方式进行数字编码。
- 所有特征列必须包含有效的数值。

""",
        "experiment_diagram_desc": "图1: 实验设置示意图",
        "dt_model_structure_desc": "图2: 决策树模型结构图"
    }
}

current_language = 'zh'  # 默认语言为中文

# ------------------------- 语言切换函数 -------------------------
def set_language(lang):
    """
    设置当前语言并更新界面文本。
    """
    global current_language
    current_language = lang
    update_language()

def update_language():
    """
    根据当前语言更新界面上的所有文本。
    """
    lang = translations[current_language]

    # 更新窗口标题
    root.title(lang["title"])

    # 更新Notebook标签
    notebook.tab(instructions_frame, text=lang["tab_instructions"])
    notebook.tab(individual_frame, text=lang["tab_individual"])
    notebook.tab(batch_frame, text=lang["tab_batch"])

    # 更新单个预测选项卡
    title_label.config(text=lang["title"])
    for i, (label, unit) in enumerate(zip(lang["labels"], lang["units"])):
        input_frame.grid_slaves(row=i, column=0)[0].config(text=label)
        if unit:
            input_frame.grid_slaves(row=i, column=2)[0].config(text=unit)

    # 更新下拉菜单选项
    fm_combobox['values'] = [f"{desc} ({val})" for desc, val in lang["FM_values"]]
    fm_combobox.current(0)

    ft_combobox['values'] = [f"{desc} ({val})" for desc, val in lang["FT_values"]]
    ft_combobox.current(0)

    bs_combobox['values'] = [f"{desc} ({val})" for desc, val in lang["BS_values"]]
    bs_combobox.current(0)

    predict_button.config(text=lang["predict_button"])

    # 更新批量预测选项卡
    batch_title.config(text=lang["batch_title"])
    file_label.config(text=lang["select_file"])
    browse_button.config(text=lang["browse"])
    batch_button.config(text=lang["batch_button"])
    batch_status.config(text=lang["batch_status_initial"])

    # 更新使用说明内容
    instructions_text.config(state='normal')
    instructions_text.delete(1.0, tk.END)
    instructions_text.insert(tk.END, lang["instructions_text"])
    instructions_text.config(state='disabled')

    # 更新图片说明
    experiment_desc_label.config(text=lang["experiment_diagram_desc"])
    dt_model_desc_label.config(text=lang["dt_model_structure_desc"])

# ------------------------- 添加菜单栏供语言选择 -------------------------
menubar = tk.Menu(root)
root.config(menu=menubar)

language_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Language", menu=language_menu)
language_menu.add_command(label="English", command=lambda: set_language('en'))
language_menu.add_command(label="中文", command=lambda: set_language('zh'))

# ------------------------- 加载图标函数 -------------------------
# 检查 Pillow 版本，确保兼容性
if hasattr(Image, 'Resampling'):
    resample_mode = Image.Resampling.LANCZOS
else:
    resample_mode = Image.ANTIALIAS  # 兼容旧版本

def load_icon(path, size=(20, 20)):
    """
    加载并调整图标大小，保持长宽比例。
    """
    try:
        image = Image.open(resource_path(path))
        image.thumbnail(size, resample_mode)  # 使用 thumbnail 保持比例
        return ImageTk.PhotoImage(image)
    except Exception as e:
        logging.error(f"加载图标 {path} 时出错: {str(e)}")
        return None

# 加载各个按钮的图标
predict_icon = load_icon("predict_icon.png")
browse_icon = load_icon("browse_icon.png")
batch_icon = load_icon("batch_icon.png")

# ------------------------- 创建Notebook -------------------------
notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True)

# ======================== 使用说明选项卡 ===========================
instructions_frame = ttk.Frame(notebook, padding="20 20 20 20", style='TFrame')
notebook.add(instructions_frame, text=translations[current_language]["tab_instructions"])

# 创建一个水平布局的子框架，左侧为文本，右侧为图片
instructions_main_frame = ttk.Frame(instructions_frame, style='TFrame')
instructions_main_frame.pack(fill=tk.BOTH, expand=True)

# 左侧：使用说明文本框（只读）
instructions_text = tk.Text(instructions_main_frame, wrap=tk.WORD, font=('Helvetica', 12), bg=frame_background,
                            fg=text_color, borderwidth=0)
instructions_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# 添加滚动条
scrollbar = ttk.Scrollbar(instructions_main_frame, orient=tk.VERTICAL, command=instructions_text.yview)
scrollbar.pack(side=tk.LEFT, fill=tk.Y)
instructions_text.configure(yscrollcommand=scrollbar.set)

# 插入初始说明文本
instructions_text.insert(tk.END, translations[current_language]["instructions_text"])
instructions_text.config(state='disabled')  # 设置为只读

# 右侧：图片及说明
images_frame = ttk.Frame(instructions_main_frame, style='TFrame')
images_frame.pack(side=tk.LEFT, padx=20, fill=tk.Y, anchor='n')

def load_image_aspect(path, max_size=(400, 300)):
    """
    加载图像并按比例调整大小，不改变原始比例。
    """
    try:
        image = Image.open(resource_path(path))
        image.thumbnail(max_size, resample_mode)  # 使用 thumbnail 以保持比例
        return ImageTk.PhotoImage(image)
    except Exception as e:
        logging.error(f"加载图像 {path} 时出错: {str(e)}")
        return None

# 加载实验示意图
experiment_diagram_img = load_image_aspect("experiment_diagram.jpg")

# 实验示意图标签
if experiment_diagram_img:
    experiment_label = ttk.Label(images_frame, image=experiment_diagram_img)
    experiment_label.pack(pady=10)
    # 图片说明标签
    experiment_desc_label = ttk.Label(images_frame, text=translations[current_language]["experiment_diagram_desc"],
                                      wraplength=400, font=('Helvetica', 10))
    experiment_desc_label.pack()
else:
    experiment_desc_label = ttk.Label(images_frame, text="", wraplength=400, font=('Helvetica', 10))
    experiment_desc_label.pack()

# 加载DT模型结构图
dt_model_structure_img = load_image_aspect("dt_model_structure.jpg")

# DT模型结构图标签
if dt_model_structure_img:
    dt_model_label = ttk.Label(images_frame, image=dt_model_structure_img)
    dt_model_label.pack(pady=10)
    # 图片说明标签
    dt_model_desc_label = ttk.Label(images_frame, text=translations[current_language]["dt_model_structure_desc"],
                                    wraplength=400, font=('Helvetica', 10))
    dt_model_desc_label.pack()
else:
    dt_model_desc_label = ttk.Label(images_frame, text="", wraplength=400, font=('Helvetica', 10))
    dt_model_desc_label.pack()

# ======================= 单个预测选项卡 =======================
individual_frame = ttk.Frame(notebook, padding="20 20 20 20", style='TFrame')
notebook.add(individual_frame, text=translations[current_language]["tab_individual"])

# 创建标题
title_label = tk.Label(
    individual_frame,
    text=translations[current_language]["title"],
    font=('Helvetica', 20, 'bold'),
    bg=frame_background
)
title_label.pack(pady=(0, 20))  # 上方不留空，下方留20

# 创建一个水平布局的主框架，将输入框和图片框架并排放置
main_input_image_frame = ttk.Frame(individual_frame, style='TFrame')
main_input_image_frame.pack(fill=tk.BOTH, expand=True)

# 左侧：输入框框架
input_frame = ttk.Frame(main_input_image_frame, padding="10 10 10 10", style='TFrame')
input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# 定义标签和单位
labels = translations[current_language]["labels"]
units = translations[current_language]["units"]

# 变量用于存储用户输入
variables = []

# 定义下拉菜单的选项
lang_fm_values = translations[current_language]["FM_values"]
lang_ft_values = translations[current_language]["FT_values"]
lang_bs_values = translations[current_language]["BS_values"]

# 定义存储Combobox控件的变量
combobox_widgets = {}

for i, (label, unit) in enumerate(zip(labels, units)):
    # 添加标签
    ttk.Label(input_frame, text=label).grid(row=i, column=0, padx=10, pady=10, sticky='e')

    # 判断当前特征是否为离散特征（FM, FT, BS）
    if label.startswith('FM') or label.startswith('FT') or label.startswith('BS'):
        var = tk.StringVar(individual_frame)
        if label.startswith('FM'):
            values = lang_fm_values
            combobox_type = 'FM'
        elif label.startswith('FT'):
            values = lang_ft_values
            combobox_type = 'FT'
        else:
            values = lang_bs_values
            combobox_type = 'BS'
        # 创建只读下拉菜单，并设置默认选择
        combo = ttk.Combobox(
            input_frame,
            textvariable=var,
            values=[f"{desc} ({val})" for desc, val in values],
            state="readonly",
            style='TCombobox'
        )
        combo.grid(row=i, column=1, padx=10, pady=10, sticky='ew')
        combo.current(0)  # 设置默认选择
        variables.append(var)

        # 保存Combobox引用以便后续更新（用于语言切换）
        combobox_widgets[combobox_type] = combo
    else:
        # 创建文本输入框用于连续特征
        entry = ttk.Entry(input_frame)
        entry.grid(row=i, column=1, padx=10, pady=10, sticky='ew')
        variables.append(entry)

    # 添加单位标签（如果有）
    if unit:
        ttk.Label(input_frame, text=unit).grid(row=i, column=2, padx=10, pady=10, sticky='w')

# 确保列0、1、2在输入框框架中有正确的权重
input_frame.columnconfigure(0, weight=1)
input_frame.columnconfigure(1, weight=3)
input_frame.columnconfigure(2, weight=1)

# 右侧：图片框架
image_frame = ttk.Frame(main_input_image_frame, style='TFrame')
image_frame.pack(side=tk.LEFT, padx=20, fill=tk.BOTH, expand=True)

# 加载实验示意图
individual_experiment_diagram_img = load_image_aspect("experiment_diagram.jpg")

# 实验示意图标签
if individual_experiment_diagram_img:
    individual_experiment_label = ttk.Label(image_frame, image=individual_experiment_diagram_img)
    individual_experiment_label.pack(pady=10)
    # 保存对图像的引用以防止被垃圾回收
    individual_experiment_label.image = individual_experiment_diagram_img

    # 添加说明标签
    individual_experiment_desc_label = ttk.Label(
        image_frame,
        text=translations[current_language]["experiment_diagram_desc"],
        wraplength=400,
        font=('Helvetica', 10)
    )
    individual_experiment_desc_label.pack()
else:
    individual_experiment_label = ttk.Label(image_frame, text="", wraplength=400, font=('Helvetica', 10))
    individual_experiment_label.pack()

def predict():
    """
    进行单个预测的函数。
    """
    lang = translations[current_language]
    result_label.config(text=lang["result_prefix"])
    individual_frame.update_idletasks()  # 立即更新GUI
    try:
        # 获取输入值并按顺序存储
        input_values = []
        for var in variables:
            if isinstance(var, tk.StringVar):
                # 处理离散特征，提取编号
                value = int(var.get().split('(')[-1].strip(')'))
            else:
                # 处理连续特征，转换为浮点数
                value = float(var.get())
            input_values.append(value)

        # 特征顺序：T, FM, FT, BS, d, la, fc, c/d
        # 确保顺序不变
        # 输入特征的顺序已在变量定义时保持

        # 缩放输入
        input_scaled = scaler.transform([input_values])

        # 进行预测
        prediction = model.predict(input_scaled)[0]

        # 显示结果
        result_label.config(text=f"{lang['result_prefix']}{prediction:.2f} MPa")
        logging.info(f"预测成功。输入: {input_values}, 预测值: {prediction:.2f}")
    except ValueError as ve:
        # 处理输入值错误
        error_msg = lang["error_input"]
        result_label.config(text=error_msg)
        logging.error(f"输入值错误: {str(ve)}")
    except Exception as e:
        # 处理其他错误
        error_msg = f"{lang['error_occurred']} {str(e)}"
        result_label.config(text=error_msg)
        logging.error(f"预测过程中出错: {str(e)}")

# 创建预测按钮（带图标）
predict_button = ttk.Button(
    individual_frame,
    text=translations[current_language]["predict_button"],
    command=predict,
    compound=tk.LEFT,
    image=predict_icon
)
predict_button.pack(pady=20)

# 创建结果标签
result_label = tk.Label(
    individual_frame,
    text="",
    font=('Helvetica', 14, 'bold'),
    bg=frame_background,
    foreground=primary_color
)
result_label.pack(pady=20)

# 保存Combobox引用以便后续更新（用于语言切换）
fm_combobox = combobox_widgets.get('FM')
ft_combobox = combobox_widgets.get('FT')
bs_combobox = combobox_widgets.get('BS')

# ------------------------- 继续其他代码 -------------------------

# ======================== 批量预测选项卡 ===========================
batch_frame = ttk.Frame(notebook, padding="20 20 20 20", style='TFrame')
notebook.add(batch_frame, text=translations[current_language]["tab_batch"])

# 批量预测的标题
batch_title = tk.Label(
    batch_frame,
    text=translations[current_language]["batch_title"],
    font=('Helvetica', 20, 'bold'),
    bg=frame_background
)
batch_title.pack(pady=(0, 20))

# 创建一个用于居中的外部帧
center_file_frame = ttk.Frame(batch_frame, style='TFrame')
center_file_frame.pack(expand=True)

# 选择的文件路径变量
selected_file = tk.StringVar()

# 文件选择的框架
file_frame = ttk.Frame(center_file_frame, padding="10 10 10 10", style='TFrame')
file_frame.pack(pady=10)

# 文件选择标签
file_label = ttk.Label(file_frame, text=translations[current_language]["select_file"])
file_label.pack(side=tk.LEFT, padx=5)

# 文件路径显示的只读输入框
file_entry = ttk.Entry(file_frame, textvariable=selected_file, width=50, state='readonly')
file_entry.pack(side=tk.LEFT, padx=5)

def browse_file():
    """
    浏览并选择Excel文件用于批量预测。
    """
    lang = translations[current_language]
    file_path = filedialog.askopenfilename(
        title=lang["browse"],
        filetypes=[("Excel文件", "*.xlsx *.xls")] if current_language == 'zh' else [("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        selected_file.set(file_path)
        logging.info(f"选择用于批量预测的文件: {file_path}")
        batch_status.config(text=lang["batch_status_select"])

# 浏览按钮（带图标）
browse_button = ttk.Button(
    file_frame,
    text=translations[current_language]["browse"],
    command=lambda: browse_file(),
    compound=tk.LEFT,
    image=browse_icon
)
browse_button.pack(side=tk.LEFT, padx=5)

def batch_predict():
    """
    进行批量预测的函数。
    """
    lang = translations[current_language]
    file_path = selected_file.get()
    if not file_path:
        # 如果未选择文件，弹出警告
        messagebox.showwarning(lang["warning_no_file"], lang["warning_no_file_msg"])
        return

    # 禁用按钮以防止重复点击
    batch_button.config(state='disabled')
    browse_button.config(state='disabled')
    batch_status.config(text=lang["batch_status_start"])
    progress_bar.start()

    def run_prediction():
        """
        实际执行批量预测的子线程函数。
        """
        try:
            # 更新状态为读取文件
            batch_status.config(text=lang["batch_status_reading"])

            # 读取Excel文件，假设没有列名
            df = pd.read_excel(file_path, header=None)
            logging.info(f"成功加载Excel文件 {file_path}。")

            # 将所有列名转换为字符串类型
            df.columns = df.columns.astype(str)

            # 检查Excel是否至少有8列
            if df.shape[1] < 8:
                raise ValueError(
                    "选择的Excel文件必须至少有8列特征数据。" if current_language == 'zh' else "The selected Excel file must contain at least 8 feature columns."
                )

            # 假设前8列是所需的特征，按顺序：T, FM, FT, BS, d, la, fc, c/d
            feature_columns = df.columns[:8]
            input_data = df[feature_columns]

            # 确保所有必要的列都是数值型
            if not all(pd.api.types.is_numeric_dtype(df[col]) for col in feature_columns):
                raise ValueError(
                    "所有特征列必须为数值型。请检查Excel文件。" if current_language == 'zh' else "All feature columns must be numeric. Please check the Excel file."
                )

            # 更新状态为缩放输入数据
            batch_status.config(text=lang["batch_status_scaling"])
            # 缩放输入
            input_scaled = scaler.transform(input_data)
            logging.info("输入数据缩放成功。")

            # 更新状态为进行预测
            batch_status.config(text=lang["batch_status_predicting"])
            # 进行预测
            predictions = model.predict(input_scaled)
            logging.info("批量预测完成。" if current_language == 'zh' else "Batch prediction completed.")

            # 将预测结果添加到DataFrame
            df['Predicted τu (MPa)'] = predictions

            # 更新状态为保存预测结果
            batch_status.config(text=lang["batch_status_saving"])
            # 保存更新后的DataFrame到新的Excel文件
            save_title = lang["browse"] if current_language == 'zh' else "Save Excel File"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx *.xls")] if current_language == 'zh' else [
                    ("Excel Files", "*.xlsx *.xls")],
                title=save_title,
                initialfile="predicted_" + os.path.basename(file_path)
            )
            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo(lang["success"], f"{lang['success_save']}{save_path}")
                logging.info(f"预测结果已保存到 {save_path}")
                batch_status.config(
                    text=lang["batch_status_start"] if current_language == 'zh' else "Batch prediction completed."
                )
            else:
                batch_status.config(
                    text="保存操作已取消。" if current_language == 'zh' else "Save operation cancelled."
                )
        except Exception as e:
            # 捕获并处理预测过程中的任何异常
            messagebox.showerror(lang["error_title"], f"{lang['error_occurred']} {str(e)}")
            logging.error(
                f"批量预测过程中出错: {str(e)}" if current_language == 'zh' else f"Error during batch prediction: {str(e)}"
            )
            batch_status.config(text=f"{lang['error_occurred']} {str(e)}")
        finally:
            # 最终步骤，无论成功与否都要执行
            progress_bar.stop()
            batch_button.config(state='normal')
            browse_button.config(state='normal')

    # 使用线程来避免阻塞GUI
    threading.Thread(target=run_prediction).start()

# 创建批量预测按钮（带图标）
batch_button = ttk.Button(
    batch_frame,
    text=translations[current_language]["batch_button"],
    command=batch_predict,
    compound=tk.LEFT,
    image=batch_icon
)
batch_button.pack(pady=20)

# 批量预测的状态标签
batch_status = tk.Label(
    batch_frame,
    text=translations[current_language]["batch_status_initial"],
    font=('Helvetica', 12),
    bg=frame_background,
    foreground=text_color
)
batch_status.pack(pady=10)

# 添加进度条
progress_bar = ttk.Progressbar(batch_frame, orient=tk.HORIZONTAL, length=400, mode='indeterminate')
progress_bar.pack(pady=10)

# ========================== 功能完成 ===========================

# ------------------------- 初始化界面语言 -------------------------
update_language()

# ------------------------- 运行主循环 -------------------------
root.mainloop()