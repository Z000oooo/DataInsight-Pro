#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：DataInsight Pro - 智能数据分析工具
@Author  ：YOLO检测与算法
@User    ：YOLO小王
@Version ：v1.0
@Date    ：2025/7/6 下午2:09
@Description：专业的数据清洗、分析与可视化一体化解决方案
'''

# 导入所需的库
import tkinter as tk
from tkinter import ttk, filedialog, messagebox  # GUI相关库
import pandas as pd  # 数据处理库
import numpy as np   # 数值计算库
# 可视化相关库将按需惰性导入以减少启动时间
import io
import os
from datetime import datetime
import webbrowser
import tempfile
# 机器学习相关库将按需惰性导入以减少启动时间
import warnings
warnings.filterwarnings('ignore')  # 忽略警告信息

# 设置matplotlib中文字体和样式
# 图表样式配置将在首次初始化可视化后设置

class DataInsightPro:
    """DataInsight Pro - 智能数据分析工具主类"""
    
    def __init__(self, root):
        """初始化应用程序"""
        self.root = root
        self.root.title("DataInsight Pro - 智能数据分析工具 v1.0")  # 设置窗口标题
        self.root.geometry("1500x950")  # 设置窗口大小
        self.root.configure(bg='#ecf0f1')  # 设置背景色
        self.root.state('zoomed')  # Windows下最大化窗口
        
        # 数据存储变量
        self.data = None  # 当前处理的数据
        self.original_data = None  # 原始数据备份
        self.cleaned_data = None  # 清洗后的数据
        
        # 惰性导入与可视化句柄
        self.plt = None
        self.matplotlib = None
        self.FigureCanvasTkAgg = None
        self.sns = None
        self.px = None
        self.ax = None
        self.fig = None
        self.canvas = None
        
        # 配置样式
        self.setup_styles()
        
        # 创建主界面
        self.create_interface()
        
        # 设置快捷键
        self.setup_shortcuts()
        
    def setup_styles(self):
        """配置GUI样式"""
        style = ttk.Style()
        
        # 尝试使用现代主题
        try:
            style.theme_use('vista')  # Windows Vista主题，更现代
        except:
            try:
                style.theme_use('xpnative')  # Windows XP主题
            except:
                style.theme_use('clam')  # 备用主题
        
        # 配置自定义样式 - 现代化设计
        style.configure('Title.TLabel', 
                       font=('微软雅黑', 20, 'bold'), 
                       background='#2c3e50',
                       foreground='white',
                       padding=(20, 10))
        
        style.configure('Header.TLabel', 
                       font=('微软雅黑', 14, 'bold'), 
                       background='#ecf0f1',
                       foreground='#2c3e50',
                       padding=(10, 5))
        
        style.configure('Custom.TButton', 
                       font=('微软雅黑', 11),
                       padding=(10, 8),
                       relief='flat')
        
        style.map('Custom.TButton',
                 background=[('active', '#3498db'),
                            ('pressed', '#2980b9')])
        
        # 配置LabelFrame样式
        style.configure('TLabelframe', 
                       background='#ecf0f1',
                       borderwidth=2,
                       relief='solid')
        
        style.configure('TLabelframe.Label', 
                       font=('微软雅黑', 12, 'bold'),
                       background='#ecf0f1',
                       foreground='#2c3e50')
        
        # 配置Notebook样式
        style.configure('TNotebook', 
                       background='#ecf0f1',
                       borderwidth=0,
                       tabmargins=[2, 8, 2, 0])
        
        style.configure('TNotebook.Tab', 
                       font=('微软雅黑', 11, 'bold'),
                       padding=(12, 12),
                       background='#bdc3c7',
                       foreground='#2c3e50',
                       relief='flat',
                       borderwidth=1,
                       focuscolor='none')
        
        style.map('TNotebook.Tab',
                 background=[('selected', '#3498db'), ('active', '#74b9ff'), ('!selected', '#bdc3c7')],
                 foreground=[('selected', '#000000'), ('active', '#2c3e50'), ('!selected', '#2c3e50')],
                 relief=[('selected', 'raised'), ('!selected', 'flat')],
                 borderwidth=[('selected', 2), ('!selected', 1)])
        
        # 配置标题框架样式
        style.configure('Title.TFrame', 
                       background='#2c3e50')
        
    def create_interface(self):
        """创建主界面"""
        # 标题栏框架
        title_frame = ttk.Frame(self.root, style='Title.TFrame')
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=0, pady=0)
        
        # 标题
        title_label = ttk.Label(title_frame, text="📊 DataInsight Pro", style='Title.TLabel')
        title_label.pack(fill='x', pady=15)
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重，使界面可以自适应调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # 左侧面板 - 控制按钮
        self.create_control_panel(main_frame)
        
        # 右侧面板 - 数据显示和可视化
        self.create_data_panel(main_frame)
        
        # 状态栏
        self.create_status_bar()
        
    def create_control_panel(self, parent):
        """创建左侧控制面板"""
        control_frame = ttk.LabelFrame(parent, text="🎛️ 控制面板", padding="15")
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        
        # 创建滚动条
        canvas = tk.Canvas(control_frame, width=280, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(control_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定鼠标滚轮事件
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # 使用scrollable_frame作为父容器
        
        # 文件操作区域
        file_frame = ttk.LabelFrame(scrollable_frame, text="📁 文件操作", padding="12")
        file_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Button(file_frame, text="📂 加载CSV/Excel文件", 
                  command=self.load_file, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(file_frame, text="💾 保存清洗后数据", 
                  command=self.save_data, style='Custom.TButton').pack(fill='x', pady=3)
        
        # 数据清洗操作区域
        cleaning_frame = ttk.LabelFrame(scrollable_frame, text="🧹 数据清洗", padding="12")
        cleaning_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Button(cleaning_frame, text="🗑️ 删除重复数据", 
                  command=self.remove_duplicates, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(cleaning_frame, text="🔧 处理缺失值", 
                  command=self.handle_missing_values, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(cleaning_frame, text="🎯 删除异常值", 
                  command=self.remove_outliers, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(cleaning_frame, text="🔄 数据类型转换", 
                  command=self.convert_data_types, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(cleaning_frame, text="⚖️ 数据标准化", 
                  command=self.normalize_data, style='Custom.TButton').pack(fill='x', pady=3)
        
        # 数据分析区域
        analysis_frame = ttk.LabelFrame(scrollable_frame, text="📊 数据分析", padding="12")
        analysis_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Button(analysis_frame, text="📈 基础统计", 
                  command=self.show_statistics, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(analysis_frame, text="🔗 相关性分析", 
                  command=self.correlation_analysis, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(analysis_frame, text="🎯 聚类分析", 
                  command=self.clustering_analysis, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(analysis_frame, text="📋 分组统计", 
                  command=self.groupby_analysis, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(analysis_frame, text="🗂️ 数据透视表", 
                  command=self.pivot_table_analysis, style='Custom.TButton').pack(fill='x', pady=3)
        
        # 数据可视化区域
        viz_frame = ttk.LabelFrame(scrollable_frame, text="📈 数据可视化", padding="12")
        viz_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Button(viz_frame, text="📊 分布图", 
                  command=self.plot_distributions, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(viz_frame, text="🔥 相关性热图", 
                  command=self.plot_correlation, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(viz_frame, text="📋 分组对比图", 
                  command=self.plot_group_comparison, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(viz_frame, text="📈 时间序列图", 
                  command=self.plot_time_series, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(viz_frame, text="🌐 交互式图表", 
                  command=self.create_interactive_plots, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(viz_frame, text="🎨 自定义可视化", 
                  command=self.custom_visualization, style='Custom.TButton').pack(fill='x', pady=3)
        
        # 数据筛选区域
        filter_frame = ttk.LabelFrame(scrollable_frame, text="🔍 数据筛选", padding="12")
        filter_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Button(filter_frame, text="🎯 条件筛选", 
                  command=self.data_filter, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(filter_frame, text="🎲 随机采样", 
                  command=self.random_sample, style='Custom.TButton').pack(fill='x', pady=3)
        ttk.Button(filter_frame, text="📶 排序数据", 
                  command=self.sort_data, style='Custom.TButton').pack(fill='x', pady=3)
        
        # 重置按钮
        ttk.Button(scrollable_frame, text="🔄 恢复原始数据", 
                  command=self.reset_data, style='Custom.TButton').pack(fill='x', pady=(15, 0))
        
        # 功能介绍区域
        intro_frame = ttk.LabelFrame(scrollable_frame, text="💡 功能介绍", padding="12")
        intro_frame.pack(fill='x', pady=(15, 0))
        
        intro_text = """功能说明：
📁 文件操作: 支持CSV、Excel文件导入导出
🧹 数据清洗: 去重、缺失值处理、异常值检测
📊 数据分析: 统计分析、相关性、聚类、分组
📈 数据可视化: 多种图表类型、交互式图表
🔍 数据筛选: 条件筛选、随机采样、排序"""
        
        intro_label = ttk.Label(intro_frame, text=intro_text, font=('微软雅黑', 10), 
                               foreground='#34495e', justify='left')
        intro_label.pack(anchor='w')
        
        # 帮助信息区域
        help_frame = ttk.LabelFrame(scrollable_frame, text="⌨️ 快捷键帮助", padding="12")
        help_frame.pack(fill='x', pady=(10, 0))
        
        help_text = """快捷键说明：
📂 Ctrl+O: 打开文件
💾 Ctrl+S: 保存数据
🔄 F5: 重置数据
🗑️ Ctrl+D: 删除重复数据
🔧 Ctrl+M: 处理缺失值
🎯 Ctrl+R: 删除异常值
📋 Ctrl+1~5: 切换选项卡"""
        
        help_label = ttk.Label(help_frame, text=help_text, font=('微软雅黑', 9), 
                              foreground='#7f8c8d', justify='left')
        help_label.pack(anchor='w')
        
        # 版本信息区域
        self.create_version_info(scrollable_frame)
        
        # 布局滚动条和画布
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def create_data_panel(self, parent):
        """创建右侧数据显示面板"""
        data_frame = ttk.LabelFrame(parent, text="📊 数据视图", padding="15")
        data_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(1, weight=1)
        
        # 信息标签，显示数据基本信息
        info_frame = ttk.LabelFrame(data_frame, text="📊 数据信息", padding="15")
        info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        info_frame.columnconfigure(0, weight=1)
        
        self.info_label = ttk.Label(info_frame, 
                                   text="📝 暂未加载数据\n请点击左侧'打开文件'按钮\n导入CSV或Excel文件", 
                                   font=('微软雅黑', 13), 
                                   foreground='#2c3e50',
                                   justify='left')
        self.info_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 选项卡容器
        self.notebook = ttk.Notebook(data_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建数据视图选项卡
        self.create_data_view_tab()
        
        # 创建可视化选项卡
        self.create_visualization_tab()
        
        # 创建统计信息选项卡
        self.create_statistics_tab()
        
        # 创建分组分析选项卡
        self.create_groupby_tab()
        
        # 创建软件信息选项卡
        self.create_info_tab()
        
    def create_data_view_tab(self):
        """创建数据视图选项卡"""
        data_tab = ttk.Frame(self.notebook)
        self.notebook.add(data_tab, text="📋 数据表格")
        
        # 创建表格显示框架
        tree_frame = ttk.Frame(data_tab)
        tree_frame.pack(fill='both', expand=True)
        
        # 滚动条
        v_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')  # 垂直滚动条
        h_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal')  # 水平滚动条
        
        # 创建树形视图用于显示数据表格
        self.data_tree = ttk.Treeview(tree_frame, 
                                     yscrollcommand=v_scrollbar.set,
                                     xscrollcommand=h_scrollbar.set)
        
        # 设置表格字体
        style = ttk.Style()
        style.configure("Treeview", font=('微软雅黑', 12), rowheight=25)
        style.configure("Treeview.Heading", font=('微软雅黑', 12, 'bold'))
        
        # 配置滚动条
        v_scrollbar.config(command=self.data_tree.yview)
        h_scrollbar.config(command=self.data_tree.xview)
        
        # 布局滚动条和表格
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        self.data_tree.pack(side='left', fill='both', expand=True)
        
    def create_visualization_tab(self):
        """创建可视化选项卡"""
        viz_tab = ttk.Frame(self.notebook)
        self.notebook.add(viz_tab, text="📊 可视化图表")
        self.viz_tab = viz_tab
        
    def create_statistics_tab(self):
        """创建统计信息选项卡"""
        stats_tab = ttk.Frame(self.notebook)
        self.notebook.add(stats_tab, text="📈 统计信息")
        
        # 创建文本控件用于显示统计信息
        stats_frame = ttk.Frame(stats_tab)
        stats_frame.pack(fill='both', expand=True)
        
        # 统计信息的滚动条
        stats_scrollbar = ttk.Scrollbar(stats_frame)
        stats_scrollbar.pack(side='right', fill='y')
        
        # 文本框显示统计信息
        self.stats_text = tk.Text(stats_frame, yscrollcommand=stats_scrollbar.set, 
                                 font=('微软雅黑', 13), bg='#f8f9fa', fg='#2c3e50',
                                 relief='flat', borderwidth=1, wrap='word',
                                 padx=15, pady=10, spacing1=2, spacing3=2)
        self.stats_text.pack(side='left', fill='both', expand=True)
        stats_scrollbar.config(command=self.stats_text.yview)
        
    def create_groupby_tab(self):
        """创建分组分析选项卡"""
        groupby_tab = ttk.Frame(self.notebook)
        self.notebook.add(groupby_tab, text="📋 分组分析")
        
        # 创建文本控件用于显示分组分析结果
        groupby_frame = ttk.Frame(groupby_tab)
        groupby_frame.pack(fill='both', expand=True)
        
        # 分组分析的滚动条
        groupby_scrollbar = ttk.Scrollbar(groupby_frame)
        groupby_scrollbar.pack(side='right', fill='y')
        
        # 文本框显示分组分析结果
        self.groupby_text = tk.Text(groupby_frame, yscrollcommand=groupby_scrollbar.set, 
                                   font=('微软雅黑', 13), bg='#f8f9fa', fg='#2c3e50',
                                   relief='flat', borderwidth=1, wrap='word',
                                   padx=15, pady=10, spacing1=2, spacing3=2)
        self.groupby_text.pack(side='left', fill='both', expand=True)
        groupby_scrollbar.config(command=self.groupby_text.yview)
    
    def _ensure_matplotlib(self):
        if self.plt is None:
            import matplotlib.pyplot as plt_local
            import matplotlib as matplotlib_local
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg as CanvasLocal
            self.plt = plt_local
            self.matplotlib = matplotlib_local
            self.FigureCanvasTkAgg = CanvasLocal
            self.plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
            self.plt.rcParams['axes.unicode_minus'] = False
            self.plt.rcParams['font.size'] = 12
            self.plt.rcParams['figure.facecolor'] = '#f8f9fa'
            self.plt.rcParams['axes.facecolor'] = '#ffffff'
            self.plt.rcParams['axes.edgecolor'] = '#2c3e50'
            self.plt.rcParams['axes.linewidth'] = 1.2
            self.plt.rcParams['grid.color'] = '#bdc3c7'
            self.plt.rcParams['grid.alpha'] = 0.3
            self.plt.rcParams['xtick.labelsize'] = 11
            self.plt.rcParams['ytick.labelsize'] = 11
        if self.canvas is None:
            self.fig, self.ax = self.plt.subplots(figsize=(12, 8))
            self.fig.patch.set_facecolor('#f8f9fa')
            self.canvas = self.FigureCanvasTkAgg(self.fig, self.viz_tab)
            self.canvas.get_tk_widget().pack(fill='both', expand=True, padx=10, pady=10)

    def _ensure_seaborn(self):
        if self.sns is None:
            import seaborn as sns_local
            self.sns = sns_local

    def _ensure_plotly(self):
        if self.px is None:
            import plotly.express as px_local
            self.px = px_local

    def _ensure_sklearn(self):
        from sklearn.preprocessing import StandardScaler as _StandardScaler
        from sklearn.cluster import KMeans as _KMeans
        from sklearn.decomposition import PCA as _PCA
        return _StandardScaler, _KMeans, _PCA

    def create_info_tab(self):
        """创建软件信息选项卡"""
        info_tab = ttk.Frame(self.notebook)
        self.notebook.add(info_tab, text="ℹ️ 软件信息")
        
        # 创建软件信息显示框架
        info_main_frame = ttk.Frame(info_tab)
        info_main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 软件信息的滚动条
        info_canvas = tk.Canvas(info_main_frame, bg='#f8f9fa')
        info_scrollbar = ttk.Scrollbar(info_main_frame, orient="vertical", command=info_canvas.yview)
        info_scrollable_frame = ttk.Frame(info_canvas)
        
        info_scrollable_frame.bind(
            "<Configure>",
            lambda e: info_canvas.configure(scrollregion=info_canvas.bbox("all"))
        )
        
        info_canvas.create_window((0, 0), window=info_scrollable_frame, anchor="nw")
        info_canvas.configure(yscrollcommand=info_scrollbar.set)
        
        # 布局
        info_canvas.pack(side="left", fill="both", expand=True)
        info_scrollbar.pack(side="right", fill="y")
        
        # 创建软件信息内容
        self.create_software_info_content(info_scrollable_frame)
    
    def create_software_info_content(self, parent):
        """创建软件信息内容"""
        import platform
        import sys
        
        # 顶部横幅
        banner_frame = tk.Frame(parent, bg='#3498db', height=120)
        banner_frame.pack(fill='x', pady=(0, 20))
        banner_frame.pack_propagate(False)
        
        # 居中标题容器
        title_container = tk.Frame(banner_frame, bg='#3498db')
        title_container.place(relx=0.5, rely=0.5, anchor='center')
        
        title_label = tk.Label(title_container, text="📊 DataInsight Pro", 
                              font=('微软雅黑', 20, 'bold'), 
                              fg='white', bg='#3498db')
        title_label.pack()
        
        subtitle_label = tk.Label(title_container, text="智能数据分析工具 v1.0", 
                                 font=('微软雅黑', 13), 
                                 fg='#ecf0f1', bg='#3498db')
        subtitle_label.pack(pady=(5, 0))
        
        # 创建左右两列布局
        content_frame = ttk.Frame(parent)
        content_frame.pack(fill='both', expand=True, padx=10)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # 左列内容
        left_column = ttk.Frame(content_frame)
        left_column.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        
        # 系统信息卡片
        system_frame = ttk.LabelFrame(left_column, text="🖥️ 系统环境", padding="15")
        system_frame.pack(fill='x', pady=(0, 15))
        
        system_info = f"""🔹 操作系统: {platform.system()} {platform.release()}
🔹 Python版本: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}
🔹 系统架构: {platform.machine()}
🔹 处理器: {platform.processor() or 'Unknown'}
🔹 作者: Z000oooo
🔹 更新日期: 2025年7月"""
        
        system_label = ttk.Label(system_frame, text=system_info, 
                                font=('微软雅黑', 11), 
                                foreground='#34495e', justify='left')
        system_label.pack(anchor='w')
        
        # 技术架构卡片
        tech_frame = ttk.LabelFrame(left_column, text="🛠️ 技术架构", padding="15")
        tech_frame.pack(fill='x', pady=(0, 15))
        
        tech_text = f"""🔧 开发框架
• GUI: Tkinter + ttk
• 数据处理: Pandas {pd.__version__}
• 可视化: Matplotlib + Seaborn (按需加载)
• 机器学习: Scikit-learn (按需加载)

💡 设计特色
• 现代化扁平UI设计
• 响应式布局
• 多选项卡界面
• 实时状态反馈
• 智能错误处理"""
        
        tech_label = ttk.Label(tech_frame, text=tech_text, 
                              font=('微软雅黑', 11), 
                              foreground='#34495e', justify='left')
        tech_label.pack(anchor='w')
        
        # 右列内容
        right_column = ttk.Frame(content_frame)
        right_column.grid(row=0, column=1, sticky='nsew', padx=(10, 0))
        
        # 核心功能卡片
        features_frame = ttk.LabelFrame(right_column, text="🌟 核心功能", padding="15")
        features_frame.pack(fill='x', pady=(0, 15))
        
        features_text = """📁 文件操作
• CSV、Excel文件支持
• 智能编码识别
• 拖拽加载

🧹 数据清洗
• 重复数据处理
• 缺失值多策略处理
• 异常值检测删除
• 数据类型转换
• 标准化处理

📊 数据分析
• 统计分析
• 相关性分析
• 聚类分析
• 分组统计
• 透视表分析"""
        
        features_label = ttk.Label(features_frame, text=features_text, 
                                  font=('微软雅黑', 11), 
                                  foreground='#2c3e50', justify='left')
        features_label.pack(anchor='w')
        
        # 可视化功能卡片
        viz_frame = ttk.LabelFrame(right_column, text="📈 数据可视化", padding="15")
        viz_frame.pack(fill='x', pady=(0, 15))
        
        viz_text = """📊 图表类型
• 散点图、折线图、柱状图
• 直方图、箱型图、小提琴图
• 相关性热图
• 聚类可视化
• 分组对比图
• 时间序列图

🎨 交互功能
• 自定义图表创建
• 交互式图表(Plotly)
• 图表导出保存"""
        
        viz_label = ttk.Label(viz_frame, text=viz_text, 
                             font=('微软雅黑', 11), 
                             foreground='#2c3e50', justify='left')
        viz_label.pack(anchor='w')
        
        # 底部快速指南 - 跨两列
        bottom_frame = ttk.Frame(content_frame)
        bottom_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(15, 0))
        
        # 快速开始指南
        guide_frame = ttk.LabelFrame(bottom_frame, text="🚀 快速开始", padding="15")
        guide_frame.pack(fill='x', pady=(0, 15))
        
        guide_text = """1️⃣ 点击"打开文件"导入数据  2️⃣ 使用左侧面板清洗数据  3️⃣ 查看右侧分析结果  4️⃣ 保存处理后的数据

⌨️ 常用快捷键: Ctrl+O(打开) | Ctrl+S(保存) | F5(重置) | Ctrl+1~5(切换选项卡)"""
        
        guide_label = ttk.Label(guide_frame, text=guide_text, 
                               font=('微软雅黑', 11), 
                               foreground='#2c3e50', justify='center')
        guide_label.pack()
        
        # 联系信息
        contact_frame = ttk.LabelFrame(bottom_frame, text="📞 开发信息", padding="10")
        contact_frame.pack(fill='x')
        
        contact_text = """👨‍💻 开发者: Z000oooo  |  📅 版本: DataInsight Pro v1.0 (2025.07)  |  🌐 GitHub: github.com/Z000oooo/DataInsight-Pro"""
        
        contact_label = ttk.Label(contact_frame, text=contact_text, 
                                 font=('微软雅黑', 10), 
                                 foreground='#7f8c8d', justify='center')
        contact_label.pack()
        
    def load_file(self):
        """加载CSV或Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            # 根据文件扩展名选择读取方法
            if file_path.endswith('.csv'):
                # 优先使用UTF-8编码，并限制低内存推断提升加载速度
                self.data = pd.read_csv(file_path, encoding='utf-8', low_memory=False)
            elif file_path.endswith(('.xlsx', '.xls')):
                # 指定engine以避免不必要的自动检测
                engine = 'openpyxl' if file_path.endswith('.xlsx') else None
                self.data = pd.read_excel(file_path, engine=engine)
            else:
                messagebox.showerror("错误", "不支持的文件格式！")
                return
                
            # 保存原始数据副本
            self.original_data = self.data.copy()
            # 更新数据显示
            self.update_data_view()
            self.update_info_label()
            messagebox.showinfo("成功", f"文件加载成功！\n数据形状: {self.data.shape}")
            
        except UnicodeDecodeError:
            # 如果UTF-8失败，尝试GBK编码
            try:
                self.data = pd.read_csv(file_path, encoding='gbk', low_memory=False)
                self.original_data = self.data.copy()
                self.update_data_view()
                self.update_info_label()
                messagebox.showinfo("成功", f"文件加载成功！\n数据形状: {self.data.shape}")
                self.update_status(f"文件加载成功: {self.data.shape[0]}行 × {self.data.shape[1]}列", "info")
            except Exception as e:
                messagebox.showerror("错误", f"文件加载失败: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"文件加载失败: {str(e)}")
            
    def update_data_view(self):
        """更新数据表格视图"""
        if self.data is None:
            return
            
        # 清除现有数据
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
            
        # 设置列
        self.data_tree['columns'] = list(self.data.columns)
        self.data_tree['show'] = 'headings'  # 只显示列标题
        
        # 配置列标题
        for col in self.data.columns:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=100, minwidth=50)
            
        # 插入数据（为性能考虑，限制显示前1000行）
        for index, row in self.data.head(1000).itertuples(index=False, name=None):
            self.data_tree.insert('', 'end', values=row)
            
    def update_info_label(self):
        """更新信息标签"""
        if self.data is not None:
            rows, cols = self.data.shape
            missing_count = self.data.isnull().sum().sum()
            numeric_cols = len(self.data.select_dtypes(include=[np.number]).columns)
            categorical_cols = len(self.data.select_dtypes(include=['object']).columns)
            memory_usage = round(self.data.memory_usage(deep=True).sum() / 1024 / 1024, 2)
            
            # 创建美观的多行信息显示
            info_text = (f"📊 数据概览\n"
                        f"📐 维度: {rows:,} 行 × {cols} 列\n"
                        f"🔢 数值列: {numeric_cols} 个  📝 文本列: {categorical_cols} 个\n"
                        f"❓ 缺失值: {missing_count:,} 个  💾 内存: {memory_usage} MB")
            self.info_label.config(text=info_text)
        else:
            self.info_label.config(text="📝 暂未加载数据\n请点击左侧'打开文件'按钮\n导入CSV或Excel文件")
            
    def save_data(self):
        """保存清洗后的数据"""
        if self.data is None:
            messagebox.showwarning("警告", "没有数据可保存！")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="保存清洗后的数据",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx")]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.endswith('.csv'):
                self.data.to_csv(file_path, index=False, encoding='utf-8-sig')  # 使用UTF-8-BOM编码确保中文正常显示
            elif file_path.endswith('.xlsx'):
                self.data.to_excel(file_path, index=False)
                
            messagebox.showinfo("成功", "数据保存成功！")
            
        except Exception as e:
            messagebox.showerror("错误", f"文件保存失败: {str(e)}")
            
    def remove_duplicates(self):
        """删除重复行"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        original_shape = self.data.shape
        self.data = self.data.drop_duplicates()
        new_shape = self.data.shape
        
        removed = original_shape[0] - new_shape[0]
        self.update_data_view()
        self.update_info_label()
        
        messagebox.showinfo("成功", f"已删除 {removed} 行重复数据！")
        
    def handle_missing_values(self):
        """处理缺失值"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建缺失值处理选项对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("处理缺失值")
        dialog.geometry("450x400")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"450x400+{x}+{y}")
        
        ttk.Label(dialog, text="选择缺失值处理方式：", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        method_var = tk.StringVar(value="drop")
        
        ttk.Radiobutton(dialog, text="删除包含缺失值的行", 
                       variable=method_var, value="drop").pack(anchor='w', padx=20, pady=5)
        ttk.Radiobutton(dialog, text="用均值填充（数值列）", 
                       variable=method_var, value="mean").pack(anchor='w', padx=20, pady=5)
        ttk.Radiobutton(dialog, text="用中位数填充（数值列）", 
                       variable=method_var, value="median").pack(anchor='w', padx=20, pady=5)
        ttk.Radiobutton(dialog, text="用众数填充（所有列）", 
                       variable=method_var, value="mode").pack(anchor='w', padx=20, pady=5)
        ttk.Radiobutton(dialog, text="前向填充", 
                       variable=method_var, value="ffill").pack(anchor='w', padx=20, pady=5)
        ttk.Radiobutton(dialog, text="后向填充", 
                       variable=method_var, value="bfill").pack(anchor='w', padx=20, pady=5)
        
        def apply_method():
            method = method_var.get()
            original_missing = self.data.isnull().sum().sum()
            
            if method == "drop":
                self.data = self.data.dropna()
            elif method == "mean":
                numeric_cols = self.data.select_dtypes(include=[np.number]).columns
                self.data[numeric_cols] = self.data[numeric_cols].fillna(self.data[numeric_cols].mean())
            elif method == "median":
                numeric_cols = self.data.select_dtypes(include=[np.number]).columns
                self.data[numeric_cols] = self.data[numeric_cols].fillna(self.data[numeric_cols].median())
            elif method == "mode":
                for col in self.data.columns:
                    self.data[col] = self.data[col].fillna(self.data[col].mode().iloc[0] if not self.data[col].mode().empty else 0)
            elif method == "ffill":
                self.data = self.data.ffill()  # 使用新的方法
            elif method == "bfill":
                self.data = self.data.bfill()  # 使用新的方法
                
            new_missing = self.data.isnull().sum().sum()
            handled = original_missing - new_missing
            
            self.update_data_view()
            self.update_info_label()
            dialog.destroy()
            
            messagebox.showinfo("成功", f"已处理 {handled} 个缺失值！")
            
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="✅ 应用", command=apply_method, 
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy, 
                  style='Custom.TButton').pack(side='right')
        
    def remove_outliers(self):
        """使用IQR方法删除异常值"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            messagebox.showwarning("警告", "没有找到数值列！")
            return
            
        original_shape = self.data.shape
        
        for col in numeric_cols:
            Q1 = self.data[col].quantile(0.25)
            Q3 = self.data[col].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR
            
            self.data = self.data[(self.data[col] >= lower_bound) & (self.data[col] <= upper_bound)]
            
        new_shape = self.data.shape
        removed = original_shape[0] - new_shape[0]
        
        self.update_data_view()
        self.update_info_label()
        
        messagebox.showinfo("成功", f"已删除 {removed} 行异常值数据！")
        
    def convert_data_types(self):
        """转换数据类型"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建数据类型转换对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("数据类型转换")
        dialog.geometry("550x500")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (550 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"550x500+{x}+{y}")
        
        ttk.Label(dialog, text="选择列和目标数据类型：", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 列选择框架
        columns_frame = ttk.Frame(dialog)
        columns_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 可滚动框架
        canvas = tk.Canvas(columns_frame)
        scrollbar = ttk.Scrollbar(columns_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 列类型选择控件
        type_vars = {}
        for i, col in enumerate(self.data.columns):
            frame = ttk.Frame(scrollable_frame)
            frame.grid(row=i, column=0, sticky='ew', pady=2)
            
            ttk.Label(frame, text=f"{col} (当前类型: {self.data[col].dtype})", width=30).pack(side='left')
            
            type_var = tk.StringVar(value=str(self.data[col].dtype))
            type_vars[col] = type_var
            
            type_combo = ttk.Combobox(frame, textvariable=type_var, 
                                    values=['int64', 'float64', 'object', 'datetime64', 'bool'], width=15)
            type_combo.pack(side='right')
            
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def apply_conversions():
            try:
                for col, type_var in type_vars.items():
                    target_type = type_var.get()
                    if target_type != str(self.data[col].dtype):
                        if target_type == 'datetime64':
                            self.data[col] = pd.to_datetime(self.data[col], errors='coerce')
                        else:
                            self.data[col] = self.data[col].astype(target_type)
                            
                self.update_data_view()
                self.update_info_label()
                dialog.destroy()
                messagebox.showinfo("成功", "数据类型转换成功！")
                
            except Exception as e:
                messagebox.showerror("错误", f"转换失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="✅ 应用转换", command=apply_conversions,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def normalize_data(self):
        """标准化数值数据"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            messagebox.showwarning("警告", "没有找到数值列！")
            return
            
        # 按需导入标准化器
        StandardScaler, _, _ = self._ensure_sklearn()
        scaler = StandardScaler()
        self.data[numeric_cols] = scaler.fit_transform(self.data[numeric_cols])
        
        self.update_data_view()
        self.update_info_label()
        
        messagebox.showinfo("成功", "数值数据标准化成功！")
        
    def show_statistics(self):
        """显示基础统计信息"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 清除统计文本
        self.stats_text.delete(1.0, tk.END)
        
        # 基本信息
        info_str = f"📊 数据集概览\n"
        info_str += f"{'=' * 50}\n"
        info_str += f"📐 数据维度: {self.data.shape[0]:,} 行 × {self.data.shape[1]} 列\n"
        info_str += f"💾 内存占用: {self.data.memory_usage(deep=True).sum() / 1024**2:.2f} MB\n"
        info_str += f"📅 创建时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        # 数据类型统计
        dtype_counts = self.data.dtypes.value_counts()
        info_str += "🔢 数据类型分布\n"
        info_str += f"{'=' * 30}\n"
        for dtype, count in dtype_counts.items():
            if dtype == 'object':
                dtype_name = '📝 文本类型'
            elif 'int' in str(dtype):
                dtype_name = '🔢 整数类型'
            elif 'float' in str(dtype):
                dtype_name = '🔢 浮点类型'
            elif 'datetime' in str(dtype):
                dtype_name = '📅 日期类型'
            else:
                dtype_name = f'❓ {dtype}'
            info_str += f"{dtype_name}: {count} 列\n"
        info_str += "\n"
        
        # 缺失值分析
        missing_info = self.data.isnull().sum()
        missing_info = missing_info[missing_info > 0]
        info_str += "❓ 缺失值分析\n"
        info_str += f"{'=' * 30}\n"
        if len(missing_info) > 0:
            for col, missing_count in missing_info.items():
                missing_pct = (missing_count / len(self.data)) * 100
                info_str += f"🔸 {col}: {missing_count:,} 个 ({missing_pct:.1f}%)\n"
            info_str += "\n"
        else:
            info_str += "✅ 无缺失值\n\n"
            
        # 数值列统计
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            info_str += "📈 数值列统计摘要\n"
            info_str += f"{'=' * 50}\n"
            describe_df = self.data[numeric_cols].describe()
            info_str += describe_df.round(2).to_string() + "\n\n"
        
        # 分类列的唯一值
        categorical_cols = self.data.select_dtypes(include=['object']).columns
        if len(categorical_cols) > 0:
            info_str += "📝 文本列唯一值统计\n"
            info_str += f"{'=' * 40}\n"
            for col in categorical_cols:
                unique_count = self.data[col].nunique()
                total_count = len(self.data[col].dropna())
                unique_pct = (unique_count / total_count * 100) if total_count > 0 else 0
                info_str += f"🔸 {col}: {unique_count:,} 个唯一值 ({unique_pct:.1f}%)\n"
                
                # 显示前5个最常见的值
                if unique_count > 0:
                    top_values = self.data[col].value_counts().head(3)
                    info_str += f"   📊 最常见值: {', '.join([f'{v}({c})' for v, c in top_values.items()])}\n"
            info_str += "\n"
        
        # 数据质量评估
        info_str += "🔍 数据质量评估\n"
        info_str += f"{'=' * 40}\n"
        total_cells = self.data.shape[0] * self.data.shape[1]
        missing_cells = self.data.isnull().sum().sum()
        completeness = ((total_cells - missing_cells) / total_cells) * 100
        info_str += f"📊 数据完整性: {completeness:.1f}%\n"
        info_str += f"🔢 重复行数量: {self.data.duplicated().sum():,} 行\n"
        
        # 推荐操作
        info_str += "\n💡 数据清洗建议\n"
        info_str += f"{'=' * 40}\n"
        if missing_cells > 0:
            info_str += "🔧 建议处理缺失值\n"
        if self.data.duplicated().sum() > 0:
            info_str += "🗑️ 建议删除重复数据\n"
        if completeness > 95:
            info_str += "✅ 数据质量良好\n"
        elif completeness > 80:
            info_str += "⚠️ 数据质量中等，建议清洗\n"
        else:
            info_str += "❌ 数据质量较差，需要重点清洗\n"
                
        self.stats_text.insert(tk.END, info_str)
        self.notebook.select(2)  # 切换到统计信息选项卡
        
    def correlation_analysis(self):
        """执行相关性分析"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) < 2:
            messagebox.showwarning("警告", "相关性分析至少需要2个数值列！")
            return
            
        # 计算相关性矩阵
        corr_matrix = self.data[numeric_cols].corr()
        
        # 确保可视化环境
        self._ensure_matplotlib()
        self._ensure_seaborn()
        
        # 清除之前的图表
        self.fig.clear()
        self.fig.patch.set_facecolor('#f8f9fa')
        self.ax = self.fig.add_subplot(111)
        
        # 创建热图
        self.sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, 
                   square=True, ax=self.ax, cbar_kws={'shrink': 0.8}, 
                   fmt='.2f', annot_kws={'size': 10})
        self.ax.set_title('相关性矩阵', fontsize=14, fontweight='bold', pad=20)
        
        # 优化标签显示
        self.ax.tick_params(axis='x', rotation=45, labelsize=10)
        self.ax.tick_params(axis='y', rotation=0, labelsize=10)
        
        self.fig.tight_layout()
        self.canvas.draw()
        self.notebook.select(1)  # 切换到可视化选项卡
        
    def clustering_analysis(self):
        """执行聚类分析"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) < 2:
            messagebox.showwarning("警告", "聚类分析至少需要2个数值列！")
            return
            
        # 准备数据
        cluster_data = self.data[numeric_cols].dropna()
        
        if len(cluster_data) < 10:
            messagebox.showwarning("警告", "聚类分析至少需要10行数据！")
            return
            
        # 标准化与聚类所需库
        StandardScaler, KMeans, PCA = self._ensure_sklearn()
        
        # 标准化数据
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(cluster_data)
        
        # 执行K均值聚类
        kmeans = KMeans(n_clusters=3, random_state=42)
        clusters = kmeans.fit_predict(scaled_data)
        
        # PCA降维用于可视化
        pca = PCA(n_components=2)
        pca_data = pca.fit_transform(scaled_data)
        
        # 确保可视化环境
        self._ensure_matplotlib()
        
        # 清除之前的图表
        self.fig.clear()
        self.fig.patch.set_facecolor('#f8f9fa')
        self.ax = self.fig.add_subplot(111)
        
        # 创建散点图
        scatter = self.ax.scatter(pca_data[:, 0], pca_data[:, 1], c=clusters, 
                                 cmap='viridis', alpha=0.8, s=60, edgecolors='white', linewidth=0.5)
        self.ax.set_xlabel(f'主成分1 ({pca.explained_variance_ratio_[0]:.2%} 方差)', fontsize=12)
        self.ax.set_ylabel(f'主成分2 ({pca.explained_variance_ratio_[1]:.2%} 方差)', fontsize=12)
        self.ax.set_title('K均值聚类 (PCA可视化)', fontsize=14, fontweight='bold', pad=20)
        self.ax.grid(True, alpha=0.3)
        self.ax.tick_params(labelsize=10)
        
        # 添加颜色条
        self.plt.colorbar(scatter, ax=self.ax)
        
        self.fig.tight_layout()
        self.canvas.draw()
        self.notebook.select(1)  # 切换到可视化选项卡
        
    def plot_distributions(self):
        """绘制数值列的分布图"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            messagebox.showwarning("警告", "没有找到数值列！")
            return
            
        # 确保可视化环境
        self._ensure_matplotlib()
        
        # 清除之前的图表
        self.fig.clear()
        self.fig.patch.set_facecolor('#f8f9fa')
        
        # 创建子图
        n_cols = min(3, len(numeric_cols))
        n_rows = (len(numeric_cols) + n_cols - 1) // n_cols
        
        for i, col in enumerate(numeric_cols[:9]):  # 限制为9个图表
            ax = self.fig.add_subplot(n_rows, n_cols, i + 1)
            self.data[col].hist(bins=25, ax=ax, alpha=0.8, color='steelblue', 
                               edgecolor='white', linewidth=0.5)
            ax.set_title(f'{col} 分布图', fontsize=11, fontweight='bold', pad=10)
            ax.set_xlabel(col, fontsize=10)
            ax.set_ylabel('频率', fontsize=10)
            ax.grid(True, alpha=0.3)
            ax.tick_params(labelsize=9)
            
        # 优化布局
        self.fig.tight_layout(pad=2.0)
        self.canvas.draw()
        self.notebook.select(1)  # 切换到可视化选项卡
        
    def plot_correlation(self):
        """绘制相关性热图"""
        self.correlation_analysis()
        
    def create_interactive_plots(self):
        """创建交互式图表"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) < 2:
            messagebox.showwarning("警告", "交互式图表至少需要2个数值列！")
            return
            
        # 创建交互式散点图矩阵（按需导入plotly）
        self._ensure_plotly()
        fig = self.px.scatter_matrix(self.data[numeric_cols], title="交互式散点图矩阵")
        
        # 保存到临时HTML文件并打开
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.html')
        fig.write_html(temp_file.name)
        webbrowser.open(f'file://{temp_file.name}')
        
        messagebox.showinfo("信息", "交互式图表已在浏览器中打开！")
        
    def custom_visualization(self):
        """创建自定义可视化对话框"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建自定义可视化对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("🎨 自定义可视化")
        dialog.geometry("450x400")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示对话框
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"450x400+{x}+{y}")
        
        # 标题框架
        title_frame = ttk.Frame(dialog)
        title_frame.pack(fill='x', pady=(0, 20))
        
        ttk.Label(title_frame, text="🎨 创建自定义可视化", 
                 font=('微软雅黑', 14, 'bold'), 
                 foreground='#2c3e50').pack(pady=15)
        
        # 主要内容框架
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill='both', expand=True, padx=20)
        
        # 图表类型选择
        ttk.Label(main_frame, text="📊 图表类型:", 
                 font=('微软雅黑', 11, 'bold')).pack(anchor='w', pady=(10, 5))
        plot_type_var = tk.StringVar(value="散点图")
        # 创建图表类型映射字典
        plot_type_mapping = {
            '📊 散点图': 'scatter',
            '📈 折线图': 'line', 
            '📊 柱状图': 'bar',
            '📊 直方图': 'histogram',
            '📦 箱型图': 'box',
            '🎻 小提琴图': 'violin'
        }
        plot_type_combo = ttk.Combobox(main_frame, textvariable=plot_type_var,
                                      values=list(plot_type_mapping.keys()),
                                      font=('微软雅黑', 10))
        plot_type_combo.pack(fill='x', pady=(0, 15))
        plot_type_combo.set('📊 散点图')  # 设置默认值
        
        # X轴选择
        ttk.Label(main_frame, text="📐 X轴数据:", 
                 font=('微软雅黑', 11, 'bold')).pack(anchor='w', pady=(5, 5))
        x_var = tk.StringVar()
        x_combo = ttk.Combobox(main_frame, textvariable=x_var, 
                              values=list(self.data.columns),
                              font=('微软雅黑', 10))
        x_combo.pack(fill='x', pady=(0, 15))
        
        # Y轴选择
        ttk.Label(main_frame, text="📏 Y轴数据 (可选):", 
                 font=('微软雅黑', 11, 'bold')).pack(anchor='w', pady=(5, 5))
        y_var = tk.StringVar()
        y_combo = ttk.Combobox(main_frame, textvariable=y_var, 
                              values=[''] + list(self.data.columns),
                              font=('微软雅黑', 10))
        y_combo.pack(fill='x', pady=(0, 20))
        
        def create_plot():
            plot_type_chinese = plot_type_var.get()
            plot_type = plot_type_mapping.get(plot_type_chinese, 'scatter')  # 获取英文类型
            # 去掉emoji以便于显示
            clean_plot_type = plot_type_chinese.split(' ', 1)[-1] if ' ' in plot_type_chinese else plot_type_chinese
            x_col = x_var.get()
            y_col = y_var.get()
            
            if not x_col:
                messagebox.showwarning("警告", "请选择X轴列！")
                return
                
            try:
                # 确保可视化环境
                self._ensure_matplotlib()
                # 清除之前的图表并重新设置
                self.fig.clear()
                self.fig.patch.set_facecolor('#f8f9fa')
                self.ax = self.fig.add_subplot(111)
                
                # 设置子图的边距，确保有足够空间显示标签
                self.fig.subplots_adjust(left=0.12, bottom=0.15, right=0.95, top=0.9)
                if plot_type == "scatter" and y_col:
                    self.ax.scatter(self.data[x_col], self.data[y_col], alpha=0.7, 
                                   color='steelblue', s=60, edgecolors='white', linewidth=0.5)
                    self.ax.set_ylabel(y_col, fontsize=12)
                    self.ax.set_xlabel(x_col, fontsize=12)
                elif plot_type == "line" and y_col:
                    # 对于折线图，如果x轴是文本类型，先进行排序
                    if self.data[x_col].dtype == 'object':
                        plot_data = self.data.sort_values(x_col)
                        self.ax.plot(range(len(plot_data)), plot_data[y_col], 
                                    marker='o', linewidth=2.5, markersize=6, color='#3498db')
                        self.ax.set_xticks(range(len(plot_data)))
                        self.ax.set_xticklabels(plot_data[x_col], rotation=45, ha='right')
                    else:
                        sorted_data = self.data.sort_values(x_col)
                        self.ax.plot(sorted_data[x_col], sorted_data[y_col], 
                                    marker='o', linewidth=2.5, markersize=6, color='#3498db')
                    self.ax.set_ylabel(y_col, fontsize=12)
                    self.ax.set_xlabel(x_col, fontsize=12)
                elif plot_type == "bar":
                    if y_col:
                        grouped_data = self.data.groupby(x_col)[y_col].mean()
                        bars = grouped_data.plot(kind='bar', ax=self.ax, color='#e74c3c', 
                                                width=0.7, edgecolor='white', linewidth=1)
                        self.ax.set_ylabel(f'{y_col} (平均值)', fontsize=12)
                        
                        # 在柱子上添加数值标签
                        for i, bar in enumerate(self.ax.patches):
                            height = bar.get_height()
                            self.ax.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                                       f'{height:.2f}', ha='center', va='bottom', fontsize=9)
                    else:
                        value_counts = self.data[x_col].value_counts()
                        bars = value_counts.plot(kind='bar', ax=self.ax, color='#27ae60',
                                               width=0.7, edgecolor='white', linewidth=1)
                        self.ax.set_ylabel('计数', fontsize=12)
                        
                        # 在柱子上添加数值标签
                        for i, bar in enumerate(self.ax.patches):
                            height = bar.get_height()
                            self.ax.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                                       f'{int(height)}', ha='center', va='bottom', fontsize=9)
                    
                    self.ax.set_xlabel(x_col, fontsize=12)
                    self.ax.tick_params(axis='x', rotation=45, labelsize=10)
                    self.ax.tick_params(axis='y', labelsize=10)
                elif plot_type == "histogram":
                    # 确保列是数值类型
                    if pd.api.types.is_numeric_dtype(self.data[x_col]):
                        n, bins, patches = self.ax.hist(self.data[x_col].dropna(), bins=30, 
                                                       alpha=0.8, color='#f39c12', 
                                                       edgecolor='white', linewidth=1)
                        self.ax.set_ylabel('频率', fontsize=12)
                        self.ax.set_xlabel(x_col, fontsize=12)
                        
                        # 添加统计信息
                        mean_val = self.data[x_col].mean()
                        self.ax.axvline(mean_val, color='red', linestyle='--', linewidth=2, 
                                       label=f'平均值: {mean_val:.2f}')
                        self.ax.legend()
                    else:
                        messagebox.showwarning("警告", "直方图需要数值类型的列！")
                        return
                elif plot_type == "box":
                    if y_col:
                        # 分组箱型图
                        groups = [group[y_col].dropna() for name, group in self.data.groupby(x_col)]
                        group_names = [str(name) for name, group in self.data.groupby(x_col)]
                        
                        box_plot = self.ax.boxplot(groups, labels=group_names, patch_artist=True)
                        # 设置颜色
                        colors = ['lightblue', 'lightgreen', 'lightcoral', 'lightyellow', 'lightpink']
                        for patch, color in zip(box_plot['boxes'], colors * len(box_plot['boxes'])):
                            patch.set_facecolor(color)
                        self.ax.set_ylabel(y_col)
                        self.ax.tick_params(axis='x', rotation=45)
                    else:
                        if pd.api.types.is_numeric_dtype(self.data[x_col]):
                            box_plot = self.ax.boxplot(self.data[x_col].dropna(), patch_artist=True)
                            box_plot['boxes'][0].set_facecolor('lightblue')
                        else:
                            messagebox.showwarning("警告", "箱型图需要数值类型的列！")
                            return
                elif plot_type == "violin" and y_col:
                    # 为小提琴图准备分组数据
                    try:
                        groups = [group[y_col].dropna() for name, group in self.data.groupby(x_col)]
                        group_names = [str(name) for name, group in self.data.groupby(x_col)]
                        
                        # 过滤掉空组
                        valid_groups = [group for group in groups if len(group) > 1]
                        valid_names = [name for name, group in zip(group_names, groups) if len(group) > 1]
                        
                        if valid_groups and len(valid_groups) > 0:
                            positions = range(1, len(valid_groups) + 1)
                            violin_parts = self.ax.violinplot(valid_groups, positions=positions)
                            
                            # 设置小提琴图颜色
                            for pc in violin_parts['bodies']:
                                pc.set_facecolor('lightsteelblue')
                                pc.set_alpha(0.7)
                                
                            self.ax.set_xticks(positions)
                            self.ax.set_xticklabels(valid_names, rotation=45)
                            self.ax.set_ylabel(y_col)
                        else:
                            messagebox.showwarning("警告", "没有足够的数据创建小提琴图！每组至少需要2个数据点。")
                            return
                    except Exception as violin_error:
                        messagebox.showerror("错误", f"小提琴图创建失败: {str(violin_error)}")
                        return
                
                # 设置轴标签（如果还没有设置的话）
                if not self.ax.get_xlabel():
                    self.ax.set_xlabel(x_col, fontsize=12)
                if not self.ax.get_ylabel() and y_col:
                    self.ax.set_ylabel(y_col, fontsize=12)
                
                # 使用中文图表类型名称作为标题
                title = f'{clean_plot_type}: {x_col}'
                if y_col and plot_type != "histogram":
                    title += f' vs {y_col}'
                self.ax.set_title(title, fontsize=14, fontweight='bold')
                
                # 设置网格
                self.ax.grid(True, alpha=0.3)
                
                # 最终设置和显示
                # 根据图表类型调整布局
                if plot_type in ["bar", "box", "violin"]:
                    # 对于可能有长标签的图表，增加底部空间
                    self.fig.subplots_adjust(left=0.12, bottom=0.25, right=0.95, top=0.85)
                else:
                    self.fig.subplots_adjust(left=0.12, bottom=0.15, right=0.95, top=0.9)
                
                # 优化布局
                try:
                    self.fig.tight_layout(pad=2.0)
                except:
                    pass  # 如果tight_layout失败，使用手动调整的布局
                
                self.canvas.draw()
                self.notebook.select(1)  # 切换到可视化选项卡
                dialog.destroy()
                
                # 更新状态栏
                self.update_status(f"自定义图表创建成功: {clean_plot_type}", "info")
                
            except Exception as e:
                messagebox.showerror("错误", f"创建图表失败: {str(e)}\n请检查所选列的数据类型是否适合该图表类型。")
                self.update_status("图表创建失败", "error")
                print(f"详细错误信息: {e}")  # 调试用
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        # 创建按钮
        create_btn = ttk.Button(button_frame, text="🎨 创建图表", 
                               command=create_plot, style='Custom.TButton')
        create_btn.pack(side='right', padx=(10, 0))
        
        # 取消按钮
        cancel_btn = ttk.Button(button_frame, text="❌ 取消", 
                               command=dialog.destroy, style='Custom.TButton')
        cancel_btn.pack(side='right')
        
    def reset_data(self):
        """重置数据到原始状态"""
        if self.original_data is None:
            messagebox.showwarning("警告", "没有原始数据可以重置！")
            return
            
        self.data = self.original_data.copy()
        self.update_data_view()
        self.update_info_label()
        
        messagebox.showinfo("成功", "数据已重置到原始状态！")
        
    def groupby_analysis(self):
        """分组统计分析"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建分组分析对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("分组统计分析")
        dialog.geometry("550x450")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (550 // 2)
        y = (dialog.winfo_screenheight() // 2) - (450 // 2)
        dialog.geometry(f"550x450+{x}+{y}")
        
        ttk.Label(dialog, text="分组统计设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 分组列选择
        ttk.Label(dialog, text="选择分组列：").pack(anchor='w', padx=20, pady=(10, 0))
        group_var = tk.StringVar()
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=list(self.data.columns))
        group_combo.pack(fill='x', padx=20, pady=5)
        
        # 聚合列选择
        ttk.Label(dialog, text="选择聚合列：").pack(anchor='w', padx=20, pady=(10, 0))
        agg_var = tk.StringVar()
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
        agg_combo = ttk.Combobox(dialog, textvariable=agg_var, values=numeric_cols)
        agg_combo.pack(fill='x', padx=20, pady=5)
        
        # 聚合函数选择
        ttk.Label(dialog, text="选择聚合函数：").pack(anchor='w', padx=20, pady=(10, 0))
        func_var = tk.StringVar(value="均值")
        func_mapping = {
            '均值': 'mean',
            '中位数': 'median',
            '求和': 'sum',
            '计数': 'count',
            '最大值': 'max',
            '最小值': 'min',
            '标准差': 'std',
            '方差': 'var'
        }
        func_combo = ttk.Combobox(dialog, textvariable=func_var, values=list(func_mapping.keys()))
        func_combo.pack(fill='x', padx=20, pady=5)
        
        def perform_groupby():
            group_col = group_var.get()
            agg_col = agg_var.get()
            func_name = func_var.get()
            func = func_mapping.get(func_name, 'mean')
            
            if not group_col:
                messagebox.showwarning("警告", "请选择分组列！")
                return
            if not agg_col:
                messagebox.showwarning("警告", "请选择聚合列！")
                return
            
            try:
                # 执行分组统计
                if func == 'count':
                    result = self.data.groupby(group_col).size().reset_index(name='计数')
                else:
                    result = self.data.groupby(group_col)[agg_col].agg(func).reset_index()
                    result.columns = [group_col, f'{agg_col}_{func_name}']
                
                # 显示结果
                self.groupby_text.delete(1.0, tk.END)
                result_str = f"📊 分组统计分析报告\n"
                result_str += f"{'=' * 60}\n"
                result_str += f"🔍 分组列: {group_col}\n"
                result_str += f"📈 聚合列: {agg_col}\n" if agg_col else ""
                result_str += f"⚙️ 聚合函数: {func_name}\n"
                result_str += f"📅 分析时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                result_str += f"{'=' * 60}\n\n"
                
                # 格式化结果显示
                result_str += f"📋 分组统计结果\n"
                result_str += f"{'-' * 40}\n"
                for i, row in result.iterrows():
                    if func == 'count':
                        result_str += f"🔸 {row[group_col]}: {row['计数']:,} 项\n"
                    else:
                        value = row.iloc[-1]  # 最后一列是聚合结果
                        if isinstance(value, (int, float)):
                            result_str += f"🔸 {row[group_col]}: {value:,.2f}\n"
                        else:
                            result_str += f"🔸 {row[group_col]}: {value}\n"
                result_str += "\n"
                
                # 添加统计摘要
                if len(result) > 1 and func != 'count':
                    values = result.iloc[:, -1]
                    if pd.api.types.is_numeric_dtype(values):
                        result_str += f"📈 统计摘要\n"
                        result_str += f"{'-' * 30}\n"
                        result_str += f"📊 最大值: {values.max():.2f}\n"
                        result_str += f"📊 最小值: {values.min():.2f}\n"
                        result_str += f"📊 平均值: {values.mean():.2f}\n"
                        result_str += f"📊 中位数: {values.median():.2f}\n"
                        result_str += f"📊 标准差: {values.std():.2f}\n"
                
                self.groupby_text.insert(tk.END, result_str)
                self.notebook.select(3)  # 切换到分组分析选项卡
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("错误", f"分组统计失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="📊 执行分组统计", command=perform_groupby,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def pivot_table_analysis(self):
        """数据透视表分析"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建透视表对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("数据透视表")
        dialog.geometry("650x550")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (650 // 2)
        y = (dialog.winfo_screenheight() // 2) - (550 // 2)
        dialog.geometry(f"650x550+{x}+{y}")
        
        ttk.Label(dialog, text="透视表设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 行索引选择
        ttk.Label(dialog, text="选择行索引：").pack(anchor='w', padx=20, pady=(10, 0))
        index_var = tk.StringVar()
        index_combo = ttk.Combobox(dialog, textvariable=index_var, values=list(self.data.columns))
        index_combo.pack(fill='x', padx=20, pady=5)
        
        # 列索引选择
        ttk.Label(dialog, text="选择列索引（可选）：").pack(anchor='w', padx=20, pady=(10, 0))
        columns_var = tk.StringVar()
        columns_combo = ttk.Combobox(dialog, textvariable=columns_var, values=list(self.data.columns))
        columns_combo.pack(fill='x', padx=20, pady=5)
        
        # 值选择
        ttk.Label(dialog, text="选择值列：").pack(anchor='w', padx=20, pady=(10, 0))
        values_var = tk.StringVar()
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
        values_combo = ttk.Combobox(dialog, textvariable=values_var, values=numeric_cols)
        values_combo.pack(fill='x', padx=20, pady=5)
        
        # 聚合函数选择
        ttk.Label(dialog, text="选择聚合函数：").pack(anchor='w', padx=20, pady=(10, 0))
        aggfunc_var = tk.StringVar(value="均值")
        aggfunc_mapping = {
            '均值': 'mean',
            '求和': 'sum',
            '计数': 'count',
            '最大值': 'max',
            '最小值': 'min'
        }
        aggfunc_combo = ttk.Combobox(dialog, textvariable=aggfunc_var, values=list(aggfunc_mapping.keys()))
        aggfunc_combo.pack(fill='x', padx=20, pady=5)
        
        def create_pivot():
            index_col = index_var.get()
            columns_col = columns_var.get() if columns_var.get() else None
            values_col = values_var.get()
            aggfunc_name = aggfunc_var.get()
            aggfunc = aggfunc_mapping.get(aggfunc_name, 'mean')
            
            if not index_col:
                messagebox.showwarning("警告", "请选择行索引！")
                return
            if not values_col:
                messagebox.showwarning("警告", "请选择值列！")
                return
            
            try:
                # 创建透视表
                if columns_col:
                    pivot_result = pd.pivot_table(self.data, 
                                                values=values_col,
                                                index=index_col,
                                                columns=columns_col,
                                                aggfunc=aggfunc,
                                                fill_value=0)
                else:
                    pivot_result = self.data.groupby(index_col)[values_col].agg(aggfunc).reset_index()
                
                # 显示结果
                self.groupby_text.delete(1.0, tk.END)
                result_str = f"📊 数据透视表分析报告\n"
                result_str += f"{'=' * 60}\n"
                result_str += f"📋 行索引: {index_col}\n"
                result_str += f"📋 列索引: {columns_col}\n" if columns_col else "📋 列索引: 无\n"
                result_str += f"📈 值列: {values_col}\n"
                result_str += f"⚙️ 聚合函数: {aggfunc_name}\n"
                result_str += f"📅 创建时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                result_str += f"{'=' * 60}\n\n"
                
                result_str += f"📋 透视表结果\n"
                result_str += f"{'-' * 40}\n"
                result_str += str(pivot_result) + "\n\n"
                
                # 添加透视表统计摘要
                if hasattr(pivot_result, 'values') and len(pivot_result) > 0:
                    if isinstance(pivot_result, pd.DataFrame):
                        numeric_data = pivot_result.select_dtypes(include=[np.number])
                        if not numeric_data.empty:
                            result_str += f"📈 数据摘要\n"
                            result_str += f"{'-' * 30}\n"
                            result_str += f"📊 数据形状: {pivot_result.shape[0]} 行 × {pivot_result.shape[1]} 列\n"
                            result_str += f"📊 总计: {numeric_data.sum().sum():.2f}\n"
                            result_str += f"📊 平均值: {numeric_data.mean().mean():.2f}\n"
                
                self.groupby_text.insert(tk.END, result_str)
                self.notebook.select(3)  # 切换到分组分析选项卡
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("错误", f"透视表创建失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="📊 创建透视表", command=create_pivot,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def data_filter(self):
        """数据条件筛选"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建筛选对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("数据筛选")
        dialog.geometry("550x450")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (550 // 2)
        y = (dialog.winfo_screenheight() // 2) - (450 // 2)
        dialog.geometry(f"550x450+{x}+{y}")
        
        ttk.Label(dialog, text="数据筛选设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 筛选列选择
        ttk.Label(dialog, text="选择筛选列：").pack(anchor='w', padx=20, pady=(10, 0))
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=list(self.data.columns))
        column_combo.pack(fill='x', padx=20, pady=5)
        
        # 筛选条件选择
        ttk.Label(dialog, text="筛选条件：").pack(anchor='w', padx=20, pady=(10, 0))
        condition_var = tk.StringVar(value="等于")
        condition_mapping = {
            '等于': '==',
            '不等于': '!=',
            '大于': '>',
            '小于': '<',
            '大于等于': '>=',
            '小于等于': '<=',
            '包含': 'contains',
            '不包含': 'not_contains'
        }
        condition_combo = ttk.Combobox(dialog, textvariable=condition_var, values=list(condition_mapping.keys()))
        condition_combo.pack(fill='x', padx=20, pady=5)
        
        # 筛选值输入
        ttk.Label(dialog, text="筛选值：").pack(anchor='w', padx=20, pady=(10, 0))
        value_var = tk.StringVar()
        value_entry = ttk.Entry(dialog, textvariable=value_var)
        value_entry.pack(fill='x', padx=20, pady=5)
        
        def apply_filter():
            col = column_var.get()
            condition_name = condition_var.get()
            condition = condition_mapping.get(condition_name)
            value = value_var.get()
            
            if not col or not value:
                messagebox.showwarning("警告", "请填写完整的筛选条件！")
                return
                
            try:
                original_rows = len(self.data)
                
                if condition in ['contains']:
                    self.data = self.data[self.data[col].astype(str).str.contains(value, na=False)]
                elif condition == 'not_contains':
                    self.data = self.data[~self.data[col].astype(str).str.contains(value, na=False)]
                else:
                    # 尝试转换为数值
                    try:
                        numeric_value = float(value)
                        query_str = f"{col} {condition} {numeric_value}"
                    except ValueError:
                        query_str = f"{col} {condition} '{value}'"
                    
                    self.data = self.data.query(query_str)
                
                filtered_rows = len(self.data)
                removed_rows = original_rows - filtered_rows
                
                self.update_data_view()
                self.update_info_label()
                dialog.destroy()
                
                messagebox.showinfo("成功", f"筛选完成！\n保留 {filtered_rows} 行，删除 {removed_rows} 行")
                
            except Exception as e:
                messagebox.showerror("错误", f"筛选失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="🔍 应用筛选", command=apply_filter,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def random_sample(self):
        """随机采样"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建采样对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("随机采样")
        dialog.geometry("450x350")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (350 // 2)
        dialog.geometry(f"450x350+{x}+{y}")
        
        ttk.Label(dialog, text="随机采样设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 采样方式选择
        ttk.Label(dialog, text="采样方式：").pack(anchor='w', padx=20, pady=(10, 0))
        method_var = tk.StringVar(value="按数量")
        method_combo = ttk.Combobox(dialog, textvariable=method_var, values=['按数量', '按比例'])
        method_combo.pack(fill='x', padx=20, pady=5)
        
        # 采样值输入
        ttk.Label(dialog, text="采样值：").pack(anchor='w', padx=20, pady=(10, 0))
        value_var = tk.StringVar()
        value_entry = ttk.Entry(dialog, textvariable=value_var)
        value_entry.pack(fill='x', padx=20, pady=5)
        
        info_label = ttk.Label(dialog, text=f"当前数据行数: {len(self.data)}", foreground='blue')
        info_label.pack(pady=5)
        
        def perform_sample():
            method = method_var.get()
            value_str = value_var.get()
            
            if not value_str:
                messagebox.showwarning("警告", "请输入采样值！")
                return
                
            try:
                if method == "按数量":
                    n = int(value_str)
                    if n > len(self.data):
                        messagebox.showwarning("警告", "采样数量不能大于数据总行数！")
                        return
                    self.data = self.data.sample(n=n, random_state=42)
                else:  # 按比例
                    frac = float(value_str)
                    if frac <= 0 or frac > 1:
                        messagebox.showwarning("警告", "采样比例必须在0到1之间！")
                        return
                    self.data = self.data.sample(frac=frac, random_state=42)
                
                self.update_data_view()
                self.update_info_label()
                dialog.destroy()
                
                messagebox.showinfo("成功", f"随机采样完成！\n当前数据行数: {len(self.data)}")
                
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数值！")
            except Exception as e:
                messagebox.showerror("错误", f"采样失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="🎲 执行采样", command=perform_sample,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def sort_data(self):
        """数据排序"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建排序对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("数据排序")
        dialog.geometry("450x350")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (350 // 2)
        dialog.geometry(f"450x350+{x}+{y}")
        
        ttk.Label(dialog, text="数据排序设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 排序列选择
        ttk.Label(dialog, text="选择排序列：").pack(anchor='w', padx=20, pady=(10, 0))
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=list(self.data.columns))
        column_combo.pack(fill='x', padx=20, pady=5)
        
        # 排序方向选择
        ttk.Label(dialog, text="排序方向：").pack(anchor='w', padx=20, pady=(10, 0))
        ascending_var = tk.StringVar(value="升序")
        direction_combo = ttk.Combobox(dialog, textvariable=ascending_var, values=['升序', '降序'])
        direction_combo.pack(fill='x', padx=20, pady=5)
        
        def perform_sort():
            col = column_var.get()
            direction = ascending_var.get()
            
            if not col:
                messagebox.showwarning("警告", "请选择排序列！")
                return
                
            try:
                ascending = True if direction == "升序" else False
                self.data = self.data.sort_values(by=col, ascending=ascending)
                
                self.update_data_view()
                dialog.destroy()
                
                messagebox.showinfo("成功", f"数据按 {col} 列{direction}排序完成！")
                
            except Exception as e:
                messagebox.showerror("错误", f"排序失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="🔄 执行排序", command=perform_sort,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def plot_group_comparison(self):
        """绘制分组对比图"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 创建分组对比图对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("分组对比图")
        dialog.geometry("450x400")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"450x400+{x}+{y}")
        
        ttk.Label(dialog, text="分组对比图设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 分组列选择
        ttk.Label(dialog, text="选择分组列：").pack(anchor='w', padx=20, pady=(10, 0))
        group_var = tk.StringVar()
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=list(self.data.columns))
        group_combo.pack(fill='x', padx=20, pady=5)
        
        # 数值列选择
        ttk.Label(dialog, text="选择数值列：").pack(anchor='w', padx=20, pady=(10, 0))
        value_var = tk.StringVar()
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
        value_combo = ttk.Combobox(dialog, textvariable=value_var, values=numeric_cols)
        value_combo.pack(fill='x', padx=20, pady=5)
        
        # 图表类型选择
        ttk.Label(dialog, text="图表类型：").pack(anchor='w', padx=20, pady=(10, 0))
        plot_type_var = tk.StringVar(value="柱状图")
        plot_type_combo = ttk.Combobox(dialog, textvariable=plot_type_var, 
                                      values=['柱状图', '箱型图', '小提琴图', '分组散点图'])
        plot_type_combo.pack(fill='x', padx=20, pady=5)
        
        def create_group_plot():
            group_col = group_var.get()
            value_col = value_var.get()
            plot_type = plot_type_var.get()
            
            if not group_col or not value_col:
                messagebox.showwarning("警告", "请选择分组列和数值列！")
                return
                
            try:
                # 确保可视化环境
                self._ensure_matplotlib()
                # 清除之前的图表
                self.ax.clear()
                
                if plot_type == "柱状图":
                    grouped_data = self.data.groupby(group_col)[value_col].mean()
                    grouped_data.plot(kind='bar', ax=self.ax)
                    self.ax.set_ylabel(f'{value_col} 平均值')
                    self.ax.set_title(f'{group_col} 分组的 {value_col} 平均值对比')
                    self.ax.tick_params(axis='x', rotation=45)
                    
                elif plot_type == "箱型图":
                    groups = [group[value_col].dropna() for name, group in self.data.groupby(group_col)]
                    group_names = [name for name, group in self.data.groupby(group_col)]
                    
                    self.ax.boxplot(groups, labels=group_names)
                    self.ax.set_ylabel(value_col)
                    self.ax.set_title(f'{group_col} 分组的 {value_col} 箱型图')
                    self.ax.tick_params(axis='x', rotation=45)
                    
                elif plot_type == "小提琴图":
                    groups = [group[value_col].dropna() for name, group in self.data.groupby(group_col)]
                    positions = range(1, len(groups) + 1)
                    
                    self.ax.violinplot(groups, positions=positions)
                    group_names = [name for name, group in self.data.groupby(group_col)]
                    self.ax.set_xticks(positions)
                    self.ax.set_xticklabels(group_names, rotation=45)
                    self.ax.set_ylabel(value_col)
                    self.ax.set_title(f'{group_col} 分组的 {value_col} 小提琴图')
                    
                elif plot_type == "分组散点图":
                    groups = self.data.groupby(group_col)
                    colors = self.plt.cm.Set1(np.linspace(0, 1, len(groups)))
                    
                    for (name, group), color in zip(groups, colors):
                        self.ax.scatter(group.index, group[value_col], label=name, alpha=0.7, color=color)
                    
                    self.ax.set_ylabel(value_col)
                    self.ax.set_xlabel('索引')
                    self.ax.set_title(f'{group_col} 分组的 {value_col} 散点图')
                    self.ax.legend()
                
                self.fig.tight_layout()
                self.canvas.draw()
                self.notebook.select(1)  # 切换到可视化选项卡
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("错误", f"创建分组对比图失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="📊 创建图表", command=create_group_plot,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def plot_time_series(self):
        """绘制时间序列图"""
        if self.data is None:
            messagebox.showwarning("警告", "没有加载数据！")
            return
            
        # 查找日期时间列
        datetime_cols = []
        for col in self.data.columns:
            if self.data[col].dtype == 'datetime64[ns]' or 'date' in col.lower() or 'time' in col.lower():
                datetime_cols.append(col)
        
        if not datetime_cols:
            messagebox.showwarning("警告", "没有找到日期时间列！请先转换数据类型。")
            return
            
        # 创建时间序列图对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("时间序列图")
        dialog.geometry("450x350")
        dialog.configure(bg='#ecf0f1')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (350 // 2)
        dialog.geometry(f"450x350+{x}+{y}")
        
        ttk.Label(dialog, text="时间序列图设置", font=('微软雅黑', 12, 'bold')).pack(pady=10)
        
        # 时间列选择
        ttk.Label(dialog, text="选择时间列：").pack(anchor='w', padx=20, pady=(10, 0))
        time_var = tk.StringVar()
        time_combo = ttk.Combobox(dialog, textvariable=time_var, values=datetime_cols)
        time_combo.pack(fill='x', padx=20, pady=5)
        
        # 数值列选择
        ttk.Label(dialog, text="选择数值列：").pack(anchor='w', padx=20, pady=(10, 0))
        value_var = tk.StringVar()
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
        value_combo = ttk.Combobox(dialog, textvariable=value_var, values=numeric_cols)
        value_combo.pack(fill='x', padx=20, pady=5)
        
        def create_time_plot():
            time_col = time_var.get()
            value_col = value_var.get()
            
            if not time_col or not value_col:
                messagebox.showwarning("警告", "请选择时间列和数值列！")
                return
                
            try:
                # 确保可视化环境
                self._ensure_matplotlib()
                # 清除之前的图表
                self.ax.clear()
                
                # 确保时间列是datetime类型
                if self.data[time_col].dtype != 'datetime64[ns]':
                    time_data = pd.to_datetime(self.data[time_col], errors='coerce')
                else:
                    time_data = self.data[time_col]
                
                # 创建时间序列图
                self.ax.plot(time_data, self.data[value_col], marker='o', linewidth=1, markersize=3)
                self.ax.set_xlabel(time_col)
                self.ax.set_ylabel(value_col)
                self.ax.set_title(f'{value_col} 时间序列图')
                
                # 设置x轴标签旋转
                self.ax.tick_params(axis='x', rotation=45)
                self.ax.grid(True, alpha=0.3)
                
                self.fig.tight_layout()
                self.canvas.draw()
                self.notebook.select(1)  # 切换到可视化选项卡
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("错误", f"创建时间序列图失败: {str(e)}")
                
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill='x', pady=20, padx=20)
        
        ttk.Button(button_frame, text="📈 创建图表", command=create_time_plot,
                  style='Custom.TButton').pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="❌ 取消", command=dialog.destroy,
                  style='Custom.TButton').pack(side='right')
        
    def setup_shortcuts(self):
        """设置键盘快捷键"""
        self.root.bind('<Control-o>', lambda e: self.load_file())  # Ctrl+O 打开文件
        self.root.bind('<Control-s>', lambda e: self.save_data())  # Ctrl+S 保存数据
        self.root.bind('<F5>', lambda e: self.reset_data())  # F5 重置数据
        self.root.bind('<Control-d>', lambda e: self.remove_duplicates())  # Ctrl+D 删除重复数据
        self.root.bind('<Control-m>', lambda e: self.handle_missing_values())  # Ctrl+M 处理缺失值
        self.root.bind('<Control-r>', lambda e: self.remove_outliers())  # Ctrl+R 删除异常值
        self.root.bind('<Control-1>', lambda e: self.notebook.select(0))  # Ctrl+1 切换到数据表格
        self.root.bind('<Control-2>', lambda e: self.notebook.select(1))  # Ctrl+2 切换到可视化
        self.root.bind('<Control-3>', lambda e: self.notebook.select(2))  # Ctrl+3 切换到统计信息
        self.root.bind('<Control-4>', lambda e: self.notebook.select(3))  # Ctrl+4 切换到分组分析
        self.root.bind('<Control-5>', lambda e: self.notebook.select(4))  # Ctrl+5 切换到软件信息
        
    def create_status_bar(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.root, relief='sunken', borderwidth=1)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=0, pady=(5, 0))
        
        # 状态标签
        self.status_label = ttk.Label(status_frame, text="🟢 准备就绪", 
                                     font=('微软雅黑', 10), foreground='#27ae60')
        self.status_label.pack(side='left', padx=10, pady=3)
        
        # 版本信息
        version_label = ttk.Label(status_frame, text="DataInsight Pro v1.0", 
                                 font=('微软雅黑', 9), foreground='#7f8c8d')
        version_label.pack(side='right', padx=10, pady=3)
    
    def update_status(self, message, status_type="info"):
        """更新状态栏信息"""
        icons = {
            "info": "🟢",
            "warning": "🟡", 
            "error": "🔴",
            "working": "🔄"
        }
        colors = {
            "info": "#27ae60",
            "warning": "#f39c12",
            "error": "#e74c3c", 
            "working": "#3498db"
        }
        
        icon = icons.get(status_type, "🟢")
        color = colors.get(status_type, "#27ae60")
        
        self.status_label.config(text=f"{icon} {message}", foreground=color)
    
    def create_version_info(self, parent):
        """创建版本信息区域"""
        import platform
        import sys
        
        version_frame = ttk.LabelFrame(parent, text="ℹ️ 版本信息", padding="12")
        version_frame.pack(fill='x', pady=(10, 0))
        
        # 获取系统信息
        system_info = platform.system()
        system_version = platform.release()
        python_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        
        version_text = f"""系统信息：
🖥️ 操作系统: {system_info} {system_version}
🐍 Python版本: {python_version}
📊 工具版本: DataInsight Pro v1.0
👨‍💻 作者: Z000oooo
📅 更新日期: 2025-07
🌐 GitHub: github.com/Z000oooo/DataInsight-Pro"""
        
        version_label = ttk.Label(version_frame, text=version_text, font=('微软雅黑', 9), 
                                 foreground='#95a5a6', justify='left')
        version_label.pack(anchor='w')
        
        # 添加查看详细信息按钮
        self.detail_btn = ttk.Button(version_frame, text="📋 查看详细信息", 
                                    command=self.show_software_info, style='Custom.TButton')
        self.detail_btn.pack(anchor='w', pady=(10, 0))
    
    def show_software_info(self):
        """显示软件详细信息"""
        # 切换到软件信息选项卡（第5个选项卡，索引为4）
        self.notebook.select(4)
        self.update_status("已切换到软件信息页面", "info")

def main():
    root = tk.Tk()
    app = DataInsightPro(root)
    root.mainloop()

if __name__ == "__main__":
    main() 