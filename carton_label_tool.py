# -*- coding: utf-8 -*-
"""
箱唛识别工具 - Windows独立版 v3.0
使用EasyOCR，无需安装Python即可运行
"""

import os
import sys
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
from pathlib import Path
import threading

# 导入OCR库
try:
    import easyocr
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    print("需要安装easyocr: pip install easyocr")

from PIL import Image
import pandas as pd

class CartonLabelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 箱唛识别工具")
        self.root.geometry("800x700")
        self.root.minsize(700, 600)
        
        # 初始化OCR（延迟加载）
        self.reader = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # 标题
        title_frame = tk.Frame(self.root)
        title_frame.pack(pady=20)
        
        tk.Label(title_frame, text="📦", font=("Segoe UI", 32)).pack()
        tk.Label(title_frame, text="箱唛识别工具", font=("微软雅黑", 20, "bold")).pack()
        tk.Label(title_frame, text="自动识别白色标签，生成Excel装箱清单", 
                font=("微软雅黑", 11), fg="gray").pack()
        
        # 选择导入方式
        import_frame = tk.LabelFrame(self.root, text="📁 选择导入", font=("微软雅黑", 10))
        import_frame.pack(pady=15, padx=30, fill=tk.X)
        
        btn_frame = tk.Frame(import_frame)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="📂 选择文件夹", command=self.select_folder,
                 font=("微软雅黑", 11), bg="#3b82f6", fg="white", 
                 width=15, height=2).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="🖼️ 选择图片", command=self.select_images,
                 font=("微软雅黑", 11), bg="#10b981", fg="white",
                 width=15, height=2).pack(side=tk.LEFT, padx=5)
        
        # 文件列表显示
        self.file_label = tk.Label(import_frame, text="未选择文件", 
                                  font=("微软雅黑", 9), fg="gray")
        self.file_label.pack()
        
        # 项目名称（自动识别，可修改）
        project_frame = tk.LabelFrame(self.root, text="🏢 项目名称（自动识别）", 
                                     font=("微软雅黑", 10))
        project_frame.pack(pady=10, padx=30, fill=tk.X)
        
        self.project_var = tk.StringVar(value="")
        self.project_entry = tk.Entry(project_frame, textvariable=self.project_var,
                                     font=("微软雅黑", 11), width=50)
        self.project_entry.pack(pady=10, padx=10, fill=tk.X)
        
        # 输出位置选择
        output_frame = tk.LabelFrame(self.root, text="💾 保存位置", font=("微软雅黑", 10))
        output_frame.pack(pady=10, padx=30, fill=tk.X)
        
        output_btn_frame = tk.Frame(output_frame)
        output_btn_frame.pack(pady=5)
        
        self.output_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        tk.Entry(output_btn_frame, textvariable=self.output_path, 
                font=("微软雅黑", 10), width=40).pack(side=tk.LEFT, padx=5)
        
        tk.Button(output_btn_frame, text="📁 浏览", command=self.select_output,
                 font=("微软雅黑", 10)).pack(side=tk.LEFT)
        
        # 开始按钮
        self.start_btn = tk.Button(self.root, text="🚀 开始识别", command=self.start_recognition,
                                  font=("微软雅黑", 14, "bold"), bg="#22c55e", fg="white",
                                  padx=40, pady=12, state=tk.DISABLED)
        self.start_btn.pack(pady=20)
        
        # 进度条
        self.progress = ttk.Progressbar(self.root, length=700, mode='determinate')
        self.progress.pack(pady=10, padx=30)
        
        self.status_label = tk.Label(self.root, text="请选择图片或文件夹", 
                                    font=("微软雅黑", 10), fg="gray")
        self.status_label.pack()
        
        # 识别结果预览
        result_frame = tk.LabelFrame(self.root, text="📋 识别结果预览", font=("微软雅黑", 10))
        result_frame.pack(pady=10, padx=30, fill=tk.BOTH, expand=True)
        
        # 创建表格
        columns = ('箱号', '明细', '数量', '楼层', '备注')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=8)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')
        
        self.tree.column('明细', width=250)
        
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 存储选择的文件
        self.selected_files = []
    
    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_files = []
            for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.JPEG', '*.PNG']:
                self.selected_files.extend(Path(folder).glob(ext))
            self.selected_files = sorted(self.selected_files, 
                                       key=lambda x: int(re.findall(r'\d+', x.name)[0]) 
                                       if re.findall(r'\d+', x.name) else 999)
            self.update_file_label()
    
    def select_images(self):
        files = filedialog.askopenfilenames(
            title="选择箱唛照片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png"), ("所有文件", "*.*")]
        )
        if files:
            self.selected_files = [Path(f) for f in files]
            self.update_file_label()
    
    def update_file_label(self):
        if self.selected_files:
            self.file_label.config(text=f"已选择 {len(self.selected_files)} 个文件", fg="green")
            self.start_btn.config(state=tk.NORMAL)
        else:
            self.file_label.config(text="未选择文件", fg="gray")
            self.start_btn.config(state=tk.DISABLED)
    
    def select_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path.set(folder)
    
    def init_ocr(self):
        """初始化OCR引擎"""
        if self.reader is None:
            self.status_label.config(text="正在加载OCR引擎（首次较慢，请等待）...")
            self.root.update()
            # 使用CPU模式，支持中英文
            self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
    
    def start_recognition(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择图片")
            return
        
        self.start_btn.config(state=tk.DISABLED, text="识别中...")
        self.tree.delete(*self.tree.get_children())
        
        # 在新线程运行
        thread = threading.Thread(target=self.process_images)
        thread.start()
    
    def process_images(self):
        try:
            # 初始化OCR
            self.init_ocr()
            
            total = len(self.selected_files)
            results = []
            project_name = ""
            
            for i, img_path in enumerate(self.selected_files, 1):
                self.root.after(0, lambda p=(i/total)*100: self.progress.config(value=p))
                self.root.after(0, lambda s=f"正在识别: {img_path.name}": 
                               self.status_label.config(text=s))
                
                # 识别图片
                result = self.recognize_image(img_path)
                
                if result:
                    # 从第一张图提取项目名称
                    if i == 1 and result.get('project'):
                        project_name = result['project']
                        self.root.after(0, lambda p=project_name: self.project_var.set(p))
                    
                    results.append(result)
                    
                    # 添加到表格
                    self.root.after(0, lambda r=result: self.add_to_table(r))
            
            # 生成Excel
            if results:
                self.root.after(0, lambda: self.status_label.config(text="正在生成Excel..."))
                output_file = self.create_excel(results)
                self.root.after(0, lambda: messagebox.showinfo("完成", 
                    f"✅ 识别完成！\n\n共识别 {len(results)} 个箱子\n已保存到:\n{output_file}"))
                
                # 尝试打开文件
                try:
                    os.startfile(output_file)
                except:
                    pass
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        
        finally:
            self.root.after(0, self.reset_ui)
    
    def recognize_image(self, img_path):
        """识别单张图片"""
        try:
            # 读取图片
            image = Image.open(img_path)
            
            # OCR识别
            ocr_result = self.reader.readtext(str(img_path), detail=1)
            
            # 解析结果
            return self.parse_ocr_result(ocr_result, img_path.name)
            
        except Exception as e:
            print(f"识别失败 {img_path}: {e}")
            return None
    
    def parse_ocr_result(self, ocr_result, filename):
        """解析OCR结果，提取白色标签内容"""
        info = {
            'box_no': '',
            'project': '',
            'item': '',
            'quantity': '',
            'floor': '',
            'remark': ''
        }
        
        # 从文件名提取箱号（备用）
        file_nums = re.findall(r'\d+', filename)
        if file_nums:
            info['box_no'] = file_nums[0]
        
        # 提取所有文字
        texts = [item[1] for item in ocr_result]
        full_text = ' '.join(texts)
        
        for text in texts:
            text = text.strip()
            if not text:
                continue
            
            # 提取箱号 NO: 1 / NO.1 / 编号:1
            box_match = re.search(r'[Nn][Oo][:.\s]*(\d+)', text)
            if box_match:
                info['box_no'] = box_match.group(1)
                continue
            
            # 提取数量 XXpcs / XX个 / XX件
            qty_match = re.search(r'(\d+)\s*(pcs|个|件|只|台|套|PC)', text, re.IGNORECASE)
            if qty_match:
                info['quantity'] = qty_match.group(1)
                # 尝试提取产品名（在同一行或前一行）
                if '：' in text or ':' in text:
                    parts = re.split(r'[:：]', text)
                    if len(parts) >= 2 and parts[0]:
                        info['item'] = parts[0].strip()
                continue
            
            # 提取项目名称（包含酒店、山庄、公寓等）
            if any(keyword in text for keyword in ['酒店', '山庄', '公寓', '温泉', '宾馆']):
                info['project'] = text.strip()
                continue
            
            # 如果还没提取到明细，且包含中文
            if not info['item'] and len(text) > 2 and re.search(r'[\u4e00-\u9fa5]', text):
                if 'NO' not in text.upper() and not re.match(r'^\d+$', text):
                    if 'pcs' not in text.lower():
                        info['item'] = text
        
        return info
    
    def add_to_table(self, result):
        """添加结果到表格"""
        self.tree.insert('', tk.END, values=(
            result.get('box_no', ''),
            result.get('item', ''),
            result.get('quantity', ''),
            result.get('floor', ''),
            result.get('remark', '')
        ))
    
    def create_excel(self, results):
        """创建Excel文件"""
        df = pd.DataFrame(results)
        
        # 删除不需要的列
        df = df[['box_no', 'item', 'quantity', 'floor', 'remark']]
        df.columns = ['箱号', '明细', '数量', '楼层', '备注']
        
        # 生成文件名
        project = self.project_var.get() or "项目"
        today = datetime.now().strftime('%Y%m%d')
        filename = f"{project}装箱清单{today}.xlsx"
        output_path = os.path.join(self.output_path.get(), filename)
        
        # 保存Excel
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        # 美化（这里简化处理，实际可以添加样式）
        return output_path
    
    def reset_ui(self):
        self.start_btn.config(state=tk.NORMAL, text="🚀 开始识别")
        self.status_label.config(text="就绪")
        self.progress.config(value=0)

def main():
    if not OCR_AVAILABLE:
        print("请先安装依赖: pip install easyocr pillow pandas openpyxl")
        input("按回车退出...")
        return
    
    root = tk.Tk()
    app = CartonLabelApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
