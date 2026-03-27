#!/usr/bin/env python3
"""
Excel分类导出工具
功能：导入Excel，按选定列分类，生成多个Excel文件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import threading


class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel分类导出工具")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        # 数据
        self.df = None
        self.columns = []
        self.file_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # 标题
        title_label = tk.Label(self.root, text="Excel分类导出工具", font=("微软雅黑", 18, "bold"))
        title_label.pack(pady=20)
        
        # 文件选择区
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill="x")
        
        self.file_label = tk.Label(file_frame, text="未选择文件", fg="gray")
        self.file_label.pack(side="left", fill="x", expand=True)
        
        btn_select = tk.Button(file_frame, text="选择Excel文件", command=self.select_file, 
                               bg="#4CAF50", fg="white", font=("微软雅黑", 10), padx=15, pady=5)
        btn_select.pack(side="right", padx=5)
        
        # 列选择区
        col_frame = tk.LabelFrame(self.root, text="选择分类列", font=("微软雅黑", 10))
        col_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.col_combo = ttk.Combobox(col_frame, state="disabled", font=("微软雅黑", 10))
        self.col_combo.pack(pady=10, padx=10, fill="x")
        
        # 预览区
        preview_frame = tk.LabelFrame(self.root, text="数据预览", font=("微软雅黑", 10))
        preview_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # 树形控件
        tree_scroll_y = tk.Scrollbar(preview_frame)
        tree_scroll_y.pack(side="right", fill="y")
        
        tree_scroll_x = tk.Scrollbar(preview_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")
        
        self.tree = ttk.Treeview(preview_frame, yscrollcommand=tree_scroll_y.set, 
                                  xscrollcommand=tree_scroll_x.set)
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        # 操作按钮
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)
        
        self.btn_export = tk.Button(btn_frame, text="开始分类导出", command=self.start_export,
                                     state="disabled", bg="#2196F3", fg="white", 
                                     font=("微软雅黑", 12), padx=20, pady=8)
        self.btn_export.pack()
        
        # 状态栏
        self.status_label = tk.Label(self.root, text="就绪", fg="gray", font=("微软雅黑", 9))
        self.status_label.pack(pady=5)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        self.file_path = file_path
        self.file_label.config(text=os.path.basename(file_path), fg="black")
        self.load_excel()
    
    def load_excel(self):
        try:
            self.status_label.config(text="正在加载...")
            self.root.update()
            
            # 读取Excel
            if self.file_path.endswith('.xlsx'):
                self.df = pd.read_excel(self.file_path, engine='openpyxl')
            else:
                self.df = pd.read_excel(self.file_path, engine='xlrd')
            
            # 获取列名
            self.columns = list(self.df.columns)
            
            # 更新下拉框
            self.col_combo.config(values=self.columns, state="readonly")
            if self.columns:
                self.col_combo.current(0)
                self.col_combo.config(state="normal")
            
            # 启用导出按钮
            self.btn_export.config(state="normal")
            
            # 更新预览
            self.update_preview()
            
            self.status_label.config(text=f"已加载 {len(self.df)} 行数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败：\n{str(e)}")
            self.status_label.config(text="加载失败")
    
    def update_preview(self):
        # 清空树
        self.tree.delete(*self.tree.get_children())
        
        # 设置列
        columns = list(self.df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")
        
        # 插入数据（最多显示100行）
        for idx, row in self.df.head(100).iterrows():
            values = [str(row[col]) for col in columns]
            self.tree.insert("", "end", values=values)
        
        if len(self.df) > 100:
            self.tree.insert("", "end", values=["..." for _ in columns])
    
    def start_export(self):
        if self.df is None or self.col_combo.get() == "":
            messagebox.showwarning("警告", "请先选择文件")
            return
        
        # 获取选中的列
        selected_col = self.col_combo.get()
        
        # 确认对话框
        categories = self.df[selected_col].unique()
        category_count = len(categories)
        
        confirm_msg = f"将按「{selected_col}」列分类\n共 {category_count} 个类别，是否继续？"
        if not messagebox.askyesno("确认", confirm_msg):
            return
        
        # 在新线程中执行导出
        self.btn_export.config(state="disabled", text="导出中...")
        self.status_label.config(text="正在导出...")
        
        thread = threading.Thread(target=self.export_files, args=(selected_col,))
        thread.start()
    
    def export_files(self):
        try:
            selected_col = self.col_combo.get()
            categories = self.df[selected_col].unique()
            
            # 获取原文件目录
            output_dir = os.path.dirname(self.file_path)
            if not output_dir:
                output_dir = os.path.dirname(os.path.abspath(__file__))
            
            # 按类别分组导出
            success_count = 0
            error_count = 0
            
            for category in categories:
                try:
                    # 筛选该类别的数据
                    df_category = self.df[self.df[selected_col] == category]
                    
                    # 处理文件名（去掉非法字符）
                    safe_name = str(category).replace("/", "-").replace("\\", "-")
                    safe_name = safe_name.replace(":", "-").replace("*", "-")
                    safe_name = safe_name.replace("?", "-").replace('"', "-")
                    safe_name = safe_name.replace("<", "-").replace(">", "-")
                    safe_name = safe_name.replace("|", "-")
                    
                    # 生成文件名
                    output_file = os.path.join(output_dir, f"{safe_name}.xlsx")
                    
                    # 导出到Excel
                    df_category.to_excel(output_file, index=False, engine='openpyxl')
                    
                    success_count += 1
                    
                except Exception as e:
                    print(f"导出类别 {category} 失败: {e}")
                    error_count += 1
            
            # 回到主线程更新UI
            self.root.after(0, self.export_complete, success_count, error_count)
            
        except Exception as e:
            self.root.after(0, messagebox.showerror, "错误", f"导出失败：\n{str(e)}")
    
    def export_complete(self, success_count, error_count):
        self.btn_export.config(state="normal", text="开始分类导出")
        
        if error_count == 0:
            self.status_label.config(text=f"导出完成！成功 {success_count} 个文件")
            messagebox.showinfo("完成", f"✅ 成功导出 {success_count} 个文件！\n文件保存在原表格目录下")
        else:
            self.status_label.config(text=f"完成，成功{success_count}个，失败{error_count}个")
            messagebox.showwarning("完成", f"✅ 成功导出 {success_count} 个文件\n❌ 失败 {error_count} 个文件")


def main():
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
