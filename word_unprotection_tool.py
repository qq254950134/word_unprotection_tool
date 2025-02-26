import os
import re
import zipfile
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import xml.etree.ElementTree as ET

class WordUnprotectionTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Word文档保护解除工具")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # 设置主题颜色
        self.bg_color = "#f0f0f0"
        self.accent_color = "#4a86e8"
        self.text_color = "#333333"
        
        self.root.configure(bg=self.bg_color)
        
        self.setup_ui()
        
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="Word文档保护解除工具", 
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=10)
        file_frame.pack(fill=tk.X, pady=10)
        
        # 文件选择按钮
        self.file_btn = ttk.Button(
            file_frame, 
            text="选择Word文件", 
            command=self.select_files
        )
        self.file_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件夹选择按钮
        self.folder_btn = ttk.Button(
            file_frame, 
            text="选择文件夹", 
            command=self.select_folder
        )
        self.folder_btn.pack(side=tk.LEFT, padx=5)
        
        # 显示已选择文件数量
        self.file_count_var = tk.StringVar(value="已选择: 0 个文件")
        file_count_label = ttk.Label(file_frame, textvariable=self.file_count_var)
        file_count_label.pack(side=tk.LEFT, padx=20)
        
        # 输出目录选择
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding=10)
        output_frame.pack(fill=tk.X, pady=10)
        
        self.output_var = tk.StringVar(value="与源文件相同目录")
        
        # 输出选项
        same_dir_radio = ttk.Radiobutton(
            output_frame, 
            text="与源文件相同目录", 
            variable=self.output_var, 
            value="与源文件相同目录"
        )
        same_dir_radio.pack(anchor=tk.W)
        
        custom_dir_radio = ttk.Radiobutton(
            output_frame, 
            text="自定义输出目录", 
            variable=self.output_var, 
            value="自定义输出目录",
            command=self.toggle_output_dir
        )
        custom_dir_radio.pack(anchor=tk.W)
        
        # 自定义输出目录框架
        self.custom_output_frame = ttk.Frame(output_frame)
        self.custom_output_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.output_path_var = tk.StringVar()
        output_entry = ttk.Entry(
            self.custom_output_frame, 
            textvariable=self.output_path_var,
            state="disabled"
        )
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(20, 5))
        
        output_browse_btn = ttk.Button(
            self.custom_output_frame, 
            text="浏览...", 
            command=self.select_output_dir,
            state="disabled"
        )
        output_browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.output_entry = output_entry
        self.output_browse_btn = output_browse_btn
        
        # 自定义文件名设置
        filename_frame = ttk.Frame(output_frame)
        filename_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.filename_var = tk.StringVar(value="原文件名_unprotected")
        
        filename_label = ttk.Label(filename_frame, text="输出文件名格式:")
        filename_label.pack(side=tk.LEFT, padx=(0, 5))
        
        filename_entry = ttk.Entry(filename_frame, textvariable=self.filename_var)
        filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        filename_hint = ttk.Label(filename_frame, text="(使用 '原文件名' 表示原始文件名)")
        filename_hint.pack(side=tk.LEFT, padx=(5, 0))
        
        # 操作按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 10))
        
        self.start_btn = ttk.Button(
            btn_frame, 
            text="开始转换", 
            command=self.start_conversion,
            style="Accent.TButton"
        )
        self.start_btn.pack(side=tk.RIGHT, padx=5)
        
        self.clear_btn = ttk.Button(
            btn_frame, 
            text="清除选择", 
            command=self.clear_selection
        )
        self.clear_btn.pack(side=tk.RIGHT, padx=5)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = ScrolledText(
            log_frame, 
            width=70, 
            height=10, 
            wrap=tk.WORD,
            font=("Consolas", 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 进度条
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient=tk.HORIZONTAL, 
            length=100, 
            mode='determinate',
            variable=self.progress_var
        )
        self.progress_bar.pack(fill=tk.X)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(
            self.root, 
            textvariable=self.status_var,
            relief=tk.SUNKEN, 
            anchor=tk.W, 
            padding=(5, 2)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 自定义按钮样式
        style = ttk.Style()
        style.configure(
            "Accent.TButton",
            background=self.accent_color,
            foreground="white"
        )
        
        # 存储文件列表
        self.file_list = []
        
    def toggle_output_dir(self):
        if self.output_var.get() == "自定义输出目录":
            self.output_entry.config(state="normal")
            self.output_browse_btn.config(state="normal")
        else:
            self.output_entry.config(state="disabled")
            self.output_browse_btn.config(state="disabled")
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择Word文件",
            filetypes=[
                ("Word文件", "*.docx *.doc"),
                ("Word 2007-2019", "*.docx"),
                ("Word 97-2003", "*.doc"),
                ("所有文件", "*.*")
            ]
        )
        
        if files:
            self.file_list.extend(files)
            self.update_file_count()
            
            self.log(f"已添加 {len(files)} 个文件")
            for file in files:
                self.log(f"  - {os.path.basename(file)}")
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="选择包含Word文件的文件夹")
        
        if folder:
            count = 0
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.endswith((".docx", ".doc")):
                        file_path = os.path.join(root, file)
                        self.file_list.append(file_path)
                        count += 1
            
            self.update_file_count()
            
            self.log(f"从文件夹添加了 {count} 个Word文件")
            
    def select_output_dir(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_path_var.set(folder)
            
    def update_file_count(self):
        self.file_count_var.set(f"已选择: {len(self.file_list)} 个文件")
        
    def clear_selection(self):
        self.file_list = []
        self.update_file_count()
        self.log("已清除所有选择的文件")
        
    def log(self, message):
        """向日志区域添加消息，并确保界面更新"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        # 强制更新界面，确保日志立即显示
        self.root.update_idletasks()
        
    def update_status(self, message):
        self.status_var.set(message)
        
    def update_progress(self, value):
        self.progress_var.set(value)
        
    def start_conversion(self):
        if not self.file_list:
            messagebox.showwarning("警告", "请先选择Word文件")
            return
            
        if self.output_var.get() == "自定义输出目录" and not self.output_path_var.get():
            messagebox.showwarning("警告", "请选择输出目录")
            return
            
        # 禁用按钮，防止重复点击
        self.start_btn.config(state=tk.DISABLED)
        self.clear_btn.config(state=tk.DISABLED)
        self.file_btn.config(state=tk.DISABLED)
        self.folder_btn.config(state=tk.DISABLED)
        
        # 重置进度条
        self.update_progress(0)
        
        # 在新线程中执行转换，避免阻塞UI
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
        
    def process_files(self):
        total_files = len(self.file_list)
        processed = 0
        success = 0
        failed = 0
        
        self.log("===== 开始批量转换 =====")
        self.update_status("正在处理...")
        
        for file_path in self.file_list:
            try:
                filename = os.path.basename(file_path)
                self.log(f"正在处理: {filename}")
                
                # 确定输出路径
                if self.output_var.get() == "自定义输出目录":
                    output_dir = self.output_path_var.get()
                else:
                    output_dir = os.path.dirname(file_path)
                    
                # 确定输出文件名
                file_base = os.path.splitext(filename)[0]
                file_ext = os.path.splitext(filename)[1]
                
                output_name = self.filename_var.get().replace("原文件名", file_base)
                output_path = os.path.join(output_dir, output_name + file_ext)
                
                # 处理文件
                result = self.process_word_file(file_path, output_path)
                
                if result:
                    self.log(f"转换成功: {os.path.basename(output_path)}")
                    success += 1
                else:
                    self.log(f"转换失败: {filename}")
                    failed += 1
                    
            except Exception as e:
                self.log(f"处理错误: {filename} - {str(e)}")
                failed += 1
                
            processed += 1
            progress = (processed / total_files) * 100
            self.update_progress(progress)
            
        # 转换完成
        self.log(f"===== 转换完成 =====")
        self.log(f"总计: {total_files} 个文件")
        self.log(f"成功: {success} 个文件")
        self.log(f"失败: {failed} 个文件")
        
        self.update_status(f"完成 - 成功: {success}, 失败: {failed}")
        
        # 恢复按钮状态
        self.root.after(0, self.enable_buttons)
        
    def enable_buttons(self):
        self.start_btn.config(state=tk.NORMAL)
        self.clear_btn.config(state=tk.NORMAL)
        self.file_btn.config(state=tk.NORMAL)
        self.folder_btn.config(state=tk.NORMAL)
        
    def process_word_file(self, input_path, output_path):
        """处理单个Word文件的核心逻辑"""
        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        
        try:
            # 处理逻辑根据文件扩展名不同而不同
            file_ext = os.path.splitext(input_path)[1].lower()
            
            if file_ext == '.docx':
                return self.process_docx(input_path, output_path, temp_dir)
            elif file_ext == '.doc':
                return self.process_doc(input_path, output_path, temp_dir)
            else:
                self.log(f"不支持的文件类型: {file_ext}")
                return False
                
        except Exception as e:
            self.log(f"处理异常: {str(e)}")
            return False
        finally:
            # 清理临时目录
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
                
    def process_docx(self, input_path, output_path, temp_dir):
        """处理DOCX文件"""
        # 复制原文件到临时目录
        temp_file = os.path.join(temp_dir, "temp.docx")
        shutil.copy2(input_path, temp_file)
        
        # 解压docx文件 (docx本质上是一个zip文件)
        extract_dir = os.path.join(temp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        
        try:
            with zipfile.ZipFile(temp_file, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
        except zipfile.BadZipFile:
            self.log("  - 错误：无法解压文件，可能不是有效的DOCX文件")
            return False
        
        # 查找并修改文档保护相关的XML文件
        # 保护标签可能出现在多个XML文件中
        word_dir = os.path.join(extract_dir, "word")
        xml_files = []
        
        # 查找word目录下所有的XML文件
        if os.path.exists(word_dir):
            for root, _, files in os.walk(word_dir):
                for file in files:
                    if file.endswith(".xml"):
                        xml_files.append(os.path.join(root, file))
        
        # 主文档和设置文件优先检查
        doc_file = os.path.join(extract_dir, "word", "document.xml")
        settings_file = os.path.join(extract_dir, "word", "settings.xml")
        
        # 确保主文档和设置文件在列表的前面
        if doc_file in xml_files:
            xml_files.remove(doc_file)
        if settings_file in xml_files:
            xml_files.remove(settings_file)
            
        xml_files = [doc_file, settings_file] + xml_files
        
        modified = False
        
        # 处理所有XML文件
        for xml_file in xml_files:
            if os.path.exists(xml_file):
                try:
                    file_name = os.path.basename(xml_file)
                    
                    with open(xml_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # 检查各种可能的文档保护标记
                    protection_found = False
                    new_content = content
                    
                    # 1. 检查并移除documentProtection标签 (多种形式)
                    if re.search(r'<w:documentProtection\s[^>]*?/?>', new_content):
                        new_content = re.sub(r'<w:documentProtection\s[^>]*?/?>', '', new_content)
                        protection_found = True
                    
                    if re.search(r'<w:documentProtection\s[^>]*?>.*?</w:documentProtection>', new_content, re.DOTALL):
                        new_content = re.sub(r'<w:documentProtection\s[^>]*?>.*?</w:documentProtection>', '', new_content, re.DOTALL)
                        protection_found = True
                    
                    # 2. 检查并移除任何保护相关属性
                    if re.search(r'w:edit\s*=\s*["\'].*?["\']', new_content):
                        new_content = re.sub(r'w:edit\s*=\s*["\'].*?["\']', '', new_content)
                        protection_found = True
                    
                    if re.search(r'w:enforcement\s*=\s*["\'].*?["\']', new_content):
                        new_content = re.sub(r'w:enforcement\s*=\s*["\'].*?["\']', '', new_content)
                        protection_found = True
                    
                    # 3. 替换任何字符串引用
                    if "DocumentProtection" in new_content:
                        new_content = new_content.replace("DocumentProtection", "unDocumentProtection")
                        protection_found = True
                    
                    if "documentProtection" in new_content:
                        new_content = new_content.replace("documentProtection", "undocumentProtection")
                        protection_found = True
                    
                    # 写回修改后的内容
                    if protection_found:
                        with open(xml_file, 'w', encoding='utf-8') as f:
                            f.write(new_content)
                        
                        modified = True
                        self.log(f"  - 已移除{file_name}中的保护")
                        
                except Exception as e:
                    self.log(f"  - 处理{file_name}时出错: {str(e)}")
        
        if not modified:
            self.log("  - 未找到文档保护标签")
        
        # 重新打包为docx
        output_zip = os.path.join(temp_dir, "output.docx")
        
        with zipfile.ZipFile(output_zip, 'w') as zip_out:
            for root, _, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zip_out.write(file_path, arcname)
        
        # 复制到最终输出位置
        shutil.copy2(output_zip, output_path)
        return True
    
    def process_doc(self, input_path, output_path, temp_dir):
        """处理DOC文件"""
        try:
            # 尝试使用win32com处理DOC文件
            import importlib.util
            if importlib.util.find_spec("win32com"):
                self.log("  - 使用win32com处理DOC文件...")
                return self.process_doc_with_win32com(input_path, output_path)
            else:
                self.log("  - 未安装win32com模块")
                self.log("  - 尝试使用替代方法...")
                
                # 尝试先将文件复制到输出位置，然后使用二进制模式替换内容
                shutil.copy2(input_path, output_path)
                
                # 二进制方式打开并替换内容
                try:
                    with open(output_path, 'rb') as f:
                        content = f.read()
                    
                    # 尝试使用二进制替换
                    if b'documentProtection' in content:
                        new_content = content.replace(b'documentProtection', b'undocumentProtection')
                        
                        with open(output_path, 'wb') as f:
                            f.write(new_content)
                        
                        self.log("  - 已尝试二进制替换保护标记")
                        return True
                    else:
                        self.log("  - 未找到二进制保护标记")
                        return True
                except Exception as e:
                    self.log(f"  - 二进制处理错误: {str(e)}")
                
                # 如果上述方法都失败，提供安装建议
                self.log("  - 建议安装win32com处理DOC文件:")
                self.log("    pip install pywin32")
                return False
        except Exception as e:
            self.log(f"  - 处理DOC文件错误: {str(e)}")
            return False
            
    def process_doc_with_win32com(self, input_path, output_path):
        """使用win32com处理DOC文件"""
        try:
            from win32com import client
            import os
            
            word = client.Dispatch("Word.Application")
            word.Visible = False
            
            self.log("  - 打开Word文档...")
            doc = word.Documents.Open(os.path.abspath(input_path))
            
            # 检查并禁用保护
            if doc.ProtectionType != -1:  # -1表示没有保护
                self.log(f"  - 检测到保护类型: {doc.ProtectionType}")
                doc.Unprotect()
                self.log("  - 已解除保护")
            else:
                self.log("  - 文档未启用保护")
            
            # 保存为新文件
            self.log("  - 保存修改后的文档...")
            doc.SaveAs(os.path.abspath(output_path))
            doc.Close()
            word.Quit()
            
            self.log("  - DOC文件处理完成")
            return True
            
        except Exception as e:
            self.log(f"  - Win32COM处理错误: {str(e)}")
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = WordUnprotectionTool(root)
    root.mainloop()
