import time
import tkinter as tk
from tkinter import N, ttk, scrolledtext, messagebox
import configparser
import os
import threading

# Import functions from your existing scripts
from playwright.sync_api import sync_playwright, Page, BrowserContext
import traceback
import os
import re
from typing import Tuple
import traceback
import datetime

# 导入浏览器检测模块
from browser_detector import BrowserDetector

# CONFIG_FILE = 'config-test.ini'
# AUTH_FILE = 'auth-test.json'
CONFIG_FILE = 'config.ini'
AUTH_FILE = 'auth.json'



class NovelPublisherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("小说自动发布工具-内测v1.4")
        self.geometry("800x800")

        # Create and configure the main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Configuration Section ---
        config_frame = ttk.LabelFrame(main_frame, text="配置", padding="10")
        config_frame.pack(fill=tk.X, pady=5)

        self.config = configparser.ConfigParser()
        self.load_config()

        # Init More Var
        self.publish_time = datetime.datetime.now()

        # URL
        ttk.Label(config_frame, text="浏览器地址:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.custom_browser_path_var = tk.StringVar(
            value=self.config.get('Settings', 'custom_browser_path', fallback=''))
        ttk.Entry(config_frame, textvariable=self.custom_browser_path_var, width=70).grid(row=0, column=1, sticky=tk.EW,
                                                                                          padx=5)
        
        # Auto Detect Browser Path
        self.auto_detect_btn = ttk.Button(config_frame, text="自动检测", command=self.auto_detect_browser)
        self.auto_detect_btn.grid(row=0, column=2, sticky=tk.W, padx=5)

        # Publish Plate
        ttk.Label(config_frame, text="发布平台:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.publish_plate_var = tk.StringVar(value=self.config.get('Settings', 'publish_plate', fallback='番茄小说'))
        ttk.Combobox(config_frame, textvariable=self.publish_plate_var, values=['番茄小说', '起点小说', '七猫小说']).grid(
            row=1, column=1, sticky=tk.W, padx=5)

        # Publish Mode
        ttk.Label(config_frame, text="发布模式:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.publish_mode_var = tk.StringVar(value=self.config.get('Settings', 'publish_mode', fallback='publish'))
        ttk.Combobox(config_frame, textvariable=self.publish_mode_var, values=['publish', 'draft', 'pre-publish']).grid(
            row=2, column=1, sticky=tk.W, padx=5)

        # Publish Time
        ttk.Label(config_frame, text="每日定时发布时间:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.publish_time_var = tk.StringVar(value=self.config.get('Settings', 'publish_time', fallback='12:00'))
        ttk.Entry(config_frame, textvariable=self.publish_time_var, width=10).grid(row=3, column=1, sticky=tk.W, padx=5)

        # Daily Publish Num
        ttk.Label(config_frame, text="每日更新章节数:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.daily_publish_num_var = tk.StringVar(value=self.config.get('Settings', 'daily_publish_num', fallback='2'))
        ttk.Entry(config_frame, textvariable=self.daily_publish_num_var, width=10).grid(row=4, column=1, sticky=tk.W,
                                                                                        padx=5)

        # Novel Title
        ttk.Label(config_frame, text="小说标题:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=2)
        self.novel_title_var = tk.StringVar(value=self.config.get('Novel', 'novel_title', fallback=''))
        ttk.Entry(config_frame, textvariable=self.novel_title_var, width=80).grid(row=5, column=1, sticky=tk.EW, padx=5)

        # Novels Folder
        ttk.Label(config_frame, text="小说文件夹:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=2)
        self.novels_folder_var = tk.StringVar(value=self.config.get('Novel', 'novels_folder', fallback=''))
        ttk.Entry(config_frame, textvariable=self.novels_folder_var, width=80).grid(row=6, column=1, sticky=tk.EW,
                                                                                    padx=5)

        # Fast Publish Mode
        ttk.Label(config_frame, text="最速开书发布:").grid(row=7, column=0, sticky=tk.W, padx=5, pady=2)
        self.fast_publish_mode_var = tk.StringVar(value="10*1+10*3+5*7")
        ttk.Combobox(config_frame, textvariable=self.fast_publish_mode_var,
                     values=["15+15*2+2*7", "10*1+15*2+3*7", "10*1+10*3+5*7", "10*1+5*6+3*10"]).grid(row=7, column=1, sticky=tk.W, padx=5)

        # Chapter Count Display
        ttk.Label(config_frame, text="章节总数:").grid(row=8, column=0, sticky=tk.W, padx=5, pady=2)
        self.chapter_count_var = tk.StringVar(value="0")
        self.chapter_count_label = ttk.Label(config_frame, textvariable=self.chapter_count_var)
        self.chapter_count_label.grid(row=8, column=1, sticky=tk.W, padx=5, pady=2)

        config_frame.columnconfigure(1, weight=1)

        # Save Config Button
        ttk.Button(config_frame, text="保存配置", command=self.save_config).grid(row=8, column=2, sticky=tk.E, padx=5,
                                                                                 pady=2)

        # --- Action Section ---
        action_frame = ttk.LabelFrame(main_frame, text="操作", padding="10")
        action_frame.pack(fill=tk.X, pady=5)

        # Last Published Chapter
        self.last_published_chapter_date_var = tk.StringVar(
            value=self.config.get('History', 'last_published_chapter_date', fallback='N/A'))
        ttk.Label(action_frame, text=f"上次更新日期:").pack(side=tk.LEFT, padx=5)
        self.last_published_chapter_date_entry = ttk.Entry(action_frame,
                                                           textvariable=self.last_published_chapter_date_var, width=10)
        self.last_published_chapter_date_entry.pack(side=tk.LEFT, padx=5)

        # Chapter Range
        ttk.Label(action_frame, text="发布章节范围:").pack(side=tk.LEFT, padx=5)
        last_published_chapter = self.config.get('History', 'last_published_chapter', fallback='0')
        daily_publish_num = int(self.daily_publish_num_var.get())

        self.start_chapter_var = tk.StringVar(value=str(int(last_published_chapter) + 1))
        self.start_chapter_entry = ttk.Entry(action_frame, textvariable=self.start_chapter_var, width=3)
        self.start_chapter_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(action_frame, text="到").pack(side=tk.LEFT)
        self.end_chapter_var = tk.StringVar(value=str(int(last_published_chapter) + daily_publish_num))
        self.end_chapter_entry = ttk.Entry(action_frame, textvariable=self.end_chapter_var, width=3)
        self.end_chapter_entry.pack(side=tk.LEFT, padx=5)

        # Keep Browser Open
        self.keep_browser_open_var = tk.BooleanVar()
        ttk.Checkbutton(action_frame, text="后台执行", variable=self.keep_browser_open_var).pack(side=tk.LEFT, padx=20)

        self.fast_once_create_book_var = tk.BooleanVar()
        ttk.Checkbutton(action_frame, text="最速开书", variable=self.fast_once_create_book_var).pack(side=tk.LEFT,
                                                                                                     padx=20)

        # Login Button
        self.login_page = None
        ttk.Button(action_frame, text="登录 (完成)", command=self.run_login_thread).pack(side=tk.RIGHT, padx=5)

        # Run Button
        self.run_button = ttk.Button(action_frame, text="开始发布", command=self.run_automation_thread)
        self.run_button.pack(side=tk.RIGHT, padx=5)

        # --- Extended functionality ---
        others_frame = ttk.LabelFrame(main_frame, text="其他", padding="10")
        others_frame.pack(fill=tk.X, pady=5)


        # Export Button
        ttk.Button(others_frame, text="导出章节", command=self.export_chapters).pack(side=tk.LEFT, padx=5)

        # Import Button
        ttk.Button(others_frame, text="导入章节", command=self.import_chapters).pack(side=tk.LEFT, padx=5)

        # Manual Browser Button
        ttk.Button(others_frame, text="手动操作", command=self.open_manual_browser).pack(side=tk.LEFT, padx=5)

        # --- Log Section ---
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Redirect stdout to the log widget
        self.redirect_logging()

        # 初始化章节计数
        self.update_chapter_count()

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)

    def redirect_logging(self):
        # Simple redirection of print statements
        import sys
        sys.stdout = self
        sys.stderr = self

    def write(self, text):
        self.after(0, self.log, text.strip())

    def flush(self):
        pass

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            self.config.read(CONFIG_FILE, encoding='utf-8')
        else:
            self.config['Settings'] = {'publish_plate': '', 'publish_mode': 'publish'}
            self.config['Novel'] = {'novel_title': '', 'novels_folder': ''}
            
    def save_config(self):
        try:
            self.config['Settings']['custom_browser_path'] = self.custom_browser_path_var.get().replace('\\', '/')
            self.config['Settings']['publish_plate'] = self.publish_plate_var.get()
            self.config['Settings']['publish_mode'] = self.publish_mode_var.get()
            self.config['Settings']['publish_time'] = self.publish_time_var.get()
            self.config['Settings']['daily_publish_num'] = self.daily_publish_num_var.get()
            self.config['Novel']['novel_title'] = self.novel_title_var.get()
            self.config['History'] = {'last_published_chapter': '', 'last_published_chapter_date': ''}
            self.config['History']['last_published_chapter'] = str(int(self.start_chapter_var.get()) - 1)
            self.config['History']['last_published_chapter_date'] = self.last_published_chapter_date_var.get()
 
            if os.path.isdir(self.novels_folder_var.get()):
                self.config['Novel']['novels_folder'] = self.novels_folder_var.get()
            else:
                messagebox.showerror("错误", "请输入正确的小说文件夹路径！拆分导入章节已独立")
                return
            # if self.novels_folder_var.get().endswith('.txt'):
            #     self.create_chapter_files_in_files(self.novels_folder_var.get())
            # self.config['Novel']['novels_folder'] = self.novels_folder_var.get().replace('.txt', '')
            self.novels_folder_var.set(self.config['Novel']['novels_folder'])
            
            # 更新章节计数
            self.update_chapter_count()
            
            with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            messagebox.showinfo("成功", "配置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")

    def update_chapter_count(self):
        """更新章节总数显示"""
        try:
            novels_folder = self.novels_folder_var.get()
            if novels_folder and os.path.isdir(novels_folder):
                chapter_files = [f for f in os.listdir(novels_folder) if f.endswith('.md') and f.startswith('Chapter_')]
                count = len(chapter_files)
                self.chapter_count_var.set(str(count))
                print(f"检测到 {count} 个章节文件")
            else:
                self.chapter_count_var.set("0")
        except Exception as e:
            print(f"更新章节计数失败: {e}")
            self.chapter_count_var.set("0")

    def auto_detect_browser(self):
        """自动检测浏览器路径"""
        try:
            print("开始自动检测浏览器...")
            self.auto_detect_btn.config(state='disabled', text="检测中...")
            
            # 在新线程中执行检测，避免阻塞界面
            def detect_worker():
                try:
                    # 获取推荐的浏览器路径
                    recommended_path = BrowserDetector.get_recommended_browser()
                    
                    if recommended_path:
                        self.after(0, lambda: self._update_browser_path(recommended_path))
                        print(f"自动检测到浏览器: {recommended_path}")
                    else:
                        self.after(0, self._show_no_browser_found)
                        print("未检测到可用的浏览器")
                        
                except Exception as e:
                    self.after(0, lambda: self._show_detection_error(str(e)))
                    print(f"浏览器检测失败: {e}")
                finally:
                    self.after(0, lambda: self.auto_detect_btn.config(state='normal', text="自动检测"))
            
            # 启动检测线程
            detect_thread = threading.Thread(target=detect_worker, daemon=True)
            detect_thread.start()
            
        except Exception as e:
            messagebox.showerror("错误", f"启动浏览器检测失败: {e}")
            self.auto_detect_btn.config(state='normal', text="自动检测")

    def _update_browser_path(self, path: str):
        """更新浏览器路径显示"""
        self.custom_browser_path_var.set(path)
        messagebox.showinfo("成功", f"已自动检测到浏览器路径:\n{path}")

    def _show_no_browser_found(self):
        """显示未找到浏览器的提示"""
        messagebox.showwarning("提示", "未检测到可用的浏览器\n请手动设置浏览器路径")

    def _show_detection_error(self, error_msg: str):
        """显示检测错误"""
        messagebox.showerror("检测失败", f"浏览器检测失败:\n{error_msg}\n\n请手动设置浏览器路径")

    def run_login_thread(self):
        self.login_event = threading.Event()
        if self.login_page:
            self.login_event.set()
        else:
            threading.Thread(target=self.login, daemon=True).start()

    def _get_site_url(self):
        """根据发布平台获取登录URL"""
        publish_plate = self.publish_plate_var.get()
        if publish_plate == '番茄小说':
            return "https://fanqienovel.com/main/writer/?enter_from=author_zone"
        elif publish_plate == '起点小说':
            return "https://write.qq.com/portal/dashboard"
        elif publish_plate == '七猫小说':
            return "https://zuozhe.qimao.com/front/index"
        else:
            raise ValueError(f"不支持发布平台: {publish_plate}")

    def login(self):
        site_url = self._get_site_url()
        print("正在启动浏览器...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, executable_path=self.custom_browser_path_var.get())
            context = browser.new_context()
            self.login_page = context.new_page()

            print(f"请在打开的浏览器窗口中手动登录网站: {site_url}，完成后再次点击登录按钮完成登录。")
            self.login_page.goto(site_url, timeout=60000)
            # self.login_page.wait_for_timeout(600000)
            # 非阻塞等待登录完成信号
            while not self.login_event.is_set():
                time.sleep(1)
            try:
                context.storage_state(path=AUTH_FILE)
                print(f"登录状态已保存到 {AUTH_FILE}")
            except Exception as e:
                print(f"保存登录状态失败: {e}")
            finally:
                if browser:
                    browser.close()
                    print("浏览器已关闭。")
                    self.login_page = None

    def run_automation_thread(self):
        self.run_button.config(state='disabled')
        threading.Thread(target=self.run_automation, daemon=True).start()

    def run_automation(self):
        try:
            if self.fast_once_create_book_var.get():
                self.new_novel_publish_once()
            else:
                self.automation_flow()
        except ValueError:
            messagebox.showerror("错误", "章节范围必须是有效的整数。")
        except Exception as e:
            messagebox.showerror("运行时错误", f"执行自动化流程时发生错误: {e}")
        finally:
            self.after(0, self.enable_run_button)

    def enable_run_button(self):
        self.run_button.config(state='normal')

    def export_chapters(self):
        """导出章节功能"""
        try:
            novels_folder = self.novels_folder_var.get()
            if not novels_folder or not os.path.isdir(novels_folder):
                messagebox.showerror("错误", "请先设置正确的小说文件夹路径")
                return

            # 获取章节文件
            chapter_files = []
            for f in os.listdir(novels_folder):
                if f.endswith('.md') and f.startswith('Chapter_'):
                    match = re.search(r'_(\d+)', f)
                    if match:
                        chapter_num = int(match.group(1))
                        chapter_files.append((chapter_num, os.path.join(novels_folder, f)))
            
            if not chapter_files:
                messagebox.showerror("错误", "未找到章节文件")
                return

            chapter_files.sort(key=lambda x: x[0])
            
            # 创建导出对话框
            export_dialog = tk.Toplevel(self)
            export_dialog.title("导出章节")
            export_dialog.geometry("400x300")
            export_dialog.transient(self)
            export_dialog.grab_set()

            # 导出范围选择
            ttk.Label(export_dialog, text="选择导出范围:").pack(pady=10)
            
            # 范围输入框
            range_frame = ttk.Frame(export_dialog)
            range_frame.pack(pady=5)
            
            ttk.Label(range_frame, text="从第").pack(side=tk.LEFT)
            start_var = tk.StringVar(value="1")
            start_entry = ttk.Entry(range_frame, textvariable=start_var, width=5)
            start_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(range_frame, text="章到第").pack(side=tk.LEFT)
            end_var = tk.StringVar(value=str(chapter_files[-1][0]))
            end_entry = ttk.Entry(range_frame, textvariable=end_var, width=5)
            end_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(range_frame, text="章").pack(side=tk.LEFT)

            # 导出选项
            export_all_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(export_dialog, text="导出全部章节", 
                            variable=export_all_var,
                            command=lambda: self.toggle_export_range(start_entry, end_entry, export_all_var)).pack(pady=5)

            # 输出文件名
            ttk.Label(export_dialog, text="输出文件名:").pack(pady=5)
            output_var = tk.StringVar(value=f"{self.novel_title_var.get()}_导出.txt")
            ttk.Entry(export_dialog, textvariable=output_var, width=40).pack(pady=5)

            def perform_export():
                try:
                    if export_all_var.get():
                        start_chapter = 1
                        end_chapter = chapter_files[-1][0]
                    else:
                        start_chapter = int(start_var.get())
                        end_chapter = int(end_var.get())
                    
                    # 过滤要导出的章节
                    export_files = []
                    for chapter_num, filepath in chapter_files:
                        if start_chapter <= chapter_num <= end_chapter:
                            export_files.append((chapter_num, filepath))
                    
                    if not export_files:
                        messagebox.showerror("错误", "没有选择要导出的章节")
                        return

                    # 导出内容
                    output_content = []
                    for chapter_num, filepath in export_files:
                        try:
                            with open(filepath, 'r', encoding='utf-8') as f:
                                content = f.read()
                                # 移除Markdown格式，转换为纯文本
                                content = content.replace('# ', '').replace('**', '')
                                output_content.append(content)
                                output_content.append('\n\n')  # 章节之间添加空行
                        except Exception as e:
                            print(f"读取章节 {chapter_num} 失败: {e}")
                    
                    # 保存文件
                    output_path = os.path.join(novels_folder, output_var.get())
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(''.join(output_content))
                    
                    messagebox.showinfo("成功", f"导出完成！共导出 {len(export_files)} 个章节\n文件保存为: {output_path}")
                    export_dialog.destroy()
                    
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的章节范围")
                except Exception as e:
                    messagebox.showerror("错误", f"导出失败: {e}")

            # 导出按钮
            ttk.Button(export_dialog, text="开始导出", command=perform_export).pack(pady=20)

        except Exception as e:
            messagebox.showerror("错误", f"导出功能出错: {e}")

    def toggle_export_range(self, start_entry, end_entry, export_all_var):
        """切换导出范围输入框的状态"""
        if export_all_var.get():
            start_entry.config(state='disabled')
            end_entry.config(state='disabled')
        else:
            start_entry.config(state='normal')
            end_entry.config(state='normal')

    def import_chapters(self):
        """导入章节功能"""
        try:
            # 创建导入对话框
            import_dialog = tk.Toplevel(self)
            import_dialog.title("导入章节")
            import_dialog.geometry("600x500")
            import_dialog.transient(self)
            import_dialog.grab_set()
            
            # 文件选择区域
            file_frame = ttk.LabelFrame(import_dialog, text="选择TXT文件", padding="10")
            file_frame.pack(fill=tk.X, padx=10, pady=10)
            
            # 文件路径输入
            ttk.Label(file_frame, text="文件路径:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
            txt_file_var = tk.StringVar()
            txt_file_entry = ttk.Entry(file_frame, textvariable=txt_file_var, width=50)
            txt_file_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
            
            # 文件浏览按钮
            def browse_txt_file():
                from tkinter import filedialog
                file_path = filedialog.askopenfilename(
                    title="选择TXT文件",
                    filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
                )
                if file_path:
                    txt_file_var.set(file_path)
            
            ttk.Button(file_frame, text="浏览", command=browse_txt_file).grid(row=0, column=2, padx=5, pady=5)
            file_frame.columnconfigure(1, weight=1)
            
            # 章节列表区域
            list_frame = ttk.LabelFrame(import_dialog, text="章节列表", padding="10")
            list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 章节列表
            chapter_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, height=10)
            chapter_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # 滚动条
            scrollbar = ttk.Scrollbar(chapter_listbox, orient=tk.VERTICAL, command=chapter_listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            chapter_listbox.config(yscrollcommand=scrollbar.set)
            
            # 分解章节按钮
            def parse_chapters():
                file_path = txt_file_var.get()
                if not file_path or not os.path.exists(file_path):
                    messagebox.showerror("错误", "请选择有效的TXT文件")
                    return
                
                if not file_path.endswith('.txt'):
                    messagebox.showerror("错误", "请选择TXT文件")
                    return
                
                try:
                    # 读取文件内容
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # 使用正则表达式分解章节
                    chapter_pattern = re.compile(r'(^\s*#*\s*第\s*[一二三四五六七八九十百千万\dIVXLCDM]+\s*章.*$)', re.MULTILINE)
                    parts = chapter_pattern.split(content)
                    
                    if len(parts) <= 1:
                        messagebox.showerror("错误", "未找到任何章节")
                        return
                    
                    # 清空列表
                    chapter_listbox.delete(0, tk.END)
                    
                    # 解析章节
                    chapters = []
                    for i in range(1, len(parts), 2):
                        title = parts[i].strip()
                        if i + 1 < len(parts):
                            chapter_content = parts[i + 1].strip()
                            chapters.append({"title": title, "content": chapter_content})
                    
                    # 倒序添加到列表
                    for i in range(len(chapters)-1, -1, -1):
                        chapter_num = i + 1
                        title = chapters[i]["title"]
                        chapter_listbox.insert(tk.END, f"{chapter_num}_{title}")
                    
                    messagebox.showinfo("成功", f"成功解析出 {len(chapters)} 个章节")
                    
                except Exception as e:
                    messagebox.showerror("错误", f"解析文件失败: {str(e)}")
            
            # 追加十章按钮功能
            def select_next_ten():
                # 获取当前所有项目
                total_items = chapter_listbox.size()
                if total_items == 0:
                    messagebox.showinfo("提示", "请先分解章节")
                    return
                
                # 获取当前选中的项目
                currently_selected = set(chapter_listbox.curselection())
                
                # 从上到下查找10个未被选中的章节
                selected_count = 0
                for i in range(total_items):
                    if i not in currently_selected:
                        chapter_listbox.selection_set(i)
                        selected_count += 1
                        if selected_count >= 10:
                            break
                
                if selected_count == 0:
                    messagebox.showinfo("提示", "所有章节都已被选中")
                else:
                    pass
                    # messagebox.showinfo("成功", f"已选择 {selected_count} 个章节")
            
            # 清除所有选择按钮功能
            def clear_all_selections():
                chapter_listbox.selection_clear(0, tk.END)
                messagebox.showinfo("成功", "已清除所有选择")
            
            # 倒序按钮功能 - 重新编号章节序号
            def reverse_order():
                # 获取当前所有项目
                total_items = chapter_listbox.size()
                if total_items == 0:
                    messagebox.showinfo("提示", "请先分解章节")
                    return
                
                # 获取当前选中的项目
                currently_selected = set(chapter_listbox.curselection())
                
                # 获取所有项目内容（提取原始标题，去掉序号）
                items = []
                for i in range(total_items):
                    item_text = chapter_listbox.get(i)
                    # 分离序号和标题（格式为：序号_标题）
                    if "_" in item_text:
                        original_title = item_text.split("_", 1)[1]  # 获取标题部分
                    else:
                        original_title = item_text
                    items.append(original_title)
                
                # 清空列表并重新添加（倒序编号）
                chapter_listbox.delete(0, tk.END)
                for i, title in enumerate(reversed(items)):
                    new_chapter_num = total_items - i  # 倒序编号
                    chapter_listbox.insert(tk.END, f"{new_chapter_num}_{title}")
                
                # 恢复选中的项目（需要转换索引）
                if currently_selected:
                    new_selected = set()
                    for old_index in currently_selected:
                        # 倒序转换索引
                        new_index = total_items - 1 - old_index
                        new_selected.add(new_index)
                    
                    # 清除所有选择并重新选择
                    chapter_listbox.selection_clear(0, tk.END)
                    for new_index in new_selected:
                        chapter_listbox.selection_set(new_index)
                
                messagebox.showinfo("成功", "章节序号已倒序重排")
            
            # 按钮容器
            button_container = ttk.Frame(list_frame)
            button_container.pack(fill=tk.X, pady=5)
            
            ttk.Button(button_container, text="分解章节", command=parse_chapters).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_container, text="追加十章", command=select_next_ten).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_container, text="倒序", command=reverse_order).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_container, text="清除所有", command=clear_all_selections).pack(side=tk.RIGHT, padx=5)
            
            # 保存路径区域
            save_frame = ttk.LabelFrame(import_dialog, text="章节保存路径", padding="10")
            save_frame.pack(fill=tk.X, padx=10, pady=10)
            
            ttk.Label(save_frame, text="保存路径:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
            save_path_var = tk.StringVar(value=self.novels_folder_var.get())
            save_path_entry = ttk.Entry(save_frame, textvariable=save_path_var, width=50)
            save_path_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
            
            # 文件夹浏览按钮
            def browse_save_folder():
                from tkinter import filedialog
                folder_path = filedialog.askdirectory(title="选择章节保存文件夹")
                if folder_path:
                    save_path_var.set(folder_path)
            
            ttk.Button(save_frame, text="浏览", command=browse_save_folder).grid(row=0, column=2, padx=5, pady=5)
            save_frame.columnconfigure(1, weight=1)
            
            # 按钮区域
            button_frame = ttk.Frame(import_dialog)
            button_frame.pack(fill=tk.X, padx=10, pady=10)
            
            # 导入全部按钮
            def import_all():
                file_path = txt_file_var.get()
                if not file_path or not os.path.exists(file_path):
                    messagebox.showerror("错误", "请选择有效的TXT文件")
                    return
                
                save_path = save_path_var.get()
                if not save_path:
                    messagebox.showerror("错误", "请选择保存路径")
                    return
                
                try:
                    # 确保目录存在
                    os.makedirs(save_path, exist_ok=True)
                    
                    # 使用现有的create_chapter_files_in_files方法，但修改输出路径
                    if not self.create_chapter_files_in_files_custom(file_path, save_path):
                        messagebox.showerror("错误", "导入章节失败")
                        return
                    
                    # 更新主界面的小说文件夹路径
                    self.novels_folder_var.set(save_path)
                    self.update_chapter_count()
                    
                    messagebox.showinfo("成功", f"成功导入全部章节到: {save_path}")
                    import_dialog.destroy()
                    
                except Exception as e:
                    messagebox.showerror("错误", f"导入失败: {str(e)}")
            
            # 追加导入按钮
            def import_selected():
                file_path = txt_file_var.get()
                if not file_path or not os.path.exists(file_path):
                    messagebox.showerror("错误", "请选择有效的TXT文件")
                    return
                
                save_path = save_path_var.get()
                if not save_path:
                    messagebox.showerror("错误", "请选择保存路径")
                    return
                
                selected_indices = chapter_listbox.curselection()
                if not selected_indices:
                    messagebox.showerror("错误", "请选择要导入的章节")
                    return
                
                try:
                    # 确保目录存在
                    os.makedirs(save_path, exist_ok=True)
                    
                    # 获取当前文件夹中最大的章节号
                    max_chapter_num = 0
                    if os.path.exists(save_path):
                        for f in os.listdir(save_path):
                            if f.startswith('Chapter_') and f.endswith('.md'):
                                match = re.search(r'Chapter_(\d+)\.md', f)
                                if match:
                                    num = int(match.group(1))
                                    if num > max_chapter_num:
                                        max_chapter_num = num
                    
                    # 读取文件内容
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # 使用正则表达式分解章节
                    chapter_pattern = re.compile(r'(^\s*#*\s*第\s*[一二三四五六七八九十百千万\dIVXLCDM]+\s*章.*$)', re.MULTILINE)
                    parts = chapter_pattern.split(content)
                    
                    if len(parts) <= 1:
                        messagebox.showerror("错误", "未找到任何章节")
                        return
                    
                    # 解析章节
                    chapters = []
                    for i in range(1, len(parts), 2):
                        title = parts[i].strip()
                        if i + 1 < len(parts):
                            chapter_content = parts[i + 1].strip()
                            chapters.append({"title": title, "content": chapter_content})
                    
                    # 获取选中的章节（倒序索引转换）
                    selected_chapters = []
                    total_chapters = len(chapters)
                    for idx in sorted(selected_indices, reverse=True):
                        # 列表是倒序的，需要转换索引
                        actual_idx = total_chapters - 1 - idx
                        if 0 <= actual_idx < total_chapters:
                            selected_chapters.append(chapters[actual_idx])
                    
                    # 创建选中的章节文件
                    for i, chapter in enumerate(selected_chapters):
                        chapter_number = max_chapter_num + i + 1
                        title = str(chapter['title'])
                        if title.find('：') == -1:
                            c_index = title.find('章') + 1
                            if c_index < len(title) and title[c_index] in [':', ' ']:
                                title = title[:c_index] + '：' + title[c_index + 1:]
                        title = title.replace(' ', '')
                        file_name = f"Chapter_{chapter_number:03d}.md"
                        file_path = os.path.join(save_path, file_name)
                        
                        with open(file_path, 'w', encoding='utf-8') as chapter_file:
                            chapter_file.write(f"# {title}\n\n")
                            chapter_file.write(chapter['content'])
                    
                    # 更新主界面的小说文件夹路径
                    self.novels_folder_var.set(save_path)
                    self.update_chapter_count()
                    
                    messagebox.showinfo("成功", f"成功导入 {len(selected_chapters)} 个章节到: {save_path}")
                    import_dialog.destroy()
                    
                except Exception as e:
                    messagebox.showerror("错误", f"导入失败: {str(e)}")
            
            
            ttk.Button(button_frame, text="导入全部", command=import_all).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="追加导入", command=import_selected).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", command=import_dialog.destroy).pack(side=tk.RIGHT, padx=5)
            
        except Exception as e:
            messagebox.showerror("错误", f"打开导入对话框失败: {str(e)}")

    def create_chapter_files_in_files_custom(self, novel_file_path: str, custom_output_dir: str):
        """从文件中分解生成章节文件到指定目录"""
        if not novel_file_path.endswith('.txt'):
            print(f"错误: 文件 {novel_file_path} 不是文本文件，暂时不支持读取")
            return False

        os.makedirs(custom_output_dir, exist_ok=True)

        try:
            with open(novel_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError:
            print(f"错误: 文件 {novel_file_path} 未找到。")
            return False

        # Regex to split text by chapter markers (e.g., "第...章").
        # The pattern captures the chapter line, which is then used as the title.
        # It handles various numeral types (Chinese, Arabic, Roman).
        chapter_pattern = re.compile(r'(^\s*#*\s*第\s*[一二三四五六七八九十百千万\dIVXLCDM]+\s*章.*$)', re.MULTILINE)
        parts = chapter_pattern.split(content)

        if len(parts) <= 1:
            print("未找到任何章节。")
            return False

        # The result of split is [prologue, ch1_title, ch1_content, ch2_title, ch2_content, ...]
        # We skip the prologue at parts[0].
        chapters = []
        # The loop starts at 1 and steps by 2 to get title/content pairs.
        for i in range(1, len(parts), 2):
            title = parts[i].strip()
            # Ensure there is content for the chapter
            if i + 1 < len(parts):
                chapter_content = parts[i + 1].strip()
                chapters.append({"title": title, "content": chapter_content})

        # As per the request, we use the chapter index for file naming,
        # as a full Chinese-to-Arabic numeral conversion is complex without external libraries.
        for i, chapter in enumerate(chapters):
            chapter_number = i + 1
            _title = str(chapter['title'])
            if _title.find('：') == -1:
                c_index = _title.find('章') + 1
                if _title[c_index] in [':', ' ']:
                    _title = _title[:c_index] + '：' + _title[c_index + 1:]
            _title = _title.replace(' ', '')
            file_name = f"Chapter_{chapter_number:03d}.md"
            file_path = os.path.join(custom_output_dir, file_name)

            with open(file_path, 'w', encoding='utf-8') as chapter_file:
                chapter_file.write(f"# {_title}\n\n")
                chapter_file.write(chapter['content'])

        print(f"成功分解 {len(chapters)} 个章节到目录: {custom_output_dir}")
        return True

    def open_manual_browser(self):
        """打开手动浏览器，持续等待用户操作"""
        try:
            custom_browser_path = self.custom_browser_path_var.get()
            site_url = self._get_site_url()
            
            print("正在打开手动浏览器...")
            print("浏览器将保持打开状态，您可以进行手动操作")
            print("需要关闭时，请直接关闭浏览器窗口")
            
            # 在新线程中运行浏览器，避免阻塞主界面
            def run_manual_browser():
                try:
                    with sync_playwright() as p:
                        browser = p.chromium.launch(
                            headless=False, 
                            executable_path=custom_browser_path if custom_browser_path else None
                        )
                        if not os.path.exists(AUTH_FILE):
                            messagebox.showerror("错误", f"未找到认证文件 {AUTH_FILE} 请先执行登录操作")
                            return
                        context = browser.new_context(storage_state=AUTH_FILE)
                        page = context.new_page()
                        
                        # 导航到网站
                        if site_url:
                            page.goto(site_url, timeout=60000)
                            print(f"已导航到: {site_url}")
                        
                        print("手动浏览器已启动，您可以自由操作")
                        print("关闭浏览器窗口即可结束会话")
                        
                        # 保持浏览器打开，直到用户关闭
                        try:
                            while True:
                                if page.is_closed():
                                    break
                                time.sleep(1)
                        except Exception:
                            pass
                        
                        print("手动浏览器已关闭")
                        
                except Exception as e:
                    print(f"手动浏览器出错: {e}")
            
            # 启动浏览器线程
            browser_thread = threading.Thread(target=run_manual_browser, daemon=True)
            browser_thread.start()
            
        except Exception as e:
            messagebox.showerror("错误", f"打开手动浏览器失败: {e}")

    @staticmethod
    def create_chapter_files_in_files(novel_file_path: str):
        """从文件中分解生成章节文件"""
        if not novel_file_path.endswith('.txt'):
            print(f"错误: 文件 {novel_file_path} 不是文本文件，暂时不支持读取")
            return False

        output_dir = os.path.splitext(novel_file_path)[0]
        os.makedirs(output_dir, exist_ok=True)

        try:
            with open(novel_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError:
            print(f"错误: 文件 {novel_file_path} 未找到。")
            return False

        # Regex to split text by chapter markers (e.g., "第...章").
        # The pattern captures the chapter line, which is then used as the title.
        # It handles various numeral types (Chinese, Arabic, Roman).
        chapter_pattern = re.compile(r'(^\s*第[一二三四五六七八九十百千万\dIVXLCDM]+章.*$)', re.MULTILINE)
        parts = chapter_pattern.split(content)

        if len(parts) <= 1:
            print("未找到任何章节。")
            return False

        # The result of split is [prologue, ch1_title, ch1_content, ch2_title, ch2_content, ...]
        # We skip the prologue at parts[0].
        chapters = []
        # The loop starts at 1 and steps by 2 to get title/content pairs.
        for i in range(1, len(parts), 2):
            title = parts[i].strip()
            # Ensure there is content for the chapter
            if i + 1 < len(parts):
                chapter_content = parts[i + 1].strip()
                chapters.append({"title": title, "content": chapter_content})

        # As per the request, we use the chapter index for file naming,
        # as a full Chinese-to-Arabic numeral conversion is complex without external libraries.
        for i, chapter in enumerate(chapters):
            chapter_number = i + 1
            _title = str(chapter['title'])
            if _title.find('：') == -1:
                c_index = _title.find('章') + 1
                if _title[c_index] in [':', ' ']:
                    _title = _title[:c_index] + '：' + _title[c_index + 1:]
            _title = _title.replace(' ', '')
            file_name = f"Chapter_{chapter_number:03d}.md"
            file_path = os.path.join(output_dir, file_name)

            with open(file_path, 'w', encoding='utf-8') as chapter_file:
                chapter_file.write(f"# {_title}\n\n")
                chapter_file.write(chapter['content'])

        print(f"成功分解 {len(chapters)} 个章节到目录: {output_dir}")
        return True

    @staticmethod
    def get_chapter_files_in_range(novels_folder: str, start_chapter: int, end_chapter: int) -> list[str]:
        """从文件夹中寻找对应章节范围的文件。"""
        if not os.path.isdir(novels_folder):
            print(f"错误: 文件夹 {novels_folder} 不存在。")
            return []

        all_files = sorted([os.path.join(novels_folder, f) for f in os.listdir(novels_folder) if f.endswith('.md')])

        chapter_files = []
        for filepath in all_files:
            basename = os.path.basename(filepath)
            filename, _ = os.path.splitext(basename)
            match = re.search(r'_(\d+)', filename)
            if match:
                chapter_num = int(match.group(1))
                if start_chapter <= chapter_num <= end_chapter:
                    chapter_files.append((chapter_num, filepath))
        chapter_files.sort(key=lambda x: x[0])

        if not chapter_files:
            print(f"在文件夹 {novels_folder} 中没有找到从第 {start_chapter} 章到第 {end_chapter} 章的小说文件。")

        return chapter_files

    @staticmethod
    def get_chapter_details(filepath: str) -> Tuple[str, str, str]:
        """从.md文件中提取章节序号、标题和内容。"""
        basename = os.path.basename(filepath)
        filename, _ = os.path.splitext(basename)

        # 从 "Chapter_001" 格式中提取序号
        match = re.search(r'_(\d+)', filename)
        if match:
            chapter_num = match.group(1)
        else:
            raise ValueError(f"文件名 {filename} 格式错误，未找到章节序号。")

        with open(filepath, 'r', encoding='utf-8') as f:
            first_line = f.readline().strip()
            if first_line.find('：') == -1:
                c_index = first_line.find('章') + 1
                chapter_title = first_line[c_index + 1:]
            else:
                # 检查是否是Markdown标题
                if first_line.startswith('#'):
                    chapter_title = first_line.lstrip('# ').split('：', 1)[-1]
                else:
                    chapter_title = first_line.split('：', 1)[-1] if '：' in first_line else first_line

            content = f.read()

        print(f"读取章节: 第{chapter_num}章 - {chapter_title}")
        return chapter_num, chapter_title, content

    def publish_single_chapter_on_fanqienovel(self, context: BrowserContext, page: Page, chapter_details: tuple):
        """发布单个章节，处理新打开的页面，并在完成后关闭。"""
        chapter_num, chapter_title, chapter_content = chapter_details
        print(f"\n--- 开始发布: 第{chapter_num}章 ---")

        publish_page = None
        try:
            # 使用更稳定的方式查找小说并点击更新
            # 先 hover 到小说条目，让按钮出现
            row = page.get_by_text(self.novel_title_var.get()).locator("..").locator("..")
            row.hover()
            # 再点击“创建章节”按钮
            row.get_by_role("button", name="创建章节").click()

            with context.expect_page() as new_page_info:
                publish_page = new_page_info.value
                publish_page.wait_for_load_state('networkidle')
                print("已切换到新的发布页面。")

            # --- 在新页面上执行操作 ---
            publish_page.locator("#app").get_by_role("button", name="下一步").click()  # 预先点击去除新手引导
            # publish_page.click('body')
            print(f"填写章节序号: {chapter_num}")
            publish_page.wait_for_selector('span.left-input > input', timeout=5000)
            publish_page.fill('span.left-input > input', chapter_num.strip())

            print(f"填写章节标题: {chapter_title}")
            publish_page.get_by_placeholder("请输入标题").fill(chapter_title.strip())

            print("粘贴章节正文...")
            edit_path = '#app > div > div > div > div.publish-body > div.editor > div.serial-editor-container.notranslate > ' \
                        'div > div > div.syl-editor-container.font-size-16.indent-2 > div > div.ProseMirror'

            publish_page.wait_for_selector(edit_path, timeout=10000)
            publish_page.fill(edit_path, chapter_content)

            if self.publish_mode_var.get() == 'draft':
                print("保存为草稿...")
                publish_page.get_by_role("button", name="存草稿").click()
            else:
                print("点击'下一步'进行发布...")
                publish_page.get_by_role("button", name="下一步").click()

                # 处理发布确认弹窗
                print("处理发布确认弹窗...")
                while not publish_page.get_by_role("button", name="确认发布").is_visible():
                    try:
                        publish_page.get_by_role("button", name="提交").click(timeout=1500)
                    except Exception:
                        pass
                    try:
                        publish_page.get_by_role("button", name="确定").click(timeout=1500)
                        print("确认风险预检")
                        publish_page.wait_for_timeout(1500)
                    except Exception:
                        try:
                            print(f"重新填写章节序号: {chapter_num}")
                            publish_page.wait_for_selector('span.left-input > input', timeout=5000)
                            publish_page.fill('span.left-input > input', chapter_num.strip())

                            print(f"重新填写章节标题: {chapter_title}")
                            publish_page.get_by_placeholder("请输入标题").fill(chapter_title.strip())

                            print("重新粘贴章节正文...")
                            edit_path = '#app > div > div > div > div.publish-body > div.editor > div.serial-editor-container.notranslate > ' \
                                        'div > div > div.syl-editor-container.font-size-16.indent-2 > div > div.ProseMirror'

                            publish_page.wait_for_selector(edit_path, timeout=10000)
                            publish_page.fill(edit_path, chapter_content)
                            publish_page.fill('span.left-input > input', chapter_num.strip())

                            publish_page.get_by_role("button", name="下一步").click(timeout=1500)
                        except Exception:
                            # 所有操作失败时，短暂等待后继续循环
                            publish_page.wait_for_timeout(500)
                            continue

                print("发布设置...AI-默认否")
                # 选择“否”单选框
                publish_page.click('div.card-content-line-control > div > label:nth-child(2)')
                # publish_page.get_by_role("radiogroup").get_by_role("radio", name="否").check()

                if self.publish_mode_var.get() == 'pre-publish':
                    # 勾选定时发布
                    publish_page.click('div > div.card-content-line-control > div > button', timeout=3000)
                    # publish_page.get_by_role("button", name="定时发布").click()

                    print("定时发布时间: ", self.publish_time.strftime('%Y-%m-%d %H:%M'))
                    # publish_page.wait_for_selector('div:nth-child(1) > div.card-content-line-control > div > div.arco-picker-input > input', timeout=3000)
                    # publish_page.fill('div:nth-child(1) > div.card-content-line-control > div > div.arco-picker-input > input', 
                    # self.publish_time.strftime('%Y-%m-%d'))
                    # publish_page.fill('div:nth-child(2) > div.card-content-line-control > div > div.arco-picker-input > input', 
                    # self.publish_time.strftime('%H:%M'))


                    # 使用 get_by_placeholder 来定位输入框，并确保先清空再填写
                    date_input = publish_page.get_by_placeholder("请选择日期")
                    date_input_value = self.publish_time.strftime('%Y-%m-%d')
                    while date_input.input_value() != date_input_value:
                        date_input.clear()
                        date_input.fill(date_input_value)
                        date_input.press("Enter")
                        publish_page.wait_for_timeout(1000)

                    time_input = publish_page.get_by_placeholder("请选择时间")
                    time_input_value = self.publish_time.strftime('%H:%M')
                    while time_input.input_value() != time_input_value:
                        time_input.clear()
                        time_input.fill(time_input_value)
                        time_input.press("Enter")
                        publish_page.wait_for_timeout(1000)

                # 点击确认发布按钮
                publish_page.get_by_role("button", name="确认发布").click()

            print(f"--- 第{chapter_num}章发布成功！---")
            publish_page.wait_for_timeout(3000)  # 等待一下，确保操作完成

        except Exception as e:
            print(f"发布第{chapter_num}章时发生错误: {e}")
            traceback.print_exc()
            return False
        else:
            if publish_page:
                print("关闭章节发布页面...")
                publish_page.close()
            return True

    def publish_single_chapter_on_qidiannovel(self, context: BrowserContext, write_button, chapter_details: tuple):
        """发布单个章节"""
        chapter_num, chapter_title, chapter_content = chapter_details
        print(f"\n--- 开始发布: 第{chapter_num}章 ---")

        try:
            # --- 在新页面上执行操作 ---
            write_button.click(timeout=3000)
            publish_page = None
            with context.expect_page() as new_page_info:
                publish_page = new_page_info.value
                publish_page.wait_for_load_state('networkidle')
                print("已切换到发布页面。")
            while not publish_page.is_visible('#chapter-body > div > div.chapter-content > div.chapter-content-placeholder', timeout=5000):
                try:
                    publish_page.click('#root > div.write-tabs.ne-sidebar.ne-sidebar-status > div.ne-sidebar-tool > a', timeout=1000)
                    print("创建章节成功")
                    publish_page.wait_for_timeout(500) 
                except:
                    publish_page.click('#root > div.write-tabs.ne-sidebar.ne-sidebar-status > div.side-coll-btn > span', timeout=1000)
                    publish_page.wait_for_timeout(500) 

            print(f"填写章节序号和标题——第{chapter_num}章：{chapter_title}")
            title_input = publish_page.get_by_placeholder("请输入章节号与章节名。示例：“第十章 天降奇缘”")
            title_input.fill(f"第{chapter_num}章：{chapter_title}")

            print("粘贴章节正文...")
            try:
                # 方法1：使用XPath定位并填充内容
                tinymce_body = publish_page.wait_for_selector('#chapter-body > div > div.chapter-content > div.chapter-content-placeholder', timeout=5000)
                
                if tinymce_body:
                    # 点击编辑器区域使其获得焦点
                    tinymce_body.click()
                    publish_page.wait_for_timeout(500)
                    
                    # 清空现有内容
                    publish_page.keyboard.press("Control+A")
                    publish_page.keyboard.press("Delete")
                    publish_page.wait_for_timeout(300)
                    
                    # 输入新内容
                    publish_page.keyboard.type(chapter_content)
                    publish_page.wait_for_timeout(500)
                    
                    print("成功填入章节内容")
                else:
                    print("未找到TinyMCE编辑器")
                    
            except Exception as e:
                print(f"填入章节内容时出错: {e}")
                return False

            if self.publish_mode_var.get() == 'draft':
                print("保存为草稿...")
                publish_page.get_by_role("button", name="保存").click()
            else:
                print("点击'发布'进行发布...")
                publish_page.get_by_role("button", name="发布").click()
                publish_page.wait_for_timeout(3000)
                print("发布设置...")
                if self.publish_mode_var.get() == 'pre-publish':
                    # 勾选定时发布
                    publish_page.click('body > div.publish-form-dialog.nu_dialog_wrap._middle._open > div > div.ui-dialog-body > form > ul > li:nth-child(6) > div > div > label', timeout=3000)
                    # publish_page.get_by_role("button", name="定时发布").click()
                    print("定时发布时间: ", self.publish_time.strftime('%Y-%m-%d %H:%M'))
                    # 先选择时间，不然不允许点击日期
                    time_input_value = self.publish_time.strftime('%H:%M')
                    # time_button = publish_page.query_selector(f'a.ui-time-item[data-time="{time_input_value}"][role="button"]')
                    time_button = publish_page.query_selector(f'//a[@class="publish-set-time" and @role="button" and text()="{time_input_value}"]')
                    if time_button:
                        time_button.click()
                        print(f"成功选择常用时间: {time_input_value}")
                    else:
                        # 测试困难 -- 正确性未知
                        # 如果时间按钮不可见，点击时间输入框展开选择器
                        while not publish_page.is_visible('span.ui-input.ui-time-input.active'):
                            try:
                                publish_page.click('span.ui-input.ui-time-input')
                                publish_page.wait_for_timeout(500)
                            except:
                                publish_page.wait_for_timeout(500)
                        
                        # 选择小时
                        hour_value = time_input_value.split(":")[0]
                        hour_element = publish_page.query_selector(f'div.ui-time-picker a.ui-time-item[data-type="hour"][role="button"]:has-text("{hour_value}")')
                        if hour_element and hour_element.is_visible():
                            hour_element.click()
                            publish_page.wait_for_timeout(300)
                        else:
                            print(f"未找到小时元素: {hour_value}")
                            return False
                            
                        # 选择分钟
                        minute_value = time_input_value.split(":")[1]
                        minute_element = publish_page.query_selector(f'div.ui-time-picker a.ui-time-item[data-type="minute"][role="button"]:has-text("{minute_value}")')
                        if minute_element and minute_element.is_visible():
                            minute_element.click()
                            publish_page.wait_for_timeout(300)
                        else:
                            print(f"未找到分钟元素: {minute_value}")
                            return False

                    while not publish_page.is_visible('span.ui-input.ui-date-input.active'):
                        try:
                            publish_page.click("span.ui-input.ui-date-input", timeout=2000)
                            publish_page.wait_for_timeout(500)
                        except:
                            publish_page.wait_for_timeout(500)
                    month_text = publish_page.query_selector('div.ui-date-head > a.ui-date-switch').inner_text()
                    while month_text != self.publish_time.strftime('%Y-%m'):
                        publish_page.click('div > div.ui-date-head > a.ui-date-next') # 目前只能选择下个月
                        publish_page.wait_for_timeout(500)
                        month_text = publish_page.query_selector('div.ui-date-head > a.ui-date-switch').inner_text()

                    date_input_value = self.publish_time.strftime('%Y-%m-%d')
                    day_number = self.publish_time.day
                    date_element = publish_page.query_selector(f'a.ui-date-item:has-text("{day_number}")')

                    if date_element:
                        try:
                            date_element.click(timeout=1000)
                        except:
                            print("无法点击")
                        print(f"成功选择日期: {date_input_value}")
                    else:
                        print(f"未找到日期元素: {date_input_value}")

                    publish_page.get_by_role("button", name="定时发布").click()

                # 点击确认发布按钮
                else:
                    publish_page.get_by_role("button", name="确认发布").click()

            print(f"--- 第{chapter_num}章发布成功！---")
            publish_page.wait_for_timeout(3000)  # 等待一下，确保操作完成

        except Exception as e:
            print(f"发布第{chapter_num}章时发生错误: {e}")
            traceback.print_exc()
            return False
        else:
            if publish_page:
                print("关闭章节发布页面...")
                publish_page.close()
            return True

    def publish_single_chapter_on_qimaonovel(self, context: BrowserContext, page: Page, chapter_details: tuple, daily_publish_num: int):
        """发布单个章节，处理新打开的页面，并在完成后关闭。"""
        chapter_num, chapter_title, chapter_content = chapter_details
        print(f"\n--- 开始发布: 第{chapter_num}章 ---")

        publish_page = page # 七猫发布页面不会弹窗
        try:
            # 使用更稳定的方式查找小说并点击更新
            print(f"填写章节序号: {chapter_num} - 七猫有默认序号")

            print(f"填写章节标题: {chapter_title}")
            publish_page.get_by_placeholder("请输入章节名称，最多20个字").fill(chapter_title.strip())

            print("粘贴章节正文...")
            edit_path = 'div.chapter-editor > div > div > div.q-contenteditable.book.font-size-16.edit-mask'
            # 'div.chapter-editor > div > div'

            publish_page.wait_for_selector(edit_path, timeout=10000)
            publish_page.fill(edit_path, chapter_content)

            if self.publish_mode_var.get() == 'draft':
                print("保存为草稿...")
                publish_page.get_by_role("button", name="存为草稿").click()
            else:
                pulish_button = None
                add_var = int(self.daily_publish_num_var.get()) - daily_publish_num - 1
                # 七猫不允许同时定时发布，要求一定要晚一点
                while not pulish_button:
                    try:
                        if self.publish_mode_var.get() == 'pre-publish':
                            # 勾选定时发布
                            # publish_page.click('div > div.card-content-line-control > div > button', timeout=3000)
                            publish_page.get_by_text("定时发布").click()
                            publish_page.wait_for_timeout(1500)
                            print("获取确认发布按钮...")
                            # 使用更可靠的选择器，避免绝对路径
                            pulish_button = publish_page.query_selector('div.el-dialog__wrapper.qm-pop div.qm-pop-tf.show-mask > a, div.el-dialog__wrapper.qm-pop div.qm-pop-tf > a, .qm-pop-tf.show-mask > a, .qm-pop-tf > a')
                            if not pulish_button:
                                print(f"尝试其他方法1")
                                pulish_button = publish_page.get_by_text("确认发布", timeout=2000)
                            
                            # 如果还找不到，尝试更通用的选择器
                            if not pulish_button:
                                print(f"尝试其他方法2")
                                pulish_button = publish_page.query_selector('div[class*="qm-pop"] a, div[class*="el-dialog__body"] a')
                            # pulish_button = publish_page.get_by_text('确认发布', timeout=1500)
                            print(f"成功获取发布按钮")
                            date_input = publish_page.get_by_placeholder("选择日期")
                            date_input_value = self.publish_time.strftime('%Y-%m-%d')
                            print(f"发布日期: {date_input_value}")
                            while date_input.input_value() != date_input_value:
                                date_input.clear()
                                date_input.fill(date_input_value)
                                date_input.press("Enter")
                                publish_page.wait_for_timeout(1000)
                        
                            # 七猫也有常用时间设计
                            # 七猫不允许同时定时发布，要求一定要晚一点
                            if add_var > 0:
                                new_publish_time = self.publish_time + datetime.timedelta(minutes=add_var)
                            else:
                                new_publish_time = self.publish_time
                            time_input_value = new_publish_time.strftime('%H:%M')
                            print(f"发布时间: {time_input_value}")
                            print(f"Ps.七猫不允许同时定时发布，要求一定要晚于上一章发布时间")
                            # time_button = publish_page.query_selector(f'a.ui-time-item[data-time="{time_input_value}"][role="button"]')
                            time_button = publish_page.get_by_text(time_input_value).first
                            if time_button:
                                time_button.click()
                                print(f"成功选择常用时间: {time_input_value}")
                            else:
                                print(f"未找到常用时间元素: {time_input_value}")
                                return False
                                # 测试困难 -- 正确性未知
                                # 选择小时
                                hour_value = time_input_value.split(":")[0]
                                publish_page.get_by_placeholder("小时").click()
                                hour_element = publish_page.query_selector(f'//li[contains(@class, "el-select-dropdown__item") and .//span[text()="{hour_value}"]]')        
                                if hour_element and hour_element.is_visible():
                                    hour_element.click()
                                    publish_page.wait_for_timeout(300)
                                else:
                                    print(f"未找到小时元素: {hour_value}")
                                    return False
                                    
                                # 选择分钟
                                minute_value = time_input_value.split(":")[1]
                                publish_page.query_selector('分钟').click()
                                minute_element = publish_page.query_selector(f'//li[contains(@class, "el-select-dropdown__item") and .//span[text()="{minute_value}"]]')
                                if minute_element and minute_element.is_visible():
                                    minute_element.click()
                                    publish_page.wait_for_timeout(300)
                                else:
                                    print(f"未找到分钟元素: {minute_value}")
                                    return False
                        else:
                            publish_page.get_by_text("立即发布", timeout=3000).click()
                            publish_page.wait_for_timeout(1500)
                            print("获取确认发布按钮...")
                            # 使用更可靠的选择器，避免绝对路径
                            pulish_button = publish_page.query_selector('div.el-dialog__wrapper.qm-pop div.qm-pop-tf.show-mask > a, div.el-dialog__wrapper.qm-pop div.qm-pop-tf > a, .qm-pop-tf.show-mask > a, .qm-pop-tf > a')
                            if not pulish_button:
                                print(f"尝试其他方法1")
                                pulish_button = publish_page.get_by_text("确认发布", timeout=2000)
                            
                            # 如果还找不到，尝试更通用的选择器
                            if not pulish_button:
                                print(f"尝试其他方法2")
                                pulish_button = publish_page.query_selector('div[class*="qm-pop"] a, div[class*="el-dialog__body"] a')
                            # pulish_button = publish_page.get_by_text('确认发布', timeout=1500)
                            print(f"成功获取发布按钮")
                    except:
                        if pulish_button is None:
                            # try:
                            #     _button = input("输入发布按钮元素")
                            #     pulish_button = publish_page.query_selector(_button, timeout=1500)
                            # except:
                            #     publish_page.wait_for_timeout(500)
                            publish_page.get_by_placeholder("请输入章节名称，最多20个字").fill(chapter_title.strip())
                            print("重新粘贴章节正文...")
                            edit_path = 'div.chapter-editor > div > div > div.q-contenteditable.book.font-size-16.edit-mask'
                            # 'div.chapter-editor > div > div'
                            publish_page.wait_for_selector(edit_path, timeout=10000)
                            publish_page.fill(edit_path, chapter_content)
                            publish_page.wait_for_timeout(3000)
                            print("重新尝试发布")
                
                # 点击确认发布按钮
                pulish_button.click()
                publish_page.wait_for_timeout(5000)
                if add_var == 0:
                    try:
                        # 需要点击已阅读并同意协议
                        has_read_button = publish_page.query_selector('div.el-dialog__wrapper.qm-pop.first-upload-dialog div.qm-pop-tf > a, div.first-upload-dialog div.qm-pop-tf > a, .first-upload-dialog .qm-pop-tf > a')
                        if has_read_button:
                            has_read_button.click(timeout=3000)
                        else:
                            # 尝试通过文本内容查找
                            has_read_button = publish_page.get_by_text("已阅读并同意协议").first
                        if has_read_button:
                            has_read_button.click(timeout=3000)
                    except:
                        pass

            print(f"--- 第{chapter_num}章发布成功！---")
            publish_page.wait_for_timeout(3000)  # 等待一下，确保操作完成

        except Exception as e:
            print(f"发布第{chapter_num}章时发生错误: {e}")
            traceback.print_exc()
            return False
        else:
            # if publish_page:
            #     print("关闭章节发布页面...")
            #     publish_page.close()
            return True

    def automation_flow(self):
        """主自动化流程，负责初始化和循环调用章节发布。"""
        start_time = time.perf_counter()
        if self.publish_plate_var.get() == '番茄小说':
            self.automation_flow_by_fanqienovel()
        elif self.publish_plate_var.get() == '起点小说':
            self.automation_flow_by_qidiannovel()
        elif self.publish_plate_var.get() == '七猫小说':
            self.automation_flow_by_qimaonovel()
        else:
            print("不支持的发布平台")
        end_time = time.perf_counter()
        spend_time = end_time - start_time
        print(f"任务总耗时: {spend_time//60:.0f}分{spend_time%60:.2f}秒")

    def automation_flow_by_fanqienovel(self):
        custom_browser_path = self.custom_browser_path_var.get()
        site_url = self._get_site_url()
        novel_title = self.novel_title_var.get()
        novels_folder = self.novels_folder_var.get()
        publish_mode = self.publish_mode_var.get()
        start_chapter = int(self.start_chapter_var.get())
        end_chapter = int(self.end_chapter_var.get())
        keep_browser_open = self.keep_browser_open_var.get()

        if not os.path.exists(AUTH_FILE):
            print(f"错误：找不到登录状态文件 {AUTH_FILE}。请先登录，允许扫码登录。")
            return

        if not os.path.exists(novels_folder):
            print(f"错误：小说文件夹 {novels_folder} 不存在。")
            return

        if start_chapter is not None and end_chapter is not None:
            novel_files = self.get_chapter_files_in_range(novels_folder, start_chapter, end_chapter)
        else:
            novel_files = sorted(
                [os.path.join(novels_folder, f) for f in os.listdir(novels_folder) if f.endswith('.md')])

        if not novel_files:
            print(f"在文件夹 {novels_folder} 中没有找到要发布的小说文件。")
            return

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=keep_browser_open, executable_path=custom_browser_path)
            context = browser.new_context(storage_state=AUTH_FILE)
            page = context.new_page()

            try:
                print("导航到作者后台...")
                page.goto(site_url, timeout=60000, wait_until='networkidle')

                # 尝试关闭初始引导浮窗
                try:
                    page.click('div.user-guide-btn > button', timeout=1000)
                    print("已关闭初始引导浮窗。")
                except Exception:
                    print("未找到初始引导浮窗，继续执行...")

                print("点击进入小说列表页面...")
                novel_list_selector = '#app > div > div.content.new-content > div.serial-affix > div > div > div > div > div:nth-child(2) > div.new-nav-children.new-nav-children-expanded > div:nth-child(1)'
                try:
                    page.wait_for_selector(novel_list_selector, timeout=60000)  # 等待元素出现
                    page.click(novel_list_selector, timeout=60000)  # 延长点击超时
                    page.wait_for_load_state('networkidle')
                except Exception as e:
                    print(f"点击小说列表导航元素时发生错误: {e}")
                    raise  # 重新抛出异常以便上层捕获

                print(f"正在查找小说: '{novel_title}'")

                while True:
                    page.wait_for_selector('[id^="long-article-table-item-"]')
                    page.wait_for_timeout(3000)  # 等待一下，确保加载完成
                    novel_items = page.query_selector_all('[id^="long-article-table-item-"]')
                    update_button = None
                    for item in novel_items:
                        title_element = item.query_selector(
                            'div > div.book-item-info > div.info-content > div.info-content-title.font-1 > div')
                        # print(title_element.inner_text())
                        if title_element and title_element.inner_text() == novel_title:
                            print(f"找到小说 '{novel_title}'。")
                            # 定位到“更新章节”按钮
                            item.hover(timeout=3000)  # 鼠标悬停到该小说条目，确保按钮可见  
                            update_button = item.query_selector(
                                'div > div.book-item-info > div.info-content > div.info-right > div > a:nth-child(3) > button')
                            update_button_unsign = item.query_selector(
                                'div > div.book-item-info > div.info-content > div.info-right > div > a:nth-child(4) > button')
                            if update_button:
                                break
                            elif update_button_unsign:
                                print("小说未签约")
                                update_button = update_button_unsign
                                break
                            else:
                                print("未找到'更新章节'按钮")

                    if not update_button:
                        next_page_button_limit = page.query_selector(
                            '#arco-tabs-3-panel-0 > div > div > div > div.arco-pagination.arco-pagination-size-default.serial-pagination.long-article-table-pagination > ul > li.arco-pagination-item.arco-pagination-item-next.arco-pagination-item-disabled')
                        next_page_button = page.query_selector(
                            "#arco-tabs-1-panel-0 > div > div > div > div.arco-pagination.arco-pagination-size-default.serial-pagination.long-article-table-pagination > ul > li.arco-pagination-item.arco-pagination-item-next")
                        if next_page_button and not next_page_button_limit:
                            next_page_button.click()
                            page.wait_for_load_state('networkidle')
                        else:
                            print(f"错误：在列表中未找到小说 '{novel_title}'")
                            return
                    else:
                        break

                # 循环发布所有章节
                if publish_mode == "pre-publish":
                    time_str = self.last_published_chapter_date_var.get() + " " + self.publish_time_var.get()
                    self.publish_time = datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M')
                    self.publish_time += datetime.timedelta(days=1)
                    print("预计定时发布时间: ", self.publish_time.strftime('%Y-%m-%d %H:%M'))
                    daily_publish_num = int(self.daily_publish_num_var.get())
                    print(f"计划：日更{daily_publish_num}章！")

                list_count = len(novel_files)
                for i, novel_file in novel_files:
                    chapter_details = self.get_chapter_details(novel_file)
                    if publish_mode == "pre-publish":
                        if daily_publish_num > 0:
                            daily_publish_num -= 1
                        else:
                            daily_publish_num = int(self.daily_publish_num_var.get()) - 1
                            self.publish_time += datetime.timedelta(days=1)
                    if not self.publish_single_chapter_on_fanqienovel(context, page, chapter_details):
                        print(f"发布章节 {chapter_details[0]} 失败。")
                        break
                    list_count -= 1
                    if list_count > 0:
                        print(f"本次任务剩余{list_count}章")
                        print("准备发布下一章...")
                    # page.reload()
                    page.wait_for_timeout(3000)
                browser.close()
                if list_count == 0:
                    print("\n所有章节发布完毕！")
                    self.config['History']['last_published_chapter'] = str(end_chapter)
                    self.config['History']['last_published_chapter_date'] = self.publish_time.strftime('%Y-%m-%d')
                    # 更新界面输入框的显示值
                    self.start_chapter_var.set(str(end_chapter + 1))
                    self.end_chapter_var.set(str(end_chapter + int(self.daily_publish_num_var.get()) + 1))
                    self.last_published_chapter_date_var.set(self.publish_time.strftime('%Y-%m-%d'))

                    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                        self.config.write(configfile)
                    return True
                else:
                    print("\n有章节发布失败，已停止发布后续章节。")
                    return False
            except Exception as e:
                traceback.print_exc()
                print("主自动化流程中发生错误:")
                print("常见问题有：登陆失效")

    def automation_flow_by_qidiannovel(self):
        custom_browser_path = self.custom_browser_path_var.get()
        site_url = self._get_site_url()
        novel_title = self.novel_title_var.get()
        novels_folder = self.novels_folder_var.get()
        publish_mode = self.publish_mode_var.get()
        start_chapter = int(self.start_chapter_var.get())
        end_chapter = int(self.end_chapter_var.get())
        keep_browser_open = self.keep_browser_open_var.get()

        if not os.path.exists(AUTH_FILE):
            print(f"错误：找不到登录状态文件 {AUTH_FILE}。请先登录，允许扫码登录。")
            return

        if not os.path.exists(novels_folder):
            print(f"错误：小说文件夹 {novels_folder} 不存在。")
            return

        if start_chapter is not None and end_chapter is not None:
            novel_files = self.get_chapter_files_in_range(novels_folder, start_chapter, end_chapter)
        else:
            novel_files = sorted(
                [os.path.join(novels_folder, f) for f in os.listdir(novels_folder) if f.endswith('.md')])

        if not novel_files:
            print(f"在文件夹 {novels_folder} 中没有找到要发布的小说文件。")
            return

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=keep_browser_open, executable_path=custom_browser_path)
            context = browser.new_context(storage_state=AUTH_FILE)
            page = context.new_page()

            try:
                print("导航到作者后台...")
                page.goto(site_url, timeout=60000, wait_until='networkidle')
                write_button = None  
                # 直接查找小说标题(因为一般起点是单开，工作台首页即所需发布的书籍)
                page.wait_for_selector('#body > div.g-row.mt24 > div.g-col.g-col-10.g-body-main > div:nth-child(2) > div.g-prodution-item.fix.pb24 > div.g-prodution-item-lf.pr > div.g-prodution-item-title > a > h3', timeout=60000)
                actual_title = page.query_selector('#body > div.g-row.mt24 > div.g-col.g-col-10.g-body-main > div:nth-child(2) > div.g-prodution-item.fix.pb24 > div.g-prodution-item-lf.pr > div.g-prodution-item-title > a > h3').inner_text()
                if actual_title == novel_title:
                    print(f"已找到小说: '{novel_title}'")
                    write_button = page.get_by_role("button", name="去写作")
                else:
                    print(f"未找到小说: '{novel_title}'，继续查找...")
                    print("点击进入小说列表页面...")
                    page.query_selector("#body > div.g-row.mt24 > div.g-col.g-col-2.g-side > div > div > ul > li:nth-child(2) > a").click(timeout=1500)
                    novel_list_selector = '#book-body > div.g-nav-con > div > ul'
                    try:
                        page.wait_for_selector(novel_list_selector, timeout=60000)  # 等待元素出现
                        page.click(novel_list_selector, timeout=60000)  # 延长点击超时
                        page.wait_for_load_state('networkidle')
                    except Exception as e:
                        print(f"点击小说列表导航元素时发生错误: {e}")
                        raise  # 重新抛出异常以便上层捕获

                    print(f"正在查找小说: '{novel_title}'")

                    while True:
                        page.wait_for_selector('.g-prodution-item', timeout=10000)
                        page.wait_for_timeout(3000)
                        
                        # 获取所有小说项目
                        novel_items = page.query_selector_all('.g-prodution-item')
                        target_item = None
                        
                        for item in novel_items:
                            # 查找标题元素
                            title_element = item.query_selector('h3.open-book')
                            
                            if title_element and title_element.inner_text().strip() == novel_title:
                                print(f"找到小说 '{novel_title}'")
                                target_item = item
                                break
                        
                        if target_item:
                            # 找到后执行相关操作
                            write_button = target_item.query_selector('a.ui-button-primary')
                            if write_button:
                                write_button.click()
                                break
                        else:
                            # 处理分页
                            next_button = page.query_selector('.pagination-next:not(.disabled)')
                            if next_button:
                                next_button.click()
                                page.wait_for_load_state('networkidle')
                                page.wait_for_timeout(2000)
                            else:
                                print(f"错误：在列表中未找到小说 '{novel_title}'")
                                return False

                # 循环发布所有章节
                if publish_mode == "pre-publish":
                    time_str = self.last_published_chapter_date_var.get() + " " + self.publish_time_var.get()
                    self.publish_time = datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M')
                    self.publish_time += datetime.timedelta(days=1)
                    print("预计定时发布时间: ", self.publish_time.strftime('%Y-%m-%d %H:%M'))
                    daily_publish_num = int(self.daily_publish_num_var.get())
                    print(f"计划：日更{daily_publish_num}章！")

                list_count = len(novel_files)
                for i, novel_file in novel_files:
                    chapter_details = self.get_chapter_details(novel_file)
                    if publish_mode == "pre-publish":
                        if daily_publish_num > 0:
                            daily_publish_num -= 1
                        else:
                            daily_publish_num = int(self.daily_publish_num_var.get()) - 1
                            self.publish_time += datetime.timedelta(days=1)
                    if not self.publish_single_chapter_on_qidiannovel(context, write_button, chapter_details):
                        print(f"发布章节 {chapter_details[0]} 失败。")
                        break
                    list_count -= 1
                    if list_count > 0:
                        print(f"本次任务剩余{list_count}章")
                        print("准备发布下一章...")
                    # page.reload()
                    page.wait_for_timeout(3000)
                browser.close()
                if list_count == 0:
                    print("\n所有章节发布完毕！")
                    self.config['History']['last_published_chapter'] = str(end_chapter)
                    self.config['History']['last_published_chapter_date'] = self.publish_time.strftime('%Y-%m-%d')
                    # 更新界面输入框的显示值
                    self.start_chapter_var.set(str(end_chapter + 1))
                    self.end_chapter_var.set(str(end_chapter + int(self.daily_publish_num_var.get())))
                    self.last_published_chapter_date_var.set(self.publish_time.strftime('%Y-%m-%d'))

                    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                        self.config.write(configfile)
                    return True
                else:
                    print("\n有章节发布失败，已停止发布后续章节。")
                    return False
            except Exception as e:
                traceback.print_exc()
                print("主自动化流程中发生错误:")
                print("常见问题有：登陆失效")

    def automation_flow_by_qimaonovel(self):
        custom_browser_path = self.custom_browser_path_var.get()
        site_url = self._get_site_url()
        novel_title = self.novel_title_var.get()
        novels_folder = self.novels_folder_var.get()
        publish_mode = self.publish_mode_var.get()
        start_chapter = int(self.start_chapter_var.get())
        end_chapter = int(self.end_chapter_var.get())
        keep_browser_open = self.keep_browser_open_var.get()

        if not os.path.exists(AUTH_FILE):
            print(f"错误：找不到登录状态文件 {AUTH_FILE}。请先登录，允许扫码登录。")
            return

        if not os.path.exists(novels_folder):
            print(f"错误：小说文件夹 {novels_folder} 不存在。")
            return

        if start_chapter is not None and end_chapter is not None:
            novel_files = self.get_chapter_files_in_range(novels_folder, start_chapter, end_chapter)
        else:
            novel_files = sorted(
                [os.path.join(novels_folder, f) for f in os.listdir(novels_folder) if f.endswith('.md')])

        if not novel_files:
            print(f"在文件夹 {novels_folder} 中没有找到要发布的小说文件。")
            return

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=keep_browser_open, executable_path=custom_browser_path)
            context = browser.new_context(storage_state=AUTH_FILE)
            page = context.new_page()

            try:
                print("导航到作者后台...")
                page.goto(site_url, timeout=60000, wait_until='networkidle')
                write_button = None  
                # 直接查找小说标题(因为一般七猫也应该是单开，先查找工作台首页所需发布的书籍)
                # page.wait_for_selector('', timeout=60000)
                actual_title = page.query_selector('#app > div.layout > div.qm-main > div.wrapper > div > div.right-col > div.right-col-content.bg-index > div > div.author-book > div.index-module > div.index-module-body > div > ul > li > div > div.book-item-pic > div.txt > div > div.p-top > a').inner_text()
                if actual_title == novel_title:
                    print(f"已找到小说: '{novel_title}'")
                    write_button = page.get_by_text('新建章节')
                    
                else:
                    print(f"未找到小说: '{novel_title}'，继续查找...")
                    print("点击进入小说列表页面...")
                    page.query_selector("#app > div.layout > div.qm-main > div.wrapper > div > div.left-col > div > div > div.el-scrollbar__wrap > div > ul > li:nth-child(2) > ul > li:nth-child(1) > div").click(timeout=1500)
                    novel_list_selector = '???' # 编辑签约网站一般一个号更一本，该功能暂时无意义
                    try:
                        page.wait_for_selector(novel_list_selector, timeout=60000)  # 等待元素出现
                        page.click(novel_list_selector, timeout=60000)  # 延长点击超时
                        page.wait_for_load_state('networkidle')
                    except Exception as e:
                        print(f"点击小说列表导航元素时发生错误: {e}")
                        raise  # 重新抛出异常以便上层捕获

                    print(f"正在查找小说: '{novel_title}'")

                    while True:
                        books_list = '#app > div.layout > div.qm-main > div.wrapper > div > div.right-col > div.right-col-content.bg-shadow > div > div > div.qm-mod > div.qm-mod-tb > ul'
                        page.wait_for_selector(books_list, timeout=10000)
                        page.wait_for_timeout(3000)
                        
                        # 获取所有小说项目
                        novel_items = page.query_selector_all(books_list)
                        target_item = None
                        
                        for item in novel_items:
                            # 查找标题元素
                            title_element = item.query_selector('li > div.pic-txt > div.txt')

                            if title_element and title_element.inner_text().strip() == novel_title:
                                print(f"找到小说 '{novel_title}'")
                                target_item = item
                                break
                        
                        if target_item:
                            # 找到后执行相关操作
                            write_button = target_item.query_selector('div:nth-child(3) > span.s-right > a:nth-child(1) > span')
                            if write_button:
                                break
                        else:
                            # 处理分页
                            next_button = page.query_selector('.pagination-next:not(.disabled)')
                            if next_button:
                                next_button.click()
                                page.wait_for_load_state('networkidle')
                                page.wait_for_timeout(2000)
                            else:
                                print(f"错误：在列表中未找到小说 '{novel_title}'")
                                return False

                # 循环发布所有章节
                if publish_mode == "pre-publish":
                    time_str = self.last_published_chapter_date_var.get() + " " + self.publish_time_var.get()
                    self.publish_time = datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M')
                    self.publish_time += datetime.timedelta(days=1)
                    print("预计定时发布时间: ", self.publish_time.strftime('%Y-%m-%d %H:%M'))
                    daily_publish_num = int(self.daily_publish_num_var.get())
                    print(f"计划：日更{daily_publish_num}章！")

                list_count = len(novel_files)
                for i, novel_file in novel_files:
                    chapter_details = self.get_chapter_details(novel_file)
                    if publish_mode == "pre-publish":
                        if daily_publish_num > 0:
                            daily_publish_num -= 1
                        else:
                            daily_publish_num = int(self.daily_publish_num_var.get()) - 1
                            self.publish_time += datetime.timedelta(days=1)
                    write_button.click() # 点击新建章节按钮
                    if not self.publish_single_chapter_on_qimaonovel(context, page, chapter_details, daily_publish_num):
                        print(f"发布章节 {chapter_details[0]} 失败。")
                        break
                    while True:
                        try:
                            back_index = page.get_by_text("专区首页").first
                            back_index.click(timeout=1500)
                            page.wait_for_timeout(3000)
                            break
                        except Exception as e:
                            print(f"等待超时错误: {e}")
                            continue

                    # 回到首页
                    list_count -= 1
                    if list_count > 0:
                        print(f"本次任务剩余{list_count}章")
                        print("准备发布下一章...")
                    # page.reload()
                    # page.wait_for_timeout(3000)
                browser.close()
                if list_count == 0:
                    print("\n所有章节发布完毕！")
                    self.config['History']['last_published_chapter'] = str(end_chapter)
                    self.config['History']['last_published_chapter_date'] = self.publish_time.strftime('%Y-%m-%d')
                    # 更新界面输入框的显示值
                    self.start_chapter_var.set(str(end_chapter + 1))
                    self.end_chapter_var.set(str(end_chapter + int(self.daily_publish_num_var.get()) + 1))
                    self.last_published_chapter_date_var.set(self.publish_time.strftime('%Y-%m-%d'))

                    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                        self.config.write(configfile)
                    return True
                else:
                    print("\n有章节发布失败，已停止发布后续章节。")
                    return False
            except Exception as e:
                traceback.print_exc()
                print("主自动化流程中发生错误:")
                print("常见问题有：登陆失效")

    def parse_fast_publish_mode(self, fast_publish_mode):
        """
        解析快速发布模式字符串，生成发布计划
        例如: "15*1+15*2+2*7" 表示:
        - 第一步: 发布1-15章
        - 第二步: 预发布16-45章，每日15章，持续2天
        - 第三步: 预发布46-60章，每日2章，持续7天
        """
        try:
            # 分割字符串获取各个阶段
            stages = fast_publish_mode.split('+')
            plan = []
            
            current_chapter = 1
            
            for i, stage in enumerate(stages):
                if '*' in stage:
                    # 解析格式如 "15*3" (每日章节数*持续天数)
                    daily_count, days = stage.split('*')
                    daily_count = int(daily_count)
                    days = int(days)
                    
                    if i == 0:
                        # 第一步总是立即发布模式
                        end_chapter = current_chapter + daily_count - 1
                        plan.append({
                            'mode': 'publish',
                            'start_chapter': current_chapter,
                            'end_chapter': end_chapter,
                            'description': f"发布第{current_chapter}-{end_chapter}章"
                        })
                        current_chapter = end_chapter + 1
                    else:
                        # 后续步骤是预发布模式
                        end_chapter = current_chapter + daily_count * days - 1
                        plan.append({
                            'mode': 'pre-publish',
                            'start_chapter': current_chapter,
                            'end_chapter': end_chapter,
                            'daily_count': daily_count,
                            'days': days,
                            'description': f"预发布第{current_chapter}-{end_chapter}章 (每日{daily_count}章)"
                        })
                        current_chapter = end_chapter + 1
                else:
                    # 处理简单数字，如 "15"
                    count = int(stage)
                    end_chapter = current_chapter + count - 1
                    
                    if i == 0:
                        # 第一步总是立即发布模式
                        plan.append({
                            'mode': 'publish',
                            'start_chapter': current_chapter,
                            'end_chapter': end_chapter,
                            'description': f"发布第{current_chapter}-{end_chapter}章"
                        })
                    current_chapter = end_chapter + 1
            
            return plan
        except Exception as e:
            print(f"解析发布模式失败: {e}")
            return None

    def execute_publish_plan(self, plan):
        """
        执行发布计划
        """
        try:
            if self.publish_plate_var.get() != '番茄小说':
                print("最速开书流程目前仅支持番茄小说")
                return False
                
            for i, step in enumerate(plan, 1):
                print(f"\n--- 步骤{i}: {step['description']} ---")
                
                # 设置发布模式
                self.publish_mode_var.set(step['mode'])
                
                # 设置章节范围
                self.start_chapter_var.set(str(step['start_chapter']))
                self.end_chapter_var.set(str(step['end_chapter']))
                
                # 如果是预发布模式，设置每日发布章节数
                if step['mode'] == 'pre-publish':
                    self.daily_publish_num_var.set(str(step['daily_count']))
                    
                    # 如果是第一步预发布，设置从明天开始
                    if i == 2:
                        today = datetime.datetime.now()
                        self.last_published_chapter_date_var.set(today.strftime('%Y-%m-%d'))
                
                # 执行自动化流程
                if not self.automation_flow():
                    raise Exception(f"执行步骤{i}失败: {step['description']}")
            
            print("\n最速开书流程完成！")
            return True
        except Exception as e:
            print(f"最速开书流程中发生错误: {e}")
            traceback.print_exc()
            return False

    def new_novel_publish_once(self):
        print("执行最速开书！")

        novels_folder = self.novels_folder_var.get()
        if novels_folder.endswith('.txt'):
            print("识别为小说文件，正在从文件中分解章节...")
            if not self.create_chapter_files_in_files(novels_folder):
                print("从小说文件中分解章节失败")
                return
            novels_folder = novels_folder.replace('.txt', '')

        all_files = sorted([os.path.join(novels_folder, f) for f in os.listdir(novels_folder) if
                            f.endswith('.md') and f.startswith('Chapter_')])
        fast_publish_mode = self.fast_publish_mode_var.get()
        need_num = eval(fast_publish_mode)
        print(need_num)
        if len(all_files) < need_num:
            print(f"错误：小说章节不足{need_num}章，无法执行最速开书-{fast_publish_mode}。\n"
                  f"当前章节数: {len(all_files)}")
            return
        
        # 解析发布模式并生成发布计划
        plan = self.parse_fast_publish_mode(fast_publish_mode)
        if not plan:
            print("解析发布模式失败")
            return
        # print(plan)
        # 执行发布计划
        self.execute_publish_plan(plan)

if __name__ == "__main__":
    if datetime.datetime.now().year == 2025:
        app = NovelPublisherApp()
        app.mainloop()
    else:
        print('免费使用，为避免他人收费盈利，内测时间已过，不予执行，可联系作者qq:2306995722，进行更新')
        print('懂代码的可以查看代码 https://github.com/zbateau/Tomato-Novel-Auto-Publish-Tool ，自行修改，进行二次开发')
