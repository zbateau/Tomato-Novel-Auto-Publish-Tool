import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
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

CONFIG_FILE = 'config.ini'
AUTH_FILE = 'auth.json'


class NovelPublisherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("番茄小说自动发布工具-内测版")
        self.geometry("800x600")

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
        ttk.Entry(config_frame, textvariable=self.custom_browser_path_var, width=80).grid(row=0, column=1, sticky=tk.EW,
                                                                                          padx=5)

        # Publish Mode
        ttk.Label(config_frame, text="发布模式:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.publish_mode_var = tk.StringVar(value=self.config.get('Settings', 'publish_mode', fallback='publish'))
        ttk.Combobox(config_frame, textvariable=self.publish_mode_var, values=['publish', 'draft', 'pre-publish']).grid(
            row=1, column=1, sticky=tk.W, padx=5)

        # Publish Time
        ttk.Label(config_frame, text="每日定时发布时间:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.publish_time_var = tk.StringVar(value=self.config.get('Settings', 'publish_time', fallback='12:00'))
        ttk.Entry(config_frame, textvariable=self.publish_time_var, width=10).grid(row=2, column=1, sticky=tk.W, padx=5)

        # Daily Publish Num
        ttk.Label(config_frame, text="每日更新章节数:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.daily_publish_num_var = tk.StringVar(value=self.config.get('Settings', 'daily_publish_num', fallback='2'))
        ttk.Entry(config_frame, textvariable=self.daily_publish_num_var, width=10).grid(row=3, column=1, sticky=tk.W,
                                                                                        padx=5)

        # Novel Title
        ttk.Label(config_frame, text="小说标题:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.novel_title_var = tk.StringVar(value=self.config.get('Novel', 'novel_title', fallback=''))
        ttk.Entry(config_frame, textvariable=self.novel_title_var, width=80).grid(row=4, column=1, sticky=tk.EW, padx=5)

        # Novels Folder
        ttk.Label(config_frame, text="小说文件/文件夹:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=2)
        self.novels_folder_var = tk.StringVar(value=self.config.get('Novel', 'novels_folder', fallback=''))
        ttk.Entry(config_frame, textvariable=self.novels_folder_var, width=80).grid(row=5, column=1, sticky=tk.EW,
                                                                                    padx=5)

        # Fast Publish Mode
        ttk.Label(config_frame, text="最速开书发布:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=2)
        self.fast_publish_mode_var = tk.StringVar(value="15*3+2*7")
        ttk.Combobox(config_frame, textvariable=self.fast_publish_mode_var,
                     values=["15*3+2*7", "10*1+15*2+3*7", "10*4+5*7"]).grid(row=6, column=1, sticky=tk.W, padx=5)

        config_frame.columnconfigure(1, weight=1)

        # Save Config Button
        ttk.Button(config_frame, text="保存配置", command=self.save_config).grid(row=6, column=1, sticky=tk.E, padx=5,
                                                                                 pady=5)

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

        # --- Log Section ---
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Redirect stdout to the log widget
        self.redirect_logging()

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
            self.config['Settings'] = {'url': '', 'publish_mode': 'publish'}
            self.config['Novel'] = {'novel_title': '', 'novels_folder': ''}

    def save_config(self):
        try:
            self.config['Settings']['custom_browser_path'] = self.custom_browser_path_var.get().replace('\\', '/')
            self.config['Settings']['publish_mode'] = self.publish_mode_var.get()
            self.config['Settings']['publish_time'] = self.publish_time_var.get()
            self.config['Settings']['daily_publish_num'] = self.daily_publish_num_var.get()
            self.config['Novel']['novel_title'] = self.novel_title_var.get()
            if self.novels_folder_var.get().endswith('.txt'):
                self.create_chapter_files_in_files(self.novels_folder_var.get())
            self.config['Novel']['novels_folder'] = self.novels_folder_var.get().replace('.txt', '')
            self.novels_folder_var.set(self.config['Novel']['novels_folder'])
            with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            messagebox.showinfo("成功", "配置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")

    def run_login_thread(self):
        self.login_event = threading.Event()
        if self.login_page:
            self.login_event.set()
        else:
            threading.Thread(target=self.login, daemon=True).start()

    def login(self):
        site_url = self.config['Settings']['url']
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

    def publish_single_chapter(self, context: BrowserContext, page: Page, chapter_details: tuple):
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

    def automation_flow(self):
        """主自动化流程，负责初始化和循环调用章节发布。"""
        custom_browser_path = self.custom_browser_path_var.get()
        site_url = self.config['Settings']['url']
        novel_title = self.novel_title_var.get()
        novels_folder = self.novels_folder_var.get()
        publish_mode = self.publish_mode_var.get()
        start_chapter = int(self.start_chapter_var.get())
        end_chapter = int(self.end_chapter_var.get())
        keep_browser_open = self.keep_browser_open_var.get()

        if not os.path.exists(AUTH_FILE):
            print(f"错误：找不到登录状态文件 {AUTH_FILE}。请先登录，允许扫码登录。")
            return

        if novels_folder.endswith('.txt'):
            print(f"识别为小说文件，正在从文件中分解章节...")
            if not self.create_chapter_files_in_files(novels_folder):
                raise Exception("从小说文件中分解章节失败")
            else:
                novels_folder = novels_folder.replace('.txt', '')

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
                    if not self.publish_single_chapter(context, page, chapter_details):
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
                print("主自动化流程中发生错误:")
                traceback.print_exc()

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
        if fast_publish_mode == "15*3+2*7":
            # 15*3+2*7 章节字数不足版
            try:
                # 第一步：当日发布1-15章 # 保证字数过2万签约
                print("\n--- 步骤1: 发布第1-15章 ---")
                self.publish_mode_var.set("publish")
                self.start_chapter_var.set("1")
                self.end_chapter_var.set("15")
                if not self.automation_flow():
                    raise Exception("发布1-15章失败")

                # 第二步：预发布16-45章，每日15章
                print("\n--- 步骤2: 预发布第16-45章 (每日15章) ---")
                self.publish_mode_var.set("pre-publish")
                self.start_chapter_var.set("16")
                self.end_chapter_var.set("45")
                self.daily_publish_num_var.set("15")
                # 设置从明天开始
                today = datetime.datetime.now()
                self.last_published_chapter_date_var.set(today.strftime('%Y-%m-%d'))
                if not self.automation_flow():
                    raise Exception("预发布16-45章失败")

                # 第三步：预发布46-60章，每日2章
                print("\n--- 步骤3: 预发布第46-60章 (每日2章) ---")
                self.publish_mode_var.set("pre-publish")
                self.start_chapter_var.set("46")
                self.end_chapter_var.set("60")
                self.daily_publish_num_var.set("2")
                # 从上一个发布任务结束后开始
                if not self.automation_flow():
                    raise Exception("预发布46-60章失败")

                print("\n最速开书流程完成！")

            except Exception as e:
                print(f"最速开书流程中发生错误: {e}")
                traceback.print_exc()

        if fast_publish_mode == "10*1+15*2+3*7":
            # 10+15*2+3*7 验证期日六版
            try:
                # 第一步：当日发布1-10章 # 保证字数过2万签约
                print("\n--- 步骤1: 当日发布第1-10章 ---")
                self.publish_mode_var.set("publish")
                self.start_chapter_var.set("1")
                self.end_chapter_var.set("10")
                if not self.automation_flow():
                    raise Exception("发布1-10章失败")

                # 第二步：预发布11-40章，每日15章
                print("\n--- 步骤2: 预发布第11-40章 (每日15章) ---")
                self.publish_mode_var.set("pre-publish")
                self.start_chapter_var.set("11")
                self.end_chapter_var.set("40")
                self.daily_publish_num_var.set("15")
                # 设置从明天开始
                today = datetime.datetime.now()
                self.last_published_chapter_date_var.set(today.strftime('%Y-%m-%d'))
                if not self.automation_flow():
                    raise Exception("预发布11-40章失败")

                # 第三步：预发布41-61章，每日3章
                print("\n--- 步骤3: 预发布第41-61章 (每日3章) ---")
                self.publish_mode_var.set("pre-publish")
                self.start_chapter_var.set("41")
                self.end_chapter_var.set("61")
                self.daily_publish_num_var.set("3")
                # 从上一个发布任务结束后开始
                if not self.automation_flow():
                    raise Exception("预发布41-61章失败")

                print("\n最速开书流程完成！")

            except Exception as e:
                print(f"最速开书流程中发生错误: {e}")
                traceback.print_exc()

        if fast_publish_mode == "10*4+5*7":
            # 10*4+5*7 验证期日万版，四天推荐
            try:
                # 第一步：当日发布1-10章 # 保证字数过2万签约
                print("\n--- 步骤1: 当日发布第1-10章 ---")
                self.publish_mode_var.set("publish")
                self.start_chapter_var.set("1")
                self.end_chapter_var.set("10")
                if not self.automation_flow():
                    raise Exception("发布1-10章失败")

                # 第二步：预发布11-40章，每日15章
                print("\n--- 步骤2: 预发布第11-40章 (每日10章) ---")
                self.publish_mode_var.set("pre-publish")
                self.start_chapter_var.set("11")
                self.end_chapter_var.set("40")
                self.daily_publish_num_var.set("10")
                # 设置从明天开始
                today = datetime.datetime.now()
                self.last_published_chapter_date_var.set(today.strftime('%Y-%m-%d'))
                if not self.automation_flow():
                    raise Exception("预发布11-40章失败")

                # 第三步：预发布41-61章，每日5章
                print("\n--- 步骤3: 预发布第41-75章 (每日5章) ---")
                self.publish_mode_var.set("pre-publish")
                self.start_chapter_var.set("41")
                self.end_chapter_var.set("75")
                self.daily_publish_num_var.set("5")
                # 从上一个发布任务结束后开始
                if not self.automation_flow():
                    raise Exception("预发布41-75章失败")

                print("\n最速开书流程完成！")

            except Exception as e:
                print(f"最速开书流程中发生错误: {e}")
                traceback.print_exc()


if __name__ == "__main__":
    if datetime.datetime.now().year == 2025:
        app = NovelPublisherApp()
        app.mainloop()
    else:
        print('免费使用，为避免他人收费盈利，内测时间已过，不予执行，可联系作者qq:2306995722，进行更新')
        print('懂代码的可以查看代码 https://github.com/zbateau/Tomato-Novel-Auto-Publish-Tool ，自行修改，进行二次开发')
