import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import os
import subprocess
from datetime import datetime
import yt_dlp
import sys
import glob
import re


class DownloadSignal:
    def __init__(self, gui):
        self.gui = gui

    def progress(self, download_id, percent, speed, eta=None):
        self.gui.update_progress(download_id, percent, speed, eta)

    def finished(self, download_id, success, message):
        self.gui.download_finished(download_id, success, message)

    def log(self, download_id, message):
        self.gui.add_log(download_id, message)


class DownloadWorker(threading.Thread):
    def __init__(self, url, download_dir, quality, signal, download_id, ytdlp_path, ffmpeg_path):
        super().__init__()
        self.url = url
        self.download_dir = download_dir
        self.quality = quality
        self.signal = signal
        self.download_id = download_id
        self.ytdlp_path = ytdlp_path
        self.ffmpeg_path = ffmpeg_path
        self._is_running = True

    def create_subprocess(self, cmd):
        """创建隐藏窗口的子进程"""
        startupinfo = None
        creationflags = 0

        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            startupinfo=startupinfo,
            creationflags=creationflags
        )
        return process

    def parse_ytdlp_progress(self, line):
        """解析 yt-dlp 命令行输出的进度信息"""
        try:
            percent = 0
            speed = 0
            eta = None

            # 匹配进度百分比：例如 [download]  65.3% of ... 或 65.3%
            percent_match = re.search(r'(\d+\.\d+|\d+)%', line)
            if percent_match:
                percent = float(percent_match.group(1))

            # 匹配下载速度：支持多种格式
            # 2.50MiB/s, 1.2MB/s, 500KiB/s, 123.4kB/s
            speed_match = re.search(r'(\d+\.\d+|\d+)\s*([KM]?i?B)/s', line, re.IGNORECASE)
            if speed_match:
                speed_value = float(speed_match.group(1))
                speed_unit = speed_match.group(2).upper()

                # 转换为字节/秒
                if speed_unit == 'B/S':
                    speed = speed_value
                elif speed_unit == 'KB/S' or speed_unit == 'KIB/S':
                    speed = speed_value * 1024
                elif speed_unit == 'MB/S' or speed_unit == 'MIB/S':
                    speed = speed_value * 1024 * 1024
                elif speed_unit == 'GB/S' or speed_unit == 'GIB/S':
                    speed = speed_value * 1024 * 1024 * 1024

            # 更精确的速度匹配（处理没有空格的情况）
            if speed == 0:
                speed_match2 = re.search(r'(\d+\.\d+|\d+)([KM]?i?B/s)', line, re.IGNORECASE)
                if speed_match2:
                    speed_value = float(speed_match2.group(1))
                    speed_unit = speed_match2.group(2).upper()

                    # 转换为字节/秒
                    if speed_unit == 'B/S':
                        speed = speed_value
                    elif speed_unit == 'KB/S' or speed_unit == 'KIB/S':
                        speed = speed_value * 1024
                    elif speed_unit == 'MB/S' or speed_unit == 'MIB/S':
                        speed = speed_value * 1024 * 1024
                    elif speed_unit == 'GB/S' or speed_unit == 'GIB/S':
                        speed = speed_value * 1024 * 1024 * 1024

            # 匹配ETA：例如 ETA 00:12 或 in 00:12 或 00:12
            eta_match = re.search(r'(ETA|in)\s+(\d+:\d+(:\d+)?)', line)
            if eta_match:
                eta = eta_match.group(2)
            else:
                # 尝试匹配简单的 00:12 格式
                eta_match_simple = re.search(r'(?<!\d)(\d+:\d+(:\d+)?)(?!\d)', line)
                if eta_match_simple and 'download' in line.lower():
                    eta = eta_match_simple.group(1)

            # 匹配文件大小和已下载大小：例如 45.7MiB/100.0MiB
            size_match = re.search(r'(\d+\.\d+|\d+)([KM]?i?B)\s*/\s*(\d+\.\d+|\d+)([KM]?i?B)', line, re.IGNORECASE)
            if size_match and percent == 0:
                # 如果百分比为0但能解析到大小，计算百分比
                downloaded = float(size_match.group(1))
                total = float(size_match.group(3))
                downloaded_unit = size_match.group(2).upper()
                total_unit = size_match.group(4).upper()

                # 转换为字节
                def convert_to_bytes(value, unit):
                    unit = unit.upper()
                    if unit == 'B':
                        return value
                    elif unit == 'KB' or unit == 'KIB':
                        return value * 1024
                    elif unit == 'MB' or unit == 'MIB':
                        return value * 1024 * 1024
                    elif unit == 'GB' or unit == 'GIB':
                        return value * 1024 * 1024 * 1024
                    else:
                        return value

                downloaded_bytes = convert_to_bytes(downloaded, downloaded_unit)
                total_bytes = convert_to_bytes(total, total_unit)

                if total_bytes > 0:
                    percent = (downloaded_bytes / total_bytes) * 100

            return percent, speed, eta

        except Exception as e:
            # 如果解析出错，返回默认值
            return 0, 0, None

    def run(self):
        try:
            # 配置 yt-dlp 选项
            ydl_opts = {
                'outtmpl': os.path.join(self.download_dir, '%(title)s.%(ext)s'),
                'noplaylist': True,
            }

            # 添加隐藏窗口的配置（Windows）
            if sys.platform == "win32":
                ydl_opts['external_downloader_args'] = ['--no-progress']

            # 如果指定了 ffmpeg 路径，设置 ffmpeg_location
            if self.ffmpeg_path and os.path.exists(self.ffmpeg_path):
                ydl_opts['ffmpeg_location'] = self.ffmpeg_path
                self.signal.log(self.download_id, f"使用 FFmpeg 路径: {self.ffmpeg_path}")

            # 根据选择的画质设置
            if self.quality == "最佳画质":
                ydl_opts['format'] = 'best'
            elif self.quality == "仅音频":
                ydl_opts['format'] = 'bestaudio/best'
                ydl_opts['postprocessors'] = [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '192',
                }]
            else:
                # 自定义画质 - 使用更灵活的格式选择
                # 对于 720p, 360p 等，使用更智能的格式选择
                if self.quality in ["2160p", "1440p", "1080p", "720p", "480p", "360p"]:
                    # 使用格式选择器来选择指定分辨率的视频
                    resolution = self.quality.replace('p', '')
                    ydl_opts['format'] = f'best[height<={resolution}]/best'
                else:
                    # 用户自定义格式
                    ydl_opts['format'] = self.quality

            # 进度回调函数（用于Python模块方式）
            def progress_hook(d):
                if not self._is_running:
                    raise Exception("下载被用户停止")

                if d['status'] == 'downloading':
                    if 'total_bytes' in d and d['total_bytes']:
                        percent = int(d['downloaded_bytes'] / d['total_bytes'] * 100)
                        self.signal.progress(self.download_id, percent, d.get('speed', 0))
                    elif 'total_bytes_estimate' in d and d['total_bytes_estimate']:
                        percent = int(d['downloaded_bytes'] / d['total_bytes_estimate'] * 100)
                        self.signal.progress(self.download_id, percent, d.get('speed', 0))
                    else:
                        self.signal.progress(self.download_id, 0, d.get('speed', 0))

                    # 记录日志
                    if 'eta' in d and d['eta']:
                        log_msg = f"下载中: {d.get('_percent_str', 'N/A')} - 速度: {d.get('_speed_str', 'N/A')} - ETA: {d.get('_eta_str', 'N/A')}"
                        self.signal.log(self.download_id, log_msg)

                elif d['status'] == 'finished':
                    self.signal.progress(self.download_id, 100, 0)
                    self.signal.log(self.download_id, f"下载完成: {d['filename']}")

            ydl_opts['progress_hooks'] = [progress_hook]

            # 如果指定了自定义 yt-dlp 路径，使用子进程调用
            if self.ytdlp_path and os.path.exists(self.ytdlp_path):
                self.signal.log(self.download_id, f"使用自定义 yt-dlp 路径: {self.ytdlp_path}")
                self.run_with_external_ytdlp(ydl_opts)
            else:
                # 使用 Python 模块
                self.signal.log(self.download_id, "使用内置 yt-dlp 模块")
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    self.signal.log(self.download_id, f"开始下载: {self.url}")
                    ydl.download([self.url])

            if self._is_running:
                self.signal.finished(self.download_id, True, "下载完成")

        except Exception as e:
            if self._is_running:  # 只有非用户停止的错误才报告
                error_msg = str(e)
                self.signal.log(self.download_id, f"错误: {error_msg}")
                self.signal.finished(self.download_id, False, error_msg)

    def run_with_external_ytdlp(self, ydl_opts):
        """使用外部 yt-dlp 可执行文件进行下载"""
        # 构建命令行参数
        cmd = [self.ytdlp_path]

        # 输出模板
        cmd.extend(['-o', os.path.join(self.download_dir, '%(title)s.%(ext)s')])

        # 添加进度显示参数
        cmd.extend(['--newline'])  # 强制每行输出
        cmd.extend(['--progress'])  # 显示进度条
        cmd.extend(['--console-title'])  # 在控制台标题显示进度

        # 画质设置
        if self.quality == "最佳画质":
            cmd.extend(['-f', 'best'])
        elif self.quality == "仅音频":
            cmd.extend(['-f', 'bestaudio', '--extract-audio', '--audio-format', 'mp3'])
        else:
            # 自定义画质 - 使用更灵活的格式选择
            if self.quality in ["2160p", "1440p", "1080p", "720p", "480p", "360p"]:
                # 使用格式选择器来选择指定分辨率的视频
                resolution = self.quality.replace('p', '')
                cmd.extend(['-f', f'best[height<={resolution}]'])
            else:
                # 用户自定义格式
                cmd.extend(['-f', self.quality])

        # 添加 URL
        cmd.append(self.url)

        # 使用隐藏窗口的方式启动进程
        process = self.create_subprocess(cmd)

        # 读取输出并解析进度
        last_progress_update = 0
        for line in process.stdout:
            if not self._is_running:
                process.terminate()
                break

            line = line.strip()
            if line:
                # 记录所有输出到日志（但过滤掉频繁的进度更新以减少日志量）
                should_log = True
                if '[download]' in line and any(x in line for x in ['%', 'ETA']):
                    # 进度行，减少日志频率
                    current_time = datetime.now().timestamp()
                    if current_time - last_progress_update < 2:  # 每2秒记录一次进度
                        should_log = False
                    else:
                        last_progress_update = current_time

                if should_log:
                    self.signal.log(self.download_id, line)

                # 解析进度信息
                if '[download]' in line or 'download' in line.lower():
                    percent, speed, eta = self.parse_ytdlp_progress(line)

                    # 更新进度显示
                    if percent > 0:
                        self.signal.progress(self.download_id, int(percent), speed, eta)

                    # 检测下载完成
                    if '100%' in line or 'already downloaded' in line.lower() or 'has already been downloaded' in line:
                        self.signal.progress(self.download_id, 100, 0)
                        self.signal.log(self.download_id, "下载完成")

                # 检测错误
                elif 'error' in line.lower() or 'failed' in line.lower():
                    self.signal.log(self.download_id, f"下载错误: {line}")

        process.wait()

        # 检查进程退出状态
        if process.returncode == 0:
            self.signal.log(self.download_id, "外部 yt-dlp 进程正常退出")
        else:
            self.signal.log(self.download_id, f"外部 yt-dlp 进程异常退出，代码: {process.returncode}")

    def stop(self):
        self._is_running = False


class VideoDownloaderApp:
    def __init__(self):
        self.download_workers = {}
        self.download_items = {}
        self.setup_gui()
        self.auto_find_ffmpeg()
        self.auto_find_ytdlp()

    def setup_gui(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("YTDLP-Gui 视频下载器")
        self.root.geometry("900x700")

        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 路径设置框架
        path_frame = ttk.LabelFrame(main_frame, text="路径设置", padding="5")
        path_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        path_frame.columnconfigure(1, weight=1)

        # yt-dlp 路径
        ttk.Label(path_frame, text="yt-dlp 路径:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ytdlp_path = tk.StringVar()
        ytdlp_entry = ttk.Entry(path_frame, textvariable=self.ytdlp_path)
        ytdlp_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 5))
        ttk.Button(path_frame, text="浏览", command=self.browse_ytdlp_path).grid(row=0, column=2, pady=2)

        # FFmpeg 路径
        ttk.Label(path_frame, text="FFmpeg 路径:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ffmpeg_path = tk.StringVar()
        ffmpeg_entry = ttk.Entry(path_frame, textvariable=self.ffmpeg_path)
        ffmpeg_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 5))
        ttk.Button(path_frame, text="浏览", command=self.browse_ffmpeg_path).grid(row=1, column=2, pady=2)

        # 测试按钮
        ttk.Button(path_frame, text="测试路径", command=self.test_paths).grid(row=2, column=1, pady=5)

        # 下载设置框架
        download_frame = ttk.LabelFrame(main_frame, text="下载设置", padding="5")
        download_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        download_frame.columnconfigure(1, weight=1)

        # 下载目录
        ttk.Label(download_frame, text="下载目录:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.download_dir = tk.StringVar(value=os.path.expanduser("~/Downloads"))
        dir_entry = ttk.Entry(download_frame, textvariable=self.download_dir)
        dir_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 5))
        ttk.Button(download_frame, text="浏览", command=self.browse_directory).grid(row=0, column=2, pady=2)

        # 画质选择框架
        quality_frame = ttk.Frame(download_frame)
        quality_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        quality_frame.columnconfigure(1, weight=1)

        ttk.Label(quality_frame, text="画质:").grid(row=0, column=0, sticky=tk.W, pady=2)

        # 画质选择组合框
        self.quality = tk.StringVar(value="最佳画质")
        quality_combo = ttk.Combobox(quality_frame, textvariable=self.quality,
                                     values=["最佳画质", "仅音频", "2160p", "1440p", "1080p", "720p", "480p", "360p", "自定义"])
        quality_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 5))
        quality_combo.bind('<<ComboboxSelected>>', self.on_quality_selected)

        # 自定义画质输入框
        ttk.Label(quality_frame, text="自定义:").grid(row=0, column=2, sticky=tk.W, pady=2, padx=(5, 0))
        self.custom_quality = tk.StringVar()
        self.custom_quality_entry = ttk.Entry(quality_frame, textvariable=self.custom_quality, width=15)
        self.custom_quality_entry.grid(row=0, column=3, sticky=tk.W, pady=2, padx=(5, 0))
        self.custom_quality_entry.config(state='disabled')

        # 自定义画质说明标签
        help_label = ttk.Label(quality_frame, text="格式示例: bestvideo+bestaudio, 137+140 等",
                               foreground="gray", font=("Arial", 8))
        help_label.grid(row=1, column=1, columnspan=3, sticky=tk.W, pady=(0, 2))

        # URL 输入框架
        url_frame = ttk.Frame(main_frame)
        url_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        url_frame.columnconfigure(0, weight=1)

        ttk.Label(url_frame, text="视频URL (支持多地址，每行一个):").grid(row=0, column=0, sticky=tk.W, pady=2)

        # 使用 ScrolledText 替代 Entry 来支持多行输入
        self.url_text = scrolledtext.ScrolledText(url_frame, height=4, wrap=tk.WORD)
        self.url_text.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)

        # 绑定右键菜单
        self.setup_url_context_menu()

        ttk.Button(url_frame, text="添加下载", command=self.add_download).grid(row=1, column=1, pady=2, padx=(5, 0))

        # 下载列表框架
        list_frame = ttk.LabelFrame(main_frame, text="下载队列", padding="5")
        list_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # 创建树形视图显示下载队列
        columns = ("url", "progress", "status", "actions")
        self.download_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)

        # 设置列
        self.download_tree.heading("url", text="URL")
        self.download_tree.heading("progress", text="进度")
        self.download_tree.heading("status", text="状态")
        self.download_tree.heading("actions", text="操作")

        self.download_tree.column("url", width=300)
        self.download_tree.column("progress", width=150)
        self.download_tree.column("status", width=100)
        self.download_tree.column("actions", width=150)

        self.download_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 添加滚动条
        tree_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.download_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.download_tree.configure(yscrollcommand=tree_scroll.set)

        # 控制按钮框架
        control_frame = ttk.Frame(list_frame)
        control_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))

        ttk.Button(control_frame, text="开始全部", command=self.start_all_downloads).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="暂停全部", command=self.pause_all_downloads).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="清除已完成", command=self.clear_completed).pack(side=tk.LEFT)

        # 日志框架
        log_frame = ttk.LabelFrame(main_frame, text="下载日志", padding="5")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 创建信号对象
        self.download_signal = DownloadSignal(self)

    def on_quality_selected(self, event):
        """画质选择变化事件"""
        selected = self.quality.get()
        if selected == "自定义":
            self.custom_quality_entry.config(state='normal')
            self.custom_quality_entry.focus()
        else:
            self.custom_quality_entry.config(state='disabled')
            self.custom_quality.set("")

    def get_selected_quality(self):
        """获取选择的画质"""
        selected = self.quality.get()
        if selected == "自定义":
            return self.custom_quality.get().strip()
        return selected

    def setup_url_context_menu(self):
        """设置URL输入框的右键菜单"""
        self.url_context_menu = tk.Menu(self.url_text, tearoff=0)
        self.url_context_menu.add_command(label="粘贴", command=self.paste_to_url)
        self.url_context_menu.add_command(label="剪切", command=self.cut_from_url)
        self.url_context_menu.add_command(label="复制", command=self.copy_from_url)
        self.url_context_menu.add_separator()
        self.url_context_menu.add_command(label="全选", command=self.select_all_url)

        # 绑定右键事件
        self.url_text.bind("<Button-3>", self.show_url_context_menu)

    def show_url_context_menu(self, event):
        """显示URL输入框的右键菜单"""
        try:
            self.url_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.url_context_menu.grab_release()

    def paste_to_url(self):
        """粘贴到URL输入框"""
        try:
            clipboard_content = self.root.clipboard_get()
            self.url_text.insert(tk.INSERT, clipboard_content)
        except tk.TclError:
            pass

    def cut_from_url(self):
        """从URL输入框剪切"""
        try:
            selected_text = self.url_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.url_text.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass

    def copy_from_url(self):
        """从URL输入框复制"""
        try:
            selected_text = self.url_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
        except tk.TclError:
            pass

    def select_all_url(self):
        """全选URL输入框内容"""
        self.url_text.tag_add(tk.SEL, "1.0", tk.END)
        self.url_text.mark_set(tk.INSERT, "1.0")
        self.url_text.see(tk.INSERT)

    def auto_find_ffmpeg(self):
        """自动查找FFmpeg可执行文件"""
        if sys.platform == "win32":
            # Windows平台下的常见FFmpeg安装位置
            possible_paths = [
                # 当前目录
                "ffmpeg.exe",
                # 程序所在目录
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "ffmpeg.exe"),
                # PATH环境变量
                *[os.path.join(path, "ffmpeg.exe") for path in os.environ.get("PATH", "").split(os.pathsep) if path],
                # 常见安装目录
                "C:\\ffmpeg\\bin\\ffmpeg.exe",
                "C:\\Program Files\\ffmpeg\\bin\\ffmpeg.exe",
                "C:\\Program Files (x86)\\ffmpeg\\bin\\ffmpeg.exe",
                os.path.expanduser("~\\ffmpeg\\bin\\ffmpeg.exe"),
            ]

            # 递归搜索当前目录和子目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            for root, dirs, files in os.walk(current_dir):
                if "ffmpeg.exe" in files:
                    possible_paths.append(os.path.join(root, "ffmpeg.exe"))

            # 检查每个可能的路径
            for path in possible_paths:
                if os.path.exists(path):
                    try:
                        # 验证是否是有效的FFmpeg可执行文件
                        result = subprocess.run([path, '-version'],
                                                capture_output=True, text=True, timeout=5)
                        if result.returncode == 0 and "ffmpeg version" in result.stdout:
                            self.ffmpeg_path.set(path)
                            self.add_log("system", f"自动找到 FFmpeg: {path}")
                            return
                    except (subprocess.SubprocessError, OSError):
                        continue

            self.add_log("system", "未找到 FFmpeg，请手动设置路径")

    def auto_find_ytdlp(self):
        """自动查找 yt-dlp 可执行文件"""
        if sys.platform == "win32":
            # Windows平台下的常见yt-dlp安装位置
            possible_paths = [
                # 当前目录
                "yt-dlp.exe",
                "yt-dlp",
                # 程序所在目录
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "yt-dlp.exe"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "yt-dlp"),
                # PATH环境变量
                *[os.path.join(path, "yt-dlp.exe") for path in os.environ.get("PATH", "").split(os.pathsep) if path],
                *[os.path.join(path, "yt-dlp") for path in os.environ.get("PATH", "").split(os.pathsep) if path],
                # 常见安装目录
                "C:\\yt-dlp\\yt-dlp.exe",
                os.path.expanduser("~\\yt-dlp\\yt-dlp.exe"),
            ]

            # 递归搜索当前目录和子目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            for root, dirs, files in os.walk(current_dir):
                if "yt-dlp.exe" in files:
                    possible_paths.append(os.path.join(root, "yt-dlp.exe"))
                if "yt-dlp" in files and not root.endswith('.exe'):
                    possible_paths.append(os.path.join(root, "yt-dlp"))

            # 检查每个可能的路径
            for path in possible_paths:
                if os.path.exists(path):
                    try:
                        # 验证是否是有效的yt-dlp可执行文件
                        startupinfo = None
                        creationflags = 0
                        if sys.platform == "win32":
                            startupinfo = subprocess.STARTUPINFO()
                            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                            startupinfo.wShowWindow = 0
                            creationflags = subprocess.CREATE_NO_WINDOW

                        result = subprocess.run([path, '--version'],
                                                capture_output=True, text=True, timeout=5,
                                                startupinfo=startupinfo, creationflags=creationflags)
                        if result.returncode == 0:
                            self.ytdlp_path.set(path)
                            self.add_log("system", f"自动找到 yt-dlp: {path}")
                            return
                    except (subprocess.SubprocessError, OSError):
                        continue

            self.add_log("system", "未找到 yt-dlp，将使用内置模块")

    def browse_ytdlp_path(self):
        file_path = filedialog.askopenfilename(
            title="选择 yt-dlp 可执行文件",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if file_path:
            self.ytdlp_path.set(file_path)

    def browse_ffmpeg_path(self):
        file_path = filedialog.askopenfilename(
            title="选择 FFmpeg 可执行文件",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if file_path:
            self.ffmpeg_path.set(file_path)

    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.download_dir.get())
        if directory:
            self.download_dir.set(directory)

    def test_paths(self):
        """测试配置的路径是否有效"""
        ytdlp_path = self.ytdlp_path.get().strip()
        ffmpeg_path = self.ffmpeg_path.get().strip()

        messages = []

        def run_hidden_process(cmd, timeout=10):
            """运行隐藏窗口的进程"""
            startupinfo = None
            creationflags = 0
            if sys.platform == "win32":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0
                creationflags = subprocess.CREATE_NO_WINDOW

            return subprocess.run(cmd,
                                  capture_output=True,
                                  text=True,
                                  timeout=timeout,
                                  startupinfo=startupinfo,
                                  creationflags=creationflags)

        # 测试 yt-dlp 路径
        if ytdlp_path:
            if os.path.exists(ytdlp_path):
                try:
                    result = run_hidden_process([ytdlp_path, '--version'])
                    if result.returncode == 0:
                        messages.append(f"✓ yt-dlp 路径有效: {ytdlp_path}")
                        messages.append(f"  yt-dlp 版本: {result.stdout.strip()}")
                    else:
                        messages.append(f"✗ yt-dlp 路径无效: {result.stderr}")
                except Exception as e:
                    messages.append(f"✗ yt-dlp 路径测试失败: {str(e)}")
            else:
                messages.append("✗ yt-dlp 路径不存在")
        else:
            messages.append("ℹ 使用内置 yt-dlp 模块")

        # 测试 FFmpeg 路径
        if ffmpeg_path:
            if os.path.exists(ffmpeg_path):
                try:
                    result = run_hidden_process([ffmpeg_path, '-version'])
                    if result.returncode == 0:
                        messages.append(f"✓ FFmpeg 路径有效: {ffmpeg_path}")
                        version_line = result.stdout.split('\n')[0] if result.stdout else "未知版本"
                        messages.append(f"  FFmpeg 版本: {version_line}")
                    else:
                        messages.append(f"✗ FFmpeg 路径无效: {result.stderr}")
                except Exception as e:
                    messages.append(f"✗ FFmpeg 路径测试失败: {str(e)}")
            else:
                messages.append("✗ FFmpeg 路径不存在")
        else:
            messages.append("ℹ 使用系统 PATH 中的 FFmpeg")

        # 显示测试结果
        messagebox.showinfo("路径测试结果", "\n".join(messages))

    def add_download(self):
        # 获取多行URL输入
        url_content = self.url_text.get("1.0", tk.END).strip()
        if not url_content:
            messagebox.showwarning("输入错误", "请输入视频URL")
            return

        # 按行分割URL
        urls = [url.strip() for url in url_content.split('\n') if url.strip()]

        # 获取选择的画质
        quality = self.get_selected_quality()
        if not quality:
            messagebox.showwarning("输入错误", "请选择或输入画质设置")
            return

        added_count = 0
        for url in urls:
            # 生成下载ID
            download_id = f"download_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{added_count}"

            # 添加到树形视图
            item_id = self.download_tree.insert("", tk.END, values=(
                url[:50] + "..." if len(url) > 50 else url,
                "0%",
                "等待中",
                "开始 停止"
            ))

            # 保存下载项信息
            self.download_items[download_id] = {
                'url': url,
                'item_id': item_id,
                'status': 'pending',
                'progress': 0,
                'quality': quality
            }

            # 添加日志
            self.add_log(download_id, f"已添加到下载队列: {url} (画质: {quality})")
            added_count += 1

        # 清空输入框
        self.url_text.delete("1.0", tk.END)

        # 显示添加结果
        self.add_log("system", f"成功添加 {added_count} 个下载任务")

    def start_download(self, download_id):
        if download_id not in self.download_items:
            return

        item_info = self.download_items[download_id]

        # 更新状态
        item_info['status'] = 'downloading'
        self.update_tree_item(download_id, status="下载中")

        # 创建下载线程
        worker = DownloadWorker(
            item_info['url'],
            self.download_dir.get(),
            item_info['quality'],  # 使用保存的画质设置
            self.download_signal,
            download_id,
            self.ytdlp_path.get().strip(),
            self.ffmpeg_path.get().strip()
        )

        self.download_workers[download_id] = worker
        worker.start()

    def stop_download(self, download_id):
        if download_id in self.download_workers:
            worker = self.download_workers[download_id]
            worker.stop()

        if download_id in self.download_items:
            item_info = self.download_items[download_id]
            item_info['status'] = 'stopped'
            self.update_tree_item(download_id, status="已停止")

    def start_all_downloads(self):
        for download_id in self.download_items:
            if self.download_items[download_id]['status'] == 'pending':
                self.start_download(download_id)

    def pause_all_downloads(self):
        for download_id in self.download_items:
            if self.download_items[download_id]['status'] == 'downloading':
                self.stop_download(download_id)

    def clear_completed(self):
        items_to_remove = []

        # 遍历所有下载项，找出已完成、错误或已停止的项
        for download_id, item_info in self.download_items.items():
            if item_info['status'] in ['completed', 'error', 'stopped']:
                # 从树形视图中删除
                try:
                    self.download_tree.delete(item_info['item_id'])
                except tk.TclError:
                    # 如果项已经被删除，忽略错误
                    pass
                items_to_remove.append(download_id)

        # 从数据结构中移除
        for download_id in items_to_remove:
            if download_id in self.download_workers:
                del self.download_workers[download_id]
            if download_id in self.download_items:
                del self.download_items[download_id]

        # 添加日志记录
        if items_to_remove:
            self.add_log("system", f"已清除 {len(items_to_remove)} 个已完成项目")
        else:
            self.add_log("system", "没有可清除的已完成项目")

    def update_tree_item(self, download_id, progress=None, status=None):
        if download_id not in self.download_items:
            return

        item_info = self.download_items[download_id]
        item_id = item_info['item_id']

        current_values = list(self.download_tree.item(item_id, 'values'))

        if progress is not None:
            current_values[1] = f"{progress}%"
            item_info['progress'] = progress

        if status is not None:
            current_values[2] = status
            item_info['status'] = status

        self.download_tree.item(item_id, values=current_values)

    def update_progress(self, download_id, percent, speed, eta=None):
        if download_id in self.download_items:
            speed_str = self.format_speed(speed)
            if eta:
                status = f"下载中 - {speed_str} - ETA: {eta}"
            else:
                status = f"下载中 - {speed_str}" if speed > 0 else "下载中"
            self.update_tree_item(download_id, progress=percent, status=status)

    def download_finished(self, download_id, success, message):
        if download_id in self.download_items:
            if success:
                self.update_tree_item(download_id, progress=100, status="已完成")
                self.download_items[download_id]['status'] = 'completed'
            else:
                self.update_tree_item(download_id, status=f"错误: {message}")
                self.download_items[download_id]['status'] = 'error'

            if download_id in self.download_workers:
                del self.download_workers[download_id]

    def add_log(self, download_id, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.log_text.insert(tk.END, log_entry + "\n")
        self.log_text.see(tk.END)

    def format_speed(self, speed_bytes):
        """格式化下载速度显示"""
        if speed_bytes == 0:
            return "0 B/s"
        elif speed_bytes >= 1024 * 1024 * 1024:
            return f"{speed_bytes / (1024 * 1024 * 1024):.2f} GB/s"
        elif speed_bytes >= 1024 * 1024:
            return f"{speed_bytes / (1024 * 1024):.2f} MB/s"
        elif speed_bytes >= 1024:
            return f"{speed_bytes / 1024:.2f} KB/s"
        else:
            return f"{speed_bytes:.0f} B/s"

    def run(self):
        self.root.mainloop()


def main():
    # 检查必要的依赖
    try:
        import yt_dlp
    except ImportError:
        print("错误: 未安装 yt-dlp")
        print("请运行: pip install yt-dlp")
        return

    app = VideoDownloaderApp()
    app.run()


if __name__ == "__main__":
    main()