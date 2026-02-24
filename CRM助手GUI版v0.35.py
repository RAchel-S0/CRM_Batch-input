import time
import random
import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================= 配置区域 =================

CONFIG_FILENAME = "reply_list.txt"
USER_DATA_DIR_NAME = "ChromeUserData"

# 默认话术
DEFAULT_POOL = [
    "推荐用户使用诺企服开票，用户已知晓",
    "客户反馈系统稳定，暂无新增需求",
    "电话无人接听，稍后再次联系",
    "已告知客户最新的优惠活动",
    "常规回访，客户表示满意"
]

SELECTORS = {
    "list_row": "div[data-fieldname='name']",
    "detail_link": "a[href*='AccountObj']",
    "last_record": ".fxeditor-render-text",
    "publish_btn": ".d-salelog_publish_btn",
    "editor": ".tiptap.ProseMirror",
    "select_input": "input.j-select-input",
    "submit_real_btn": "span.j-ok[action-type='dialogEnter']"
}


# ================= 业务逻辑类 =================

class LogicHandler:
    def __init__(self, log_callback):
        self.log = log_callback
        self.driver = None
        self.is_running = False
        self.reply_pool = []

    def load_reply_pool(self):
        """加载话术库 (增加GUI日志)"""
        if getattr(sys, 'frozen', False):
            app_path = os.path.dirname(sys.executable)
        else:
            app_path = os.path.dirname(os.path.abspath(__file__))

        config_path = os.path.join(app_path, CONFIG_FILENAME)

        if not os.path.exists(config_path):
            self.log(f"未找到配置文件，生成默认文件: {CONFIG_FILENAME}")
            try:
                with open(config_path, 'w', encoding='utf-8') as f:
                    for line in DEFAULT_POOL: f.write(line + "\n")
                return DEFAULT_POOL
            except Exception as e:
                self.log(f"生成失败: {e}")
                return DEFAULT_POOL

        try:
            custom_pool = []
            with open(config_path, 'r', encoding='utf-8') as f:
                for line in f:
                    txt = line.strip()
                    if txt: custom_pool.append(txt)

            if custom_pool:
                self.log(f"✅ 成功加载话术库: {len(custom_pool)} 条")
                return custom_pool
            else:
                self.log("⚠️ 配置文件为空，使用默认值")
                return DEFAULT_POOL
        except Exception as e:
            self.log(f"❌ 读取配置出错: {e}")
            return DEFAULT_POOL

    def get_version_info(self, chrome_path, driver_path):
        """Win7 兼容的版本检测"""
        c_ver, d_ver = "未检测到", "未检测到"
        if os.path.exists(chrome_path):
            try:
                escaped_path = chrome_path.replace("\\", "\\\\")
                cmd = f'wmic datafile where name="{escaped_path}" get Version /value'
                result = subprocess.check_output(cmd, shell=True).decode().strip()
                if "Version=" in result:
                    parts = result.split("Version=")
                    if len(parts) > 1: c_ver = parts[1].split()[0]
            except:
                c_ver = "存在 (读取失败)"

        if os.path.exists(driver_path):
            try:
                result = subprocess.check_output([driver_path, "--version"]).decode().strip()
                if " " in result: d_ver = result.split(" ")[1]
            except:
                d_ver = "存在 (读取失败)"
        return c_ver, d_ver

    def start_browser(self, chrome_path, driver_path):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--ignore-certificate-errors")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        # 用户数据目录 (保持登录状态)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        user_data_dir = os.path.join(base_path, USER_DATA_DIR_NAME)
        if not os.path.exists(user_data_dir): os.makedirs(user_data_dir)
        options.add_argument(f"--user-data-dir={user_data_dir}")

        if chrome_path and os.path.exists(chrome_path):
            options.binary_location = chrome_path
        else:
            self.log("错误：Chrome 路径不存在")
            return None

        service = Service(driver_path)
        try:
            driver = webdriver.Chrome(service=service, options=options)
            return driver
        except Exception as e:
            self.log(f"启动浏览器失败: {e}")
            self.log("提示：请先关闭所有已打开的 Chrome 窗口！")
            return None

    def process_logic(self, params):
        chrome_path = params['chrome_path']
        driver_path = params['driver_path']
        target_url = params['url']
        limit = params['limit']
        t_min = params.get('min_wait', 10)
        t_max = params.get('max_wait', 20)

        # 重新加载话术
        self.log("-" * 30)
        self.reply_pool = self.load_reply_pool()
        self.log("-" * 30)

        self.driver = self.start_browser(chrome_path, driver_path)
        if not self.driver: return

        try:
            self.log("正在打开浏览器...")
            self.driver.get(target_url)

            # 交互式引导
            if not messagebox.askokcancel("准备就绪",
                                          "请确认：\n1. 已登录 CRM\n2. 已进入[客户-未回访]列表页\n\n点击【确定】开始任务。"):
                self.driver.quit()
                return

            self.is_running = True
            success_count = 0

            while success_count < limit and self.is_running:
                list_window_handle = self.driver.current_window_handle

                self.log("正在扫描当前页...")
                time.sleep(2)
                rows = self.driver.find_elements(By.CSS_SELECTOR, SELECTORS["list_row"])

                page_tasks = []
                for row in rows:
                    if not row.is_displayed(): continue
                    try:
                        links = row.find_elements(By.TAG_NAME, "a")
                        t_url = ""
                        name = "未知客户"
                        for a in links:
                            href = a.get_attribute("href")
                            if a.text.strip(): name = a.text.strip()
                            if href and "AccountObj" in href:
                                t_url = href
                        if t_url:
                            if not t_url.startswith("http"):
                                t_url = "https://www.fxiaoke.com" + t_url
                            page_tasks.append({'name': name, 'url': t_url})
                    except:
                        continue

                self.log(f"当前页提取到 {len(page_tasks)} 个任务。")

                if not page_tasks:
                    self.log(">>> 请手动翻页，等待加载完成后点击弹窗确认...")
                    top = tk.Toplevel()
                    top.attributes('-topmost', True)
                    top.withdraw()
                    if not messagebox.askyesno("翻页提示", "当前页无数据。\n请手动翻页后，点击【是】继续，【否】退出。",
                                               parent=top):
                        top.destroy()
                        break
                    top.destroy()
                    continue

                # === 处理流程 ===
                self.log(">>> 开始处理 (您可以最小化浏览器)...")
                self.driver.switch_to.new_window('tab')

                for i, task in enumerate(page_tasks):
                    if success_count >= limit or not self.is_running: break

                    name = task['name']
                    url = task['url']
                    self.log(f"[{success_count + 1}/{limit}] 处理: {name}")

                    try:
                        # 使用白屏过渡
                        self.driver.get("about:blank")
                        self.driver.get(url)

                        if not self.wait_for_page_load():
                            self.log("   -> [跳过] 页面加载超时")
                            continue

                        # 执行录入逻辑
                        res = self.process_detail_page()

                        if res:
                            self.log(f"   -> [成功]")
                            success_count += 1
                        else:
                            self.log(f"   -> [失败] 录入流程未完成")

                        delay = random.uniform(t_min, t_max)
                        self.log(f"   -> [休息] {delay:.1f} 秒...")
                        time.sleep(delay)

                    except Exception as e:
                        self.log(f"   -> [异常] {e}")

                self.driver.close()
                self.driver.switch_to.window(list_window_handle)

                if success_count >= limit: break
                if not self.is_running: break

                # 翻页提示
                top = tk.Toplevel()
                top.attributes('-topmost', True)
                top.withdraw()
                if not messagebox.askyesno("翻页提示", f"进度: {success_count}/{limit}\n请手动翻页，完成后点击【是】继续。",
                                           parent=top):
                    top.destroy()
                    break
                top.destroy()

            messagebox.showinfo("完成", f"任务结束，共录入 {success_count} 条。")

        except Exception as e:
            self.log(f"程序错误: {e}")
        finally:
            self.is_running = False

    def wait_for_page_load(self):
        wait = WebDriverWait(self.driver, 15)
        try:
            wait.until(EC.url_contains("AccountObj"))
            # 增加等待时间，确保DOM结构完全渲染，这对读取last_record至关重要
            time.sleep(3)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, SELECTORS["publish_btn"])))
            return True
        except:
            return False

    def process_detail_page(self):

        wait = WebDriverWait(self.driver, 15)

        # === 1. 读取上一条  ===
        last_content = ""
        try:
            time.sleep(2)
            elements = self.driver.find_elements(By.CSS_SELECTOR, SELECTORS["last_record"])
            for el in elements:
                if el.text.strip():
                    last_content = el.text.strip()
                    break

            if last_content:
                self.log(f"   -> 历史: {last_content[:10]}...")
            else:
                self.log("   -> 无历史记录")
        except:
            self.log("   -> 读取历史记录异常")

        # === 2. 准备内容 ===
        final_content = ""
        is_standard = False
        clean_last = last_content.strip()

        for pool_txt in self.reply_pool:
            # 双向包含检查
            if pool_txt in clean_last or clean_last in pool_txt:
                is_standard = True
                break

        if not clean_last or is_standard:
            # 随机选一条
            candidates = [t for t in self.reply_pool if t != clean_last]
            if not candidates: candidates = self.reply_pool
            final_content = random.choice(candidates)
            self.log("   -> 策略: 随机库内容")
        else:
            # 复制历史
            final_content = clean_last
            self.log("   -> 策略: 复制上一条")

        # === 3. 点击按钮 ===
        try:
            publish_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SELECTORS["publish_btn"])))
            self.driver.execute_script("arguments[0].click();", publish_btn)
            time.sleep(2)  # 增加一点弹窗等待
        except:
            self.log("   -> 错误: 无法点击写跟进")
            return False

        # === 4. 填写文字 ===
        try:
            editor = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, SELECTORS["editor"])))
            editor.click()
            editor.send_keys(final_content)
            time.sleep(1)
        except:
            self.log("   -> 错误: 找不到输入框")
            return False

        # === 5. 下拉框选择  ===
        try:
            all_selects = self.driver.find_elements(By.CSS_SELECTOR, SELECTORS["select_input"])
            target_input = None

            if len(all_selects) >= 2:
                target_input = all_selects[1]
            elif len(all_selects) == 1:
                target_input = all_selects[0]

            if target_input:
                self.driver.execute_script("arguments[0].click();", target_input)
                time.sleep(1)
                target_input.send_keys("客户反馈")
                time.sleep(1.5)
                target_input.send_keys(Keys.ARROW_DOWN)
                time.sleep(0.5)
                target_input.send_keys(Keys.ENTER)
                time.sleep(1)
            else:
                self.log("   -> 警告: 未找到下拉框")
        except Exception as e:
            self.log(f"   -> 警告: 下拉框异常 {e}")

        # === 6. 提交 ===
        try:
            submit = self.driver.find_element(By.CSS_SELECTOR, SELECTORS["submit_real_btn"])
            self.driver.execute_script("arguments[0].click();", submit)

            for _ in range(8):
                time.sleep(1)
                try:
                    if not submit.is_displayed(): return True
                except:
                    return True

            self.log("   -> 失败: 弹窗未关闭")
            return False
        except Exception as e:
            self.log(f"   -> 提交出错: {e}")
            return False


# ================= GUI 界面类 =================

class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CRM 自动化助手v2.35 bySTL")
        self.root.geometry("600x680")

        self.logic = LogicHandler(self.log_message)

        # 默认变量
        self.chrome_path_var = tk.StringVar()
        self.url_var = tk.StringVar(
            value="https://www.fxiaoke.com/proj/page/login?returnUrl=https%3A%2F%2Fwww.fxiaoke.com%2FXV%2FUI%2FHome#paasapp/list/=/appId_CRM/AccountObj")
        self.count_var = tk.StringVar(value="50")
        self.min_wait_var = tk.StringVar(value="10")
        self.max_wait_var = tk.StringVar(value="20")
        self.chrome_ver_var = tk.StringVar(value="检测中...")
        self.driver_ver_var = tk.StringVar(value="检测中...")

        self.init_ui()
        self.auto_detect_paths()

    def init_ui(self):
        # 配置区
        frame_config = ttk.LabelFrame(self.root, text="运行配置", padding=10)
        frame_config.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_config, text="Chrome路径:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_config, textvariable=self.chrome_path_var, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(frame_config, text="...", width=3, command=self.browse_chrome).grid(row=0, column=2)

        ttk.Label(frame_config, text="登录网址:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame_config, textvariable=self.url_var, width=40).grid(row=1, column=1, padx=5)

        ttk.Label(frame_config, text="录入数量:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frame_config, textvariable=self.count_var, width=10).grid(row=2, column=1, sticky="w", padx=5)

        ttk.Label(frame_config, text="随机等待(秒):").grid(row=3, column=0, sticky="w")
        frame_time = ttk.Frame(frame_config)
        frame_time.grid(row=3, column=1, sticky="w", padx=5)
        ttk.Entry(frame_time, textvariable=self.min_wait_var, width=5).pack(side="left")
        ttk.Label(frame_time, text="-").pack(side="left", padx=5)
        ttk.Entry(frame_time, textvariable=self.max_wait_var, width=5).pack(side="left")
        ttk.Label(frame_time, text="(范围)").pack(side="left", padx=5)

        # 状态区
        frame_info = ttk.Frame(self.root, padding=5)
        frame_info.pack(fill="x", padx=10)
        ttk.Label(frame_info, text="Chrome: ").pack(side="left")
        ttk.Label(frame_info, textvariable=self.chrome_ver_var, foreground="blue").pack(side="left")
        ttk.Label(frame_info, text=" | Driver: ").pack(side="left")
        ttk.Label(frame_info, textvariable=self.driver_ver_var, foreground="green").pack(side="left")
        ttk.Button(frame_info, text="刷新版本", command=self.check_versions).pack(side="right")

        # 按钮区
        frame_action = ttk.Frame(self.root, padding=10)
        frame_action.pack(fill="x", padx=10)

        self.btn_start = ttk.Button(frame_action, text="开始运行", command=self.start_thread)
        self.btn_start.pack(side="left", fill="x", expand=True, padx=5)

        self.btn_stop = ttk.Button(frame_action, text="停止", command=self.stop_task, state="disabled")
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=5)

        # 日志区
        frame_log = ttk.LabelFrame(self.root, text="日志", padding=10)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_area = scrolledtext.ScrolledText(frame_log, height=15, state='disabled')
        self.log_area.pack(fill="both", expand=True)

    def log_message(self, msg):
        def _log():
            timestamp = time.strftime("%H:%M:%S", time.localtime())
            self.log_area.configure(state='normal')
            self.log_area.insert(tk.END, f"[{timestamp}] {msg}\n")
            self.log_area.see(tk.END)
            self.log_area.configure(state='disabled')

        self.root.after(0, _log)

    def auto_detect_paths(self):
        paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe")
        ]
        base_path = os.path.dirname(os.path.abspath(__file__))
        if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)

        saved = ""
        try:
            with open(os.path.join(base_path, "chrome_path.txt"), 'r') as f:
                saved = f.read().strip().replace('"', '')
        except:
            pass

        if saved and os.path.exists(saved):
            self.chrome_path_var.set(saved)
        else:
            for p in paths:
                if os.path.exists(p):
                    self.chrome_path_var.set(p)
                    break
        self.check_versions()

    def check_versions(self):
        base_path = os.path.dirname(os.path.abspath(__file__))
        if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
        d_path = os.path.join(base_path, "chromedriver.exe")
        c_ver, d_ver = self.logic.get_version_info(self.chrome_path_var.get(), d_path)
        self.chrome_ver_var.set(c_ver)
        self.driver_ver_var.set(d_ver)

    def browse_chrome(self):
        filename = filedialog.askopenfilename(filetypes=[("Exe", "*.exe")])
        if filename:
            self.chrome_path_var.set(filename)
            self.check_versions()
            try:
                base_path = os.path.dirname(os.path.abspath(__file__))
                if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
                with open(os.path.join(base_path, "chrome_path.txt"), "w") as f:
                    f.write(filename)
            except:
                pass

    def start_thread(self):
        if not self.chrome_path_var.get():
            messagebox.showerror("错误", "请指定 Chrome 路径")
            return

        try:
            limit = int(self.count_var.get())
        except:
            limit = 50

        try:
            t_min = float(self.min_wait_var.get())
            t_max = float(self.max_wait_var.get())
            if t_min < 0 or t_max < 0: raise ValueError
        except:
            messagebox.showerror("错误", "随机等待时间必须是大于0的数字")
            return

        base_path = os.path.dirname(os.path.abspath(__file__))
        if getattr(sys, 'frozen', False): base_path = os.path.dirname(sys.executable)
        driver_path = os.path.join(base_path, "chromedriver.exe")

        params = {
            'chrome_path': self.chrome_path_var.get(),
            'driver_path': driver_path,
            'url': self.url_var.get(),
            'limit': limit,
            'min_wait': t_min,
            'max_wait': t_max
        }

        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")

        t = threading.Thread(target=self.run_logic_wrapper, args=(params,))
        t.daemon = True
        t.start()

    def run_logic_wrapper(self, params):
        self.logic.process_logic(params)
        self.root.after(0, lambda: self.btn_start.config(state="normal"))
        self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

    def stop_task(self):
        self.logic.is_running = False
        self.log_message("正在停止... (请等待当前操作完成)")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()