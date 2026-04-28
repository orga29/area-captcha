import os
import time
import threading
import pyautogui
import keyboard
from PIL import Image, ImageChops
import shutil
import tkinter as tk
import ctypes
from ctypes import wintypes

# フェイルセーフを無効にする
pyautogui.FAILSAFE = False

# グローバルなフラグ
stop_event = threading.Event()
pause_event = threading.Event()

def set_focus_to_window(x, y):
    """指定した座標にあるウィンドウをアクティブにする（クリックなし）"""
    try:
        # 座標からウィンドウハンドルを取得
        point = wintypes.POINT(x, y)
        hwnd = ctypes.windll.user32.WindowFromPoint(point)
        if hwnd:
            # 親ウィンドウを遡って取得
            root_hwnd = ctypes.windll.user32.GetAncestor(hwnd, 2) # GA_ROOT=2
            # ウィンドウを前面に持ってくる
            ctypes.windll.user32.SetForegroundWindow(root_hwnd)
            return True
    except Exception as e:
        print(f"フォーカス設定エラー: {e}")
    return False

class ControlPanel:
    def __init__(self, stop_evt, pause_evt):
        self.root = tk.Toplevel()
        self.root.title("Control")
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.9)
        
        screen_width = self.root.winfo_screenwidth()
        x = (screen_width // 2) - 100
        y = 50
        self.root.geometry(f"200x50+{x}+{y}")
        self.root.overrideredirect(True)

        self.stop_evt = stop_evt
        self.pause_evt = pause_evt
        self.just_resumed = False 

        self.root.bind("<Button-1>", self.start_move)
        self.root.bind("<B1-Motion>", self.do_move)

        frame = tk.Frame(self.root, bg="#222", padx=5, pady=5)
        frame.pack(fill="both", expand=True)

        self.btn_toggle = tk.Button(frame, text="一時停止", command=self.toggle_pause, bg="#f39c12", fg="white", font=("Arial", 9, "bold"), width=10)
        self.btn_toggle.pack(side="left", padx=2)

        self.btn_stop = tk.Button(frame, text="強制終了", command=self.stop, bg="#e74c3c", fg="white", font=("Arial", 9, "bold"), width=10)
        self.btn_stop.pack(side="left", padx=2)

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def toggle_pause(self):
        if self.pause_evt.is_set():
            self.just_resumed = True
            self.pause_evt.clear()
            self.btn_toggle.config(text="一時停止", bg="#f39c12")
            print(">> [USER] 再開")
        else:
            self.pause_evt.set()
            self.btn_toggle.config(text="再 開", bg="#27ae60")
            print(">> [USER] 一時停止")

    def stop(self):
        print(">> [USER] 強制終了")
        self.stop_evt.set()

    def hide(self):
        self.root.withdraw()
        self.root.update()

    def show(self):
        self.root.deiconify()
        self.root.update()

    def update(self):
        try:
            self.root.update()
        except:
            pass

    def close(self):
        try:
            self.root.destroy()
        except:
            pass

class Capture:
    def __init__(self):
        self.output_folder = 'screenshots'
        os.makedirs(self.output_folder, exist_ok=True)
        self.clear_folder(self.output_folder)
        self.capture_region = None

    def clear_folder(self, folder):
        for filename in os.listdir(folder):
            if filename.endswith(".pdf"): continue
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except:
                pass

    def show_message(self):
        root = tk.Tk()
        root.withdraw()
        message_window = tk.Toplevel(root)
        message_window.title("お知らせ")
        message_window.geometry("350x120") # 少し広げる
        message_window.attributes("-topmost", True)
        label = tk.Label(message_window, text="キャプチャ対象をアクティブにして\nキャプチャ領域を指定してください。", font=("Arial", 11))
        label.pack(pady=20)
        button = tk.Button(message_window, text="OK", command=message_window.destroy)
        button.pack()
        root.wait_window(message_window)
        return root

    def select_capture_region(self):
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.SetWindowPos(hwnd, 1, 0, 0, 0, 0, 0x0013)
            ctypes.windll.user32.ShowWindow(hwnd, 6)

        root = self.show_message()
        self.capture_region = RegionSelector().select_region()
        return root

    def window_capture(self, gui_root):
        if self.capture_region is None:
            return []

        # ターミナルを最小化
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 6)
            ctypes.windll.user32.SetWindowPos(hwnd, 1, 0, 0, 0, 0, 0x0013)

        panel = ControlPanel(stop_event, pause_event)
        time.sleep(1)
        
        x1, y1, x2, y2 = self.capture_region
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        
        # クリックせずにフォーカスを設定
        set_focus_to_window(cx, cy)

        previous_screenshot = None
        page_number = 0
        img_files = []
        same_page_count = 0
        max_same_page_count = 5

        print("キャプチャプロセス開始...")
        try:
            while not stop_event.is_set():
                panel.update()

                while pause_event.is_set() and not stop_event.is_set():
                    panel.update()
                    time.sleep(0.2)

                if panel.just_resumed:
                    print(">> 再開：フォーカス復旧中...")
                    if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 6)
                    # クリックせずにフォーカスだけ戻す
                    set_focus_to_window(cx, cy)
                    time.sleep(0.5)
                    panel.just_resumed = False
                    same_page_count = 0
                    previous_screenshot = None 

                if stop_event.is_set(): break

                time.sleep(0.1)
                panel.hide()
                screenshot = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))
                panel.show()
                screenshot = screenshot.convert("RGB")

                if previous_screenshot and self.images_are_same(previous_screenshot, screenshot):
                    same_page_count += 1
                    print(f">> 同一ページ検出中 ({same_page_count}/{max_same_page_count})")
                    if same_page_count >= max_same_page_count:
                        print(">> 自動終了：同一ページ連続を検出。")
                        break
                else:
                    img_file_path = os.path.join(self.output_folder, f'page_{page_number}.jpg')
                    screenshot.save(img_file_path, "JPEG", quality=85)
                    print(f'Saved: {img_file_path}')
                    img_files.append(img_file_path)
                    same_page_count = 0
                    page_number += 1

                previous_screenshot = screenshot
                pyautogui.press('pgdn')
                
                for _ in range(10):
                    if stop_event.is_set(): break
                    panel.update()
                    time.sleep(0.1)
        except Exception as e:
            print(f">> エラー: {e}")

        panel.close()
        return img_files

    def images_are_same(self, img1, img2):
        return ImageChops.difference(img1, img2).getbbox() is None

class RegionSelector:
    def __init__(self):
        self.root = tk.Tk()
        self.root.attributes("-alpha", 0.4)
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.start_x = self.start_y = 0
        self.rect = None
        self.region = None
        self.canvas = tk.Canvas(self.root, cursor="cross", bg="gray")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect: self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='#ff69b4', fill='#ff69b4', width=2)

    def on_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_button_release(self, event):
        end_x, end_y = event.x, event.y
        self.region = (min(self.start_x, end_x), min(self.start_y, end_y), max(self.start_x, end_x), max(self.start_y, end_y))
        self.root.quit()

    def select_region(self):
        self.root.mainloop()
        self.root.destroy()
        return self.region

class PDFConverter:
    @staticmethod
    def images_to_pdf(image_files, output_pdf_path):
        if not image_files: return
        print(f"PDF作成中... ({len(image_files)}枚)")
        images = [Image.open(img_file) for img_file in image_files]
        try:
            images[0].save(output_pdf_path, save_all=True, append_images=images[1:], dpi=(150, 150))
            print(f"成功: {output_pdf_path}")
        except PermissionError:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            new_path = output_pdf_path.replace(".pdf", f"_{timestamp}.pdf")
            print(f"別名保存: {new_path}")
            images[0].save(new_path, save_all=True, append_images=images[1:], dpi=(150, 150))

def check_for_stop_key():
    while not stop_event.is_set():
        if keyboard.is_pressed('esc'):
            print(">> [USER] ESCキー中断")
            stop_event.set()
        time.sleep(0.1)

def main():
    global stop_event
    stop_event = threading.Event()

    esc_thread = threading.Thread(target=check_for_stop_key, daemon=True)
    esc_thread.start()

    capture = Capture()
    pdf_converter = PDFConverter()

    gui_root = capture.select_capture_region()
    image_files = capture.window_capture(gui_root)

    output_pdf_path = os.path.join(capture.output_folder, 'output.pdf')
    pdf_converter.images_to_pdf(image_files, output_pdf_path)

    os.startfile(capture.output_folder)
    stop_event.set()
    esc_thread.join()
    print("完了。")

if __name__ == "__main__":
    main()
