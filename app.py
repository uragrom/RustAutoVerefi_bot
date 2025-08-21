import os
import re
import json
import time
import threading
from datetime import datetime
from queue import Queue, Empty

import numpy as np
import cv2
import mss
from PIL import Image, ImageTk
import pytesseract
import pyautogui
import keyboard  # используем только его для ввода
# pyperclip не используем, чтобы не трогать буфер

import tkinter as tk
from tkinter import ttk, messagebox

# =============== КОНСТАНТЫ ===============
CONFIG_FILE = "verify_config.json"
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

DEFAULT_CHECK_INTERVAL = 0.30   # сек
DEFAULT_COOLDOWN = 1            # сек (антиспам). Можно 0 в GUI
DEFAULT_CHAT_KEY = "t"          # открыть чат в Rust (обычно 't')

# /verify + 3–4 цифры, пробел может отсутствовать
VERIFY_REGEX = re.compile(r"/\s*verify\D*?(\d{3,4})", re.IGNORECASE)

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
pyautogui.FAILSAFE = False


# =============== OCR ===============
def preprocess_for_ocr(img_bgr):
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    thr = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                cv2.THRESH_BINARY, 31, 9)
    kernel = np.ones((2, 2), np.uint8)
    thr = cv2.morphologyEx(thr, cv2.MORPH_CLOSE, kernel, iterations=1)
    return thr


def ocr_verify_code(img_bgr):
    proc = preprocess_for_ocr(img_bgr)
    cfg = "--oem 3 --psm 6 -c tessedit_char_whitelist=/verifyVERIFY0123456789"
    text = pytesseract.image_to_string(proc, lang="eng", config=cfg) or ""
    norm = (text.replace("O", "0").replace("o", "0")
                 .replace("I", "1").replace("|", "1")
                 .replace("S", "5").replace("B", "8")
                 .replace("Z", "2"))
    m = VERIFY_REGEX.search(norm)
    return (m.group(1) if m else None), norm


# =============== Снимок области ===============
def grab_region(sct, region):
    left, top, width, height = [int(x) for x in region]
    left = max(0, left); top = max(0, top)
    width = max(1, width); height = max(1, height)
    frame = np.array(sct.grab({"left": left, "top": top, "width": width, "height": height}))
    return cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)


# =============== Ввод БЕЗ смены фокуса ===============
def type_in_chat_no_focus(chat_key, cmd):
    """
    Никаких переключений окон.
    Отправляем чисто «клавиатурой» в текущее активное окно:
      chat_key -> Ctrl+A -> Backspace -> печать cmd -> Enter
    """
    # открыть чат
    keyboard.send(chat_key)
    time.sleep(0.10)

    # очистить строку (на случай, если там что-то осталось)
    keyboard.send("ctrl+a")
    time.sleep(0.02)
    keyboard.send("backspace")
    time.sleep(0.02)

    # печатаем как физическая клавиатура (раскладка не важна)
    keyboard.write(cmd, delay=0.0)

    time.sleep(0.03)
    keyboard.send("enter")


# =============== Конфиг ===============
def load_config():
    if not os.path.isfile(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =============== Выбор области (оверлей) ===============
class RegionSelector(tk.Toplevel):
    """Полупрозрачный оверлей: ЛКМ тянуть, отпустить — выбрать, Esc — отмена."""
    def __init__(self, master, on_done):
        super().__init__(master)
        self.on_done = on_done
        self.withdraw()
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.25)
        self.overrideredirect(True)
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")

        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)

        self.start = None
        self.rect = None

        self.bind("<Escape>", self.cancel)
        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        self.deiconify()

    def on_press(self, e):
        self.start = (e.x, e.y)
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None

    def on_drag(self, e):
        if not self.start:
            return
        x1, y1 = self.start
        x2, y2 = e.x, e.y
        if self.rect:
            self.canvas.coords(self.rect, x1, y1, x2, y2)
        else:
            self.rect = self.canvas.create_rectangle(
                x1, y1, x2, y2, outline="white", width=2, dash=(4, 2)
            )

    def on_release(self, e):
        if not self.start:
            self.cancel()
            return
        x1, y1 = self.start
        x2, y2 = e.x, e.y
        left, top = min(x1, x2), min(y1, y2)
        w, h = abs(x2 - x1), abs(y2 - y1)
        self.destroy()
        if w < 5 or h < 5:
            return
        self.on_done((left, top, w, h))

    def cancel(self, *_):
        self.destroy()


# =============== Мониторинг ===============
class Worker(threading.Thread):
    def __init__(self, params_getter, log_q):
        super().__init__(daemon=True)
        self.params_getter = params_getter
        self.log_q = log_q
        self._stop = threading.Event()
        self.last_code = None
        self.last_time = 0.0

    def log(self, msg):
        self.log_q.put(msg)

    def stop(self):
        self._stop.set()

    def run(self):
        self.log("Мониторинг запущен.")
        with mss.mss() as sct:
            while not self._stop.is_set():
                try:
                    p = self.params_getter()
                    region = p["region"]
                    interval = p["interval"]
                    cooldown = p["cooldown"]
                    chat_key = p["chat_key"]

                    frame = grab_region(sct, region)
                    code, raw = ocr_verify_code(frame)

                    if raw.strip():
                        self.log(f"OCR: {raw.strip()}")

                    if code:
                        now = time.time()
                        is_recent = (now - self.last_time) < max(0, cooldown)
                        is_same = (code == self.last_code)
                        if not (is_recent and is_same):
                            cmd = f"/verify {code}"
                            self.log(f"Отправляю: {cmd}")
                            type_in_chat_no_focus(chat_key, cmd)
                            self.last_code = code
                            self.last_time = now
                        else:
                            self.log(f"Код {code} уже вводился недавно — пропуск.")
                    time.sleep(max(0.05, float(interval)))
                except Exception as e:
                    self.log(f"Ошибка: {e}")
                    time.sleep(0.5)
        self.log("Мониторинг остановлен.")


# =============== GUI ===============
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rust Verify Bot — без смены фокуса")
        self.geometry("900x640")
        self.minsize(860, 620)

        cfg = load_config()
        if not cfg.get("region"):
            sw, sh = pyautogui.size()
            cfg["region"] = [10, int(sh*0.62), int(sw*0.45), int(sh*0.22)]

        self.region = [int(v) for v in cfg["region"]]
        self.interval = tk.DoubleVar(value=float(cfg.get("interval", DEFAULT_CHECK_INTERVAL)))
        self.cooldown = tk.IntVar(value=int(cfg.get("cooldown", DEFAULT_COOLDOWN)))
        self.chat_key = tk.StringVar(value=cfg.get("chat_key", DEFAULT_CHAT_KEY))
        self.always_on_top = tk.BooleanVar(value=bool(cfg.get("always_on_top", False)))

        self.worker = None
        self.log_q = Queue()
        self.preview_imgtk = None

        self._build_ui()
        self._apply_aot()
        self.after(120, self._poll_log)

        keyboard.add_hotkey("F12", lambda: self.safe_quit())

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}
        ttk.Label(self, text="Авто-ввод /verify #### (не переключает окна)", font=("Segoe UI", 13, "bold")).pack(anchor="w", **pad)

        grid = ttk.Frame(self)
        grid.pack(fill="x", **pad)

        # Поля зоны
        self.x_var = tk.IntVar(value=self.region[0])
        self.y_var = tk.IntVar(value=self.region[1])
        self.w_var = tk.IntVar(value=self.region[2])
        self.h_var = tk.IntVar(value=self.region[3])

        row = 0
        ttk.Label(grid, text="X:").grid(row=row, column=0, sticky="e")
        ttk.Spinbox(grid, from_=0, to=10000, textvariable=self.x_var, width=8, command=self._on_region_change).grid(row=row, column=1, sticky="w")
        ttk.Label(grid, text="Y:").grid(row=row, column=2, sticky="e")
        ttk.Spinbox(grid, from_=0, to=10000, textvariable=self.y_var, width=8, command=self._on_region_change).grid(row=row, column=3, sticky="w")
        ttk.Label(grid, text="W:").grid(row=row, column=4, sticky="e")
        ttk.Spinbox(grid, from_=50, to=10000, textvariable=self.w_var, width=8, command=self._on_region_change).grid(row=row, column=5, sticky="w")
        ttk.Label(grid, text="H:").grid(row=row, column=6, sticky="e")
        ttk.Spinbox(grid, from_=30, to=10000, textvariable=self.h_var, width=8, command=self._on_region_change).grid(row=row, column=7, sticky="w")

        row += 1
        ttk.Label(grid, text="Интервал (сек):").grid(row=row, column=0, sticky="e")
        ttk.Spinbox(grid, from_=0.00, to=2.0, increment=0.05, textvariable=self.interval, width=8, command=self._save).grid(row=row, column=1, sticky="w")
        ttk.Label(grid, text="Повтор (сек):").grid(row=row, column=2, sticky="e")
        ttk.Spinbox(grid, from_=0, to=60, increment=1, textvariable=self.cooldown, width=8, command=self._save).grid(row=row, column=3, sticky="w")
        ttk.Label(grid, text="Клавиша чата:").grid(row=row, column=4, sticky="e")
        ttk.Entry(grid, textvariable=self.chat_key, width=10).grid(row=row, column=5, sticky="w")
        ttk.Checkbutton(grid, text="Поверх всех окон", variable=self.always_on_top, command=self._apply_aot).grid(row=row, column=6, columnspan=2, sticky="w")

        # Кнопки
        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        self.pick_btn = ttk.Button(btns, text="Выбрать область", command=self.on_pick_region)
        self.start_btn = ttk.Button(btns, text="Старт (свернуть)", command=self.on_start)
        self.stop_btn = ttk.Button(btns, text="Стоп (показать)", command=self.on_stop, state="disabled")
        self.test_btn = ttk.Button(btns, text="Тест OCR", command=self.on_test_ocr)
        self.preview_btn = ttk.Button(btns, text="Показать кадр", command=self.on_preview_once)
        self.exit_btn = ttk.Button(btns, text="Выход (F12)", command=self.safe_quit)

        self.pick_btn.grid(row=0, column=0, padx=6)
        self.start_btn.grid(row=0, column=1, padx=6)
        self.stop_btn.grid(row=0, column=2, padx=6)
        self.test_btn.grid(row=0, column=3, padx=6)
        self.preview_btn.grid(row=0, column=4, padx=6)
        self.exit_btn.grid(row=0, column=5, padx=6)

        # Превью области
        prev_frame = ttk.LabelFrame(self, text="Превью области (разово)")
        prev_frame.pack(fill="both", expand=False, **pad)
        self.preview_lbl = ttk.Label(prev_frame)
        self.preview_lbl.pack(padx=6, pady=6, anchor="w")

        # Лог
        log_frame = ttk.LabelFrame(self, text="Лог")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(log_frame, height=12, wrap="word")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

        self._log("Запусти «Старт (свернуть)», вернись в Rust. Бот не меняет фокус и печатает как клавиатура.")

    # ----- helpers -----
    def _poll_log(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self._log(msg)
        except Empty:
            pass
        self.after(120, self._poll_log)

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{ts}] {msg}\n")
        self.log.see("end")

    def _apply_aot(self):
        self.attributes("-topmost", self.always_on_top.get())
        self._save()

    def _on_region_change(self):
        self.region = [self.x_var.get(), self.y_var.get(), self.w_var.get(), self.h_var.get()]
        self._save()

    def _save(self):
        save_config({
            "region": self.region,
            "interval": float(self.interval.get()),
            "cooldown": int(self.cooldown.get()),
            "chat_key": self.chat_key.get().strip() or DEFAULT_CHAT_KEY,
            "always_on_top": bool(self.always_on_top.get()),
        })

    def params_getter(self):
        return {
            "region": self.region,
            "interval": float(self.interval.get()),
            "cooldown": int(self.cooldown.get()),
            "chat_key": self.chat_key.get().strip() or DEFAULT_CHAT_KEY,
        }

    # ----- actions -----
    def on_pick_region(self):
        # на время выбора областью спрячем окно, чтобы не мешало
        self.withdraw()
        def done(region):
            self.deiconify()
            if not region:
                return
            self.region = list(map(int, region))
            self.x_var.set(self.region[0]); self.y_var.set(self.region[1])
            self.w_var.set(self.region[2]); self.h_var.set(self.region[3])
            self._save()
            self._log(f"Новая область: {self.region}")
            self.on_preview_once()
        self.after(150, lambda: RegionSelector(self, done))

    def on_start(self):
        if self.worker and self.worker.is_alive():
            return
        if not os.path.isfile(TESSERACT_PATH):
            messagebox.showwarning("Tesseract",
                "Не найден Tesseract:\n"
                f"{TESSERACT_PATH}\nУстанови его или поправь путь в начале файла.")
        self.worker = Worker(self.params_getter, self.log_q)
        self.worker.start()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        # Свернём окно, чтобы клавиатура точно шла в игру (мы не меняем фокус)
        self.iconify()
        self._log("Старт мониторинга. Вернись в игру (Alt+Tab).")

    def on_stop(self):
        if self.worker:
            self.worker.stop()
            self.worker = None
        self.deiconify()
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._log("Стоп мониторинга.")

    def on_test_ocr(self):
        try:
            with mss.mss() as sct:
                frame = grab_region(sct, self.region)
            code, raw = ocr_verify_code(frame)
            self._log(f"Тест OCR → текст: {raw.strip() or '(пусто)'}")
            self._log("Тест OCR → " + (f"код найден: {code}" if code else "код НЕ найден."))
        except Exception as e:
            messagebox.showerror("Ошибка OCR", str(e))

    def on_preview_once(self):
        try:
            with mss.mss() as sct:
                frame = grab_region(sct, self.region)
            preview = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            h, w = preview.shape[:2]
            max_w = 820
            scale = min(1.0, max_w / float(w))
            if scale < 1.0:
                preview = cv2.resize(preview, (int(w * scale), int(h * scale)))
            img = Image.fromarray(preview)
            self.preview_imgtk = ImageTk.PhotoImage(image=img)
            self.preview_lbl.configure(image=self.preview_imgtk)
            self._log("Превью обновлено.")
        except Exception as e:
            messagebox.showerror("Ошибка превью", str(e))

    def safe_quit(self):
        try:
            if self.worker:
                self.worker.stop()
                time.sleep(0.2)
        finally:
            self.destroy()


if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except KeyboardInterrupt:
        pass
