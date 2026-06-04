import sys
import tkinter as tk
from tkinter import ttk
import time
import threading
import os
import queue
import math
import pythoncom
import win32com.client
from PIL import Image, ImageTk
import cv2
import numpy as np

def resource_path(rel):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.dirname(__file__)
    return os.path.join(base, rel)


class TTSManager:
    def __init__(self):
        self._queue = queue.Queue()
        self._voice = None
        self._start_worker()

    def _start_worker(self):
        def _worker():
            pythoncom.CoInitialize()
            try:
                v = win32com.client.Dispatch("SAPI.SpVoice")
                v.Rate = 3
                v.Volume = 100
                v.Speak("", 1)
                self._voice = v
                while True:
                    item = self._queue.get()
                    if item is None:
                        v = None
                        self._voice = None
                        break
                    text, callback = item
                    try:
                        v.Speak(text, 0)
                    except Exception:
                        pass
                    if callback:
                        callback()
            finally:
                pythoncom.CoUninitialize()
        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def speak(self, text, callback=None):
        self._queue.put((text, callback))

    def estimate_ms(self, text):
        chars = len(text)
        return max(2500, chars * 55)

    def pause(self):
        if self._voice:
            try:
                self._voice.Pause()
            except Exception:
                pass

    def resume(self):
        if self._voice:
            try:
                self._voice.Resume()
            except Exception:
                pass

    def stop(self):
        self.resume()
        self._queue.queue.clear()


TIMINGS = {
    "step0_welcome":         6000,
    "step1_show_resistors": 4000,
    "step2_show_meters":    4000,
    "step3_show_battery":   3000,
    "step4_build_wires":    7000,
    "step5_initial_values": 7000,
    "step6_property1":      7000,
    "step6_property2":      7000,
    "step6_property3":      7000,
    "step7_r2_150":         5000,
    "step7_r2_100":         5000,
    "step7_r2_50":          5000,
    "step7_r2_25":          5000,
    "step7_r2_0":           5000,
    "step8_finish":         8000,
}

NARRATION = [
    "Здравствуйте! В этом видеоуроке мы познакомимся с параллельным соединением резисторов и его свойствами. Мы соберём электрическую цепь, проведём измерения и убедимся в справедливости законов параллельного соединения.",
    "Свойства параллельного соединения можно продемонстрировать с помощью несложного опыта. Для его проведения нам понадобятся два резистора: первый резистор с сопротивлением 100 Ом, второй резистор с переменным сопротивлением.",
    "Также для определения силы тока будут использоваться электроизмерительные приборы: три миллиамперметра и один вольтметр.",
    "Для получения экспериментальных данных будем использовать источник постоянного напряжения 12 Вольт.",
    "Собираем электрическую цепь с параллельным соединением резисторов. Следует помнить, что амперметры включаются последовательно с элементами, ток через которые требуется измерить, а вольтметр включается параллельно.",
    "Первоначально установим сопротивление переменного резистора R2 равным 200 Ом. При этом первый амперметр будет показывать силу тока 120 миллиампер, второй амперметр показывает силу тока 60 миллиампер, третий амперметр показывает 180 миллиампер, вольтметр показывает 12 Вольт.",
    "Первое свойство параллельного соединения: общий ток равен сумме токов в ветвях, то есть 120 плюс 60 равно 180 миллиампер.",
    "Второе свойство: токи распределяются по ветвям обратно пропорционально их сопротивлениям. Сопротивление второго резистора в 2 раза больше сопротивления первого, поэтому ток через него в два раза меньше.",
    "По показаниям приборов, используя закон Ома, можно рассчитать эквивалентное сопротивление цепи. Оно равно 66,7 Ом, что меньше наименьшего сопротивления 100 Ом.",
    "При уменьшении сопротивления переменного резистора R2 до 150 Ом ток I2 увеличивается до 80 миллиампер, общий ток — до 200 миллиампер.",
    "R2 равно 100 Ом. I2 равен 120 миллиампер, общий ток — 240 миллиампер.",
    "R2 равно 50 Ом. Ток I2 значительно возрос до 240 миллиампер, общий ток — 360 миллиампер.",
    "R2 равно 25 Ом. Почти короткое замыкание: ток I2 равен 480 миллиампер, общий ток — 600 миллиампер.",
    "R2 равно 0 Ом. Короткое замыкание! Ток I2 стремится к бесконечности.",
    "Повторим свойства параллельного соединения: напряжение на всех участках одинаково, общий ток равен сумме токов в ветвях, эквивалентное сопротивление меньше наименьшего сопротивления.",
]

C_BG          = "#f5f0e8"
C_RESISTOR    = "#D32F2F"
C_RESISTOR_BG = "#FFCDD2"
C_WIRE_SERIES = "#1565C0"
C_WIRE_BRANCH = "#2E7D32"
C_METER_BG    = "#E0E0E0"
C_METER_BORDER= "#424242"
C_BATTERY     = "#E65100"
C_TEXT        = "#212121"
C_PANEL_BG    = "#FFFFFF"
C_PANEL_BORDER= "#BDBDBD"
C_R2_ARROW    = "#E65100"
C_TABLE_HEADER= "#37474F"
C_SPEECH_BG   = "#FFFFFF"
C_SPEECH_BORDER = "#FFB300"

_REF_SCALE = 4/3

DX, DY = 60, -150

E1_X = 150; E1_Y1 = 420; E1_Y2 = 620

PA3_X = 300; PA3_Y = 368
NODE_A_X = 420; NODE_A_Y = 368
NODE_B_X = 900; NODE_B_Y = 368

PA1_X = 450; PA1_Y = 210
R1_L = 570; R1_R = 730; R1_Y = 210

PA2_X = 450; PA2_Y = 526
R2_L = 570; R2_R = 730; R2_Y = 526

PV1_X = 660; PV1_Y = 368

PANEL_Y1=65; PANEL_Y2=925

SW = []
def w(x1,y1,x2,y2): SW.append((x1+DX,y1+DY,x2+DX,y2+DY))

w(150,368, 258,368)
w(258,368, 342,368)
w(342,368, 420,368)
w(420,368, 420,210)
w(420,210, 408,210)
w(408,210, 492,210)
w(492,210, 570,210)
w(730,210, 900,210)
w(900,210, 900,368)
w(420,368, 420,526)
w(420,526, 408,526)
w(408,526, 492,526)
w(492,526, 570,526)
w(730,526, 900,526)
w(900,526, 900,368)
w(900,368, 1150,368)
w(1150,368, 1150,750)
w(1150,750, 150,750)
w(150,750, 150,620)

BW = []
def g(x1,y1,x2,y2): BW.append((x1+DX,y1+DY,x2+DX,y2+DY))

g(420,368, 638,368)
g(683,368, 900,368)

JUNCTION_DOTS = [(NODE_A_X+DX,NODE_A_Y+DY), (NODE_B_X+DX,NODE_B_Y+DY)]

ALL_WIRE_STEPS = SW + BW

MEDIA_DIR = resource_path("media")
PHOTOS_DIR = os.path.join(MEDIA_DIR, "photos")
CUBE_IMG_PATH = os.path.join(MEDIA_DIR, "kubik.png")
CUBE_VIDEO_PATH = os.path.join(MEDIA_DIR, "kubik.webm")

SUB_PHASE_GROUPS = {
    "step1": [
        {"trigger": "первый резистор", "tags": ["r1"]},
        {"trigger": "второй резистор", "tags": ["r2"]},
    ],
    "step2": [
        {"trigger": "три миллиамперметра", "tags": ["pa1", "pa2", "pa3"]},
        {"trigger": "один вольтметр", "tags": ["pv1"]},
    ],
    "step3": [
        {"trigger": "источник постоянного", "tags": ["e1", "e1_wire", "battery_photo"]},
    ],
    "step5": [
        {"trigger": "200 Ом",          "tags": []},
        {"trigger": "120 миллиампер",  "tags": ["pa1_val"]},
        {"trigger": "60 миллиампер",   "tags": ["pa2_val"]},
        {"trigger": "180 миллиампер",  "tags": ["pa3_val"]},
        {"trigger": "12 Вольт",        "tags": ["pv1_val"]},
    ],
}


class ResistorCartoon:
    def __init__(self, root):
        self.root = root
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        self._scale = min(sw / 1600.0, (sh - 40) / 1110.0)
        self.px2 = int(sw / self._scale)
        self.px1 = self.px2 - 550
        self.root.title("Параллельное соединение резисторов - обучающий видеоурок")
        self.root.state("zoomed")
        self.root.resizable(False, False)

        self.paused      = False
        self.stopped     = True
        self.phase_idx   = 0
        self.phase_timer = None
        self.wire_timer  = None
        self.phase_start = 0.0
        self.phase_remaining = 0
        self.wire_substep = 0
        self.wire_in_progress = False
        self._schedule_time = 0
        self._timer_update_id = None
        self._speech_ended = False
        self._speech_phase_idx = -1
        self._min_phase_dur_ms = 0
        self._junction_hl_timers = []
        self._junction_hl_items = []
        self.wires_visible = 0
        self.active_phases = set()
        self.active_sub_groups = {}
        self._sub_phase_timers = []
        self._sub_phase_remaining = []
        self._in_transition = False

        self.phase_items = {}
        self.wire_items  = []
        self.formula_items = []
        self.f7_row_items = []
        self.photo_items = []
        self.cube_items = []
        self.speech_items = []

        self.tts = TTSManager()
        self._speech_delay_id = None

        self._video_cap = None
        self._video_playing = False
        self._video_after_id = None
        self._video_poster = None

        self.setup_ui()
        self.load_images()
        self.create_circuit()
        self.create_component_photos()
        self.create_cube_character()
        s = self._scale
        self.canvas.scale("all", 0, 0, s, s)
        if hasattr(self, 'speech_text'):
            self.canvas.itemconfig(self.speech_text, width=max(1, int(640 * s)))
        self.canvas.config(width=int(self.px2 * s), height=int(920 * s))
        self.hide_all()

    def setup_ui(self):
        self.main_frame = tk.Frame(self.root, bg=C_BG)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(
            self.main_frame, bg=C_BG, width=1600, height=990,
            highlightthickness=0
        )
        self.canvas.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH)

        self.ctrl = tk.Frame(self.root, bg="#37474F", height=max(1, int(130 * self._scale)))
        self.ctrl.pack(side=tk.BOTTOM, fill=tk.X)
        self.ctrl.pack_propagate(False)

        btnf = tk.Frame(self.ctrl, bg="#37474F")
        btnf.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(8, 2))
        _bf = ("Arial", max(1, int(14 * self._scale)), "bold")
        _bf2 = ("Arial", max(1, int(14 * self._scale)))
        self.btn_play = tk.Button(
            btnf, text="  \u25b6 СТАРТ  ", font=_bf,
            command=self.on_play, bg="#43A047", fg="white",
            bd=0, padx=18, pady=4, cursor="hand2"
        )
        self.btn_play.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_pause = tk.Button(
            btnf, text="  \u23f8 ПАУЗА  ", font=_bf2,
            command=self.on_pause, bg="#FB8C00", fg="white",
            bd=0, padx=18, pady=4, cursor="hand2", state="disabled"
        )
        self.btn_pause.pack(side=tk.LEFT, padx=6)
        self.btn_reset = tk.Button(
            btnf, text="  \u21ba СБРОС  ", font=_bf2,
            command=self.on_reset, bg="#E53935", fg="white",
            bd=0, padx=18, pady=4, cursor="hand2", state="disabled"
        )
        self.btn_reset.pack(side=tk.LEFT, padx=6)

        step_names = [s[2] for s in self.PHASES]
        self.step_combo = ttk.Combobox(
            btnf, values=step_names, font=_bf2,
            state="readonly", width=45
        )
        self.step_combo.pack(side=tk.LEFT, padx=(20, 0))
        self.step_combo.set("Выберите шаг...")
        self.step_combo.bind("<<ComboboxSelected>>", self.on_step_selected)

        infof = tk.Frame(self.ctrl, bg="#37474F")
        infof.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(2, 8))
        self.prog = tk.Canvas(infof, height=max(1, int(14 * self._scale)), bg="#546E7A", highlightthickness=0)
        self.prog.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))
        self.step_label = tk.Label(
            infof, text="", font=("Consolas", max(1, int(12 * self._scale))),
            bg="#37474F", fg="#B0BEC5", width=20, anchor="w"
        )
        self.step_label.pack(side=tk.LEFT, padx=(0, 12))
        self.timer_label = tk.Label(
            infof, text="", font=("Consolas", max(1, int(14 * self._scale)), "bold"),
            bg="#37474F", fg="#FFF", width=8, anchor="e"
        )
        self.timer_label.pack(side=tk.RIGHT, padx=(0, 4))
        self.status_label = tk.Label(
            self.ctrl,
            text="Нажмите «СТАРТ» для начала",
            font=("Arial", max(1, int(13 * self._scale))), bg="#37474F", fg="#ECEFF1", anchor="w"
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 4))
        self.prog_bg = self.prog.create_rectangle(0,0,2000,20, fill="#546E7A", outline="")
        self.prog_fill = self.prog.create_rectangle(0,0,0,20, fill="#66BB6A", outline="")

    def load_images(self):
        self.photo_imgs = {}
        sizes = {"resistor":(130, 90), "battery":(150, 240), "voltmeter":(110, 110), "ampermetr":(110, 110)}
        rotate_names = {"battery": -90}
        for name in ["resistor","battery","voltmeter","ampermetr"]:
            p = os.path.join(PHOTOS_DIR, name + ".png")
            try:
                pil = Image.open(p).convert("RGBA")
                pil = pil.resize(sizes[name], Image.LANCZOS)
                if name in rotate_names:
                    pil = pil.rotate(rotate_names[name], expand=True, fillcolor=(0,0,0,0))
                self.photo_imgs[name] = ImageTk.PhotoImage(pil)
            except:
                try:
                    img = tk.PhotoImage(file=p)
                    self.photo_imgs[name] = img
                except:
                    self.photo_imgs[name] = None

    def draw_resistor(self, x1, y, x2, label, tag, r2_mode=False):
        h = 28
        items = []
        cx = (x1 + x2) // 2
        img = self.photo_imgs.get("resistor")
        if img:
            items.append(self.canvas.create_image(cx, y, image=img, anchor="center", tags=tag))
        else:
            items.append(self.canvas.create_rectangle(
                x1, y-h, x2, y+h, fill=C_RESISTOR_BG, outline=C_RESISTOR, width=2, tags=tag
            ))
            w = x2 - x1
            pts = [x1, y]
            for i in range(1, 10, 2):
                pts.extend([x1 + i*w/10,   y - h + 5])
                pts.extend([x1 + (i+1)*w/10, y + h - 5])
            pts.extend([x2, y])
            items.append(self.canvas.create_line(*pts, fill=C_RESISTOR, width=2, tags=tag))
        items.append(self.canvas.create_text(
            (x1+x2)//2, y+h+18, text=label,
            font=("Arial", self._sf(16), "bold"), fill=C_TEXT, anchor="n", tags=tag
        ))
        if r2_mode:
            items.append(self.canvas.create_text(
                x2+22, y, text="\u21c5", font=("Arial", self._sf(20), "bold"),
                fill=C_R2_ARROW, anchor="center", tags=tag
            ))
            val = self.canvas.create_text(
                (x1+x2)//2, y-h-18, text="R2 = 200 \u03a9",
                font=("Arial", self._sf(16), "bold"), fill=C_TEXT, anchor="s", tags=tag
            )
            items.append(val)
            self.r2_label = val
        return items

    def draw_meter(self, cx, cy, kind, label, tag, specific_tag=None):
        items = []
        t = (tag, specific_tag) if specific_tag else tag
        img_key = "ampermetr" if kind == "A" else "voltmeter"
        img = self.photo_imgs.get(img_key)
        if img:
            items.append(self.canvas.create_image(cx, cy, image=img, anchor="center", tags=t))
            dy = 55
        else:
            r = 18
            items.append(self.canvas.create_oval(
                cx-r, cy-r, cx+r, cy+r,
                fill=C_METER_BG, outline=C_METER_BORDER, width=2, tags=t
            ))
            items.append(self.canvas.create_text(
                cx, cy, text=kind, font=("Arial", self._sf(16), "bold"),
                fill=C_TEXT, anchor="center", tags=t
            ))
            dy = 0
        items.append(self.canvas.create_text(
            cx, cy+dy+r+14 if not img else cy+48, text=label,
            font=("Arial", self._sf(11), "bold"), fill=C_TEXT, anchor="n", tags=t
        ))
        val_tags = (tag, specific_tag + "_val") if specific_tag else tag
        val = self.canvas.create_text(
            cx, cy+dy+r+32 if not img else cy+63, text="",
            font=("Consolas", self._sf(11)), fill="#D32F2F", anchor="n", tags=val_tags
        )
        items.append(val)
        return items, val

    def draw_battery(self, x, y1, y2, label, tag):
        items = []
        items.append(self.canvas.create_line(x-22, y1, x+22, y1, fill=C_BATTERY, width=2, tags=tag))
        items.append(self.canvas.create_line(x-16, y2, x+16, y2, fill=C_BATTERY, width=5, tags=tag))
        items.append(self.canvas.create_rectangle(
            x-10, y1+4, x+10, y2-4,
            fill="#FFF3E0", outline=C_BATTERY, width=2, tags=tag
        ))
        items.append(self.canvas.create_text(
            x+24, y1+4, text="+", font=("Arial", self._sf(16), "bold"),
            fill=C_BATTERY, anchor="w", tags=tag
        ))
        items.append(self.canvas.create_text(
            x+24, y2-4, text="\u2212", font=("Arial", self._sf(16), "bold"),
            fill=C_BATTERY, anchor="w", tags=tag
        ))
        items.append(self.canvas.create_text(
            x-50, (y1+y2)//2, text=label,
            font=("Arial", self._sf(14), "bold"), fill=C_TEXT,
            anchor="center", angle=90, tags=tag
        ))
        return items

    def junction_dot(self, x, y, tag):
        return self.canvas.create_oval(x-4, y-4, x+4, y+4, fill=C_TEXT, outline="", tags=tag)

    def create_circuit(self):
        p = {}
        self.sub_phase_items = {}

        p["step1"] = []
        r1_items = self.draw_resistor(R1_L+DX, R1_Y+DY, R1_R+DX, "R1 = 100 \u03a9", "r1")
        p["step1"].extend(r1_items)
        self.sub_phase_items[("step1", 0)] = r1_items

        r2_items = self.draw_resistor(R2_L+DX, R2_Y+DY, R2_R+DX, "R2 = 0\u2026200 \u03a9", "r2", r2_mode=True)
        p["step1"].extend(r2_items)
        self.sub_phase_items[("step1", 1)] = r2_items

        p["step2"] = []
        ammeter_items = []
        items, self.pa1_val = self.draw_meter(PA1_X+DX, PA1_Y+DY, "A", "PA1", "meter", "pa1")
        ammeter_items.extend(items)
        p["step2"].extend(items)
        items, self.pa2_val = self.draw_meter(PA2_X+DX, PA2_Y+DY, "A", "PA2", "meter", "pa2")
        ammeter_items.extend(items)
        p["step2"].extend(items)
        items, self.pa3_val = self.draw_meter(PA3_X+DX, PA3_Y+DY, "A", "PA3", "meter", "pa3")
        ammeter_items.extend(items)
        p["step2"].extend(items)
        self.sub_phase_items[("step2", 0)] = ammeter_items

        pv_items = []
        items, self.pv1_val = self.draw_meter(PV1_X+DX, PV1_Y+DY, "V", "PV1", "meter", "pv1")
        pv_items.extend(items)
        p["step2"].extend(items)
        self.sub_phase_items[("step2", 1)] = pv_items

        p["step3"] = []
        bx, by1, by2 = E1_X+DX, E1_Y1+DY, E1_Y2+DY
        bat_items = self.draw_battery(bx, by1, by2, "12 \u0412", "e1")
        p["step3"].extend(bat_items)
        p["step3"].append(self.canvas.create_line(bx, NODE_A_Y+DY, bx, by1+70, fill=C_WIRE_SERIES, width=3, tags="e1_wire"))
        p["step3"].append(self.canvas.create_line(bx, by2-70, bx, 750+DY, fill=C_WIRE_SERIES, width=3, tags="e1_wire"))

        p["step4"] = []
        for jx, jy in JUNCTION_DOTS:
            p["step4"].append(self.junction_dot(jx, jy, "junc"))

        self.wire_items = []
        for i, (x1,y1,x2,y2) in enumerate(ALL_WIRE_STEPS):
            is_br = i >= len(SW)
            col = C_WIRE_BRANCH if is_br else C_WIRE_SERIES
            w = 5 if not is_br else 3
            item = self.canvas.create_line(x1, y1, x2, y2, fill=col, width=w, tags="w%d" % i)
            self.wire_items.append(item)
        self.canvas.tag_raise("meter")
        self.canvas.tag_raise("junc")
        self.canvas.tag_raise("hl_junc")

        p["step5"] = []
        self.phase_items = p
        self.create_formula_panel()

    def create_formula_panel(self):
        px = self.px1 + 20
        fy0 = PANEL_Y1 + 70
        fi = []
        fi.append(self.canvas.create_rectangle(
            self.px1, PANEL_Y1, self.px2, PANEL_Y2,
            fill=C_PANEL_BG, outline=C_PANEL_BORDER, width=2, tags="f_all"
        ))
        fi.append(self.canvas.create_text(
            (self.px1+self.px2)//2, PANEL_Y1+25,
            text="\u0421\u0432\u043e\u0439\u0441\u0442\u0432\u0430 \u0446\u0435\u043f\u0438",
            font=("Arial", self._sf(24), "bold"), fill=C_TABLE_HEADER,
            anchor="center", tags="f_all"
        ))
        self.f_panel_bg_items = fi
        self.formula_items.extend(fi)

        self.f5_items = []
        ty = fy0
        self.f5_items.append(self.canvas.create_text(
            px, ty,
            text="{:15}{:>8}".format("\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440", "\u0417\u043d\u0430\u0447\u0435\u043d\u0438\u0435"),
            font=("Consolas", self._sf(20), "bold"), fill=C_TABLE_HEADER,
            anchor="w", tags="f_step5"
        ))
        for i, param in enumerate(["R1","R2","I1","I2","I\u043e\u0431\u0449","U"]):
            y = ty + 40 + i*36
            self.f5_items.append(self.canvas.create_text(
                px, y, text="{:15}{:>8}".format(param, ""),
                font=("Consolas", self._sf(20)), fill=C_TEXT,
                anchor="w", tags="f_step5"
            ))
        self.formula_items.extend(self.f5_items)

        self.f6_groups = []
        for lines in [
            ["I\u043e\u0431\u0449 = I\u2081 + I\u2082", "180 \u043c\u0410 = 120 \u043c\u0410 + 60 \u043c\u0410",
             "\u041e\u0431\u0449\u0438\u0439 \u0442\u043e\u043a \u0440\u0430\u0432\u0435\u043d \u0441\u0443\u043c\u043c\u0435",
             "\u0442\u043e\u043a\u043e\u0432 \u0432 \u0432\u0435\u0442\u0432\u044f\u0445."],
            ["R\u2082 / R\u2081 = I\u2081 / I\u2082", "200 / 100 = 120 / 60 = 2",
             "\u0422\u043e\u043a\u0438 \u043e\u0431\u0440\u0430\u0442\u043d\u043e",
             "\u043f\u0440\u043e\u043f\u043e\u0440\u0446\u0438\u043e\u043d\u0430\u043b\u044c\u043d\u044b",
             "\u0441\u043e\u043f\u0440\u043e\u0442\u0438\u0432\u043b\u0435\u043d\u0438\u044f\u043c."],
            ["R\u044d\u043a\u0432 = 66,7 \u041e\u043c",
             "66,7 \u041e\u043c < 100 \u041e\u043c = R\u2081",
             "\u042d\u043a\u0432\u0438\u0432\u0430\u043b\u0435\u043d\u0442\u043d\u043e\u0435",
             "\u0441\u043e\u043f\u0440\u043e\u0442\u0438\u0432\u043b\u0435\u043d\u0438\u0435 \u043c\u0435\u043d\u044c\u0448\u0435",
             "\u043d\u0430\u0438\u043c\u0435\u043d\u044c\u0448\u0435\u0433\u043e."],
        ]:
            grp = []
            y0 = fy0
            for li, line in enumerate(lines):
                fs = self._sf(22 if li < 2 else 18)
                fw = "bold" if li < 2 else "normal"
                fc = "#D32F2F" if li < 2 else C_TEXT
                y = y0 + li*40 + len(self.f6_groups)*160
                grp.append(self.canvas.create_text(
                    px, y, text=line,
                    font=("Arial", fs, fw), fill=fc,
                    anchor="w", tags="f_step6"
                ))
            self.f6_groups.append(grp)
            for item in grp:
                self.formula_items.append(item)

        y7 = fy0
        col_x = [px + i*135 for i in range(4)]
        col_hdr = ["R2","I2","I\u043e\u0431\u0449","U"]
        self.f7_header = []
        for i, txt in enumerate(col_hdr):
            t = self.canvas.create_text(
                col_x[i], y7, text=txt,
                font=("Arial", self._sf(22), "bold"), fill="#D32F2F",
                anchor="w", tags="f_step7"
            )
            self.f7_header.append(t)
            self.formula_items.append(t)
        self.f7_row_items = []
        for ri in range(6):
            y = y7 + 70 + ri*65
            for ci in range(4):
                t = self.canvas.create_text(
                    col_x[ci], y, text="",
                    font=("Arial", self._sf(23)), fill=C_TEXT,
                    anchor="w", tags="f_step7"
                )
                self.f7_row_items.append(t)
                self.formula_items.append(t)

        self.f8_items = []
        for li, line in enumerate([
            "\u0421\u0432\u043e\u0439\u0441\u0442\u0432\u0430 \u043f\u0430\u0440\u0430\u043b\u043b\u0435\u043b\u044c\u043d\u043e\u0433\u043e",
            "\u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u044f \u0440\u0435\u0437\u0438\u0441\u0442\u043e\u0440\u043e\u0432:", "",
            "1.  U = const = 12 \u0412",
            "    (\u043d\u0430\u043f\u0440\u044f\u0436\u0435\u043d\u0438\u0435 \u043e\u0434\u0438\u043d\u0430\u043a\u043e\u0432\u043e)", "",
            "2.  I\u043e\u0431\u0449 = I\u2081 + I\u2082",
            "    (\u0442\u043e\u043a \u0432\u0435\u0442\u0432\u0438 \u043d\u0435 \u0437\u0430\u0432\u0438\u0441\u0438\u0442",
            "    \u043e\u0442 \u0434\u0440\u0443\u0433\u043e\u0439 \u0432\u0435\u0442\u0432\u0438)", "",
            "3.  R\u044d\u043a\u0432 < R\u2081, R\u044d\u043a\u0432 < R\u2082",
            "    (\u044d\u043a\u0432\u0438\u0432\u0430\u043b\u0435\u043d\u0442\u043d\u043e\u0435",
            "    \u0441\u043e\u043f\u0440\u043e\u0442\u0438\u0432\u043b\u0435\u043d\u0438\u0435",
            "    \u043c\u0435\u043d\u044c\u0448\u0435 \u043d\u0430\u0438\u043c\u0435\u043d\u044c\u0448\u0435\u0433\u043e)", "",
            "4.  I\u2081 / I\u2082 = R\u2082 / R\u2081",
            "    (\u0442\u043e\u043a\u0438 \u043e\u0431\u0440\u0430\u0442\u043d\u043e \u043f\u0440\u043e\u043f\u043e\u0440\u0446.",
            "    \u0441\u043e\u043f\u0440\u043e\u0442\u0438\u0432\u043b\u0435\u043d\u0438\u044f\u043c)",
        ]):
            is_numbered = line and line[0].isdigit()
            fs = self._sf(22 if is_numbered else 18)
            fw = "bold" if is_numbered else "normal"
            fc = "#D32F2F" if "=" in line else C_TEXT
            y = fy0 + li*34
            self.f8_items.append(self.canvas.create_text(
                px, y, text=line,
                font=("Arial", fs, fw), fill=fc,
                anchor="w", tags="f_step8"
            ))
        self.formula_items.extend(self.f8_items)

    def create_component_photos(self):
        img = self.photo_imgs.get("battery")
        if img:
            bx = E1_X + DX - 130
            by = (E1_Y1 + DY + E1_Y2 + DY) // 2
            item = self.canvas.create_image(bx, by, image=img, anchor="center", tags=("e1", "battery_photo"))
            if "step3" in self.phase_items:
                self.phase_items["step3"].append(item)

    def create_cube_character(self):
        cx, cy = 210, 820
        self._cube_pos = (cx, cy)

        self._video_poster = None
        if os.path.exists(CUBE_IMG_PATH):
            try:
                pil = Image.open(CUBE_IMG_PATH).convert("RGBA")
                cs = max(1, round(502 * self._scale / _REF_SCALE))
                self._video_poster = ImageTk.PhotoImage(pil.resize((cs, cs), Image.LANCZOS))
            except:
                pass

        self.cube_img_id = self.canvas.create_image(
            cx, cy, image=self._video_poster, anchor="center", tags="cube"
        )
        self.cube_items.append(self.cube_img_id)

        self._open_video()

        bubble_x1 = cx + 140
        bubble_y1 = 620
        bubble_x2 = cx + 850
        bubble_y2 = 710
        self.speech_bg = self.canvas.create_rectangle(
            bubble_x1, bubble_y1, bubble_x2, bubble_y2,
            fill=C_SPEECH_BG, outline=C_SPEECH_BORDER, width=2,
            tags="speech"
        )
        self.speech_items.append(self.speech_bg)
        tri_cx = cx + 80
        tri_y  = bubble_y2 + 2
        self.speech_tri = self.canvas.create_polygon(
            tri_cx, tri_y + 8,
            bubble_x1 - 2, tri_y - 4,
            bubble_x1 - 2, tri_y + 4,
            fill=C_SPEECH_BG, outline=C_SPEECH_BORDER, width=2,
            tags="speech"
        )
        self.speech_items.append(self.speech_tri)
        self.speech_text = self.canvas.create_text(
            (bubble_x1+bubble_x2)//2, (bubble_y1+bubble_y2)//2,
            text="", font=("Arial", self._sf(17)), fill=C_TEXT,
            anchor="center", width=int((bubble_x2 - bubble_x1) * 0.85), tags="speech"
        )
        self.speech_items.append(self.speech_text)
        self.canvas.tag_raise("speech")

    def _open_video(self):
        if self._video_cap:
            self._video_cap.release()
            self._video_cap = None
        if os.path.exists(CUBE_VIDEO_PATH):
            try:
                self._video_cap = cv2.VideoCapture(CUBE_VIDEO_PATH)
            except:
                self._video_cap = None

    def _video_update(self):
        if not self._video_playing or self.stopped:
            return
        if self._video_cap is None or not self._video_cap.isOpened():
            return
        ret, frame = self._video_cap.read()
        if not ret:
            self._video_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ret, frame = self._video_cap.read()
            if not ret:
                return
        try:
            frame_rgb = self._frame_to_rgba(frame)
            pil_img = Image.fromarray(frame_rgb, 'RGBA')
            cs = max(1, round(502 * self._scale / _REF_SCALE))
            pil_img = pil_img.resize((cs, cs), Image.LANCZOS)
            photo = ImageTk.PhotoImage(pil_img)
            self.canvas.itemconfig(self.cube_img_id, image=photo)
            self._video_current_photo = photo
        except:
            pass
        self._video_after_id = self.root.after(5, self._video_update)

    def _frame_to_rgba(self, bgr_frame):
        is_bg = np.all(bgr_frame <= 8, axis=2)
        alpha = np.where(is_bg, 0, 255).astype(np.uint8)
        rgb = bgr_frame[...,::-1]
        return np.dstack((rgb, alpha))

    def _get_trigger_delay(self, phase_key, trigger):
        idx = next((i for i, (pk,_,_) in enumerate(self.PHASES) if pk == phase_key), None)
        if idx is None or idx >= len(NARRATION):
            return 0
        text = NARRATION[idx]
        pos = text.find(trigger)
        if pos < 0:
            return 0
        chars_before = len(text[:pos])
        total_chars = max(len(text), 1)
        total_dur = self.tts.estimate_ms(text)
        delay = int(total_dur * chars_before / total_chars)
        return max(0, delay - 600)

    def _start_video(self):
        self._video_playing = True
        if self._video_cap is None or not self._video_cap.isOpened():
            self._open_video()
        if self._video_after_id:
            self.root.after_cancel(self._video_after_id)
            self._video_after_id = None
        self._video_update()

    def _stop_video(self):
        self._video_playing = False
        if self._video_after_id:
            self.root.after_cancel(self._video_after_id)
            self._video_after_id = None

    def hide_all(self):
        for lst in self.phase_items.values():
            for item in lst:
                self.canvas.itemconfig(item, state="hidden")
        for item in self.wire_items:
            self.canvas.itemconfig(item, state="hidden")
        for item in self.formula_items:
            self.canvas.itemconfig(item, state="hidden")
        for item in self.photo_items:
            self.canvas.itemconfig(item, state="hidden")
        for item in self.cube_items:
            self.canvas.itemconfig(item, state="hidden")
        for item in self.speech_items:
            self.canvas.itemconfig(item, state="hidden")
        self.wires_visible = 0
        self.active_phases.clear()
        self.active_sub_groups.clear()
        for tid, _, _, _ in self._sub_phase_timers:
            self.root.after_cancel(tid)
        self._sub_phase_timers.clear()
        self._sub_phase_remaining.clear()

    def show_photo_panel(self, visible=True):
        state = "normal" if visible else "hidden"
        for item in self.photo_items:
            self.canvas.itemconfig(item, state=state)

    def show_cube(self, visible=True):
        state = "normal" if visible else "hidden"
        for item in self.cube_items:
            self.canvas.itemconfig(item, state=state)
        if visible:
            if not self._video_playing:
                self._start_video()
        else:
            self._stop_video()

    def show_speech(self, text):
        state = "normal"
        for item in self.speech_items:
            self.canvas.itemconfig(item, state=state)
        if hasattr(self, 'speech_text'):
            self.canvas.itemconfig(self.speech_text, text=text)

    def hide_speech(self):
        for item in self.speech_items:
            self.canvas.itemconfig(item, state="hidden")

    def _activate_sub_group(self, phase_key, sub_idx):
        if phase_key in self.active_sub_groups:
            self.active_sub_groups[phase_key].add(sub_idx)
        self._highlight_formula_data(phase_key, sub_idx)
        try:
            self.refresh_display()
        except Exception:
            pass

    def refresh_display(self):
        for lst in self.phase_items.values():
            for item in lst:
                self.canvas.itemconfig(item, state="hidden")

        for key in self.active_phases:
            if key in self.phase_items:
                for item in self.phase_items[key]:
                    self.canvas.itemconfig(item, state="normal")

        for phase_key, sub_indices in self.active_sub_groups.items():
            for sub_idx in sub_indices:
                key = (phase_key, sub_idx)
                if key in self.sub_phase_items:
                    for item in self.sub_phase_items[key]:
                        self.canvas.itemconfig(item, state="normal")
                else:
                    if phase_key in SUB_PHASE_GROUPS and sub_idx < len(SUB_PHASE_GROUPS[phase_key]):
                        tags = SUB_PHASE_GROUPS[phase_key][sub_idx]["tags"]
                        for tag in tags:
                            for item in self.canvas.find_withtag(tag):
                                self.canvas.itemconfig(item, state="normal")

        for i, item in enumerate(self.wire_items):
            state = "normal" if i < self.wires_visible else "hidden"
            self.canvas.itemconfig(item, state=state)
        if self.phase_idx > 0 or self.active_phases or self.active_sub_groups:
            self.show_photo_panel(True)
        self.show_cube(True)

    def show_formula(self, key):
        for item in self.formula_items:
            self.canvas.itemconfig(item, state="hidden")
        for item in self.f_panel_bg_items:
            self.canvas.itemconfig(item, state="normal")
        if key == "step5":
            for item in self.f5_items:
                self.canvas.itemconfig(item, state="normal")
        elif key == "step6_p1":
            for item in self.f6_groups[0]:
                self.canvas.itemconfig(item, state="normal")
        elif key == "step6_p2":
            for item in self.f6_groups[1]:
                self.canvas.itemconfig(item, state="normal")
        elif key == "step6_p3":
            for item in self.f6_groups[2]:
                self.canvas.itemconfig(item, state="normal")
        elif key == "step7":
            for item in self.f7_header:
                self.canvas.itemconfig(item, state="normal")
            for item in self.f7_row_items:
                self.canvas.itemconfig(item, state="normal")
        elif key == "step8":
            for item in self.f8_items:
                self.canvas.itemconfig(item, state="normal")

    def calc_values(self, r2):
        i1 = 12.0 / 100.0
        u = 12.0
        if r2 <= 0:
            return i1, float("inf"), float("inf"), u
        i2 = u / r2
        i3 = i1 + i2
        return i1, i2, i3, u

    def _update_f5_table(self, r2):
        i1, i2, i3, u = self.calc_values(r2)
        rows = [
            ("R1", "100 \u041e\u043c"),
            ("R2", "%d \u03a9" % r2 if r2 > 0 else "0 \u03a9 (\u043a\u0437)"),
            ("I1", "%d \u043c\u0410" % (i1*1000)),
        ]
        if math.isinf(i2):
            rows.append(("I2", "\u221e \u043c\u0410"))
            rows.append(("I\u043e\u0431\u0449", "\u221e \u043c\u0410"))
        else:
            rows.append(("I2", "%d \u043c\u0410" % (i2*1000)))
            rows.append(("I\u043e\u0431\u0449", "%d \u043c\u0410" % (i3*1000)))
        rows.append(("U", "%d \u0412" % u))
        for i, (param, val) in enumerate(rows):
            idx = i + 1
            if idx < len(self.f5_items):
                self.canvas.itemconfig(self.f5_items[idx], text="{:15}{:>8}".format(param, val))

    def _update_f7_table(self):
        r2_vals = [200, 150, 100, 50, 25, 0]
        for ri, r2 in enumerate(r2_vals):
            i1, i2, i3, u = self.calc_values(r2)
            i2s = "\u221e" if math.isinf(i2) else "%d" % (i2*1000)
            i3s = "\u221e" if math.isinf(i3) else "%d" % (i3*1000)
            row_vals = [str(r2), "{:>6}".format(i2s), "{:>6}".format(i3s), "%d\u0412" % u]
            for ci, txt in enumerate(row_vals):
                idx = ri * 4 + ci
                if idx < len(self.f7_row_items):
                    self.canvas.itemconfig(self.f7_row_items[idx], text=txt)

    def update_readings(self, r2):
        i1, i2, i3, u = self.calc_values(r2)

        def fmt(v):
            if math.isinf(v):
                return "\u221e"
            v_ma = v * 1000.0
            if abs(v_ma - round(v_ma)) < 0.05:
                return "%d" % v_ma
            return "%.1f" % v_ma

        self.canvas.itemconfig(self.pa1_val, text="%d \u043c\u0410" % (i1*1000))
        if math.isinf(i2):
            self.canvas.itemconfig(self.pa2_val, text="\u221e \u043c\u0410")
            self.canvas.itemconfig(self.pa3_val, text="\u221e \u043c\u0410")
        else:
            self.canvas.itemconfig(self.pa2_val, text="%s \u043c\u0410" % fmt(i2))
            self.canvas.itemconfig(self.pa3_val, text="%s \u043c\u0410" % fmt(i3))
        self.canvas.itemconfig(self.pv1_val, text="%d \u0412" % u)
        if hasattr(self, "r2_label"):
            self.canvas.itemconfig(self.r2_label, text="R2 = %d \u03a9" % r2 if r2 > 0 else "R2 = 0 \u03a9 (\u043a\u0437)")
        self._update_f5_table(r2)
        self._update_f7_table()

    PHASES = [
        ("step0",       "\u0428\u0430\u0433 1 \u0438\u0437 9","\u041f\u0440\u0438\u0432\u0435\u0442\u0441\u0442\u0432\u0438\u0435"),
        ("step1",       "\u0428\u0430\u0433 2 \u0438\u0437 9","\u041f\u043e\u044f\u0432\u043b\u0435\u043d\u0438\u0435 \u0440\u0435\u0437\u0438\u0441\u0442\u043e\u0440\u043e\u0432 R1 = 100 \u041e\u043c \u0438 R2 = 0\u2026200 \u041e\u043c"),
        ("step2",       "\u0428\u0430\u0433 3 \u0438\u0437 9","\u0414\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u0438\u0435 \u0438\u0437\u043c\u0435\u0440\u0438\u0442\u0435\u043b\u044c\u043d\u044b\u0445 \u043f\u0440\u0438\u0431\u043e\u0440\u043e\u0432 PA1\u2013PA3, PV1"),
        ("step3",       "\u0428\u0430\u0433 4 \u0438\u0437 9","\u041f\u043e\u0434\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0430 E1 = 12 \u0412"),
        ("step4",       "\u0428\u0430\u0433 5 \u0438\u0437 9","\u0421\u0431\u043e\u0440\u043a\u0430 \u0446\u0435\u043f\u0438: \u043f\u0430\u0440\u0430\u043b\u043b\u0435\u043b\u044c\u043d\u043e\u0435 \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u0435"),
        ("step5",       "\u0428\u0430\u0433 6 \u0438\u0437 9","\u041f\u043e\u043a\u0430\u0437\u0430\u043d\u0438\u044f \u043f\u0440\u0438\u0431\u043e\u0440\u043e\u0432 \u043f\u0440\u0438 R2 = 200 \u041e\u043c"),
        ("step6_p1",    "\u0428\u0430\u0433 7 \u0438\u0437 9","\u0421\u0432\u043e\u0439\u0441\u0442\u0432\u043e 1: I\u043e\u0431\u0449 = I\u2081 + I\u2082 (180 = 120 + 60)"),
        ("step6_p2",    "\u0428\u0430\u0433 7 \u0438\u0437 9","\u0421\u0432\u043e\u0439\u0441\u0442\u0432\u043e 2: R\u2082/R\u2081 = I\u2081/I\u2082 = 2"),
        ("step6_p3",    "\u0428\u0430\u0433 7 \u0438\u0437 9","\u0421\u0432\u043e\u0439\u0441\u0442\u0432\u043e 3: R\u044d\u043a\u0432 = 66,7 \u041e\u043c < 100 \u041e\u043c"),
        ("step7_r2_150","\u0428\u0430\u0433 8 \u0438\u0437 9","R\u2082 = 150 \u041e\u043c \u2013 \u043f\u043e\u043a\u0430\u0437\u0430\u043d\u0438\u044f \u043c\u0435\u043d\u044f\u044e\u0442\u0441\u044f"),
        ("step7_r2_100","\u0428\u0430\u0433 8 \u0438\u0437 9","R\u2082 = 100 \u041e\u043c \u2013 \u0442\u043e\u043a I\u2082 \u0440\u0430\u0441\u0442\u0451\u0442"),
        ("step7_r2_50", "\u0428\u0430\u0433 8 \u0438\u0437 9","R\u2082 = 50 \u041e\u043c \u2013 I\u2082 = 240 \u043c\u0410"),
        ("step7_r2_25", "\u0428\u0430\u0433 8 \u0438\u0437 9","R\u2082 = 25 \u041e\u043c \u2013 \u043f\u043e\u0447\u0442\u0438 \u043a\u0437"),
        ("step7_r2_0",  "\u0428\u0430\u0433 8 \u0438\u0437 9","R\u2082 = 0 \u041e\u043c \u2013 \u043a\u043e\u0440\u043e\u0442\u043a\u043e\u0435 \u0437\u0430\u043c\u044b\u043a\u0430\u043d\u0438\u0435!"),
        ("step8",       "\u0428\u0430\u0433 9 \u0438\u0437 9","\u0418\u0442\u043e\u0433: \u0441\u0432\u043e\u0439\u0441\u0442\u0432\u0430 \u043f\u0430\u0440\u0430\u043b\u043b\u0435\u043b\u044c\u043d\u043e\u0433\u043e \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u044f"),
    ]
    PHASE_DURATIONS = [TIMINGS[k] for k in (
        "step0_welcome",
        "step1_show_resistors","step2_show_meters","step3_show_battery",
        "step4_build_wires","step5_initial_values",
        "step6_property1","step6_property2","step6_property3",
        "step7_r2_150","step7_r2_100","step7_r2_50","step7_r2_25","step7_r2_0",
        "step8_finish",
    )]
    PHASE_FORMULA = [
        None,
        None,None,None,None,"step5",
        "step6_p1","step6_p2","step6_p3",
        "step7","step7","step7","step7","step7",
        "step8",
    ]

    def enter_phase(self, idx):
        if idx >= len(self.PHASES):
            self.show_final()
            return

        self.phase_idx = idx
        key, step_txt, status = self.PHASES[idx]
        dur = self.PHASE_DURATIONS[idx]

        self.step_label.config(text=step_txt)
        self.status_label.config(text=status)
        self.update_progress(idx)

        if idx < len(NARRATION):
            if self._speech_delay_id:
                self.root.after_cancel(self._speech_delay_id)
            self.show_speech(NARRATION[idx])
            self._speech_delay_id = None
            self._speech_ended = False
            self._speech_phase_idx = idx
            self._min_phase_dur_ms = dur
            cb_idx = idx
            self.tts.speak(NARRATION[idx], callback=lambda: self.root.after_idle(lambda: self._on_speech_ended(cb_idx)))
        else:
            self._speech_ended = True
            self.show_speech(status)

        for tid, _, _, _ in self._sub_phase_timers:
            self.root.after_cancel(tid)
        self._sub_phase_timers.clear()
        self._sub_phase_remaining.clear()

        for i in range(idx):
            pk = self.PHASES[i][0]
            if pk in SUB_PHASE_GROUPS:
                self.active_sub_groups[pk] = set(range(len(SUB_PHASE_GROUPS[pk])))
            elif pk in self.phase_items:
                self.active_phases.add(pk)

        if key in SUB_PHASE_GROUPS:
            groups = SUB_PHASE_GROUPS[key]
            if key not in self.active_sub_groups:
                self.active_sub_groups[key] = set()
            for sub_idx, sg in enumerate(groups):
                delay = self._get_trigger_delay(key, sg["trigger"])
                if delay <= 50:
                    self.active_sub_groups[key].add(sub_idx)
                else:
                    deadline = time.time() + delay / 1000.0
                    tid = self.root.after(delay, lambda k=key, si=sub_idx: self._activate_sub_group(k, si))
                    self._sub_phase_timers.append((tid, deadline, key, sub_idx))
        else:
            self.active_phases.add(key)

        if idx == 5:
            self.update_readings(200)

        self.refresh_display()

        fkey = self.PHASE_FORMULA[idx]
        if fkey:
            self.show_formula(fkey)
        else:
            for item in self.formula_items:
                self.canvas.itemconfig(item, state="hidden")
            for item in self.f_panel_bg_items:
                self.canvas.itemconfig(item, state="normal")

        if idx == 6:
            self.highlight_property1()
            self._highlight_formula_data("step6_p1", 0)
        elif idx == 7:
            self.highlight_property2()
            self._highlight_formula_data("step6_p2", 0)
        elif idx == 8:
            self.highlight_property3()
            self._highlight_formula_data("step6_p3", 0)
        elif 9 <= idx <= 13:
            r2_vals = [150, 100, 50, 25, 0]
            sub = idx - 9
            if sub < len(r2_vals):
                self.update_readings(r2_vals[sub])
                self.highlight_r2_table_row(sub + 1)
                self.highlight_junction_points(True)

        if idx not in (6, 7, 8):
            self.clear_highlights()

        if idx == 4:
            self.wire_substep = 0
            self.wire_in_progress = True
            self.wires_visible = 0
            self.refresh_display()
            self.root.after(50, self.advance_wires)
        else:
            self.wire_in_progress = False

        self.phase_start = time.time()
        if self._speech_ended:
            self.phase_remaining = dur + 500
            self.schedule_next(self.phase_remaining)
        else:
            self.phase_remaining = dur + 30000
            self.schedule_next(self.phase_remaining)

    def _on_speech_ended(self, phase_idx):
        if phase_idx != self._speech_phase_idx:
            return
        if self._in_transition or self.stopped or self.paused:
            return
        self._speech_ended = True
        if self.phase_timer:
            self.root.after_cancel(self.phase_timer)
            self.phase_timer = None
        elapsed = (time.time() - self.phase_start) * 1000
        if elapsed >= self._min_phase_dur_ms:
            self.root.after_idle(self.on_timeout)
        else:
            remaining = int(self._min_phase_dur_ms - elapsed + 300)
            self.schedule_next(remaining)

    def schedule_next(self, delay_ms):
        if self.phase_timer:
            self.root.after_cancel(self.phase_timer)
        if self._timer_update_id:
            self.root.after_cancel(self._timer_update_id)
            self._timer_update_id = None
        self._schedule_time = delay_ms
        self.phase_timer = self.root.after(delay_ms, self.on_timeout)
        self.timer_label.config(text="%d\u0441" % (max(1, delay_ms//1000)))
        self._update_timer_display()

    def _update_timer_display(self):
        if self.paused or self.stopped or self.phase_timer is None:
            self._timer_update_id = None
            return
        elapsed = (time.time() - self.phase_start) * 1000
        remaining = max(0, self._schedule_time - elapsed)
        self.timer_label.config(text="%d\u0441" % (max(1, remaining//1000)))
        self._timer_update_id = self.root.after(500, self._update_timer_display)

    def on_timeout(self):
        self.phase_timer = None
        self.highlight_junction_points(False)
        if self._in_transition:
            self._in_transition = False
            try:
                self.enter_phase(self.phase_idx + 1)
            except Exception:
                pass
            return
        if self.wire_in_progress:
            self.wire_in_progress = False
            if self.wire_timer:
                self.root.after_cancel(self.wire_timer)
                self.wire_timer = None
            self.wires_visible = len(self.wire_items)
            try:
                self.refresh_display()
            except Exception:
                pass
        self._in_transition = True
        self.phase_remaining = 600
        self.phase_start = time.time()
        self.status_label.config(text="\u23ec")
        self.schedule_next(600)

    def advance_wires(self):
        if self.paused:
            return
        try:
            if self.wire_substep >= len(self.wire_items):
                self.wire_in_progress = False
                self.wires_visible = len(self.wire_items)
                return
            self.wire_substep += 1
            self.wires_visible = self.wire_substep
            self.refresh_display()
            if self.wire_substep < len(self.wire_items):
                remaining = len(self.wire_items) - self.wire_substep
                delay = 200 if remaining > 6 else 160
                self.wire_timer = self.root.after(delay, self.advance_wires)
            else:
                self.wire_in_progress = False
                self.wires_visible = len(self.wire_items)
        except Exception:
            self.wire_in_progress = False
            self.wires_visible = len(self.wire_items)

    def _sf(self, size):
        return max(1, round(size * self._scale / _REF_SCALE))

    def clear_highlights(self):
        self.highlight_junction_points(False)
        self._clear_formula_highlights()
        for item in self.canvas.find_withtag("hl"):
            self.canvas.delete(item)

    def _hl_frame(self, x1, y1, x2, y2, colour):
        s = self._scale
        self.canvas.create_rectangle(
            round((x1-6)*s), round((y1-6)*s), round((x2+6)*s), round((y2+6)*s),
            outline=colour, width=max(1, round(3*s)),
            dash=(round(6*s), round(3*s)), tags="hl"
        )

    def _hl_text(self, text, colour):
        s = self._scale
        self.canvas.create_text(
            round(620*s), round(750*s), text=text,
            font=("Arial", self._sf(18), "bold"), fill=colour,
            anchor="center", tags="hl"
        )

    def highlight_property1(self):
        self.clear_highlights()
        for cx,cy in [(PA1_X+DX, PA1_Y+DY), (PA2_X+DX, PA2_Y+DY), (PA3_X+DX, PA3_Y+DY)]:
            self._hl_frame(cx-45, cy-45, cx+45, cy+45, "#E65100")
        self._hl_text("I\u043e\u0431\u0449 = I\u2081 + I\u2082 = 120 + 60 = 180 \u043c\u0410", "#D32F2F")

    def highlight_property2(self):
        self.clear_highlights()
        self._hl_frame(R1_L+DX-10, R1_Y+DY-32, R1_R+DX+10, R1_Y+DY+32, "#1565C0")
        self._hl_frame(R2_L+DX-10, R2_Y+DY-32, R2_R+DX+10, R2_Y+DY+32, "#1565C0")
        self._hl_text("R\u2082/R\u2081 = I\u2081/I\u2082 = 200/100 = 120/60 = 2", "#1565C0")

    def highlight_property3(self):
        self.clear_highlights()
        self._hl_frame(R1_L+DX-10, R1_Y+DY-32, R1_R+DX+10, R1_Y+DY+32, "#E65100")
        self._hl_frame(R2_L+DX-10, R2_Y+DY-32, R2_R+DX+10, R2_Y+DY+32, "#E65100")
        self._hl_text("R\u044d\u043a\u0432 = 66,7 \u041e\u043c  <  100 \u041e\u043c = R\u2081", "#E65100")

    def highlight_junction_points(self, visible=True):
        for item in getattr(self, '_junction_hl_items', []):
            self.canvas.delete(item)
        self._junction_hl_items = []
        if not visible:
            return
        for jx, jy in JUNCTION_DOTS:
            dot = self.canvas.create_oval(
                jx-8, jy-8, jx+8, jy+8,
                fill="#212121", outline="", tags="hl_junc"
            )
            self._junction_hl_items.append(dot)
        self.canvas.tag_raise("speech")

    def _clear_formula_highlights(self):
        for item in self.canvas.find_withtag("f_hl"):
            self.canvas.delete(item)

    def _highlight_formula_data(self, phase_key, sub_idx):
        self._clear_formula_highlights()
        s = self._scale
        if phase_key == "step5":
            row_map = {0: 2, 1: 3, 2: 4, 3: 5, 4: 6}
            if sub_idx in row_map:
                ri = row_map[sub_idx]
                if ri < len(self.f5_items):
                    bbox = self.canvas.bbox(self.f5_items[ri])
                    if bbox:
                        x1, y1, x2, y2 = bbox
                        pad = round(4 * s)
                        self.canvas.create_rectangle(
                            x1 - pad, y1 - pad, x2 + pad, y2 + pad,
                            fill="#FFF9C4", outline="", tags="f_hl"
                        )
                        self.canvas.tag_lower("f_hl")
        elif phase_key in ("step6_p1", "step6_p2", "step6_p3"):
            grp_idx = {"step6_p1": 0, "step6_p2": 1, "step6_p3": 2}.get(phase_key, 0)
            if grp_idx < len(self.f6_groups):
                for li in (0, 1):
                    if li < len(self.f6_groups[grp_idx]):
                        bbox = self.canvas.bbox(self.f6_groups[grp_idx][li])
                        if bbox:
                            x1, y1, x2, y2 = bbox
                            pad = round(4 * s)
                            self.canvas.create_rectangle(
                                x1 - pad, y1 - pad, x2 + pad, y2 + pad,
                                fill="#E3F2FD", outline="", tags="f_hl"
                            )
                            self.canvas.tag_lower("f_hl")

    def highlight_r2_table_row(self, row_idx):
        cols = 4
        for i, item in enumerate(self.f7_row_items):
            ri = i // cols
            fc = "#D32F2F" if ri == row_idx else C_TEXT
            fw = "bold" if ri == row_idx else "normal"
            self.canvas.itemconfig(item, fill=fc, font=("Arial", self._sf(23), fw))
        self._clear_formula_highlights()
        s = self._scale
        for i, item in enumerate(self.f7_row_items):
            ri = i // cols
            if ri == row_idx:
                bbox = self.canvas.bbox(item)
                if bbox:
                    x1, y1, x2, y2 = bbox
                    pad = round(5 * s)
                    self.canvas.create_rectangle(
                        x1 - pad, y1 - pad, x2 + pad, y2 + pad,
                        fill="#E8F5E9", outline="", tags="f_hl"
                    )
                    self.canvas.tag_lower("f_hl")

    def update_progress(self, idx):
        total = len(self.PHASES)
        frac = (idx + 1) / total if idx >= 0 else 0
        bw = max(self.prog.winfo_width() - 4, 400)
        fw = int(bw * frac)
        self.prog.coords(self.prog_fill, 2, 2, 2 + fw, 16)

    def on_play(self):
        if not self.stopped:
            return
        self.stopped = False
        self.paused = False
        self.btn_play.config(state="disabled")
        self.btn_pause.config(state="normal", text="  \u23f8 ПАУЗА  ")
        self.btn_reset.config(state="normal")
        self.show_photo_panel(True)
        self.show_cube(True)
        self.enter_phase(0)

    def on_pause(self):
        if self.stopped:
            return
        if not self.paused:
            self.paused = True
            self.btn_pause.config(text="  \u25b6 ПРОДОЛЖИТЬ  ")
            self.tts.pause()
            self._stop_video()
            self._show_poster_frame()
            if self.phase_timer:
                self.root.after_cancel(self.phase_timer)
                self.phase_timer = None
            if self._timer_update_id:
                self.root.after_cancel(self._timer_update_id)
                self._timer_update_id = None
            if self.wire_timer:
                self.root.after_cancel(self.wire_timer)
                self.wire_timer = None
            if self._speech_delay_id:
                self.root.after_cancel(self._speech_delay_id)
                self._speech_delay_id = None
            self._sub_phase_remaining.clear()
            for tid, deadline, key, sub_idx in self._sub_phase_timers:
                self.root.after_cancel(tid)
                remaining = max(0, (deadline - time.time()) * 1000)
                if remaining > 50:
                    self._sub_phase_remaining.append((remaining, key, sub_idx))
                else:
                    self._sub_phase_remaining.append((0, key, sub_idx))
            self._sub_phase_timers.clear()
            elapsed = (time.time() - self.phase_start) * 1000
            self.phase_remaining = max(0, self.phase_remaining - elapsed)
            self.timer_label.config(text="\u23f8")
        else:
            self.paused = False
            self.btn_pause.config(text="  \u23f8 ПАУЗА  ")
            self.tts.resume()
            self._start_video()
            self.phase_start = time.time()
            if self.wire_in_progress:
                self.wire_timer = self.root.after(50, self.advance_wires)
            for remaining, key, sub_idx in self._sub_phase_remaining:
                if remaining > 50:
                    deadline = time.time() + remaining / 1000.0
                    tid = self.root.after(int(remaining), lambda k=key, si=sub_idx: self._activate_sub_group(k, si))
                    self._sub_phase_timers.append((tid, deadline, key, sub_idx))
                else:
                    self._activate_sub_group(key, sub_idx)
            self._sub_phase_remaining.clear()
            self.schedule_next(int(self.phase_remaining))

    def on_reset(self):
        self.stopped = True
        self.paused = False
        self.tts.stop()
        self._stop_video()
        if self._video_cap:
            self._video_cap.release()
            self._video_cap = None
        if self.phase_timer:
            self.root.after_cancel(self.phase_timer)
            self.phase_timer = None
        if self._timer_update_id:
            self.root.after_cancel(self._timer_update_id)
            self._timer_update_id = None
        if self.wire_timer:
            self.root.after_cancel(self.wire_timer)
            self.wire_timer = None
        if self._speech_delay_id:
            self.root.after_cancel(self._speech_delay_id)
            self._speech_delay_id = None
        self.wire_in_progress = False
        for tid, _, _, _ in self._sub_phase_timers:
            self.root.after_cancel(tid)
        self._sub_phase_timers.clear()
        self._sub_phase_remaining.clear()
        self._in_transition = False
        self.clear_highlights()
        self.highlight_junction_points(False)
        self.hide_all()
        self._open_video()
        self._show_poster_frame()
        self.btn_play.config(state="normal")
        self.btn_pause.config(state="disabled", text="  \u23f8 ПАУЗА  ")
        self.btn_reset.config(state="disabled")
        self.step_label.config(text="")
        self.status_label.config(text="Нажмите «СТАРТ» для начала")
        self.timer_label.config(text="")
        self.update_progress(-1)
        self.hide_speech()

    def on_step_selected(self, event):
        sel = self.step_combo.get()
        for idx, (_, _, s) in enumerate(self.PHASES):
            if s == sel:
                self.tts.stop()
                self._stop_video()
                if self.phase_timer:
                    self.root.after_cancel(self.phase_timer)
                    self.phase_timer = None
                if self._timer_update_id:
                    self.root.after_cancel(self._timer_update_id)
                    self._timer_update_id = None
                if self.wire_timer:
                    self.root.after_cancel(self.wire_timer)
                    self.wire_timer = None
                if self._speech_delay_id:
                    self.root.after_cancel(self._speech_delay_id)
                    self._speech_delay_id = None
                for tid, _, _, _ in self._sub_phase_timers:
                    self.root.after_cancel(tid)
                self._sub_phase_timers.clear()
                self._sub_phase_remaining.clear()
                self._in_transition = False
                self.wire_in_progress = False
                self.clear_highlights()
                self.highlight_junction_points(False)
                self.hide_all()
                self.wires_visible = len(self.wire_items)
                self.stopped = False
                self.paused = False
                self.btn_play.config(state="disabled")
                self.btn_pause.config(state="normal", text="  \u23f8 ПАУЗА  ")
                self.btn_reset.config(state="normal")
                self.show_photo_panel(True)
                self.show_cube(True)
                self.enter_phase(idx)
                break

    def _show_poster_frame(self):
        if self._video_poster:
            self.canvas.itemconfig(self.cube_img_id, image=self._video_poster)
            self._video_current_photo = self._video_poster
        elif self._video_cap and self._video_cap.isOpened():
            self._video_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ret, frame = self._video_cap.read()
            if ret:
                try:
                    frame_rgba = self._frame_to_rgba(frame)
                    pil_img = Image.fromarray(frame_rgba, 'RGBA')
                    cs = max(1, round(502 * self._scale / _REF_SCALE))
                    pil_img = pil_img.resize((cs, cs), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(pil_img)
                    self.canvas.itemconfig(self.cube_img_id, image=photo)
                    self._video_current_photo = photo
                except:
                    pass

    def show_final(self):
        self.status_label.config(
            text="Видеоурок завершён! Нажмите «СБРОС» для повтора."
        )
        self.timer_label.config(text="\u2714")
        self.step_label.config(text="Шаг 9 из 9")
        self.btn_pause.config(state="disabled")
        self._stop_video()
        self._show_poster_frame()
        if self._timer_update_id:
            self.root.after_cancel(self._timer_update_id)
            self._timer_update_id = None
        self.show_speech("Видеоурок завершён! Спасибо за внимание.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ResistorCartoon(root)

    def _on_closing():
        app.tts.stop()
        app.stopped = True
        if app.phase_timer:
            root.after_cancel(app.phase_timer)
        if app.wire_timer:
            root.after_cancel(app.wire_timer)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_closing)

    try:
        root.mainloop()
    except SystemExit:
        pass
    except Exception:
        import traceback
        traceback.print_exc()
        try:
            root.destroy()
        except Exception:
            pass
