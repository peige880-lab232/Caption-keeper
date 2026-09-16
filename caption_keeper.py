"""Caption Keeper — local screen-caption logger for Microsoft Teams on Windows.

The program captures only a user-selected rectangle, OCRs it locally, and writes
deduplicated speaker-labelled captions to Markdown and JSON files.
"""

from __future__ import annotations

import json
import os
import queue
import re
import threading
import time
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

import mss
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from rapidocr_onnxruntime import RapidOCR


APP_DIR = Path.home() / "Documents" / "CaptionKeeper"


def clean_caption(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    # OCR occasionally leaves a standalone timestamp; it is not useful content.
    return "" if re.fullmatch(r"\d{1,2}:\d{2}(?::\d{2})?", text) else text


# These occur when an OCR word box contains both words, particularly after a
# Teams caption has just scrolled.  This is deliberately a small, conservative
# list: guessing arbitrary missing spaces is more harmful than leaving a rare
# technical term untouched.
COMMON_JOINED_WORDS = (
    "and", "are", "because", "can", "could", "does", "for", "from", "have",
    "not", "our", "that", "the", "their", "then", "this", "what", "with",
    "would", "you",
)


def restore_obvious_spaces(text: str) -> str:
    """Repair only very likely English word joins without inventing words."""
    for word in COMMON_JOINED_WORDS:
        # Avoid splitting a genuine short word or proper noun.  Requiring at
        # least three letters on both sides keeps this correction conservative.
        text = re.sub(rf"(?<=[a-z]{{3}}){word}(?=[a-z]{{3}})", f" {word} ", text,
                      flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()


def similar(a: str, b: str) -> bool:
    """Treat the scrolling overlap of two Teams caption frames as a duplicate."""
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    return SequenceMatcher(None, a.lower(), b.lower()).ratio() >= 0.86


def looks_like_speaker_name(text: str) -> bool:
    """Conservative recognition of the name line displayed above a Teams caption."""
    text = format_speaker_name(text)
    if not 2 <= len(text) <= 42 or re.search(r"[0-9.!?,;:，。！？]", text):
        return False
    # Chinese names are generally 2–6 Han characters. Permit a middle dot for
    # international names written in Chinese.
    if re.fullmatch(r"[\u3400-\u9fff·]{2,8}", text):
        return True
    # English display names normally have two or more title-cased name parts.
    parts = re.split(r"[\s'-]+", text)
    return len(parts) >= 2 and all(part and (part[0].isupper() or part.isupper()) for part in parts)


def format_speaker_name(text: str) -> str:
    """Restore spaces in a CamelCase display name, e.g. ScottGraybeal."""
    text = text.strip()
    return re.sub(r"(?<=[a-z])(?=[A-Z])", " ", text)


def split_known_speakers(lines: list[str], known_speakers: set[str]) -> list[str]:
    """Separate a known speaker label when OCR has removed its inner spaces."""
    output: list[str] = []
    for line in lines:
        # Prefer longer names so one name cannot split part of another.
        for speaker in sorted(known_speakers, key=len, reverse=True):
            compact = re.sub(r"\s+", "", speaker)
            if len(compact) < 5:
                continue
            pattern = re.compile(re.escape(compact), re.IGNORECASE)
            if not pattern.search(line):
                continue
            # A name embedded inside a caption is a boundary from an overlapping
            # screen frame, not spoken text.
            pieces = pattern.split(line)
            for index, piece in enumerate(pieces):
                if piece.strip():
                    output.append(piece.strip())
                if index < len(pieces) - 1:
                    output.append(speaker)
            break
        else:
            output.append(line)
    return output


def utterances_from_lines(lines: list[str], previous_speaker: str | None) -> tuple[list[dict[str, str]], str | None]:
    """Convert Teams' alternating name/content rows into speaker-labelled turns."""
    speaker = previous_speaker
    content: list[str] = []
    utterances: list[dict[str, str]] = []

    def emit():
        nonlocal content
        message = " ".join(content).strip()
        if message:
            utterances.append({"speaker": speaker or "未辨識說話者", "text": message})
        content = []

    for line in lines:
        if looks_like_speaker_name(line):
            emit()
            speaker = format_speaker_name(line)
        else:
            content.append(line)
    emit()
    return utterances, speaker


class RegionPicker(tk.Toplevel):
    def __init__(self, root: tk.Tk, callback):
        super().__init__(root)
        self.callback = callback
        self.start = None
        self.rect = None
        self.attributes("-fullscreen", True)
        self.attributes("-alpha", 0.25)
        self.attributes("-topmost", True)
        self.configure(bg="black")
        self.canvas = tk.Canvas(self, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Escape>", lambda _event: self.cancel())
        self.focus_force()

    def on_press(self, event):
        self.start = (event.x_root, event.y_root)
        self.rect = self.canvas.create_rectangle(event.x, event.y, event.x, event.y,
                                                 outline="#57b6ff", width=3)

    def on_drag(self, event):
        if self.start and self.rect:
            self.canvas.coords(self.rect, self.start[0], self.start[1], event.x_root, event.y_root)

    def on_release(self, event):
        if not self.start:
            return
        x1, y1 = self.start
        x2, y2 = event.x_root, event.y_root
        left, top = min(x1, x2), min(y1, y2)
        width, height = abs(x2 - x1), abs(y2 - y1)
        if width >= 80 and height >= 25:
            self.callback({"left": left, "top": top, "width": width, "height": height})
        self.destroy()

    def cancel(self):
        self.master.deiconify()
        self.destroy()


class CaptionKeeper:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Caption Keeper — Teams 字幕記錄")
        self.root.geometry("780x660")
        self.root.minsize(650, 520)

        self.region = None
        self.running = False
        self.ocr = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.entries: list[dict[str, str]] = []
        self.known_speakers: set[str] = set()
        self.last_frame = ""
        self.current_speaker: str | None = None
        self.session_file: Path | None = None
        self.stop_event = threading.Event()
        self.capture_interval = 2
        # Teams normally reserves about 40–55 physical pixels for the avatar.
        # Start just to its right so name labels remain available to OCR.
        self.left_crop = tk.IntVar(value=52)
        self.capture_region = None
        self.interval = tk.IntVar(value=2)
        self.status = tk.StringVar(value="先選取 Teams 即時字幕的區域。")
        self._build_ui()
        self.root.after(250, self.drain_events)
        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Caption Keeper", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(outer, text="只讀取你框選的 Teams 即時字幕區域；所有紀錄只保存在你的電腦。",
                  wraplength=720).pack(anchor="w", pady=(2, 10))

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=4)
        self.pick_button = ttk.Button(controls, text="1. 選取字幕區域", command=self.pick_region)
        self.pick_button.pack(side="left")
        self.start_button = ttk.Button(controls, text="2. 開始記錄", command=self.toggle_recording, state="disabled")
        self.start_button.pack(side="left", padx=8)
        self.open_folder_button = ttk.Button(
            controls, text="開啟最新紀錄資料夾", command=self.open_recording_folder
        )
        self.open_folder_button.pack(side="left", padx=(0, 12))
        ttk.Label(controls, text="擷取頻率（秒）").pack(side="left", padx=(12, 2))
        ttk.Spinbox(controls, from_=1, to=10, width=4, textvariable=self.interval).pack(side="left")
        ttk.Label(controls, text="略過頭像左欄（px）").pack(side="left", padx=(12, 2))
        ttk.Spinbox(controls, from_=0, to=500, increment=10, width=5, textvariable=self.left_crop).pack(side="left")

        ttk.Label(outer, textvariable=self.status, foreground="#246b3b").pack(anchor="w", pady=(2, 7))
        self.preview = scrolledtext.ScrolledText(outer, height=18, wrap="word", state="disabled", font=("Segoe UI", 10))
        self.preview.pack(fill="both", expand=True)

    def pick_region(self):
        self.root.withdraw()
        def selected(region):
            self.region = region
            self.root.deiconify()
            self.start_button.config(state="normal")
            self.status.set(f"已選取字幕區域：{region['width']} × {region['height']}。請確認字幕面板保持在此位置。")
        picker = RegionPicker(self.root, selected)
        picker.protocol("WM_DELETE_WINDOW", lambda: (picker.destroy(), self.root.deiconify()))

    def toggle_recording(self):
        if self.running:
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        if not self.region:
            return
        APP_DIR.mkdir(parents=True, exist_ok=True)
        started = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.session_file = APP_DIR / f"teams-captions_{started}.md"
        self.entries, self.last_frame, self.current_speaker = [], "", None
        self.known_speakers = set()
        self._write_file(header=True)
        self.stop_event.clear()
        self.capture_interval = max(1, self.interval.get())
        trim = min(max(0, self.left_crop.get()), max(0, self.region["width"] - 80))
        self.capture_region = {
            "left": self.region["left"] + trim,
            "top": self.region["top"],
            "width": self.region["width"] - trim,
            "height": self.region["height"],
        }
        self.running = True
        self.start_button.config(text="停止記錄")
        self.status.set(f"記錄中：略過左側 {trim}px 的頭像欄；姓名與字幕會被保留。")
        threading.Thread(target=self.capture_loop, daemon=True).start()

    def open_recording_folder(self):
        """Open the folder that contains this session (or the normal save folder)."""
        folder = self.session_file.parent if self.session_file else APP_DIR
        folder.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(folder)  # type: ignore[attr-defined]  # Windows only
        except OSError as exc:
            messagebox.showerror("無法開啟資料夾", str(exc))

    def stop_recording(self):
        self.running = False
        self.stop_event.set()
        self.start_button.config(text="2. 開始記錄")
        self.status.set(f"已停止。已保存 {len(self.entries)} 則字幕至 {self.session_file}")

    def capture_loop(self):
        try:
            self.ocr = self.ocr or RapidOCR()
            with mss.mss() as screen:
                while not self.stop_event.is_set():
                    frame = screen.grab(self.capture_region)
                    image = Image.frombytes("RGB", frame.size, frame.rgb)
                    # A larger, autocontrasted image gives RapidOCR more detail
                    # on Teams' small anti-aliased caption font.  Sharpening is
                    # mild so it does not turn punctuation into letters.
                    image = ImageOps.grayscale(image).resize((image.width * 3, image.height * 3))
                    image = ImageOps.autocontrast(image, cutoff=1)
                    image = ImageEnhance.Contrast(image).enhance(1.35)
                    image = image.filter(ImageFilter.UnsharpMask(radius=1, percent=110, threshold=3))
                    result, _elapsed = self.ocr(image)
                    # Keep only credible text.  In particular, low-confidence
                    # character soup from avatars and caption-panel ornaments is
                    # much worse than a missed frame because it is written to the
                    # permanent log.  RapidOCR already orders the boxes top-down.
                    recognized: list[str] = []
                    if result:
                        for item in result:
                            text = clean_caption(item[1])
                            confidence = item[2] if len(item) >= 3 else 1.0
                            letters = sum(char.isalpha() for char in text)
                            noisy = sum(not (char.isalpha() or char.isspace() or char in ".,?!'-:#%") for char in text)
                            if not text or confidence < 0.62:
                                continue
                            if len(text) >= 10 and letters / len(text) < 0.45:
                                continue
                            if noisy > max(2, len(text) // 4):
                                continue
                            recognized.append(restore_obvious_spaces(text))
                    lines = split_known_speakers(recognized, self.known_speakers)
                    utterances, self.current_speaker = utterances_from_lines(lines, self.current_speaker)
                    for entry in utterances:
                        if entry["speaker"] != "未辨識說話者":
                            self.known_speakers.add(entry["speaker"])
                        seen = any(
                            old["speaker"] == entry["speaker"] and similar(old["text"], entry["text"])
                            for old in self.entries[-80:]
                        )
                        if not seen:
                            self.entries.append(entry)
                            self._write_file()
                            self.events.put(("caption", entry))
                    self.stop_event.wait(self.capture_interval)
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _write_file(self, header=False):
        if not self.session_file:
            return
        lines = ["# Teams 即時字幕紀錄", ""]
        for item in self.entries:
            lines.extend([f"**{item['speaker']}**", item["text"], ""])
        self.session_file.write_text("\n".join(lines) + "\n", encoding="utf-8")
        self.session_file.with_suffix(".json").write_text(json.dumps(self.entries, ensure_ascii=False, indent=2), encoding="utf-8")

    def drain_events(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == "caption":
                    self.preview.config(state="normal")
                    self.preview.insert("end", f"{data['speaker']}\n{data['text']}\n\n")
                    self.preview.see("end")
                    self.preview.config(state="disabled")
                else:
                    self.stop_recording()
                    messagebox.showerror("字幕擷取停止", str(data))
        except queue.Empty:
            pass
        self.root.after(250, self.drain_events)

    def close(self):
        self.stop_event.set()
        self.root.destroy()


if __name__ == "__main__":
    app = tk.Tk()
    CaptionKeeper(app)
    app.mainloop()

