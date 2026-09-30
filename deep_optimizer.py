import sys
import os
import re
import json
import subprocess
import time
import threading
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QFileDialog, QMessageBox, QLineEdit, QTextEdit,
                             QProgressBar, QSlider, QComboBox)
from PyQt6.QtCore import QThread, pyqtSignal, Qt

# --- pywin32 ---
try:
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("WARNING: pywin32 is not installed. Run: pip install pywin32")


# =====================================================================
#                          OUTPUT FORMATS
# =====================================================================
OUTPUT_FORMATS = {
    "FBX (*.fbx)":     {"ext": ".fbx"},
    "OBJ (*.obj)":     {"ext": ".obj"},
    "STL (*.stl)":     {"ext": ".stl"},
    "3DS (*.3ds)":     {"ext": ".3ds"},
    "Collada (*.dae)": {"ext": ".dae"},
    "Alembic (*.abc)": {"ext": ".abc"},
    "glTF (*.gltf)":   {"ext": ".gltf"},
    "DXF (*.dxf)":     {"ext": ".dxf"},
    "DWG (*.dwg)":     {"ext": ".dwg"},
    "VRML97 (*.wrl)":  {"ext": ".wrl"},
    "IGES (*.igs)":    {"ext": ".igs"},
    "ASE (*.ase)":     {"ext": ".ase"},
    "SAT (*.sat)":     {"ext": ".sat"},
}
DEFAULT_FORMAT = "FBX (*.fbx)"


# =====================================================================
#                          TRANSLATIONS
# =====================================================================
TRANSLATIONS = {
    "fa": {
        "window_title": "بهینه‌ساز عمیق 3ds Max",
        "lang_button": "🌐 English",
        "select_root": "پوشه ریشه (Root) را انتخاب کنید:",
        "browse_button": "📁 انتخاب پوشه و جستجوی زیرپوشه‌ها",
        "format_label": "فرمت خروجی:",
        "percent_label": "درصد بهینه‌سازی (Vertex Percent):",
        "percent_hint": "ℹ️ درصد کمتر = کاهش وجه بیشتر (۱٪ حداکثر کاهش، ۱۰۰٪ بدون تغییر).",
        "start_button": "▶️ شروع عملیات تبدیل سراسری",
        "cancel_button": "⛔ لغو عملیات",
        "progress_format": "{done} / {total} فایل",
        "log_label": "گزارش زنده:",
        "msg_no_max_title": "خطا",
        "msg_no_max": "نرم‌افزار 3ds Max پیدا نشد.",
        "msg_done_title": "پایان",
        "msg_error_title": "خطا",
        "log_scanning": ">>> در حال اسکن پوشه‌ها و فراخوانی 3ds Max...",
        "log_selected_format": ">>> فرمت خروجی انتخابی: {format}",
        "log_selected_percent": ">>> درصد بهینه‌سازی انتخابی: {percent}%",
        "log_pywin32_missing": "⚠️ pywin32 نصب نیست؛ دیالوگ‌های V-Ray بسته نمی‌شوند.",
        "log_watcher_started": ">>> DialogWatcher فعال شد.",
        "log_watcher_dialog_closed": ">>> دیالوگ «{title}» بسته شد (کلیک روی «{button}»)",
        "log_cancel_requested": ">>> درخواست لغو ارسال شد...",
        "log_cancelled": "عملیات توسط کاربر لغو شد.",
        "log_no_max_files": "هیچ فایل .max پیدا نشد.",
        "log_finished": "پردازش {done} از {total} فایل به اتمام رسید.\nفرمت خروجی: {format}\nدرصد بهینه‌سازی: {percent}%\nگزارش در پوشه اصلی ذخیره شد.",
        "log_timeout": "عملیات به دلیل تجاوز از زمان مجاز ({minutes} دقیقه) متوقف شد.\n{done} از {total} فایل پردازش شده بود.",
        "log_max_error": "3ds Max با کد خطای {code} بسته شد.\n{done} از {total} فایل پردازش شده بود.",
        "log_system_error": "خطای سیستم: {error}",
        "report_title": "=== گزارش عملیات جستجوی عمیق ===",
        "report_files_found": "تعداد فایل‌های یافته شده: {count}",
        "report_target_percent": "درصد بهینه‌سازی هدف: {percent}",
        "report_output_format": "فرمت خروجی: {format}",
        "select_folder_dialog": "انتخاب پوشه اصلی",
    },
    "en": {
        "window_title": "3ds Max Deep Optimizer",
        "lang_button": "🌐 فارسی",
        "select_root": "Select root folder:",
        "browse_button": "📁 Browse & scan subfolders",
        "format_label": "Output format:",
        "percent_label": "Optimization percentage (Vertex Percent):",
        "percent_hint": "ℹ️ Lower percentage = more reduction (1% = max, 100% = no change).",
        "start_button": "▶️ Start global conversion",
        "cancel_button": "⛔ Cancel operation",
        "progress_format": "{done} / {total} files",
        "log_label": "Live log:",
        "msg_no_max_title": "Error",
        "msg_no_max": "3ds Max not found.",
        "msg_done_title": "Done",
        "msg_error_title": "Error",
        "log_scanning": ">>> Scanning folders and launching 3ds Max...",
        "log_selected_format": ">>> Selected output format: {format}",
        "log_selected_percent": ">>> Selected optimization percentage: {percent}%",
        "log_pywin32_missing": "⚠️ pywin32 is not installed; V-Ray dialogs won't be closed.",
        "log_watcher_started": ">>> DialogWatcher started.",
        "log_watcher_dialog_closed": ">>> Dialog «{title}» closed (clicked «{button}»)",
        "log_cancel_requested": ">>> Cancel requested...",
        "log_cancelled": "Operation cancelled by user.",
        "log_no_max_files": "No .max files found.",
        "log_finished": "Processed {done} of {total} files.\nOutput format: {format}\nOptimization: {percent}%\nReport saved in the root folder.",
        "log_timeout": "Operation stopped due to timeout ({minutes} min).\n{done} of {total} files were processed.",
        "log_max_error": "3ds Max closed with error code {code}.\n{done} of {total} files were processed.",
        "log_system_error": "System error: {error}",
        "report_title": "=== Deep Search Operation Report ===",
        "report_files_found": "Files found: {count}",
        "report_target_percent": "Target optimization percent: {percent}",
        "report_output_format": "Output format: {format}",
        "select_folder_dialog": "Select root folder",
    }
}


# =====================================================================
#                          SETTINGS
# =====================================================================
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".deep_optimizer_settings.json")


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return {
                "language": data.get("language", "fa"),
                "format": data.get("format", DEFAULT_FORMAT),
                "percent": data.get("percent", 10),
            }
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {"language": "fa", "format": DEFAULT_FORMAT, "percent": 10}


def save_settings(lang, fmt, percent):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump({"language": lang, "format": fmt, "percent": percent},
                      f, ensure_ascii=False, indent=2)
    except OSError:
        pass


PER_FILE_TIMEOUT_SEC = 180
MIN_TOTAL_TIMEOUT_SEC = 600

_LOG_LINE_PATTERN = re.compile(
    r"^(OK|NO GEOMETRY|FAILED TO LOAD|ERROR|NO REDUCTION|EXPORT ERROR):",
    re.MULTILINE
)


# =====================================================================
#                          DialogWatcher
# =====================================================================
class DialogWatcher(threading.Thread):
    TARGET_TITLES = [
        "V-Ray warning", "V-Ray Warning", "V-Ray Message", "V-Ray error",
        "V-Ray", "Chaos V-Ray", "Chaos", "V-Ray License",
    ]
    BUTTON_PRIORITY = ["&No", "No", "&NO", "NO", "&OK", "OK", "&Yes", "Yes"]

    def __init__(self, interval=0.8, log_callback=None, tr=None):
        super().__init__(daemon=True)
        self.interval = interval
        self._stop_event = threading.Event()
        self._log_callback = log_callback
        self.tr = tr if tr else (lambda key, **kw: key)

    def stop(self):
        self._stop_event.set()

    def _log(self, msg):
        if self._log_callback:
            try:
                self._log_callback(msg)
            except Exception:
                pass

    def _click_child_button(self, parent_hwnd, button_texts):
        clicked = {"value": False, "text": None}

        def enum_child(child_hwnd, _):
            if clicked["value"]:
                return False
            try:
                cls = win32gui.GetClassName(child_hwnd)
                if cls != "Button":
                    return True
                text = win32gui.GetWindowText(child_hwnd).strip()
                if text in button_texts:
                    win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                    clicked["value"] = True
                    clicked["text"] = text
                    return False
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(parent_hwnd, enum_child, None)
        except Exception:
            pass
        return clicked

    def run(self):
        if not WIN32_AVAILABLE:
            return
        while not self._stop_event.is_set():
            try:
                found_hwnds = []

                def enum_top(hwnd, _):
                    try:
                        if not win32gui.IsWindowVisible(hwnd):
                            return True
                        wt = win32gui.GetWindowText(hwnd)
                        if not wt:
                            return True
                        for title in self.TARGET_TITLES:
                            if title.lower() in wt.lower():
                                found_hwnds.append((hwnd, wt))
                                break
                    except Exception:
                        pass
                    return True

                try:
                    win32gui.EnumWindows(enum_top, None)
                except Exception:
                    pass

                for hwnd, wt in found_hwnds:
                    for btn in self.BUTTON_PRIORITY:
                        res = self._click_child_button(hwnd, [btn])
                        if res["value"]:
                            self._log(self.tr("log_watcher_dialog_closed",
                                              title=wt, button=res['text']))
                            time.sleep(0.3)
                            break
            except Exception:
                pass
            self._stop_event.wait(self.interval)


# =====================================================================
#                          ConversionWorker
# =====================================================================
class ConversionWorker(QThread):
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(int, int)
    log_message = pyqtSignal(str)

    def __init__(self, max_exe, root_folder, target_percent, output_format, lang="fa"):
        super().__init__()
        self.max_exe = max_exe
        self.root_folder = root_folder
        self.target_percent = target_percent
        self.output_format = output_format
        self.lang = lang
        self._cancel_requested = False

    def _tr(self, key, **kwargs):
        text = TRANSLATIONS.get(self.lang, TRANSLATIONS["fa"]).get(key, key)
        if kwargs:
            try:
                return text.format(**kwargs)
            except (KeyError, IndexError):
                return text
        return text

    def cancel(self):
        self._cancel_requested = True

    def run(self):
        watcher = None
        try:
            max_files = []
            for root, dirs, files in os.walk(self.root_folder):
                for file in files:
                    if file.lower().endswith(".max"):
                        full_path = os.path.join(root, file).replace("\\", "/")
                        max_files.append(full_path)

            if not max_files:
                self.finished.emit(False, self._tr("log_no_max_files"))
                return

            fmt_info = OUTPUT_FORMATS.get(self.output_format, OUTPUT_FORMATS[DEFAULT_FORMAT])
            out_ext = fmt_info["ext"]

            files_array = "#(" + ", ".join([f'"{f}"' for f in max_files]) + ")"
            log_file_path = os.path.join(self.root_folder, "deep_process_report.txt").replace("\\", "/")

            report_title = self._tr("report_title")
            report_files = self._tr("report_files_found", count="%")
            report_percent = self._tr("report_target_percent", percent="%")
            report_format = self._tr("report_output_format", format=self.output_format)

            ms_content = f"""
(
    local filesList = {files_array}
    local logPath = "{log_file_path}"
    local logFile = createFile logPath
    local targetPercent = {self.target_percent}
    local outputExt = "{out_ext}"

    format "{report_title}\\n" to:logFile
    format "{report_files}\\n" filesList.count to:logFile
    format "{report_percent}\\n" targetPercent to:logFile
    format "{report_format}\\n\\n" to:logFile

    fn countFaces node = (
        local snap = snapshotAsMesh node
        local n = snap.numFaces
        delete snap
        n
    )

    -- تابع خروجی: FBX با تنظیمات خاص، بقیه فرمت‌ها با تشخیص خودکار Max
    fn doExport outPath selOnly = (
        local ok = false
        if outputExt == ".fbx" then (
            try ( FBXExporterSetParam "ResetExport" ) catch ()
            try (
                exportFile outPath #noPrompt selectedOnly:selOnly using:FBXEXP
                ok = true
            ) catch (
                ok = false
            )
        ) else (
            try (
                exportFile outPath #noPrompt selectedOnly:selOnly
                ok = true
            ) catch (
                ok = false
            )
        )
        ok
    )

    for f in filesList do (
        try (
            if (loadMaxFile f quiet:true) then (

                -- غیرفعال‌سازی V-Ray
                try ( renderers.current = Default_Scanline_Renderer() ) catch ()
                try ( pluginManager.unloadPlugin "vray" ) catch ()
                try ( pluginManager.unloadPlugin "V_Ray_Adv" ) catch ()
                try ( pluginManager.unloadPlugin "V_Ray_GPU" ) catch ()

                local baseDir = getFilenamePath f
                local fileName = getFilenameFile f
                local geo = for g in geometry where (isKindOf g GeometryClass) collect g

                if geo.count > 0 then (
                    for g in geo do (
                        for mod in g.modifiers do (
                            if (classOf mod == ProOptimizer) do deleteModifier g mod
                        )
                    )

                    -- ====== ۱. خروجی Normal ======
                    local normalOut = baseDir + fileName + "_Normal" + outputExt
                    if not (doExport normalOut false) then (
                        format "EXPORT ERROR (normal): %\\n" f to:logFile
                    )

                    local facesBefore = 0
                    for g in geo do facesBefore += countFaces g

                    -- ====== ۲. بهینه‌سازی ======
                    setCommandPanelTaskMode #modify
                    for g in geo do (
                        try (
                            select g
                            local opt = ProOptimizer()
                            addModifier g opt
                            modPanel.setCurrentObject opt
                            opt.Calculate = true
                            forceCompleteRedraw()
                            opt.VertexPercent = targetPercent
                            forceCompleteRedraw()
                        ) catch (
                            format "ERROR (optimize): % -> %\\n" g.name (getCurrentException()) to:logFile
                        )
                    )

                    local facesAfter = 0
                    for g in geo do facesAfter += countFaces g

                    -- ====== ۳. خروجی Optimized ======
                    if (facesAfter < facesBefore) then (
                        select geo
                        local optOut = baseDir + fileName + "_Optimized" + outputExt
                        local okOpt = doExport optOut true
                        if not okOpt then (
                            okOpt = doExport optOut false
                        )
                        if okOpt then (
                            format "OK: % | Faces: % -> %\\n" f facesBefore facesAfter to:logFile
                        ) else (
                            format "EXPORT ERROR (optimized): %\\n" f to:logFile
                        )
                    ) else (
                        format "NO REDUCTION: % | Faces: % -> %\\n" f facesBefore facesAfter to:logFile
                    )
                ) else (
                    format "NO GEOMETRY: % \\n" f to:logFile
                )
            ) else (
                format "FAILED TO LOAD: % \\n" f to:logFile
            )
        ) catch (
            format "ERROR: % -> %\\n" f (getCurrentException()) to:logFile
        )
    )
    close logFile
    quitMax #noPrompt
)
            """

            ms_script_path = os.path.join(os.environ['TEMP'], "max_deep_search.ms")
            with open(ms_script_path, "w", encoding="utf-16") as fs:
                fs.write(ms_content)

            total_count = len(max_files)
            total_timeout = max(MIN_TOTAL_TIMEOUT_SEC, total_count * PER_FILE_TIMEOUT_SEC)

            if WIN32_AVAILABLE:
                watcher = DialogWatcher(
                    interval=0.8,
                    log_callback=lambda msg: self.log_message.emit(msg),
                    tr=self._tr
                )
                watcher.start()
                self.log_message.emit(self._tr("log_watcher_started"))

            process = subprocess.Popen(
                [self.max_exe, "-silent", "-U", "MAXScript", ms_script_path]
            )

            start_time = time.monotonic()
            timed_out = False

            while True:
                done_count = self._count_processed(log_file_path)
                self.progress.emit(done_count, total_count)

                if self._cancel_requested:
                    process.kill()
                    if watcher:
                        watcher.stop()
                    self.finished.emit(False, self._tr("log_cancelled"))
                    return

                try:
                    process.wait(timeout=1)
                    break
                except subprocess.TimeoutExpired:
                    pass

                if time.monotonic() - start_time > total_timeout:
                    timed_out = True
                    process.kill()
                    break

            if watcher:
                watcher.stop()
                time.sleep(0.3)

            final_done = self._count_processed(log_file_path)
            self.progress.emit(final_done, total_count)

            if timed_out:
                self.finished.emit(False, self._tr("log_timeout",
                    minutes=total_timeout // 60,
                    done=final_done, total=total_count))
                return

            if process.returncode != 0:
                self.finished.emit(False, self._tr("log_max_error",
                    code=process.returncode,
                    done=final_done, total=total_count))
                return

            self.finished.emit(True, self._tr("log_finished",
                done=final_done, total=total_count,
                format=self.output_format,
                percent=self.target_percent))

        except Exception as e:
            if watcher:
                try:
                    watcher.stop()
                except Exception:
                    pass
            self.finished.emit(False, self._tr("log_system_error", error=str(e)))

    @staticmethod
    def _count_processed(log_file_path):
        try:
            with open(log_file_path, "r", encoding="utf-8", errors="ignore") as lf:
                content = lf.read()
            return len(_LOG_LINE_PATTERN.findall(content))
        except (FileNotFoundError, OSError):
            return 0


# =====================================================================
#                          DeepOptimizerApp
# =====================================================================
class DeepOptimizerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.max_exe_path = r"C:\Program Files\Autodesk\3ds Max 2023\3dsmax.exe"
        settings = load_settings()
        self.current_lang = settings["language"]
        self.initUI()
        self.apply_language()

        if settings["format"] in OUTPUT_FORMATS:
            self.format_combo.setCurrentText(settings["format"])
        self.percent_slider.setValue(settings["percent"])

    def tr(self, key, **kwargs):
        text = TRANSLATIONS.get(self.current_lang, TRANSLATIONS["fa"]).get(key, key)
        if kwargs:
            try:
                return text.format(**kwargs)
            except (KeyError, IndexError):
                return text
        return text

    def initUI(self):
        self.setMinimumSize(700, 620)
        self.setStyleSheet("""
            QWidget { background-color: #252526; color: #cccccc; font-family: 'Segoe UI'; }
            QPushButton { background-color: #007acc; color: white; border: none; padding: 10px; border-radius: 4px; }
            QPushButton:hover { background-color: #0062a3; }
            QPushButton:disabled { background-color: #3e3e42; color: #777; }
            QLineEdit { background-color: #333337; border: 1px solid #434346; padding: 6px; color: #9cdcfe; border-radius: 3px; }
            QComboBox { background-color: #333337; border: 1px solid #434346; padding: 6px; color: #9cdcfe; border-radius: 3px; }
            QComboBox:hover { border: 1px solid #007acc; }
            QComboBox QAbstractItemView { background-color: #2d2d30; color: #cccccc; selection-background-color: #007acc; }
            QTextEdit { background-color: #1e1e1e; border: 1px solid #3e3e42; color: #d7ba7d; }
            QProgressBar { background-color: #1e1e1e; border: 1px solid #3e3e42; color: #cccccc; text-align: center; height: 22px; border-radius: 3px; }
            QProgressBar::chunk { background-color: #007acc; border-radius: 3px; }
            QSlider::groove:horizontal { border: 1px solid #3e3e42; height: 8px; background: #1e1e1e; border-radius: 4px; }
            QSlider::handle:horizontal { background: #007acc; border: 1px solid #0062a3; width: 18px; margin: -6px 0; border-radius: 9px; }
            QSlider::handle:horizontal:hover { background: #1a8cd8; }
            QSlider::sub-page:horizontal { background: #007acc; border-radius: 4px; }
            QPushButton#langButton { background-color: #3e3e42; color: #4ec9b0; font-weight: bold; max-width: 130px; }
            QPushButton#langButton:hover { background-color: #505054; }
        """)

        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.addStretch()
        self.btn_lang = QPushButton()
        self.btn_lang.setObjectName("langButton")
        self.btn_lang.clicked.connect(self.toggle_language)
        top_row.addWidget(self.btn_lang)
        self.main_layout.addLayout(top_row)

        self.lbl_select_root = QLabel()
        self.main_layout.addWidget(self.lbl_select_root)

        self.path_input = QLineEdit()
        self.main_layout.addWidget(self.path_input)

        self.btn_browse = QPushButton()
        self.btn_browse.clicked.connect(self.get_folder)
        self.main_layout.addWidget(self.btn_browse)

        self.lbl_format = QLabel()
        self.main_layout.addWidget(self.lbl_format)

        self.format_combo = QComboBox()
        for fmt_name in OUTPUT_FORMATS.keys():
            self.format_combo.addItem(fmt_name)
        self.format_combo.setCurrentText(DEFAULT_FORMAT)
        self.format_combo.currentTextChanged.connect(self.on_format_changed)
        self.main_layout.addWidget(self.format_combo)

        self.lbl_percent = QLabel()
        self.main_layout.addWidget(self.lbl_percent)

        slider_row = QHBoxLayout()
        self.percent_slider = QSlider(Qt.Orientation.Horizontal)
        self.percent_slider.setMinimum(1)
        self.percent_slider.setMaximum(100)
        self.percent_slider.setValue(10)
        self.percent_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.percent_slider.setTickInterval(10)
        self.percent_slider.valueChanged.connect(self.update_percent_label)
        slider_row.addWidget(self.percent_slider)

        self.percent_label = QLabel("10%")
        self.percent_label.setMinimumWidth(60)
        self.percent_label.setStyleSheet("color: #4ec9b0; font-weight: bold; font-size: 14px;")
        self.percent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        slider_row.addWidget(self.percent_label)
        self.main_layout.addLayout(slider_row)

        self.lbl_percent_hint = QLabel()
        self.lbl_percent_hint.setStyleSheet("color: #808080; font-size: 11px;")
        self.lbl_percent_hint.setWordWrap(True)
        self.main_layout.addWidget(self.lbl_percent_hint)

        self.btn_start = QPushButton()
        self.btn_start.setEnabled(False)
        self.btn_start.clicked.connect(self.start_conversion)
        self.main_layout.addWidget(self.btn_start)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.main_layout.addWidget(self.progress_bar)

        self.btn_cancel = QPushButton()
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_conversion)
        self.main_layout.addWidget(self.btn_cancel)

        self.lbl_log = QLabel()
        self.main_layout.addWidget(self.lbl_log)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.main_layout.addWidget(self.log_view)

        self.setLayout(self.main_layout)

    def apply_language(self):
        lang = self.current_lang
        is_rtl = (lang == "fa")
        layout_dir = Qt.LayoutDirection.RightToLeft if is_rtl else Qt.LayoutDirection.LeftToRight
        self.setLayoutDirection(layout_dir)

        self.setWindowTitle(self.tr("window_title"))
        self.btn_lang.setText(self.tr("lang_button"))
        self.lbl_select_root.setText(self.tr("select_root"))
        self.btn_browse.setText(self.tr("browse_button"))
        self.lbl_format.setText(self.tr("format_label"))
        self.lbl_percent.setText(self.tr("percent_label"))
        self.lbl_percent_hint.setText(self.tr("percent_hint"))
        self.btn_start.setText(self.tr("start_button"))
        self.btn_cancel.setText(self.tr("cancel_button"))
        self.lbl_log.setText(self.tr("log_label"))
        self.progress_bar.setFormat(self.tr("progress_format", done=0, total=0))

    def toggle_language(self):
        self.current_lang = "en" if self.current_lang == "fa" else "fa"
        self.apply_language()
        self._save_current_settings()

    def _save_current_settings(self):
        save_settings(self.current_lang,
                      self.format_combo.currentText(),
                      self.percent_slider.value())

    def on_format_changed(self, _):
        self._save_current_settings()

    def update_percent_label(self, value):
        self.percent_label.setText(f"{value}%")
        self._save_current_settings()

    def get_folder(self):
        folder = QFileDialog.getExistingDirectory(self, self.tr("select_folder_dialog"))
        if folder:
            self.path_input.setText(folder)
            self.btn_start.setEnabled(True)

    def start_conversion(self):
        if not os.path.exists(self.max_exe_path):
            for year in ["2024", "2025", "2022", "2021", "2020"]:
                alt = f"C:\\Program Files\\Autodesk\\3ds Max {year}\\3dsmax.exe"
                if os.path.exists(alt):
                    self.max_exe_path = alt
                    break
            else:
                QMessageBox.critical(self, self.tr("msg_no_max_title"), self.tr("msg_no_max"))
                return

        target_percent = self.percent_slider.value()
        output_format = self.format_combo.currentText()

        self.log_view.clear()
        self.log_view.append(self.tr("log_scanning"))
        self.log_view.append(self.tr("log_selected_format", format=output_format))
        self.log_view.append(self.tr("log_selected_percent", percent=target_percent))
        if not WIN32_AVAILABLE:
            self.log_view.append(self.tr("log_pywin32_missing"))

        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(self.tr("progress_format", done=0, total=0))

        self.worker = ConversionWorker(
            self.max_exe_path,
            self.path_input.text(),
            target_percent,
            output_format,
            lang=self.current_lang
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.process_complete)
        self.worker.log_message.connect(self.append_log)
        self.worker.start()

    def cancel_conversion(self):
        if hasattr(self, "worker") and self.worker.isRunning():
            self.worker.cancel()
            self.btn_cancel.setEnabled(False)
            self.log_view.append(self.tr("log_cancel_requested"))

    def update_progress(self, done, total):
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(done)
            self.progress_bar.setFormat(self.tr("progress_format", done=done, total=total))

    def append_log(self, msg):
        self.log_view.append(msg)

    def process_complete(self, success, message):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.log_view.append("")
        self.log_view.append(message)
        if success:
            QMessageBox.information(self, self.tr("msg_done_title"), message)
        else:
            QMessageBox.critical(self, self.tr("msg_error_title"), message)


# =====================================================================
#                          Main
# =====================================================================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DeepOptimizerApp()
    window.show()
    sys.exit(app.exec())