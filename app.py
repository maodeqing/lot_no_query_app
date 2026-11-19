# -*- coding: utf-8 -*-
"""
批号批量查询器
================
从局域网共享 Excel 中按批号批量查询，结果自动复制到剪贴板。

主要特性：
- 配置文件外部化（config.ini），支持设置界面修改
- 网络路径带超时检查，避免 UI 卡死
- 查询在后台线程执行，主线程不阻塞
- 日志轮转（app.log）
- 输入实时校验
"""
import sys
import os
import re
import shutil
import logging
import datetime
import subprocess
import concurrent.futures
import configparser
from pathlib import Path
from logging.handlers import RotatingFileHandler

import pandas as pd
import pandas.errors

from PyQt5 import QtWidgets, QtCore, QtGui


# ===================== 路径处理（兼容 exe + 内置资源） =====================
def get_app_dir() -> Path:
    """返回 exe 同目录（打包后）或脚本所在目录（开发时）。"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def get_resource_path(relative_path: str) -> Path:
    """返回打包后 _MEIPASS 内资源路径，或开发时脚本目录下的路径。"""
    if getattr(sys, 'frozen', False):
        base = Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parent
    return base / relative_path


APP_DIR = get_app_dir()
CONFIG_FILE = APP_DIR / "config.ini"
LOG_FILE = APP_DIR / "app.log"
BUNDLED_CONFIG = get_resource_path("config.ini")
BUNDLED_HELP = get_resource_path("Help.txt")
BUNDLED_REQ = get_resource_path("requirements.txt")
STYLE_FILE = get_resource_path("styles.qss")


# ===================== 常量（消除魔术数字） =====================
DEFAULT_CONFIG = {
    "EXCEL": {
        "base_path": r"\\192.168.1.12\生管",
        "file_name": "KDJL.xls",
    },
    "QUERY": {
        "start_year": "2017",
    },
}
LOT_NO_MAX_LEN = 50
INVALID_FILE_CHARS = r'<>:"/\|?*'
NETWORK_TIMEOUT = 5          # 查询时网络检查超时（秒）
PREVIEW_TIMEOUT = 3          # 设置预览网络检查超时（秒）
DEBOUNCE_MS = 500            # 输入防抖（毫秒）
LOG_MAX_BYTES = 10 * 1024 * 1024   # 10 MB
LOG_BACKUP_COUNT = 5
WINDOW_WIDTH = 640
WINDOW_HEIGHT = 520
MIN_YEAR = 2000


# ===================== 日志（RotatingFileHandler） =====================
def setup_logger() -> logging.Logger:
    """配置日志：文件轮转 + 控制台输出。"""
    logger = logging.getLogger("BatchQuery")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    try:
        fh = RotatingFileHandler(
            LOG_FILE, maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT, encoding="utf-8"
        )
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except OSError as e:
        print(f"无法创建日志文件 {LOG_FILE}: {e}")

    sh = logging.StreamHandler()
    sh.setFormatter(formatter)
    logger.addHandler(sh)
    return logger


logger = setup_logger()


# ===================== 样式加载（外部优先，内置兜底） =====================
def load_stylesheet() -> str:
    """优先读 exe 同目录 styles.qss，其次读内置。失败返回空串。"""
    external = APP_DIR / "styles.qss"
    path = external if external.exists() else STYLE_FILE
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except OSError as e:
        logger.warning(f"加载样式失败 {path}: {e}")
        return ""


# ===================== 校验函数（统一返回 (bool, str)） =====================
def validate_path(path: str):
    """校验共享路径。返回 (是否有效, 错误信息)。"""
    if not path:
        return False, "路径不能为空"
    invalid = re.search(r'[<>:"|?*]', path)
    if invalid:
        return False, f"路径不能包含字符: {invalid.group()}"
    if not (path.startswith("\\\\") or (len(path) >= 2 and path[1] == ":")):
        return False, "路径格式不正确（应为 \\\\服务器\\共享 或 X:\\）"
    return True, ""


def validate_file_name(name: str):
    """校验 Excel 文件名。返回 (是否有效, 错误信息)。"""
    if not name:
        return False, "文件名不能为空"
    if any(c in INVALID_FILE_CHARS for c in name):
        return False, f"文件名不能包含字符: {INVALID_FILE_CHARS}"
    if not name.lower().endswith((".xls", ".xlsx")):
        return False, "文件名必须以 .xls 或 .xlsx 结尾"
    return True, ""


def validate_lot_no(text: str):
    """校验批号输入。返回 (是否有效, 错误信息)。"""
    if not text:
        return False, "批号不能为空"
    if len(text) > LOT_NO_MAX_LEN:
        return False, f"批号长度不能超过 {LOT_NO_MAX_LEN} 个字符"
    if not re.match(r'^[\w\-.,;\s]+$', text, re.UNICODE):
        return False, "批号只能包含字母、数字、-、_、.、空格、逗号、分号"
    return True, ""


def validate_year(text: str):
    """校验起始年份。返回 (是否有效, 错误信息)。"""
    if not text:
        return False, "起始年份不能为空"
    try:
        y = int(text)
    except ValueError:
        return False, "起始年份必须是整数"
    current = datetime.datetime.now().year
    if y < MIN_YEAR or y > current:
        return False, f"起始年份应在 {MIN_YEAR}–{current} 之间"
    return True, ""


# ===================== 网络路径超时检查 =====================
def check_network_path(path, timeout: float = NETWORK_TIMEOUT) -> bool:
    """在线程池里跑 Path.exists，带超时，避免网络共享无响应阻塞 UI。"""
    def _check():
        return Path(path).exists()

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(_check)
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            logger.warning(f"路径检查超时（>{timeout}s）：{path}")
            return False


# ===================== 资源释放管理器（记录释放的文件，退出时清理） =====================
class ResourceReleaseManager:
    """
    管理“内置资源 → 释放到 APP_DIR”的逻辑。
    - 首次释放后记录路径
    - 程序退出时删除释放出来的文件（config.ini 除外，需保留用户配置）
    """
    _released_files = set()

    def __init__(self, bundled_path: Path, target_name: str):
        self.bundled_path = Path(bundled_path)
        self.target_name = target_name

    @property
    def target_path(self) -> Path:
        return APP_DIR / self.target_name

    def ensure_released(self):
        target = self.target_path
        if target.exists():
            self._released_files.add(target)
            return target
        if not self.bundled_path.exists():
            logger.warning(f"内置资源不存在: {self.bundled_path}")
            return None
        try:
            shutil.copyfile(self.bundled_path, target)
            self._released_files.add(target)
            logger.info(f"已释放资源: {target}")
            return target
        except OSError as e:
            logger.error(f"释放资源失败 {self.target_name}: {type(e).__name__}: {e}")
            return None

    @classmethod
    def cleanup(cls):
        """程序退出时清理释放出来的临时文件（config.ini 除外）。"""
        for f in list(cls._released_files):
            try:
                if f.exists() and f.name != "config.ini":
                    f.unlink()
                    logger.info(f"已清理临时文件: {f}")
            except OSError as e:
                logger.warning(f"清理临时文件失败 {f}: {e}")
        cls._released_files.clear()


# ===================== 配置管理（DRY） =====================
class ConfigManager:
    """统一管理 config.ini 的读取与保存。"""
    def __init__(self, config_path: Path, defaults: dict):
        self.config_path = Path(config_path)
        self.defaults = defaults
        self.cfg = configparser.ConfigParser()
        self.cfg.read_dict(defaults)

    def load(self):
        if not self.config_path.exists():
            self._release_or_create()
        self.cfg.read(self.config_path, encoding="utf-8")
        base_path = self.cfg.get("EXCEL", "base_path",
                                 fallback=self.defaults["EXCEL"]["base_path"]).strip()
        file_name = self.cfg.get("EXCEL", "file_name",
                                 fallback=self.defaults["EXCEL"]["file_name"]).strip()
        start_year = self.cfg.getint("QUERY", "start_year",
                                     fallback=int(self.defaults["QUERY"]["start_year"]))
        logger.info(f"加载配置: base_path={base_path!r}, file_name={file_name!r}, start_year={start_year}")
        return base_path, file_name, start_year

    def _release_or_create(self):
        config_dir = self.config_path.parent
        if BUNDLED_CONFIG.exists():
            try:
                config_dir.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(BUNDLED_CONFIG, self.config_path)
                logger.info(f"已从内置模板释放配置: {self.config_path}")
                return
            except OSError as e:
                logger.error(f"复制内置配置失败: {type(e).__name__}: {e}")
        try:
            config_dir.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.cfg.write(f)
            logger.info(f"已写入默认配置: {self.config_path}")
        except OSError as e:
            logger.error(f"写入默认配置失败: {type(e).__name__}: {e}")

    def save(self, base_path, file_name, start_year, parent=None) -> bool:
        self.cfg["EXCEL"] = {"base_path": base_path, "file_name": file_name}
        self.cfg["QUERY"] = {"start_year": str(start_year)}
        config_dir = self.config_path.parent
        try:
            config_dir.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.cfg.write(f)
            logger.info(f"配置已保存: {self.config_path}")
            return True
        except PermissionError as e:
            logger.error(f"配置保存失败（权限）: {e}")
            QtWidgets.QMessageBox.critical(
                parent, "保存失败",
                f"无法写入配置文件:\n{self.config_path}\n\n"
                f"请检查该文件夹是否只读，或以管理员身份运行。"
            )
            return False
        except OSError as e:
            logger.error(f"配置保存失败（系统）: {type(e).__name__}: {e}")
            QtWidgets.QMessageBox.critical(
                parent, "保存失败",
                f"配置保存失败:\n{e}\n\n路径: {self.config_path}"
            )
            return False


config_manager = ConfigManager(CONFIG_FILE, DEFAULT_CONFIG)


# ===================== 自定义 LineEdit（重写双击事件） =====================
class ExampleLineEdit(QtWidgets.QLineEdit):
    """带双击填入今日示例的输入框。"""
    doubleClicked = QtCore.pyqtSignal()

    def mouseDoubleClickEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.doubleClicked.emit()
        super().mouseDoubleClickEvent(event)


# ===================== 设置对话框 =====================
class SettingsDialog(QtWidgets.QDialog):
    """设置界面：路径、文件名、起始年份，实时校验 + 防抖网络检查。"""

    def __init__(self, base_path, file_name, start_year, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置 · 数据源路径")
        self.setMinimumWidth(520)
        self.setStyleSheet(load_stylesheet())

        # 防抖定时器：输入停止 500ms 后再检查网络
        self._debounce_timer = QtCore.QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(DEBOUNCE_MS)
        self._debounce_timer.timeout.connect(self._check_preview_exists)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(10)

        tip = QtWidgets.QLabel("修改后立即生效，并写入 config.ini 文件。")
        tip.setObjectName("hintLabel")
        layout.addWidget(tip)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)

        # 路径
        self.path_edit = QtWidgets.QLineEdit(base_path)
        self.path_edit.setPlaceholderText(r"例如：\\192.168.1.12\生管 或 Z:\生管")
        self.path_error = QtWidgets.QLabel()
        self.path_error.setObjectName("errorLabel")
        form.addRow("共享路径：", self.path_edit)
        form.addRow("", self.path_error)

        # 文件名
        self.file_edit = QtWidgets.QLineEdit(file_name)
        self.file_edit.setPlaceholderText("例如：KDJL.xls 或 KDJL.xlsx")
        self.file_error = QtWidgets.QLabel()
        self.file_error.setObjectName("errorLabel")
        form.addRow("Excel 文件名：", self.file_edit)
        form.addRow("", self.file_error)

        # 起始年份
        self.year_edit = QtWidgets.QLineEdit(str(start_year))
        self.year_edit.setPlaceholderText("例如：2017")
        self.year_error = QtWidgets.QLabel()
        self.year_error.setObjectName("errorLabel")
        form.addRow("起始年份：", self.year_edit)
        form.addRow("", self.year_error)

        layout.addLayout(form)

        self.preview = QtWidgets.QLabel()
        self.preview.setObjectName("hintLabel")
        self.preview.setWordWrap(True)
        layout.addWidget(self.preview)

        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        self.btn_ok = btn_box.button(QtWidgets.QDialogButtonBox.Ok)
        self.btn_ok.setEnabled(False)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        # 信号：输入变化 → 即时语法校验 + 防抖网络检查
        self.path_edit.textChanged.connect(self._on_path_changed)
        self.file_edit.textChanged.connect(self._on_file_changed)
        self.year_edit.textChanged.connect(self._validate_all)
        self._validate_all()

    # ---------- 输入清理 ----------
    def _on_path_changed(self, text):
        cleaned = text.strip()
        if cleaned != text:
            self.path_edit.blockSignals(True)
            self.path_edit.setText(cleaned)
            self.path_edit.blockSignals(False)
        self._validate_all()

    def _on_file_changed(self, text):
        filtered = "".join(c for c in text if c not in INVALID_FILE_CHARS)
        if filtered != text:
            self.file_edit.blockSignals(True)
            cursor_pos = self.file_edit.cursorPosition()
            self.file_edit.setText(filtered)
            self.file_edit.setCursorPosition(min(cursor_pos, len(filtered)))
            self.file_edit.blockSignals(False)
        self._validate_all()

    # ---------- 校验 ----------
    def _validate_all(self):
        path_ok, path_msg = validate_path(self.path_edit.text())
        file_ok, file_msg = validate_file_name(self.file_edit.text())
        year_ok, year_msg = validate_year(self.year_edit.text())

        self.path_error.setText("" if path_ok else path_msg)
        self.file_error.setText("" if file_ok else file_msg)
        self.year_error.setText("" if year_ok else year_msg)
        self._set_invalid(self.path_edit, not path_ok)
        self._set_invalid(self.file_edit, not file_ok)
        self._set_invalid(self.year_edit, not year_ok)

        # 路径 + 文件名合法时，防抖检查网络
        if path_ok and file_ok:
            full = str(Path(self.path_edit.text().strip()) / self.file_edit.text().strip())
            self.preview.setText(f"⏳ 检查中… {full}")
            self._debounce_timer.start()
        else:
            self._debounce_timer.stop()
            full = str(Path(self.path_edit.text().strip()) / self.file_edit.text().strip())
            self.preview.setText(f"<b>完整路径：</b>{full}")

        self.btn_ok.setEnabled(path_ok and file_ok and year_ok)

    def _check_preview_exists(self):
        """防抖触发：在后台线程检查网络路径是否存在。"""
        path = self.path_edit.text().strip()
        name = self.file_edit.text().strip()
        full = str(Path(path) / name)

        # 用 QThreadPool 避免阻塞 UI
        future = QtCore.QThreadPool.globalInstance()
        # 简单起见，这里用线程池 + 信号回调
        worker = _PathCheckWorker(full, PREVIEW_TIMEOUT)
        worker.signals.finished.connect(self._on_preview_checked)
        future.start(worker)

    def _on_preview_checked(self, exists: bool, full: str):
        icon = "✅" if exists else "❌"
        self.preview.setText(f"{icon} <b>完整路径：</b>{full}")

    @staticmethod
    def _set_invalid(widget, invalid):
        widget.setObjectName("invalidInput" if invalid else "")
        widget.style().unpolish(widget)
        widget.style().polish(widget)

    def values(self):
        return (self.path_edit.text().strip(),
                self.file_edit.text().strip(),
                int(self.year_edit.text().strip()))


# ===================== 后台网络检查 Worker（QThreadPool） =====================
class _PathCheckWorker(QtCore.QRunnable):
    class Signals(QtCore.QObject):
        finished = QtCore.pyqtSignal(bool, str)

    def __init__(self, path: str, timeout: float):
        super().__init__()
        self.path = path
        self.timeout = timeout
        self.signals = self.Signals()

    def run(self):
        exists = check_network_path(self.path, self.timeout)
        self.signals.finished.emit(exists, self.path)


# ===================== 查询 Worker（QThread） =====================
class QueryWorker(QtCore.QObject):
    """后台执行 query_by_lotno，避免阻塞 UI。"""
    finished = QtCore.pyqtSignal(str)   # 查询结果字符串

    def __init__(self, base_path, file_name, start_year, lot_no):
        super().__init__()
        self.base_path = base_path
        self.file_name = file_name
        self.start_year = start_year
        self.lot_no = lot_no

    @QtCore.pyqtSlot()
    def run(self):
        try:
            result = self._query()
        except Exception:
            logger.exception("查询后台线程发生未知异常")
            result = "错误: 查询时发生未知异常，请查看 app.log 日志。"
        self.finished.emit(result)

    def _query(self) -> str:
        full_path = str(Path(self.base_path) / self.file_name)
        current_year = datetime.datetime.now().year

        if self.start_year > current_year:
            logger.warning(f"起始年份 {self.start_year} 大于当前年份 {current_year}")
            return f"错误: 起始年份 {self.start_year} 大于当前年份 {current_year}，请检查设置。"

        if not check_network_path(full_path, timeout=NETWORK_TIMEOUT):
            logger.warning(f"数据源不可访问: {full_path}")
            return (
                f"错误: 找不到文件或网络无响应\n{full_path}\n\n"
                f"请点『⚙️ 设置』检查路径，或『📂 打开配置』手动编辑 config.ini。"
            )

        lot_list = [
            x.strip()
            for x in self.lot_no.replace(",", " ").replace(";", " ").split()
            if x.strip()
        ]
        years_to_query = range(self.start_year, current_year + 1)
        all_results = []
        engine = "xlrd" if self.file_name.lower().endswith(".xls") else "openpyxl"

        try:
            with pd.ExcelFile(full_path, engine=engine) as xls:
                sheet_names = set(xls.sheet_names)
                logger.info(f"工作表列表: {sorted(sheet_names)}")

                for year in years_to_query:
                    sheet_name = f"{year}年"
                    if sheet_name not in sheet_names:
                        continue

                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                    except pandas.errors.EmptyDataError:
                        logger.warning(f"工作表 {sheet_name} 为空，跳过")
                        continue
                    except pandas.errors.ParserError:
                        logger.error(f"工作表 {sheet_name} 解析失败，跳过")
                        continue
                    # 其他异常向上抛，由 run() 统一记录
                    try:
                        if all(c in df.columns for c in ('批號', 'MONO', '款號')):
                            matches = df[
                                df['批號'].astype(str).str.strip().isin(lot_list)
                            ]
                            if not matches.empty:
                                combined = (
                                    matches['批號'].fillna('').astype(str).str.strip()
                                    + matches['MONO'].fillna('').astype(str).str.strip()
                                    + matches['款號'].fillna('').astype(str).str.strip()
                                )
                                all_results.extend(combined.tolist())
                                logger.info(f"{sheet_name} 命中 {len(combined)} 行")
                        else:
                            df2 = pd.read_excel(
                                xls, sheet_name=sheet_name, header=None
                            )
                            matches2 = df2[
                                df2[0].astype(str).str.strip().isin(lot_list)
                            ]
                            if not matches2.empty:
                                combined2 = (
                                    matches2[0].fillna('').astype(str).str.strip()
                                    + "  "
                                    + matches2[4].fillna('').astype(str).str.strip()
                                    + "  "
                                    + matches2[2].fillna('').astype(str).str.strip()
                                )
                                all_results.extend(combined2.tolist())
                                logger.info(f"{sheet_name}（无表头兜底）命中 {len(combined2)} 行")
                    except KeyError as e:
                        logger.error(f"工作表 {sheet_name} 缺少列: {e}")
                        continue

        except FileNotFoundError:
            logger.error(f"文件不存在: {full_path}")
            return f"错误: 找不到文件\n{full_path}"
        except PermissionError:
            logger.error(f"无权限访问: {full_path}")
            return "错误: 没有足够的权限访问该文件，请检查共享盘权限或文件是否被占用。"
        except pandas.errors.EmptyDataError:
            logger.error(f"Excel 为空: {full_path}")
            return "错误: Excel 文件为空（没有任何工作表）。"
        except pandas.errors.ParserError:
            logger.error(f"Excel 格式损坏: {full_path}")
            return "错误: Excel 文件格式损坏，无法解析。"
        except OSError as e:
            logger.error(f"系统错误: {type(e).__name__}: {e}")
            return f"错误: 系统错误 - {e.strerror or e}"
        # 其他异常不捕获，向上抛给 run()

        return "\n".join(all_results) if all_results else "未查询到数据"


# ===================== 主窗口 =====================
class BatchQueryApp(QtWidgets.QWidget):
    """主窗口。"""

    def __init__(self):
        super().__init__()
        self.base_path, self.file_name, self.start_year = config_manager.load()
        self._query_thread = None
        self._query_worker = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('批號查询器 · Python版')
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.setStyleSheet(load_stylesheet())

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(10)

        # 标题行
        header = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("🔍 批号批量查询器")
        title.setObjectName("titleLabel")
        header.addWidget(title)
        header.addStretch()

        for text, slot in [
            ("⚙️ 设置", self.open_settings),
            ("📂 打开配置", self.open_config_file),
            ("❓ 帮助", self.open_help),
            ("📦 依赖", self.open_requirements),
        ]:
            btn = QtWidgets.QPushButton(text)
            btn.setObjectName("secondaryBtn")
            btn.clicked.connect(slot)
            header.addWidget(btn)

        main_layout.addLayout(header)

        # 路径状态
        self.path_label = QtWidgets.QLabel()
        self.path_label.setObjectName("hintLabel")
        self.path_label.setWordWrap(True)
        main_layout.addWidget(self.path_label)
        self.refresh_path_label()

        # 提示
        hint = QtWidgets.QLabel(
            f"请输入批号 (LotNo)，按 Enter 查询；双击输入框可填入今日示例（最多 {LOT_NO_MAX_LEN} 字符）"
        )
        hint.setObjectName("hintLabel")
        main_layout.addWidget(hint)

        # 输入框（子类化，重写双击）
        self.input_line = ExampleLineEdit()
        self.input_line.setPlaceholderText(
            f"例如：{datetime.datetime.now().strftime('%y%m%d')}-01（双击填入今日示例）"
        )
        self.input_line.setClearButtonEnabled(True)
        self.input_line.setMaxLength(LOT_NO_MAX_LEN)
        self.input_line.returnPressed.connect(self.on_query_triggered)
        self.input_line.textChanged.connect(self._on_lot_input_changed)
        self.input_line.doubleClicked.connect(self._fill_today_example)
        main_layout.addWidget(self.input_line)

        self.input_error = QtWidgets.QLabel()
        self.input_error.setObjectName("errorLabel")
        main_layout.addWidget(self.input_error)

        # 结果
        result_label = QtWidgets.QLabel("查询结果")
        result_label.setObjectName("hintLabel")
        main_layout.addWidget(result_label)

        self.output_text = QtWidgets.QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setPlaceholderText("查询结果会显示在这里，并自动复制到剪贴板…")
        main_layout.addWidget(self.output_text, stretch=1)

        self.input_line.setFocus()

    # ---------- 输入校验（清空同步结果） ----------
    def _on_lot_input_changed(self, text):
        if not text:
            self.input_error.setText("")
            self._set_invalid(self.input_line, False)
            self.output_text.clear()
            return
        ok, msg = validate_lot_no(text)
        self.input_error.setText("" if ok else msg)
        self._set_invalid(self.input_line, not ok)

    @staticmethod
    def _set_invalid(widget, invalid):
        widget.setObjectName("invalidInput" if invalid else "")
        widget.style().unpolish(widget)
        widget.style().polish(widget)

    def refresh_path_label(self):
        full_path = str(Path(self.base_path) / self.file_name)
        exists = check_network_path(full_path, timeout=NETWORK_TIMEOUT)
        icon = "✅" if exists else "❌"
        status = "可访问" if exists else "不可访问或无响应"
        self.path_label.setText(f"{icon} 数据源（{status}）：\n{full_path}")

    def _fill_today_example(self):
        today = datetime.datetime.now().strftime('%y%m%d')
        self.input_line.setText(f"{today}-01")
        self.input_line.selectAll()

    def _open_with_default(self, path: Path):
        """用系统默认程序打开文件，避免 Shell 注入。"""
        if not path.exists():
            QtWidgets.QMessageBox.warning(self, "文件不存在", f"找不到文件：\n{path}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))   # Windows 专用，安全
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except OSError as e:
            logger.error(f"打开文件失败: {type(e).__name__}: {e}")
            QtWidgets.QMessageBox.warning(self, "打开失败", f"{e}")

    # ---------- 设置 / 配置 / 帮助 / 依赖 ----------
    def open_settings(self):
        dlg = SettingsDialog(self.base_path, self.file_name, self.start_year, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            base_path, file_name, start_year = dlg.values()
            path_ok, path_msg = validate_path(base_path)
            file_ok, file_msg = validate_file_name(file_name)
            if not path_ok or not file_ok:
                QtWidgets.QMessageBox.warning(self, "校验失败", path_msg or file_msg)
                return
            if not config_manager.save(base_path, file_name, start_year, parent=self):
                return
            self.base_path, self.file_name, self.start_year = base_path, file_name, start_year
            self.refresh_path_label()
            QtWidgets.QMessageBox.information(self, "已保存", "配置已更新并写入 config.ini。")

    def open_config_file(self):
        if not CONFIG_FILE.exists():
            if not config_manager.save(self.base_path, self.file_name, self.start_year, parent=self):
                return
        self._open_with_default(CONFIG_FILE)

    def open_help(self):
        manager = ResourceReleaseManager(BUNDLED_HELP, "Help.txt")
        target = manager.ensure_released()
        if target is None:
            QtWidgets.QMessageBox.warning(self, "无法打开", "释放 Help.txt 失败，请查看 app.log")
            return
        self._open_with_default(target)

    def open_requirements(self):
        manager = ResourceReleaseManager(BUNDLED_REQ, "requirements.txt")
        target = manager.ensure_released()
        if target is None:
            QtWidgets.QMessageBox.warning(self, "无法打开", "释放 requirements.txt 失败，请查看 app.log")
            return
        self._open_with_default(target)

    # ---------- 查询（后台线程） ----------
    def on_query_triggered(self):
        lot_no = self.input_line.text().strip()
        ok, msg = validate_lot_no(lot_no)
        if not ok:
            self.input_error.setText(msg)
            self._set_invalid(self.input_line, True)
            logger.warning(f"批号校验失败: {lot_no!r} - {msg}")
            return

        if self._query_thread is not None and self._query_thread.isRunning():
            logger.warning("上一次查询尚未结束，忽略本次触发")
            return

        logger.info(f"开始查询批号: {lot_no}")
        self.output_text.setText("⏳ 正在查询…")
        QtWidgets.QApplication.processEvents()   # 刷新 UI

        self._query_thread = QtCore.QThread()
        self._query_worker = QueryWorker(
            self.base_path, self.file_name, self.start_year, lot_no
        )
        self._query_worker.moveToThread(self._query_thread)
        self._query_thread.started.connect(self._query_worker.run)
        self._query_worker.finished.connect(self._on_query_finished)
        self._query_worker.finished.connect(self._query_thread.quit)
        self._query_thread.finished.connect(self._query_thread.deleteLater)
        self._query_thread.start()

    def _on_query_finished(self, result: str):
        if result == "未查询到数据":
            self.output_text.setText(f"⚠️  {result}")
        elif result.startswith("错误"):
            self.output_text.setText(f"❌  {result}")
        else:
            self.output_text.setText(result)
            QtGui.QGuiApplication.clipboard().setText(result)
            logger.info(f"查询成功，结果已复制到剪贴板（{len(result)} 字符）")

    # ---------- 关闭时清理 ----------
    def closeEvent(self, event):
        if self._query_thread is not None and self._query_thread.isRunning():
            self._query_thread.quit()
            self._query_thread.wait(2000)
        ResourceReleaseManager.cleanup()
        super().closeEvent(event)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    ex = BatchQueryApp()
    ex.show()
    sys.exit(app.exec())