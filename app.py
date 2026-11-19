import sys
import os
import re
import shutil
import logging
import datetime
import concurrent.futures
import configparser
import pandas as pd
import pandas.errors
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QTextEdit, QLabel, QFrame, QPushButton,
    QDialog, QDialogButtonBox, QFormLayout, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QGuiApplication


# ===================== 路径处理（兼容 exe + 内置资源） =====================
def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


APP_DIR = get_app_dir()
CONFIG_FILE = os.path.join(APP_DIR, "config.ini")
LOG_FILE = os.path.join(APP_DIR, "app.log")
BUNDLED_CONFIG = get_resource_path("config.ini")
BUNDLED_HELP = get_resource_path("Help.txt")
BUNDLED_REQ = get_resource_path("requirements.txt")
STYLE_FILE = get_resource_path("styles.qss")

# ===================== 常量（消除硬编码） =====================
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
NETWORK_TIMEOUT = 5


# ===================== 日志 =====================
def setup_logger():
    logger = logging.getLogger("BatchQuery")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    try:
        fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
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
def load_stylesheet():
    external = os.path.join(APP_DIR, "styles.qss")
    path = external if os.path.exists(external) else STYLE_FILE
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except OSError as e:
        logger.warning(f"加载样式失败 {path}: {e}")
        return ""


# ===================== 校验函数 =====================
def validate_path(path):
    if not path:
        return False, "路径不能为空"
    invalid = re.search(r'[<>:"|?*]', path)
    if invalid:
        return False, f"路径不能包含字符: {invalid.group()}"
    if not (path.startswith("\\\\") or (len(path) >= 2 and path[1] == ":")):
        return False, "路径格式不正确（应为 \\\\服务器\\共享 或 X:\\）"
    return True, ""


def validate_file_name(name):
    if not name:
        return False, "文件名不能为空"
    if any(c in INVALID_FILE_CHARS for c in name):
        return False, f"文件名不能包含字符: {INVALID_FILE_CHARS}"
    if not name.lower().endswith((".xls", ".xlsx")):
        return False, "文件名必须以 .xls 或 .xlsx 结尾"
    return True, ""


def validate_lot_no(text):
    if not text:
        return False, "批号不能为空"
    if len(text) > LOT_NO_MAX_LEN:
        return False, f"批号长度不能超过 {LOT_NO_MAX_LEN} 个字符"
    if not re.match(r'^[\w\-.,;\s]+$', text, re.UNICODE):
        return False, "批号只能包含字母、数字、-、_、.、空格、逗号、分号"
    return True, ""


# ===================== 网络路径超时检查 =====================
def check_network_path(path, timeout=NETWORK_TIMEOUT):
    def _check():
        return os.path.exists(path)

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(_check)
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            logger.warning(f"路径检查超时（>{timeout}s）：{path}")
            return False


# ===================== 资源释放管理器（DRY） =====================
class ResourceReleaseManager:
    def __init__(self, bundled_path, target_name):
        self.bundled_path = bundled_path
        self.target_name = target_name

    @property
    def target_path(self):
        return os.path.join(APP_DIR, self.target_name)

    def ensure_released(self):
        target = self.target_path
        if os.path.exists(target):
            return target
        if not os.path.exists(self.bundled_path):
            logger.warning(f"内置资源不存在: {self.bundled_path}")
            return None
        try:
            shutil.copyfile(self.bundled_path, target)
            logger.info(f"已释放资源: {target}")
            return target
        except OSError as e:
            logger.error(f"释放资源失败 {self.target_name}: {type(e).__name__}: {e}")
            return None


# ===================== 配置管理（DRY） =====================
class ConfigManager:
    def __init__(self, config_path, defaults):
        self.config_path = config_path
        self.defaults = defaults
        self.cfg = configparser.ConfigParser()
        self.cfg.read_dict(defaults)

    def load(self):
        if not os.path.exists(self.config_path):
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
        config_dir = os.path.dirname(self.config_path)
        if os.path.exists(BUNDLED_CONFIG):
            try:
                if config_dir:
                    os.makedirs(config_dir, exist_ok=True)
                shutil.copyfile(BUNDLED_CONFIG, self.config_path)
                logger.info(f"已从内置模板释放配置: {self.config_path}")
                return
            except OSError as e:
                logger.error(f"复制内置配置失败: {type(e).__name__}: {e}")
        try:
            if config_dir:
                os.makedirs(config_dir, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.cfg.write(f)
            logger.info(f"已写入默认配置: {self.config_path}")
        except OSError as e:
            logger.error(f"写入默认配置失败: {type(e).__name__}: {e}")

    def save(self, base_path, file_name, start_year) -> bool:
        self.cfg["EXCEL"] = {"base_path": base_path, "file_name": file_name}
        self.cfg["QUERY"] = {"start_year": str(start_year)}
        config_dir = os.path.dirname(self.config_path)
        try:
            if config_dir:
                os.makedirs(config_dir, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.cfg.write(f)
            logger.info(f"配置已保存: {self.config_path}")
            return True
        except PermissionError as e:
            logger.error(f"配置保存失败（权限）: {e}")
            QMessageBox.critical(
                None, "保存失败",
                f"无法写入配置文件:\n{self.config_path}\n\n"
                f"请检查该文件夹是否只读，或以管理员身份运行。"
            )
            return False
        except OSError as e:
            logger.error(f"配置保存失败（系统）: {type(e).__name__}: {e}")
            QMessageBox.critical(
                None, "保存失败",
                f"配置保存失败:\n{e}\n\n路径: {self.config_path}"
            )
            return False


config_manager = ConfigManager(CONFIG_FILE, DEFAULT_CONFIG)


# ===================== 设置对话框 =====================
class SettingsDialog(QDialog):
    def __init__(self, base_path, file_name, start_year, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置 · 数据源路径")
        self.setMinimumWidth(520)
        self.setStyleSheet(load_stylesheet())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(10)

        tip = QLabel("修改后立即生效，并写入 config.ini 文件。")
        tip.setObjectName("hintLabel")
        layout.addWidget(tip)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)

        self.path_edit = QLineEdit(base_path)
        self.path_edit.setPlaceholderText(r"例如：\\192.168.1.12\生管 或 Z:\生管")
        self.path_error = QLabel()
        self.path_error.setObjectName("errorLabel")
        form.addRow("共享路径：", self.path_edit)
        form.addRow("", self.path_error)

        self.file_edit = QLineEdit(file_name)
        self.file_edit.setPlaceholderText("例如：KDJL.xls 或 KDJL.xlsx")
        self.file_error = QLabel()
        self.file_error.setObjectName("errorLabel")
        form.addRow("Excel 文件名：", self.file_edit)
        form.addRow("", self.file_error)

        self.year_edit = QLineEdit(str(start_year))
        self.year_edit.setPlaceholderText("例如：2017")
        self.year_error = QLabel()
        self.year_error.setObjectName("errorLabel")
        form.addRow("起始年份：", self.year_edit)
        form.addRow("", self.year_error)

        layout.addLayout(form)

        self.preview = QLabel()
        self.preview.setObjectName("hintLabel")
        self.preview.setWordWrap(True)
        layout.addWidget(self.preview)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.btn_ok = btn_box.button(QDialogButtonBox.Ok)
        self.btn_ok.setEnabled(False)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.path_edit.textChanged.connect(self._on_path_changed)
        self.file_edit.textChanged.connect(self._on_file_changed)
        self.year_edit.textChanged.connect(self._validate_all)
        self._validate_all()

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

    def _validate_all(self):
        path_ok, path_msg = validate_path(self.path_edit.text())
        file_ok, file_msg = validate_file_name(self.file_edit.text())

        year_text = self.year_edit.text().strip()
        year_ok = True
        year_msg = ""
        if not year_text:
            year_ok, year_msg = False, "起始年份不能为空"
        else:
            try:
                y = int(year_text)
                current = datetime.datetime.now().year
                if y < 2000 or y > current:
                    year_ok, year_msg = False, f"起始年份应在 2000–{current} 之间"
            except ValueError:
                year_ok, year_msg = False, "起始年份必须是整数"

        self.path_error.setText("" if path_ok else path_msg)
        self.file_error.setText("" if file_ok else file_msg)
        self.year_error.setText("" if year_ok else year_msg)
        self._set_invalid(self.path_edit, not path_ok)
        self._set_invalid(self.file_edit, not file_ok)
        self._set_invalid(self.year_edit, not year_ok)

        full = os.path.join(self.path_edit.text().strip(),
                            self.file_edit.text().strip())
        exists = check_network_path(full, timeout=3)
        icon = "✅" if exists else "❌"
        self.preview.setText(f"{icon} <b>完整路径：</b>{full}")

        self.btn_ok.setEnabled(path_ok and file_ok and year_ok)

    @staticmethod
    def _set_invalid(widget, invalid):
        widget.setObjectName("invalidInput" if invalid else "")
        widget.style().unpolish(widget)
        widget.style().polish(widget)

    def values(self):
        return (self.path_edit.text().strip(),
                self.file_edit.text().strip(),
                int(self.year_edit.text().strip()))


# ===================== 主窗口 =====================
class BatchQueryApp(QWidget):
    def __init__(self):
        super().__init__()
        self.base_path, self.file_name, self.start_year = config_manager.load()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('批號查询器 · Python版')
        self.resize(640, 520)
        self.setStyleSheet(load_stylesheet())

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(10)

        header = QHBoxLayout()
        title = QLabel("🔍 批号批量查询器")
        title.setObjectName("titleLabel")
        header.addWidget(title)
        header.addStretch()

        for text, slot in [
            ("⚙️ 设置", self.open_settings),
            ("📂 打开配置", self.open_config_file),
            ("❓ 帮助", self.open_help),
            ("📦 依赖", self.open_requirements),
        ]:
            btn = QPushButton(text)
            btn.setObjectName("secondaryBtn")
            btn.clicked.connect(slot)
            header.addWidget(btn)

        main_layout.addLayout(header)

        self.path_label = QLabel()
        self.path_label.setObjectName("hintLabel")
        self.path_label.setWordWrap(True)
        main_layout.addWidget(self.path_label)
        self.refresh_path_label()

        hint = QLabel(
            f"请输入批号 (LotNo)，按 Enter 查询；双击输入框可填入今日示例（最多 {LOT_NO_MAX_LEN} 字符）"
        )
        hint.setObjectName("hintLabel")
        main_layout.addWidget(hint)

        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText(
            f"例如：{datetime.datetime.now().strftime('%y%m%d')}-01（双击填入今日示例）"
        )
        self.input_line.setClearButtonEnabled(True)
        self.input_line.setMaxLength(LOT_NO_MAX_LEN)
        self.input_line.returnPressed.connect(self.on_query_triggered)
        self.input_line.textChanged.connect(self._on_lot_input_changed)
        self.input_line.mouseDoubleClickEvent = self._fill_today_example
        main_layout.addWidget(self.input_line)

        self.input_error = QLabel()
        self.input_error.setObjectName("errorLabel")
        main_layout.addWidget(self.input_error)

        result_label = QLabel("查询结果")
        result_label.setObjectName("hintLabel")
        main_layout.addWidget(result_label)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setPlaceholderText("查询结果会显示在这里，并自动复制到剪贴板…")
        main_layout.addWidget(self.output_text, stretch=1)

        self.setLayout(main_layout)
        self.input_line.setFocus()

    # ---------------- 输入实时校验（含清空同步） ----------------
    def _on_lot_input_changed(self, text):
        # ✅ 输入框为空 → 同步清空结果区、错误提示、红色边框
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
        full_path = os.path.join(self.base_path, self.file_name)
        exists = check_network_path(full_path, timeout=NETWORK_TIMEOUT)
        icon = "✅" if exists else "❌"
        status = "可访问" if exists else "不可访问或无响应"
        self.path_label.setText(f"{icon} 数据源（{status}）：\n{full_path}")

    def _fill_today_example(self, event):
        today = datetime.datetime.now().strftime('%y%m%d')
        self.input_line.setText(f"{today}-01")
        self.input_line.selectAll()

    def _open_with_default(self, path):
        if not os.path.exists(path):
            QMessageBox.warning(self, "文件不存在", f"找不到文件：\n{path}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except OSError as e:
            logger.error(f"打开文件失败: {type(e).__name__}: {e}")
            QMessageBox.warning(self, "打开失败", f"{e}")

    # ---------------- 设置 / 配置 / 帮助 / 依赖 ----------------
    def open_settings(self):
        dlg = SettingsDialog(self.base_path, self.file_name, self.start_year, self)
        if dlg.exec_() == QDialog.Accepted:
            base_path, file_name, start_year = dlg.values()
            path_ok, path_msg = validate_path(base_path)
            file_ok, file_msg = validate_file_name(file_name)
            if not path_ok or not file_ok:
                QMessageBox.warning(self, "校验失败", path_msg or file_msg)
                return
            if not config_manager.save(base_path, file_name, start_year):
                return
            self.base_path, self.file_name, self.start_year = base_path, file_name, start_year
            self.refresh_path_label()
            QMessageBox.information(self, "已保存", "配置已更新并写入 config.ini。")

    def open_config_file(self):
        if not os.path.exists(CONFIG_FILE):
            if not config_manager.save(self.base_path, self.file_name, self.start_year):
                return
        self._open_with_default(CONFIG_FILE)

    def open_help(self):
        manager = ResourceReleaseManager(BUNDLED_HELP, "Help.txt")
        target = manager.ensure_released()
        if target is None:
            QMessageBox.warning(self, "无法打开", "释放 Help.txt 失败，请查看 app.log")
            return
        self._open_with_default(target)

    def open_requirements(self):
        manager = ResourceReleaseManager(BUNDLED_REQ, "requirements.txt")
        target = manager.ensure_released()
        if target is None:
            QMessageBox.warning(self, "无法打开", "释放 requirements.txt 失败，请查看 app.log")
            return
        self._open_with_default(target)

    # ---------------- 查询 ----------------
    def on_query_triggered(self):
        lot_no = self.input_line.text().strip()
        ok, msg = validate_lot_no(lot_no)
        if not ok:
            self.input_error.setText(msg)
            self._set_invalid(self.input_line, True)
            logger.warning(f"批号校验失败: {lot_no!r} - {msg}")
            return

        logger.info(f"开始查询批号: {lot_no}")
        result = self.query_by_lotno(lot_no)

        if result == "未查询到数据":
            self.output_text.setText(f"⚠️  {result}")
        elif result.startswith("错误"):
            self.output_text.setText(f"❌  {result}")
        else:
            self.output_text.setText(result)

        if result and result != "未查询到数据":
            QGuiApplication.clipboard().setText(result)
            logger.info(f"查询成功，结果已复制到剪贴板（{len(result)} 字符）")

    def query_by_lotno(self, lot_no):
        full_path = os.path.join(self.base_path, self.file_name)
        current_year = datetime.datetime.now().year

        if self.start_year > current_year:
            logger.warning(f"起始年份 {self.start_year} 大于当前年份 {current_year}")
            return f"错误: 起始年份 {self.start_year} 大于当前年份 {current_year}，请检查设置。"
        years_to_query = range(self.start_year, current_year + 1)

        if not check_network_path(full_path, timeout=NETWORK_TIMEOUT):
            logger.warning(f"数据源不可访问: {full_path}")
            return (
                f"错误: 找不到文件或网络无响应\n{full_path}\n\n"
                f"请点『⚙️ 设置』检查路径，或『📂 打开配置』手动编辑 config.ini。"
            )

        lot_list = [
            x.strip()
            for x in lot_no.replace(",", " ").replace(";", " ").split()
            if x.strip()
        ]

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
                    except Exception:
                        logger.exception(f"读取工作表 {sheet_name} 时发生未知异常，跳过")
                        continue

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
                    except Exception:
                        logger.exception(f"处理工作表 {sheet_name} 数据时发生未知异常，跳过")
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
        except Exception:
            logger.exception(f"打开 Excel 时发生未知异常: {full_path}")
            return "错误: 打开 Excel 时发生未知异常，请查看 app.log 日志。"

        return "\n".join(all_results) if all_results else "未查询到数据"


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    ex = BatchQueryApp()
    ex.show()
    sys.exit(app.exec_())