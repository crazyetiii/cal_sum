import sys
import os
import re
import json
import subprocess
import win32com.client
from time import sleep
import pythoncom
import psutil

from PySide2.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QLineEdit,
    QTextEdit,
    QMessageBox,
)
from PySide2.QtCore import QThread, Signal, Qt
from PySide2.QtGui import (
    QDragEnterEvent,
    QTextCursor,
    QIcon,
    QFont,
    QGuiApplication,
    QDropEvent,
)

__version__ = "v1.0.4 by lhy"
en_out_file = "en.txt"
cn_out_file = "cn.txt"

if getattr(sys, "frozen", False):
    beyond_compare_name = os.path.join(sys._MEIPASS, "BCompare\BCompare.exe")
else:
    beyond_compare_name = "BCompare\BCompare.exe"

default_keywords_config = {
    "chinese": ["中文", "chinese"],
    "english": ["英文", "english"],
}


def load_keywords_config():
    """从外部文件加载关键词配置"""
    config_path = os.path.join(get_current_file_path(), "compare.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        # 如果配置文件不存在，则创建它并写入默认配置
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default_keywords_config, f, ensure_ascii=False, indent=4)
            print(f"配置文件已创建：{config_path}")
        return default_keywords_config  # 返回默认配置


def get_file_name(target):
    return os.path.basename(target)


def get_current_file_path():
    # 如果是从可执行文件运行，获取可执行文件的路径
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        # 如果是从源代码运行，获取当前脚本的路径
        return os.path.dirname(os.path.abspath(__file__))


def matches_parentheses(s):
    # 匹配以左小括号开始和右小括号结束的字符串
    pattern = r"^\(.*$"
    return bool(re.match(pattern, s))


def is_number(s):
    # 匹配整数、浮点数、带有 '%' 的数值，并支持千分符
    pattern = r"^-?\d{1,3}(,\d{3})*(\.\d+)?%?$"
    return bool(re.match(pattern, s.strip()))


def write_file(data, file_path):
    """将嵌套列表写入到文件"""
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(
            data, f, ensure_ascii=False, indent=4
        )  # 使用 indent 参数以便格式化输出


class FileComparator(QThread):
    log_signal = Signal(str)  # 用于传递日志消息的信号
    finished_signal = Signal()  # 用于指示比较完成的信号
    error_signal = Signal(str)  # 用于传递错误消息的信号

    def __init__(self, file1, file2):
        super().__init__()
        self.file1 = file1
        self.file2 = file2

    def convert_docx_to_doc(self, docx_path):
        """
        将 .docx 文件转换为 .doc 格式，并返回新的文件路径。
        """
        self.log_signal.emit(f"正在转换为doc格式，路径为：【{docx_path}】 ")
        # 初始化 Word 应用程序
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 不显示 Word 窗口

        docx_path = os.path.normpath(docx_path)
        # 打开 .docx 文件
        doc = word.Documents.Open(docx_path)

        # 获取转换后的文件路径（将 .docx 后缀替换为 .doc）
        doc_path = docx_path.replace(".docx", ".doc")

        try:
            # 另存为 .doc 格式，FileFormat=0 对应 .doc 格式
            doc.SaveAs(doc_path, FileFormat=0)
            self.log_signal.emit(
                f"已将【{get_file_name(docx_path)}】转换为【{get_file_name(doc_path)}】"
            )
        except Exception as e:
            self.log_signal.emit(f"转换失败: {e}")
        finally:
            # 关闭文档并退出 Word 应用
            doc.Close(False)
            word.Quit()
            sleep(5)

        return doc_path

    def run(self):
        # 初始化 COM 库
        pythoncom.CoInitialize()
        # 如果文件是 .docx 格式，先转换为 .doc 格式
        if self.file1.endswith(".docx"):
            self.file1 = self.convert_docx_to_doc(self.file1)

        if self.file2.endswith(".docx"):
            self.file2 = self.convert_docx_to_doc(self.file2)

        self.log_signal.emit(f"开始比较文件")
        try:
            # 读取中英文文档中的表格数据
            self.read_table_data(self.file1, en_out_file)
            self.read_table_data(self.file2, cn_out_file)
            self.compare_with_beyond_compare()
            self.log_signal.emit("文件比较完成。")
        except Exception as e:
            self.log_signal.emit(f"文件比较出错: {str(e)}")
            self.error_signal.emit(f"文件比较出错: {str(e)}")  # 发送错误消息
        finally:
            self.finished_signal.emit()

    def compare_with_beyond_compare(self):
        self.log_signal.emit(f"正在调用对比软件...")
        en_out_file_total = os.path.join(get_current_file_path(), en_out_file)
        cn_out_file_total = os.path.join(get_current_file_path(), cn_out_file)
        beyond_compare_path = os.path.join(get_current_file_path(), beyond_compare_name)
        # 构建命令
        command = [beyond_compare_path, en_out_file_total, cn_out_file_total]

        try:
            # 调用 Beyond Compare 进行比较
            subprocess.run(command)
        except Exception as e:
            self.log_signal.emit(f"调用 Beyond Compare 失败: {e}")
            self.error_signal.emit(f"调用 Beyond Compare 失败: {e}")  # 发送错误消息

    def read_table_data(self, doc_path, file_name):
        self.log_signal.emit(f"正在处理文件【{doc_path}】")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 让Word在后台运行

        doc = None
        try:
            doc_path = os.path.normpath(doc_path)
            doc = word.Documents.Open(doc_path)
            table_data = []

            if doc.Tables.Count == 0:
                self.log_signal.emit(f"文档 {doc_path} 中没有表格")
                return []

            for table_index, table in enumerate(doc.Tables):
                table_content = []
                self.log_signal.emit(
                    f"正在处理表格【{table_index + 1}】...行数: {table.Rows.Count}, 列数: {table.Columns.Count}"
                )
                for row in range(1, table.Rows.Count + 1):
                    row_content = []
                    for col in range(1, table.Columns.Count + 1):
                        try:
                            cell_text = table.Cell(row, col).Range.Text
                            cell_text = cell_text.strip("\r\x07")
                            if cell_text == "":
                                continue
                            if is_number(cell_text):
                                row_content.append(cell_text)
                            if matches_parentheses(cell_text):
                                cell_text = cell_text.replace(")", "")
                                cell_text = cell_text.replace("(", "-")
                                row_content.append(cell_text)
                        except Exception as e:
                            # self.log_signal.emit(f"处理单元格时出错: {e}")
                            pass
                    if len(row_content) > 0:
                        table_content.append(row_content)
                if len(table_content) > 0:
                    table_data.append(table_content)

        except Exception as e:
            self.log_signal.emit(f"打开文档时出错: {e}")
            print(f"打开文档时出错: {e}")
            return []

        finally:
            if doc:
                doc.Close(False)
            word.Quit()
            sleep(5)

        write_file(table_data, file_name)
        return table_data


class FileComparisonApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.comparator_thread = None  # 初始化线程
        load_keywords_config()  # 没有配置文件时,在这里创建

    def initUI(self):
        # 获取屏幕的可用几何信息
        screen_geometry = QGuiApplication.primaryScreen().availableGeometry()

        # 计算窗口的宽度和高度为屏幕的一半
        window_width = screen_geometry.width() // 2
        window_height = screen_geometry.height() // 2

        # 设置窗口的几何信息（居中）
        self.setGeometry(
            (screen_geometry.width() - window_width) // 2,
            (screen_geometry.height() - window_height) // 2,
            window_width,
            window_height,
        )
        layout = QVBoxLayout()
        # 设置主窗口字体
        self.setFont(QFont("Arial", 16))

        # 第一组文件选择
        self.label1 = QLabel("选择英文文件:")
        self.label1.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.label1)

        file1_layout = QHBoxLayout()
        self.file1_path = QLineEdit(self)
        self.file1_path.setPlaceholderText(
            "拖动.doc或.docx文件到此处或点击浏览按钮选择文件"
        )
        self.file1_path.setAcceptDrops(True)
        self.file1_path.dragEnterEvent = self.dragEnterEvent
        self.file1_path.dropEvent = self.dropEvent_file1
        file1_layout.addWidget(self.file1_path)

        self.browse_button1 = QPushButton("浏览", self)
        self.browse_button1.clicked.connect(self.select_file1)
        file1_layout.addWidget(self.browse_button1)

        layout.addLayout(file1_layout)

        # 第二组文件选择
        self.label2 = QLabel("选择中文文件:")
        self.label2.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.label2)

        file2_layout = QHBoxLayout()
        self.file2_path = QLineEdit(self)
        self.file2_path.setPlaceholderText(
            "拖动.doc或.docx文件到此处或点击浏览按钮选择文件"
        )
        self.file2_path.setAcceptDrops(True)
        self.file2_path.dragEnterEvent = self.dragEnterEvent
        self.file2_path.dropEvent = self.dropEvent_file2
        file2_layout.addWidget(self.file2_path)

        self.browse_button2 = QPushButton("浏览", self)
        self.browse_button2.clicked.connect(self.select_file2)
        file2_layout.addWidget(self.browse_button2)

        layout.addLayout(file2_layout)

        # 日志框
        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet(
            "font-family: Courier New; background-color: #f0f0f0;font-size: 14;"
        )
        layout.addWidget(self.log_box)

        # 开始比较按钮
        self.compare_button = QPushButton("开始比较", self)
        self.compare_button.clicked.connect(self.start_comparison)
        self.compare_button.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.compare_button)

        # 设置布局
        self.setLayout(layout)

        # 设置窗口
        self.setWindowTitle(f"Doc中英文比较工具 {__version__}")

        # 启用窗口的拖拽功能
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖拽进入事件，检查是否拖入了有效的文件类型"""
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            urls = mime_data.urls()
            if all(
                url.toLocalFile().lower().endswith((".doc", ".docx")) for url in urls
            ):
                event.acceptProposedAction()
                event.setDropAction(Qt.CopyAction)  # 设置为拷贝操作
            else:
                self.show_error_message("仅支持 .doc 和 .docx 文件")
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        """处理文件拖放事件"""
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        self.keywords = load_keywords_config()

        # 只接受两个文件
        if len(files) != 2:
            self.show_error_message("只允许拖动两个文件哦~~~")
            return

        chinese_file = None
        english_file = None

        for file_path in files:
            file_name = os.path.basename(file_path).lower()

            if any(kw.lower() in file_name.lower() for kw in self.keywords["chinese"]):
                if chinese_file is None:
                    chinese_file = file_path
                    self.file2_path.setText(file_path)  # 设置中文文件路径
                    self.log_message(f"中文文件自动识别：{file_path}")
            elif any(
                kw.lower() in file_name.lower() for kw in self.keywords["english"]
            ):
                if english_file is None:
                    english_file = file_path
                    self.file1_path.setText(file_path)  # 设置英文文件路径
                    self.log_message(f"英文文件自动识别：{file_path}")

        if chinese_file is None:
            self.show_error_message("未选择中文文件，请检查文件名是否包含正确的关键词")
        if english_file is None:
            self.show_error_message("未选择英文文件，请检查文件名是否包含正确的关键词")

    def log_message(self, message):
        """将消息写入日志框"""
        self.log_box.append(message)
        self.log_box.moveCursor(QTextCursor.End)  # 滚动到日志框的底部

    def show_error_message(self, message):
        """显示错误消息弹窗"""
        QMessageBox.warning(self, "错误", message)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """接受 .doc 和 .docx 文件的拖拽"""
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            url = mime_data.urls()[0].toLocalFile()
            if url.endswith(".doc") or url.endswith(".docx"):
                event.acceptProposedAction()
            else:
                self.show_error_message(f"拒绝文件：{url}（仅支持 .doc 和 .docx 文件）")
                event.ignore()
        else:
            event.ignore()

    def dropEvent_file1(self, event):
        """处理英文文件的拖放事件"""
        file_path = event.mimeData().urls()[0].toLocalFile()
        if file_path.lower().endswith(".doc") or file_path.lower().endswith(".docx"):
            self.file1_path.setText(file_path)
            self.log_message(f"英文文件已选择：{file_path}")
        else:
            self.show_error_message(f"拒绝文件：{file_path}（支持 .doc和.docx 文件）")

    def dropEvent_file2(self, event):
        """处理中文文件的拖放事件"""
        file_path = event.mimeData().urls()[0].toLocalFile()
        if file_path.lower().endswith(".doc") or file_path.lower().endswith(".docx"):
            self.file2_path.setText(file_path)
            self.log_message(f"中文文件已选择：{file_path}")
        else:
            self.show_error_message(f"拒绝文件：{file_path}（支持 .doc和.docx 文件）")

    def select_file1(self):
        file1, _ = QFileDialog.getOpenFileName(
            self, "选择英文文件", "", "Word 文档 (*.doc *.docx)"
        )
        if file1:
            self.file1_path.setText(file1)
            self.log_message(f"英文文件已选择：{file1}")

    def select_file2(self):
        file2, _ = QFileDialog.getOpenFileName(
            self, "选择中文文件", "", "Word 文档 (*.doc *.docx)"
        )
        if file2:
            self.file2_path.setText(file2)
            self.log_message(f"中文文件已选择：{file2}")

    def start_comparison(self):
        self.log_box.clear()
        self.log_message("开始处理...")
        write_file([], en_out_file)
        write_file([], cn_out_file)
        self.log_message("清空输出文件...")

        file1 = self.file1_path.text()
        file2 = self.file2_path.text()

        # 禁用比较按钮
        self.compare_button.setEnabled(False)

        # 开始比较文件的逻辑
        self.comparator_thread = FileComparator(file1, file2)
        self.comparator_thread.log_signal.connect(self.log_message)
        self.comparator_thread.finished_signal.connect(self.comparison_finished)
        self.comparator_thread.error_signal.connect(self.show_error_message)
        # self.comparator_thread.read_table_data(
        #     r"C:\Users\Administrator\Desktop\cal_sum\08兰妮比加利审计报告-英文附注 2022-0330.doc",
        #     "en.txt",
        # )
        self.comparator_thread.start()  # 启动线程

    def comparison_finished(self):
        # self.log_message("比较线程已完成。")
        # 恢复比较按钮
        self.compare_button.setEnabled(True)

    def close_beyond_compare(self):
        """关闭 Beyond Compare 进程"""
        for proc in psutil.process_iter():
            try:
                if proc.name() == "BCompare.exe":
                    proc.kill()  # 结束进程
                    self.log_message("已关闭 Beyond Compare。")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

    def closeEvent(self, event):
        """重载关闭事件"""
        self.close_beyond_compare()  # 关闭 Beyond Compare
        event.accept()  # 允许关闭事件


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 获取图标文件的路径
    if getattr(sys, "frozen", False):
        icon_path = os.path.join(sys._MEIPASS, "logo.ico")
    else:
        icon_path = "logo.ico"

    # 设置程序的图标
    app.setWindowIcon(QIcon(icon_path))
    ex = FileComparisonApp()
    # 默认最大化窗口
    ex.showNormal()
    ex.raise_()  # 提升窗口到最前
    ex.activateWindow()  # 激活窗口
    sys.exit(app.exec_())
