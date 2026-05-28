#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import json
import threading
from pathlib import Path
from docx import Document

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QProgressBar, QFileDialog,
    QMessageBox, QHBoxLayout, QVBoxLayout, QGridLayout, QHeaderView,
    QScrollArea, QFrame, QSizePolicy, QSpacerItem, QGroupBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPalette, QColor, QBrush


class WordProcessor:
    MAX_FILE_SIZE = 50 * 1024 * 1024
    MAX_REPLACE_TEXT_LENGTH = 10000

    def __init__(self):
        pass

    def get_word_files(self, path):
        word_files = []
        try:
            if not path:
                return []

            if isinstance(path, (str, Path)):
                path = Path(path)
                if not path.exists():
                    return []

                try:
                    path = path.resolve()
                except Exception:
                    return []

                if path.is_dir():
                    for ext in ['*.doc', '*.docx']:
                        for file in path.glob(ext):
                            try:
                                if self._validate_file(file):
                                    word_files.append(str(file))
                            except (OSError, IOError):
                                continue
                elif path.is_file() and path.suffix.lower() in ['.doc', '.docx']:
                    try:
                        if self._validate_file(path):
                            word_files.append(str(path))
                    except (OSError, IOError):
                        pass
            elif isinstance(path, list):
                for p in path:
                    p = Path(p)
                    if p.is_file() and p.suffix.lower() in ['.doc', '.docx']:
                        try:
                            if self._validate_file(p):
                                word_files.append(str(p))
                        except (OSError, IOError):
                            continue
        except Exception:
            return []

        return word_files

    def _validate_file(self, file_path):
        if not file_path.exists():
            return False
        if not file_path.is_file():
            return False
        if file_path.stat().st_size > self.MAX_FILE_SIZE:
            return False
        if not os.access(file_path, os.R_OK):
            return False
        return True

    def _normalize_path(self, path_str):
        try:
            path_obj = Path(path_str)
            if not path_obj.exists():
                return None
            resolved_path = path_obj.resolve()
            if '..' in str(path_obj):
                return None
            return resolved_path
        except Exception:
            return None

    def replace_text_in_paragraph(self, paragraph, replacements):
        sorted_replacements = sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True)

        for old_text, new_text in sorted_replacements:
            if not old_text:
                continue
            if len(old_text) > self.MAX_REPLACE_TEXT_LENGTH:
                continue
            if new_text and len(new_text) > self.MAX_REPLACE_TEXT_LENGTH:
                continue

            paragraph_text = paragraph.text
            if old_text not in paragraph_text:
                continue

            self._replace_text_in_paragraph_clean(paragraph, old_text, new_text)

    def _replace_text_in_paragraph_clean(self, paragraph, old_text, new_text):
        runs = paragraph.runs
        if not runs:
            return

        full_text = paragraph.text
        start_pos = full_text.find(old_text)
        if start_pos == -1:
            return
        end_pos = start_pos + len(old_text)

        run_positions = []
        current_pos = 0
        for i, run in enumerate(runs):
            run_len = len(run.text)
            run_positions.append((i, current_pos, current_pos + run_len, run))
            current_pos += run_len

        first_run_idx = None
        first_local_start = None
        last_run_idx = None
        last_local_end = None

        for idx, run_start, run_end, run in run_positions:
            overlap_start = max(run_start, start_pos)
            overlap_end = min(run_end, end_pos)
            if overlap_start < overlap_end:
                if first_run_idx is None:
                    first_run_idx = idx
                    first_local_start = overlap_start - run_start
                last_run_idx = idx
                last_local_end = overlap_end - run_start

        if first_run_idx is None or last_run_idx is None:
            return

        first_run = runs[first_run_idx]
        first_run.text = first_run.text[:first_local_start] + new_text + first_run.text[last_local_end:]

        runs_to_remove = []
        for idx in range(first_run_idx + 1, last_run_idx + 1):
            runs_to_remove.append(runs[idx])

        for run in runs_to_remove:
            try:
                run._element.getparent().remove(run._element)
            except Exception:
                run.text = ""

    def process_document(self, doc_path, replacements, output_path, callback=None):
        try:
            if not doc_path.lower().endswith('.docx'):
                return False

            file_size = Path(doc_path).stat().st_size
            if file_size > self.MAX_FILE_SIZE:
                print(f"文件过大：{doc_path} ({file_size / 1024 / 1024:.2f}MB)")
                return False

            output_path_obj = Path(output_path).resolve()
            if not str(output_path_obj).lower().endswith('.docx'):
                return False

            doc = Document(doc_path)

            for paragraph in doc.paragraphs:
                self.replace_text_in_paragraph(paragraph, replacements)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            self.replace_text_in_paragraph(paragraph, replacements)

            doc.save(str(output_path_obj))

            if callback:
                callback()

            return True
        except FileNotFoundError as e:
            print(f"文件不存在 {doc_path}: {str(e)}")
            return False
        except PermissionError as e:
            print(f"权限错误 {doc_path}: {str(e)}")
            return False
        except Exception as e:
            print(f"处理文档失败 {doc_path}: {str(e)}")
            import traceback
            traceback.print_exc()
            return False


class ExcelProcessor:
    def __init__(self):
        pass

    def write_number_to_excel(self, excel_path, output_path, number_text, use_fixed_mode=False, fixed_number="", number_format="", callback=None):
        try:
            from openpyxl import load_workbook
            from openpyxl.styles import Font, Alignment

            wb = load_workbook(excel_path)

            header_font = Font(name='Times New Roman', bold=True, size=11)
            center_align = Alignment(horizontal='center', vertical='center')

            if use_fixed_mode and fixed_number:
                sheet_num = 1
                for ws in wb.worksheets:
                    ws.insert_rows(1)

                    ws.merge_cells('E1:F1')
                    cell_e1 = ws['E1']
                    cell_e1.font = header_font
                    cell_e1.alignment = center_align
                    cell_e1.value = fixed_number

                    ws.merge_cells('E2:F2')
                    cell_e2 = ws['E2']
                    cell_e2.font = header_font
                    cell_e2.alignment = center_align
                    sheet_num_str = str(sheet_num).zfill(2) if sheet_num <= 99 else str(sheet_num)
                    cell_e2.value = number_format.replace("01", sheet_num_str)
                    sheet_num += 1
            else:
                sheet_num = 1
                for ws in wb.worksheets:
                    ws.insert_rows(1)

                    ws.merge_cells('E1:F1')
                    cell_e1 = ws['E1']
                    cell_e1.font = header_font
                    cell_e1.alignment = center_align
                    sheet_num_str = str(sheet_num).zfill(2) if sheet_num <= 99 else str(sheet_num)
                    cell_e1.value = number_text.replace("01", sheet_num_str)
                    sheet_num += 1

            wb.save(output_path)

            if callback:
                callback()

            return True
        except ImportError:
            print("缺少 openpyxl 库，请安装：pip install openpyxl")
            return False
        except Exception as e:
            print(f"处理 Excel 文件失败 {excel_path}: {str(e)}")
            import traceback
            traceback.print_exc()
            return False


class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, str)
    processing_complete = pyqtSignal(int, int)

    def __init__(self, processor, files, replacements, output_folder):
        super().__init__()
        self.processor = processor
        self.files = files
        self.replacements = replacements
        self.output_folder = output_folder

    def run(self):
        total_files = len(self.files)
        success_count = 0

        for i, file_path in enumerate(self.files):
            file_name = Path(file_path).name
            base_name = Path(file_path).stem

            self.progress_updated.emit(i, f"正在处理：{file_name} ({i + 1}/{total_files})")

            if self.output_folder:
                output_dir = Path(self.output_folder)
            else:
                file_dir = Path(file_path).parent
                output_dir = file_dir / "修改版"

            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / f"{base_name}-修改版.docx"

            success = self.processor.process_document(
                file_path, self.replacements, str(output_path)
            )

            if success:
                success_count += 1

            progress = int(((i + 1) / total_files) * 100)
            self.progress_updated.emit(progress, "")

        self.processing_complete.emit(success_count, total_files)


class WriteNumberThread(QThread):
    progress_updated = pyqtSignal(int, str)
    write_complete = pyqtSignal(bool, str)

    def __init__(self, excel_processor, excel_file, output_path, number_text, use_fixed_mode, fixed_number, number_format):
        super().__init__()
        self.excel_processor = excel_processor
        self.excel_file = excel_file
        self.output_path = output_path
        self.number_text = number_text
        self.use_fixed_mode = use_fixed_mode
        self.fixed_number = fixed_number
        self.number_format = number_format

    def run(self):
        self.progress_updated.emit(0, f"正在写入编号：{self.number_text}")

        success = self.excel_processor.write_number_to_excel(
            self.excel_file,
            self.output_path,
            self.number_text,
            use_fixed_mode=self.use_fixed_mode,
            fixed_number=self.fixed_number,
            number_format=self.number_format
        )

        self.progress_updated.emit(100, "")
        self.write_complete.emit(success, self.number_text)


class WordProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word 文本处理工具")
        self.setMinimumSize(900, 800)

        screen_geometry = QApplication.desktop().availableGeometry()
        window_width = min(1000, screen_geometry.width() - 100)
        window_height = min(900, screen_geometry.height() - 100)
        self.resize(window_width, window_height)

        self.processor = WordProcessor()
        self.excel_processor = ExcelProcessor()
        self.selected_files = []
        self.output_folder = ""

        self.settings_file = Path(__file__).parent / "settings.json"
        self.number_settings = self._load_number_settings()

        self._init_styles()
        self._create_main_layout()

    def _init_styles(self):
        self.primary_color = "#14B8A6"
        self.primary_dark = "#0D9488"
        self.primary_light = "#F0FDFA"
        self.primary_hover = "#0F766E"
        self.title_color = "#0A7A6E"
        self.success_color = "#10B981"
        self.warning_color = "#F59E0B"
        self.error_color = "#EF4444"
        self.text_color = "#000000"
        self.label_color = "#808080"
        self.default_text_color = "#000000"
        self.bg_color = "#E0F2FE"
        self.card_bg = "#F8FAFC"
        self.input_bg = "#FFFFFF"
        self.border_color = "#CBD5E1"
        self.button_bg = "#E5E7EB"
        self.button_hover = "#D1D5DB"
        self.button_pressed = "#9CA3AF"

        font = QFont("Microsoft YaHei UI", 9)
        QApplication.setFont(font)

        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {self.bg_color};
            }}
            QWidget {{
                font-family: Microsoft YaHei UI;
                font-size: 9pt;
                color: {self.text_color};
            }}
            QTabWidget::pane {{
                border: 1px solid {self.border_color};
                background-color: {self.card_bg};
                border-radius: 6px;
            }}
            QTabBar::tab {{
                background-color: #DBEAFE;
                color: {self.text_color};
                padding: 8px 20px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 100px;
                font-size: 9pt;
            }}
            QTabBar::tab:selected {{
                background-color: {self.card_bg};
                color: {self.primary_color};
                font-weight: bold;
            }}
            QTabBar::tab:hover:!selected {{
                background-color: #BFDBFE;
            }}
            QGroupBox {{
                background-color: {self.card_bg};
                border: 1px solid {self.border_color};
                border-radius: 6px;
                margin-top: 8px;
                padding: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                color: {self.primary_color};
                font-weight: bold;
                padding-left: 6px;
                padding-right: 6px;
                subcontrol-origin: margin;
                subcontrol-position: top left;
            }}
            QPushButton {{
                background-color: {self.button_bg};
                color: {self.text_color};
                border: 1px solid {self.border_color};
                border-radius: 6px;
                padding: 8px 18px;
                font-weight: bold;
                font-size: 9pt;
                min-width: 85px;
            }}
            QPushButton:hover {{
                background-color: {self.button_hover};
            }}
            QPushButton:pressed {{
                background-color: {self.button_pressed};
            }}
            QPushButton:disabled {{
                background-color: #F3F4F6;
                color: #9CA3AF;
            }}
            QLineEdit {{
                border: 1px solid {self.border_color};
                border-radius: 6px;
                padding: 6px 10px;
                background-color: {self.input_bg};
                font-size: 9pt;
                color: {self.text_color};
            }}
            QLineEdit:focus {{
                border-color: {self.primary_color};
            }}
            QLineEdit:readonly {{
                background-color: #F1F5F9;
                color: {self.text_color};
            }}
            QTableWidget {{
                border: 1px solid {self.border_color};
                border-radius: 6px;
                background-color: {self.input_bg};
                gridline-color: {self.border_color};
                font-size: 9pt;
                color: {self.text_color};
            }}
            QTableWidget::item {{
                padding: 6px 8px;
                border-bottom: 1px solid {self.border_color};
                color: {self.text_color};
            }}
            QTableWidget::item:selected {{
                background-color: {self.primary_light};
                color: {self.text_color};
            }}
            QHeaderView::section {{
                background-color: #BFDBFE;
                color: {self.text_color};
                font-weight: bold;
                padding: 6px 8px;
                border: none;
                border-bottom: 1px solid {self.primary_color};
            }}
            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: {self.border_color};
                height: 12px;
                text-align: center;
                color: {self.text_color};
            }}
            QProgressBar::chunk {{
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 {self.primary_color}, stop:1 {self.primary_dark});
                border-radius: 4px;
            }}
            QScrollBar::vertical {{
                width: 8px;
                background-color: transparent;
                border-radius: 4px;
            }}
            QScrollBar::handle:vertical {{
                background-color: {self.border_color};
                border-radius: 4px;
                min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: #9CA3AF;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QLabel {{
                color: {self.text_color};
            }}
        """)

    def _load_number_settings(self):
        try:
            if self.settings_file.exists():
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass

        return {
            "prefix": "HT合同编号",
            "current_number": 1,
            "number_format": "HT001-2026-XM001-01",
            "fixed_number": ""
        }

    def _save_number_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.number_settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存设置失败: {e}")

    def _generate_number(self):
        number = self.number_settings["current_number"]
        if number <= 99:
            number_str = str(number).zfill(2)
        else:
            number_str = str(number)

        number_format = self.number_settings["number_format"]
        full_number = number_format.replace("01", number_str)

        self.number_settings["current_number"] = number + 1
        self._save_number_settings()

        return full_number

    def _generate_preview(self):
        number = self.number_settings["current_number"]
        if number <= 99:
            number_str = str(number).zfill(2)
        else:
            number_str = str(number)

        number_format = self.number_settings.get("number_format", "HT001-2026-XM001-01")
        return number_format.replace("01", number_str)

    def _create_main_layout(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(10, 8, 10, 8)
        main_layout.setSpacing(6)

        title_label = QLabel("Word 文本处理工具")
        title_label.setStyleSheet(f"font-size: 14pt; font-weight: bold; color: {self.title_color};")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        desc_label = QLabel("批量替换 Word 文档中的文本，保持原有格式不变")
        desc_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        desc_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(desc_label)

        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("QTabWidget::tab-bar { alignment: center; }")

        self.text_replace_tab = QWidget()
        self.number_record_tab = QWidget()

        self.tab_widget.addTab(self.text_replace_tab, "文本替换")
        self.tab_widget.addTab(self.number_record_tab, "记录编号")

        main_layout.addWidget(self.tab_widget)

        self._create_text_replace_tab()
        self._create_number_record_tab()

    def _create_text_replace_tab(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        self._create_file_section(layout)
        self._create_file_list_section(layout)
        self._create_file_buttons(layout)
        self._create_replacement_section(layout)
        self._create_output_section(layout)
        self._create_progress_section(layout)
        self._create_button_section(layout)

        layout.addStretch()

        scroll_area.setWidget(scroll_content)

        tab_layout = QVBoxLayout(self.text_replace_tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.setSpacing(0)
        tab_layout.addWidget(scroll_area)

    def _create_file_section(self, parent_layout):
        group_box = QGroupBox("文件选择（可选择文件夹或文件）")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        browse_layout = QHBoxLayout()
        browse_layout.setSpacing(5)

        browse_layout.addWidget(QLabel("选择文件/文件夹:"))

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.file_path_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        browse_layout.addWidget(self.file_path_edit)

        browse_file_btn = QPushButton("浏览文件")
        browse_file_btn.clicked.connect(self._browse_files)
        browse_layout.addWidget(browse_file_btn)

        browse_folder_btn = QPushButton("浏览文件夹")
        browse_folder_btn.clicked.connect(self._browse_folder)
        browse_layout.addWidget(browse_folder_btn)

        group_layout.addLayout(browse_layout)

        info_layout = QHBoxLayout()
        info_layout.setSpacing(5)

        self.file_count_label = QLabel("未选择文件")
        self.file_count_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        info_layout.addWidget(self.file_count_label)

        self.file_path_detail_label = QLabel("")
        self.file_path_detail_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        self.file_path_detail_label.setWordWrap(True)
        info_layout.addWidget(self.file_path_detail_label)
        info_layout.addStretch()

        group_layout.addLayout(info_layout)

        parent_layout.addWidget(group_box)

    def _create_file_list_section(self, parent_layout):
        group_box = QGroupBox("已选文件列表")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        self.file_table = QTableWidget()
        self.file_table.setColumnCount(1)
        self.file_table.setHorizontalHeaderLabels(["文件名"])
        self.file_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.file_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.file_table.doubleClicked.connect(self._remove_selected_file)
        self.file_table.setMinimumHeight(100)
        self.file_table.setMaximumHeight(120)
        self.file_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        scroll_area = QScrollArea()
        scroll_area.setWidget(self.file_table)
        scroll_area.setWidgetResizable(True)

        group_layout.addWidget(scroll_area)

        parent_layout.addWidget(group_box)

    def _create_file_buttons(self, parent_layout):
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(6)

        add_btn = QPushButton("+ 添加文件")
        add_btn.clicked.connect(self._add_file)
        add_btn.setFixedHeight(29)
        btn_layout.addWidget(add_btn)

        remove_btn = QPushButton("- 删除文件")
        remove_btn.clicked.connect(self._remove_selected_file)
        remove_btn.setFixedHeight(29)
        btn_layout.addWidget(remove_btn)

        btn_layout.addStretch()

        parent_layout.addLayout(btn_layout)

    def _create_replacement_section(self, parent_layout):
        group_box = QGroupBox("替换规则设置")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(5)

        self.replacement_table = QTableWidget()
        self.replacement_table.setColumnCount(3)
        self.replacement_table.setHorizontalHeaderLabels(["类别", "查找", "替换"])
        self.replacement_table.setColumnWidth(0, 100)
        self.replacement_table.setColumnWidth(1, 300)
        self.replacement_table.setColumnWidth(2, 300)
        self.replacement_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.replacement_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.replacement_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.replacement_table.verticalHeader().setVisible(False)
        self.replacement_table.setSelectionBehavior(QTableWidget.SelectItems)
        self.replacement_table.cellDoubleClicked.connect(self._edit_cell)
        self.replacement_table.horizontalHeader().setMinimumSectionSize(80)

        scroll_area = QScrollArea()
        scroll_area.setWidget(self.replacement_table)
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(150)

        group_layout.addWidget(scroll_area)

        btn_layout = QHBoxLayout()
        reset_btn = QPushButton("重置为默认值")
        reset_btn.clicked.connect(self._reset_to_defaults)
        reset_btn.setFixedHeight(29)
        btn_layout.addWidget(reset_btn)
        btn_layout.addStretch()

        group_layout.addLayout(btn_layout)

        parent_layout.addWidget(group_box)

        self._init_default_rules()

    def _create_output_section(self, parent_layout):
        group_box = QGroupBox("保存设置")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        browse_layout = QHBoxLayout()
        browse_layout.setSpacing(5)

        browse_layout.addWidget(QLabel("保存路径:"))

        self.output_path_edit = QLineEdit()
        self.output_path_edit.setReadOnly(True)
        self.output_path_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        browse_layout.addWidget(self.output_path_edit)

        browse_btn = QPushButton("浏览")
        browse_btn.clicked.connect(self._browse_output)
        browse_btn.setFixedHeight(29)
        browse_layout.addWidget(browse_btn)

        group_layout.addLayout(browse_layout)

        self.output_info_label = QLabel("未选择保存路径时将按规则自动创建文件夹，默认在同级目录下新建修改版文件夹")
        self.output_info_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        self.output_info_label.setWordWrap(True)
        group_layout.addWidget(self.output_info_label)

        parent_layout.addWidget(group_box)

    def _create_progress_section(self, parent_layout):
        group_box = QGroupBox("处理进度")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        group_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("就绪")
        self.progress_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        group_layout.addWidget(self.progress_label)

        parent_layout.addWidget(group_box)

    def _create_button_section(self, parent_layout):
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        self.start_btn = QPushButton("开始处理")
        self.start_btn.clicked.connect(self._start_processing)
        self.start_btn.setFixedHeight(33)
        self.start_btn.setMinimumWidth(100)

        clear_btn = QPushButton("清空设置")
        clear_btn.clicked.connect(self._clear_all)
        clear_btn.setFixedHeight(33)
        clear_btn.setMinimumWidth(100)

        btn_layout.addStretch()
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(clear_btn)
        btn_layout.addStretch()

        parent_layout.addLayout(btn_layout)

    def _init_default_rules(self):
        self.default_rules = [
            ("项目编号", "HT001-2026-XM001-01", "HT001-2026-XM001-01"),
            ("公司名称", "", ""),
            ("系统名称", "", ""),
            ("人员名称", "", "")
        ]

        self.replacement_table.setRowCount(4)
        for row, (category, original, replacement) in enumerate(self.default_rules):
            category_item = QTableWidgetItem(category)
            category_item.setFlags(category_item.flags() & ~Qt.ItemIsEditable)
            category_item.setForeground(QColor(self.default_text_color))
            category_item.setTextAlignment(Qt.AlignCenter)
            self.replacement_table.setItem(row, 0, category_item)

            original_item = QTableWidgetItem(original)
            if original:
                original_item.setForeground(QColor(self.default_text_color))
            self.replacement_table.setItem(row, 1, original_item)
            self.replacement_table.setItem(row, 2, QTableWidgetItem(replacement))

    def _reset_to_defaults(self):
        self.replacement_table.clearContents()
        self._init_default_rules()

    def _edit_cell(self, row, column):
        if column == 0:
            return

        item = self.replacement_table.item(row, column)
        if item:
            self.replacement_table.editItem(item)

    def _browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 Word 文件",
            "",
            "Word 文件 (*.docx *.doc);;所有文件 (*.*)"
        )
        if files:
            self.selected_files = list(files)
            self._update_file_display()

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含 Word 文件的文件夹")
        if folder:
            word_files = self.processor.get_word_files(folder)
            if word_files:
                self.selected_files = word_files
                self._update_file_display()
            else:
                QMessageBox.warning(self, "警告", "该文件夹下没有找到 Word 文件")

    def _add_file(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "添加 Word 文件",
            "",
            "Word 文件 (*.docx *.doc);;所有文件 (*.*)"
        )
        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
            self._update_file_display()

    def _remove_selected_file(self):
        selected_items = self.file_table.selectedItems()
        if selected_items:
            rows_to_remove = set()
            for item in selected_items:
                rows_to_remove.add(item.row())

            for row in sorted(rows_to_remove, reverse=True):
                file_path = self.file_table.item(row, 0).text()
                if file_path in self.selected_files:
                    self.selected_files.remove(file_path)
                self.file_table.removeRow(row)

            self._update_file_display()

    def _update_file_display(self):
        self.file_table.setRowCount(0)

        if self.selected_files:
            self.file_count_label.setText(f"已选择 {len(self.selected_files)} 个文件")
            self.file_count_label.setStyleSheet(f"color: {self.success_color}; font-weight: bold;")
            self.file_path_edit.setText(f"已选择 {len(self.selected_files)} 个文件")

            paths = "\n".join(self.selected_files[:5])
            if len(self.selected_files) > 5:
                paths += f"\n... 及其他 {len(self.selected_files) - 5} 个文件"
            self.file_path_detail_label.setText(paths)

            for file_path in self.selected_files:
                row = self.file_table.rowCount()
                self.file_table.insertRow(row)
                self.file_table.setItem(row, 0, QTableWidgetItem(file_path))
        else:
            self.file_count_label.setText("未选择文件")
            self.file_count_label.setStyleSheet(f"color: {self.label_color};")
            self.file_path_edit.setText("")
            self.file_path_detail_label.setText("")

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "选择保存路径")
        if folder:
            self.output_folder = folder
            self.output_path_edit.setText(folder)
            self.output_info_label.setText("已选择自定义保存路径")
            self.output_info_label.setStyleSheet(f"color: {self.success_color}; font-size: 9pt;")

    def _get_all_replacements(self):
        replacements = {}
        for row in range(self.replacement_table.rowCount()):
            original_item = self.replacement_table.item(row, 1)
            replacement_item = self.replacement_table.item(row, 2)

            if original_item:
                original = original_item.text().strip()
                if original:
                    replacement = replacement_item.text().strip() if replacement_item else ""
                    replacements[original] = replacement

        return replacements

    def _clear_all(self):
        self.selected_files = []
        self.output_folder = ""
        self.file_path_edit.setText("")
        self.output_path_edit.setText("")
        self.file_count_label.setText("未选择文件")
        self.file_count_label.setStyleSheet(f"color: {self.label_color};")
        self.output_info_label.setText("未选择保存路径时将按规则自动创建文件夹，默认在同级目录下新建修改版文件夹")
        self.output_info_label.setStyleSheet(f"color: {self.label_color}; font-size: 9pt;")
        self.progress_bar.setValue(0)
        self.progress_label.setText("就绪")

        self.file_table.setRowCount(0)
        self._reset_to_defaults()

    def _start_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择要处理的 Word 文件")
            return

        replacements = self._get_all_replacements()
        if not replacements:
            QMessageBox.warning(self, "警告", "请至少添加一条替换规则")
            return

        self.start_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_label.setText("准备处理...")

        self.processing_thread = ProcessingThread(
            self.processor,
            self.selected_files,
            replacements,
            self.output_folder
        )
        self.processing_thread.progress_updated.connect(self._update_progress)
        self.processing_thread.processing_complete.connect(self._processing_complete)
        self.processing_thread.start()

    def _update_progress(self, progress, message):
        if message:
            self.progress_label.setText(message)
        if progress >= 0:
            self.progress_bar.setValue(progress)

    def _processing_complete(self, success_count, total_files):
        self.start_btn.setEnabled(True)

        message = f"处理完成！\n成功：{success_count}/{total_files}"
        if success_count == total_files:
            message += "\n\n修改完成，已保存"
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.warning(self, "完成", message)

        self.progress_label.setText("处理完成")

    def _create_number_record_tab(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        self._create_number_mode_section(layout)
        self._create_number_format_section(layout)
        self._create_number_fixed_section(layout)
        self._create_number_file_section(layout)
        self._create_number_progress_section(layout)
        self._create_number_button_section(layout)

        layout.addStretch()

        scroll_area.setWidget(scroll_content)

        tab_layout = QVBoxLayout(self.number_record_tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.setSpacing(0)
        tab_layout.addWidget(scroll_area)

    def _create_number_mode_section(self, parent_layout):
        group_box = QGroupBox("编号模式")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(6)

        self.number_mode = "auto"

        mode_layout = QHBoxLayout()
        mode_layout.setSpacing(15)

        self.auto_radio = QPushButton("自动编号模式")
        self.auto_radio.setCheckable(True)
        self.auto_radio.setChecked(True)
        self.auto_radio.clicked.connect(lambda: self._on_mode_change("auto"))
        self.auto_radio.setFixedHeight(30)
        self.auto_radio.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.button_bg};
                border: 1px solid {self.border_color};
                border-radius: 6px;
                padding: 6px 14px;
                text-align: center;
                font-size: 9pt;
                color: {self.text_color};
            }}
            QPushButton:hover {{
                background-color: {self.button_hover};
            }}
            QPushButton:checked {{
                background-color: {self.primary_light};
                border-color: {self.primary_color};
                font-weight: bold;
            }}
        """)
        mode_layout.addWidget(self.auto_radio)

        self.fixed_radio = QPushButton("固定编号模式")
        self.fixed_radio.setCheckable(True)
        self.fixed_radio.clicked.connect(lambda: self._on_mode_change("fixed"))
        self.fixed_radio.setFixedHeight(30)
        self.fixed_radio.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.button_bg};
                border: 1px solid {self.border_color};
                border-radius: 6px;
                padding: 6px 14px;
                text-align: center;
                font-size: 9pt;
                color: {self.text_color};
            }}
            QPushButton:hover {{
                background-color: {self.button_hover};
            }}
            QPushButton:checked {{
                background-color: {self.primary_light};
                border-color: {self.primary_color};
                font-weight: bold;
            }}
        """)
        mode_layout.addWidget(self.fixed_radio)

        group_layout.addLayout(mode_layout)

        parent_layout.addWidget(group_box)

    def _create_number_format_section(self, parent_layout):
        group_box = QGroupBox("编号格式设置")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        format_layout = QHBoxLayout()
        format_layout.setSpacing(5)

        format_layout.addWidget(QLabel("编号格式:"))

        self.number_format_edit = QLineEdit()
        self.number_format_edit.setText(self.number_settings.get("number_format", "HT001-2026-XM001-01"))
        self.number_format_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        format_layout.addWidget(self.number_format_edit)

        group_layout.addLayout(format_layout)

        desc_label = QLabel("说明：使用 01 作为序号占位符，序号自动递增，超过99后自动变为3位数")
        desc_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        group_layout.addWidget(desc_label)

        parent_layout.addWidget(group_box)

    def _create_number_fixed_section(self, parent_layout):
        self.fixed_number_group_box = QGroupBox("固定编号（仅在固定编号模式下使用）")
        group_layout = QVBoxLayout(self.fixed_number_group_box)
        group_layout.setSpacing(4)

        fixed_layout = QHBoxLayout()
        fixed_layout.setSpacing(5)

        fixed_layout.addWidget(QLabel("固定编号:"))

        self.fixed_number_edit = QLineEdit()
        self.fixed_number_edit.setText(self.number_settings.get("fixed_number", ""))
        self.fixed_number_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        fixed_layout.addWidget(self.fixed_number_edit)

        group_layout.addLayout(fixed_layout)

        info_label = QLabel("固定编号将写入 E1:F1 合并单元格，编号格式将写入 E2:F2 合并单元格")
        info_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        group_layout.addWidget(info_label)

        parent_layout.addWidget(self.fixed_number_group_box)

        self.fixed_number_group_box.setVisible(self.number_mode == "fixed")

    def _on_mode_change(self, mode):
        self.number_mode = mode
        if mode == "auto":
            self.auto_radio.setChecked(True)
            self.fixed_radio.setChecked(False)
            self.fixed_number_group_box.setVisible(False)
        else:
            self.fixed_radio.setChecked(True)
            self.auto_radio.setChecked(False)
            self.fixed_number_group_box.setVisible(True)

    def _create_number_file_section(self, parent_layout):
        group_box = QGroupBox("选择 Excel 文件")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        browse_layout = QHBoxLayout()
        browse_layout.setSpacing(5)

        browse_layout.addWidget(QLabel("Excel 文件:"))

        self.excel_file_edit = QLineEdit()
        self.excel_file_edit.setReadOnly(True)
        self.excel_file_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        browse_layout.addWidget(self.excel_file_edit)

        browse_btn = QPushButton("浏览")
        browse_btn.clicked.connect(self._browse_excel_file)
        browse_btn.setFixedHeight(29)
        browse_layout.addWidget(browse_btn)

        group_layout.addLayout(browse_layout)

        self.excel_file_info_label = QLabel("未选择文件")
        self.excel_file_info_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        group_layout.addWidget(self.excel_file_info_label)

        parent_layout.addWidget(group_box)

        group_box2 = QGroupBox("保存设置")
        group_layout2 = QVBoxLayout(group_box2)
        group_layout2.setSpacing(4)

        save_layout = QHBoxLayout()
        save_layout.setSpacing(5)

        save_layout.addWidget(QLabel("保存路径:"))

        self.number_output_path_edit = QLineEdit()
        self.number_output_path_edit.setReadOnly(True)
        self.number_output_path_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        save_layout.addWidget(self.number_output_path_edit)

        save_browse_btn = QPushButton("浏览")
        save_browse_btn.clicked.connect(self._browse_number_output)
        save_browse_btn.setFixedHeight(29)
        save_layout.addWidget(save_browse_btn)

        group_layout2.addLayout(save_layout)

        self.number_output_info_label = QLabel("未选择保存路径时将自动保存在源文件同级目录，文件名后添加“-修改版”")
        self.number_output_info_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        self.number_output_info_label.setWordWrap(True)
        group_layout2.addWidget(self.number_output_info_label)

        parent_layout.addWidget(group_box2)

    def _create_number_progress_section(self, parent_layout):
        group_box = QGroupBox("处理进度")
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(4)

        self.number_progress_bar = QProgressBar()
        self.number_progress_bar.setRange(0, 100)
        self.number_progress_bar.setValue(0)
        group_layout.addWidget(self.number_progress_bar)

        self.number_progress_label = QLabel("就绪")
        self.number_progress_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")
        group_layout.addWidget(self.number_progress_label)

        parent_layout.addWidget(group_box)

    def _create_number_button_section(self, parent_layout):
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.write_number_btn = QPushButton("写入编号")
        self.write_number_btn.clicked.connect(self._start_write_number)
        self.write_number_btn.setFixedHeight(33)
        self.write_number_btn.setMinimumWidth(100)

        reset_number_btn = QPushButton("重置编号")
        reset_number_btn.clicked.connect(self._reset_number)
        reset_number_btn.setFixedHeight(33)
        reset_number_btn.setMinimumWidth(100)

        btn_layout.addStretch()
        btn_layout.addWidget(self.write_number_btn)
        btn_layout.addWidget(reset_number_btn)
        btn_layout.addStretch()

        parent_layout.addLayout(btn_layout)

    def _browse_excel_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 文件",
            "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if file:
            self.excel_file_path = file
            self.excel_file_edit.setText(file)
            self.excel_file_info_label.setText(f"已选择：{Path(file).name}")
            self.excel_file_info_label.setStyleSheet(f"color: {self.success_color};")
            self.number_output_path_edit.setText("")
            self.number_output_info_label.setText("未选择保存路径时将自动保存在源文件同级目录，文件名后添加“-修改版”")
            self.number_output_info_label.setStyleSheet(f"color: {self.label_color}; font-size: 8pt;")

    def _browse_number_output(self):
        folder = QFileDialog.getExistingDirectory(self, "选择保存路径")
        if folder:
            self.number_output_folder = folder
            self.number_output_path_edit.setText(folder)
            self.number_output_info_label.setText("已选择自定义保存路径")
            self.number_output_info_label.setStyleSheet(f"color: {self.success_color}; font-size: 8pt;")

    def _start_write_number(self):
        if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
            QMessageBox.warning(self, "警告", "请先选择要写入的 Excel 文件")
            return

        number_format = self.number_format_edit.text().strip()
        if not number_format:
            QMessageBox.warning(self, "警告", "请输入编号格式")
            return

        self.number_settings["number_format"] = number_format

        use_fixed_mode = self.number_mode == "fixed"
        fixed_number = ""

        if use_fixed_mode:
            fixed_number = self.fixed_number_edit.text().strip()
            if not fixed_number:
                QMessageBox.warning(self, "警告", "请输入固定编号")
                return
            self.number_settings["fixed_number"] = fixed_number

        self._save_number_settings()

        output_path = getattr(self, 'number_output_folder', None)
        if output_path:
            src_file = Path(self.excel_file_path)
            output_file = Path(output_path) / f"{src_file.stem}-修改版{src_file.suffix}"
        else:
            src_file = Path(self.excel_file_path)
            output_file = src_file.parent / f"{src_file.stem}-修改版{src_file.suffix}"

        self.write_number_btn.setEnabled(False)
        self.number_progress_label.setText("正在处理...")
        self.number_progress_bar.setValue(0)

        number_text = self._generate_number()

        self.write_thread = WriteNumberThread(
            self.excel_processor,
            self.excel_file_path,
            str(output_file),
            number_text,
            use_fixed_mode,
            fixed_number,
            number_format
        )
        self.write_thread.progress_updated.connect(self._update_number_progress)
        self.write_thread.write_complete.connect(self._write_number_complete)
        self.write_thread.start()

    def _update_number_progress(self, progress, message):
        if message:
            self.number_progress_label.setText(message)
        if progress >= 0:
            self.number_progress_bar.setValue(progress)

    def _write_number_complete(self, success, number_text):
        self.write_number_btn.setEnabled(True)

        use_fixed_mode = self.number_mode == "fixed"
        fixed_number = getattr(self, 'fixed_number_edit', None)
        fixed_number = fixed_number.text().strip() if fixed_number else ""

        if success:
            if use_fixed_mode:
                self.number_progress_label.setText(f"写入成功！固定编号：{fixed_number}，格式编号：{number_text}")
                QMessageBox.information(self, "完成", f"编号写入成功！\n\n固定编号（E1:F1）：{fixed_number}\n格式编号（E2:F2）：{number_text}")
            else:
                self.number_progress_label.setText(f"写入成功！编号：{number_text}")
                QMessageBox.information(self, "完成", f"编号写入成功！\n\n编号（E1:F1）：{number_text}")
        else:
            self.number_progress_label.setText("写入失败")
            self.number_progress_label.setStyleSheet(f"color: {self.error_color};")
            QMessageBox.critical(self, "错误", "编号写入失败，请检查文件是否被占用或格式是否正确")

    def _reset_number(self):
        reply = QMessageBox.question(
            self,
            "确认",
            "确定要重置编号吗？当前编号将恢复为 01",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.number_settings["current_number"] = 1
            self._save_number_settings()
            QMessageBox.information(self, "完成", "编号已重置为 01")


def main():
    app = QApplication(sys.argv)

    font = QFont("Microsoft YaHei UI", 10)
    app.setFont(font)

    window = WordProcessorGUI()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()