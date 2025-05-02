import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFormLayout, QMessageBox, QComboBox, QHBoxLayout
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from core.budget_utils import process_budget
from config.constants import LOCAL_FILE, DEFAULT_FONT, FONT_SIZE
from core.logger import log_action

class Earmark_Monitoring(QWidget):
    def __init__(self, switch_page):
        super().__init__()
        self.switch_page = switch_page
        self.initUI()

    def initUI(self):
        if not os.path.exists(LOCAL_FILE):
            QMessageBox.critical(self, "Error", f"{LOCAL_FILE} not found.")
            sys.exit()

        df = pd.read_excel(LOCAL_FILE)
        self.expenses = df["Expense"].dropna().tolist()

        form_layout = QFormLayout()
        form_layout.setSpacing(15)

        self.expense_dropdown = QComboBox()
        self.expense_dropdown.addItems(self.expenses)
        form_layout.addRow(QLabel("Select Expense:"), self.expense_dropdown)

        self.inputs = {}
        for label in ["Allotment", "Realignment", "Obligations"]:
            line_edit = QLineEdit()
            line_edit.setPlaceholderText(f"Enter {label.lower()} (in PHP)")
            line_edit.setFont(QFont(DEFAULT_FONT, FONT_SIZE))
            self.inputs[label] = line_edit
            form_layout.addRow(QLabel(label + ":"), line_edit)

        self.earmarked_input = QLineEdit()
        self.earmarked_input.setPlaceholderText("Enter earmarked value (in PHP)")
        form_layout.addRow(QLabel("Earmarked:"), self.earmarked_input)

        submit_btn = QPushButton("Submit")
        submit_btn.setFont(QFont(DEFAULT_FONT, FONT_SIZE))
        submit_btn.clicked.connect(self.on_submit)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(submit_btn)
        btn_layout.addStretch()

        btn_back = QPushButton("⬅ Back")
        btn_back.setFont(QFont(DEFAULT_FONT, FONT_SIZE))
        btn_back.clicked.connect(lambda: self.switch_page("main"))

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(30, 30, 30, 30)

        title = QLabel("ERC Earmark Management System")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(QFont("Segoe UI", 25, QFont.Bold))
        main_layout.addWidget(title)
        main_layout.addSpacing(20)

        main_layout.addLayout(form_layout)
        main_layout.addSpacing(20)
        main_layout.addLayout(btn_layout)
        main_layout.addSpacing(20)
        main_layout.addWidget(btn_back, alignment=Qt.AlignCenter)
        main_layout.addStretch()

        self.setLayout(main_layout)

    def on_submit(self):
        try:
            allotment = float(self.inputs["Allotment"].text())
            realignment = float(self.inputs["Realignment"].text())
            obligations = float(self.inputs["Obligations"].text())
            earmarked = float(self.earmarked_input.text())
            expense = self.expense_dropdown.currentText()

            process_budget(expense, allotment, realignment, obligations, earmarked)
            QMessageBox.information(self, "Success", f"Expense '{expense}' updated in {LOCAL_FILE}.")
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Please enter valid numeric values.")
        except Exception as e:
            QMessageBox.critical(self, "Processing Error", f"An error occurred:\n{e}")
