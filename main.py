import sys
import pandas as pd
import re
from docxtpl import DocxTemplate
import platform
import subprocess

import os
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QTableWidget, QTableWidgetItem, QPushButton, \
    QFileDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QDateEdit, QComboBox

regex_fio = r"^[А-ЯЁ][а-яё]*([-][А-ЯЁ][а-яё]*)?\s[А-ЯЁ][а-яё]*\s[А-ЯЁ][а-яё]*$"

def open_file(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])

class ExampleWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Успешно")
        self.setGeometry(100, 100, 400, 300)

        label = QLabel("Успешно!", self)
        label.move(150, 50)

        button = QPushButton("Открыть папку", self)
        button.setToolTip("Click to open the directory with files in a file explorer window.")
        button.move(150, 100)
        button.clicked.connect(self.open_folder)

    def open_folder(self):
        open_file(f"{os.getcwd()}/result")
        exit(0)

class InputFields(QWidget):
    change_db = pyqtSignal()

    def __init__(self):
        super().__init__()
        with open("db", 'r') as file:
            self.db = file.read().split("\n")

        self.fields_layout = QVBoxLayout()

        self.add_field_button = QPushButton("Добавить поле")
        self.add_field_button.clicked.connect(self.add_field)

        self.remove_button = QPushButton("Удалить поле")
        self.remove_button.clicked.connect(self.remove_field)
        if len(self.db) != 0:
            self.remove_button.setEnabled(True)
        else:
            self.remove_button.setEnabled(False)

        self.buttons_layout = QHBoxLayout()
        self.buttons_layout.addWidget(self.remove_button)

        self.layout = QVBoxLayout()
        self.layout.addLayout(self.fields_layout)
        self.layout.addWidget(self.add_field_button)
        self.layout.addLayout(self.buttons_layout)

        for value in self.db:
            line = QLineEdit()

            line.textChanged.connect(self.on_text_changed)

            line.setText(value)
            self.fields_layout.addWidget(line)

        self.setLayout(self.layout)

    def on_text_changed(self):
        values = [self.fields_layout.itemAt(i).widget().text() for i in range(self.fields_layout.count())]
        with open("db", 'w') as file:
            file.write("\n".join(values))

        self.change_db.emit()

    def add_field(self):
        line = QLineEdit()
        line.textChanged.connect(self.on_text_changed)
        self.fields_layout.addWidget(line)
        self.remove_button.setEnabled(True)

    def remove_field(self):
        if self.fields_layout.count() > 0:
            self.fields_layout.itemAt(self.fields_layout.count() - 1).widget().setParent(None)

        if self.fields_layout.count() == 0:
            self.remove_button.setEnabled(False)

        values = [self.fields_layout.itemAt(i).widget().text() for i in range(self.fields_layout.count())]
        with open("db", 'w') as file:
            file.write("\n".join(values))

        self.change_db.emit()


class ExcelParser(QMainWindow):
    def __init__(self):
        super().__init__()

        self.people = []
        self.group = "Не найдено"

        self.combos = []
        self.combos_type_practise = []

        self.file_path = ""

        # Set window title and dimensions
        self.setWindowTitle("Excel Парсер")
        self.setGeometry(100, 100, 1200, 800)

        # Create a widget to hold the table and buttons
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        # Create a table widget to display the Excel data
        self.table = QTableWidget()
        self.central_layout = QVBoxLayout(self.central_widget)
        self.central_layout.addWidget(self.table)

        # Create a widget to hold the labels and input fields
        self.input_widget = QWidget(self.central_widget)
        self.input_layout = QHBoxLayout(self.input_widget)
        self.central_layout.addWidget(self.input_widget)

        # Create a label and input field for the second text input
        self.label2 = QLabel("Группа:")
        self.input2 = QLineEdit()
        self.input_layout.addWidget(self.label2)
        self.input_layout.addWidget(self.input2)

        self.label3 = QLabel('Дата начала:', self)
        # self.label3.move(10, 420)
        self.date1 = QDateEdit(self)
        self.date1.setCalendarPopup(True)

        self.input_layout.addWidget(self.label3)
        self.input_layout.addWidget(self.date1)
        # self.date1.move(100, 420)

        self.label4 = QLabel('Дата конца:', self)
        # self.label4.move(10, 450)
        self.date2 = QDateEdit(self)
        self.date2.setCalendarPopup(True)

        self.input_layout.addWidget(self.label4)
        self.input_layout.addWidget(self.date2)
        # self.date2.move(100, 450)

        # Create a button to load the Excel file
        self.load_button = QPushButton("Загрузить exel файл", self.central_widget)
        self.load_button.clicked.connect(self.load_file)
        self.central_layout.addWidget(self.load_button)

        # Create a button to generate the docs
        self.generate_button = QPushButton("Сгенерировать", self.central_widget)
        self.generate_button.clicked.connect(self.generate_docs)
        self.central_layout.addWidget(self.generate_button)

        # Create a button to show the db
        self.show_db = QPushButton("Редактировать списки практик", self.central_widget)
        self.show_db.clicked.connect(self.show_db_window)
        self.central_layout.addWidget(self.show_db)

    def change_db_handler(self):
        for combo in self.combos:
            self.load_values_in_combo(combo)

    def show_db_window(self):
        self.new_window = InputFields()
        self.new_window.change_db.connect(self.change_db_handler)
        self.new_window.show()

        print("Test")

    def load_values_in_combo(self, combo, type_load="db"):
        if type_load == "db":
            with open("db", "r") as file:
                combo.clear()
                for value in file.read().split("\n"):
                    combo.addItem(value)

            return combo

        combo.addItem("производственная")
        combo.addItem("учебная")

        return combo

    def load_file(self):
        # Show a file dialog to choose the Excel file
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")

        # If a file was chosen, load its data into the table
        if file_path:
            df = pd.read_excel(file_path)
            self.file_path = file_path

            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    value = str(df.iloc[i, j])
                    if "группы" in value.lower().strip():
                        self.group = value.split(" ")[-1]
                    if re.match(regex_fio, value) is not None:
                        self.people.append({
                            "name": value,
                            "dest": df.iloc[i, j + 1]
                        })

            self.table.setRowCount(len(self.people))
            self.table.setColumnCount(6)

            for i in range(len(self.people)):

                box = self.load_values_in_combo(QComboBox(self))
                self.combos.append(box)

                fio = self.people[i]['name'].split(" ")

                for j in range(len(fio)):
                    item = QTableWidgetItem(fio[j])
                    self.table.setItem(i, j, item)

                item = QTableWidgetItem(self.people[i]['dest'])
                self.table.setItem(i, 3, item)
                self.table.setCellWidget(i, 4, box)

                box = self.load_values_in_combo(QComboBox(self), type_load="practice")
                self.combos_type_practise.append(box)
                self.table.setCellWidget(i, 5, box)

            self.table.resizeColumnsToContents()
            self.table.resizeRowsToContents()

            self.input2.setText(self.group)

    def generate_docs(self):
        people = {
            "budget_people": [],
            "target_people": [],
            "paid_people": []
        }
        data = {
            "people": []
        }
        for i in range(len(self.people)):
            fio = self.table.item(i, 0).text() + " " + self.table.item(i, 1).text() + " " + self.table.item(i, 2).text()
            dest = self.table.item(i, 3).text()
            practice_name = self.combos[i].currentText()
            practice_type = self.combos_type_practise[i].currentText()

            if dest.lower().strip() == "бюджет":
                selector = "budget_people"
            elif dest.lower().strip() == "платное":
                selector = "paid_people"
            else:
                selector = "target_people"

            if practice_type.lower().strip() == "учебная":
                practice_type = "учебная (тип:  практика по получению первичных профессиональных умений и навыков, в том числе  первичных умений и навыков научно-исследовательской деятельности, 3 з.е.)"
            else:
                practice_type = "производственная  (тип:  практика по получению профессиональных умений и опыта профессиональной деятельности, 5 з.е.)"

            date1 = self.date1.date().toPyDate()
            date2 = self.date2.date().toPyDate()

            person = {
                "fio": fio,
                "dest": dest,
                "practice_name": practice_name,
                "practice_type": practice_type,
                "practice_type_short": self.combos_type_practise[i].currentText(),
                "start_time": f"{date1.day}.{date1.month}.{date1.year}",
                "end_time": f"{date2.day}.{date2.month}.{date2.year}",
                "group": self.input2.text()
            }
            data['people'].append(person)
            people[selector].append(person)

            people["group"] = self.input2.text()
            data['group'] = self.input2.text()

        if len(self.people) == 0:
            return

        tpl = DocxTemplate("templates/report.docx")

        tpl.render(people)

        tpl.save("result/Отчет.docx")

        tpl = DocxTemplate("templates/dest.docx")

        tpl.render(data)

        tpl.save("result/Направление.docx")

        self.success_window = ExampleWindow()
        self.success_window.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelParser()
    ex.show()
    sys.exit(app.exec_())
