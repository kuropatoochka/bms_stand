import sys
import threading
import random
import time
import json
import hashlib
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QLineEdit, QMessageBox, QDialog, QFormLayout, QInputDialog,
    QTabWidget, QHeaderView, QListWidget, QListWidgetItem, QDateEdit
)
from PySide6.QtCore import QTimer, Qt, Signal, QObject, QDate
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
import logging
import hashlib
from shutil import copyfile

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

font_path = "C:/Windows/Fonts/times.ttf"
pdfmetrics.registerFont(TTFont("TimesNewRoman", font_path))

logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def log_event(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("log.txt", "a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")

class SerialEmulator(QObject):
    ping_response = Signal(bool)
    test_started = Signal(str, float)

    def __init__(self):
        super().__init__()
        self.running = False
        self.enabled = True

    def start_emulation(self):
        QTimer.singleShot(1000, lambda: self.ping_response.emit(True))
        self.running = True
        threading.Thread(target=self.listen_for_start, daemon=True).start()

    def listen_for_start(self):
        while self.running:
            if self.enabled:
                time.sleep(random.randint(3, 10))
                timestamp = datetime.now().strftime("%H:%M:%S")
                duration = round(random.uniform(0.1, 1.0), 3)
                self.test_started.emit(timestamp, duration)
            else:
                time.sleep(1)

class LoginDialog(QDialog):
    def __init__(self, users):
        super().__init__()
        self.setWindowTitle("Авторизация")
        self.setModal(True)

        self.users = users
        self.selected_user = None
        self.login_attempts = 0
        self.max_attempts = 3

        layout = QFormLayout()

        self.user_selector = QComboBox()
        self.update_user_list()
        layout.addRow("Выберите пользователя:", self.user_selector)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addRow("Пароль:", self.password_input)

        self.login_button = QPushButton("Войти")
        self.login_button.clicked.connect(self.try_login)
        layout.addWidget(self.login_button)

        self.setLayout(layout)

    def update_user_list(self):
        self.user_selector.clear()
        self.user_selector.addItems(self.users.keys())
    
    def try_login(self):
        username = self.user_selector.currentText()
        password = self.password_input.text()
        hashed_input = hash_password(password)
        if self.users.get(username, {}).get("password") == hashed_input:
            self.selected_user = username
            super().accept()
        else:
            self.login_attempts += 1
            if self.login_attempts >= self.max_attempts:
                QMessageBox.critical(self, "Ошибка", "Превышено количество попыток входа. Программа будет закрыта.")
                sys.exit(1)
            else:
                QMessageBox.warning(self, "Ошибка", f"Неверный пароль. Осталось попыток: {self.max_attempts - self.login_attempts}")

class AddUserDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить пользователя")
        self.setMinimumWidth(350)

        # Поля ввода
        self.user_id = QLineEdit()
        self.lastname = QLineEdit()
        self.firstname = QLineEdit()
        self.middlename = QLineEdit()
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.confirm_password = QLineEdit()
        self.confirm_password.setEchoMode(QLineEdit.Password)

        layout = QVBoxLayout()

        layout.addLayout(self._labeled("Идентификатор пользователя", self.user_id))
        layout.addLayout(self._labeled("Фамилия", self.lastname))
        layout.addLayout(self._labeled("Имя", self.firstname))
        layout.addLayout(self._labeled("Отчество", self.middlename))
        layout.addLayout(self._labeled("Пароль", self.password))
        layout.addLayout(self._labeled("Повтор пароля", self.confirm_password))

        # Кнопки
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        cancel_btn = QPushButton("Отмена")
        save_btn.clicked.connect(self.validate_and_accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _labeled(self, text, widget):
        layout = QHBoxLayout()
        layout.addWidget(QLabel(text))
        layout.addWidget(widget)
        return layout

    def validate_and_accept(self):
        uid = self.user_id.text().strip()
        lastname = self.lastname.text().strip()
        firstname = self.firstname.text().strip()
        password = self.password.text()
        confirm = self.confirm_password.text()

        if not uid or len(uid) < 3:
            QMessageBox.warning(self, "Ошибка", "Идентификатор должен содержать минимум 3 символа.")
            return
        if not lastname or not firstname:
            QMessageBox.warning(self, "Ошибка", "Фамилия и Имя обязательны.")
            return
        if not password:
            QMessageBox.warning(self, "Ошибка", "Пароль обязателен.")
            return
        if password != confirm:
            QMessageBox.warning(self, "Ошибка", "Пароли не совпадают.")
            return

        self.accept()

    def get_user_data(self):
        return {
            "user_id": self.user_id.text().strip(),
            "lastname": self.lastname.text().strip(),
            "firstname": self.firstname.text().strip(),
            "middlename": self.middlename.text().strip(),
            "password": self.password.text(),
        }

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Стенд электрических испытаний СКУ ЛИАБ")
        self.resize(700, 500)

        self.users = self.load_users()
        if not self.users:
            self.users = {"Default": {"password": hash_password("admin"), "role": "admin"}}
            self.save_users()

        self.current_user = ""

        self.serial = SerialEmulator()
        self.serial.ping_response.connect(self.on_device_connected)
        self.serial.test_started.connect(self.on_test_received)

        self.device_connected = False
        self.reports = []

        self.results_received = False

        self.setup_ui()
        self.serial.start_emulation()
        self.show_login_dialog()

    def setup_ui(self):
        self.user_label = QLabel("Пользователь: не выбран")
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.main_tab = QWidget()
        self.settings_tab = QWidget()
        self.reports_tab = QWidget()

        self.tabs.addTab(self.main_tab, "Испытания")
        self.tabs.addTab(self.settings_tab, "Настройки")
        self.tabs.addTab(self.reports_tab, "Отчетность")

        self.tabs.setTabPosition(QTabWidget.North)

        self.init_main_tab()
        self.init_settings_tab()
        self.init_reports_tab()

    def init_main_tab(self):
        layout = QVBoxLayout(self.main_tab)

        self.status_label = QLabel("Ожидание подключения устройства...")
        layout.addWidget(self.status_label)

        self.detailed_table = QTableWidget(8, 4)
        self.detailed_table.setHorizontalHeaderLabels([
            "Работа\nпри напряжении\nниже 4,2 В",
            "Отключение\nпри напряжении\nвыше 4,3 В",
            "Работа\nпри напряжении\nвыше 2,9 В",
            "Отключение\nпри напряжении\nниже 2,8 В"
        ])
        self.detailed_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.detailed_table.horizontalHeader().setFixedHeight(60)

        for i in range(8):
            self.detailed_table.setVerticalHeaderItem(i, QTableWidgetItem(f"Канал {i+1}"))
        self.detailed_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.detailed_table.setStyleSheet("QTableWidget::item { padding: 2px; }")

        layout.addWidget(QLabel("Результаты испытаний по каналам:"))
        layout.addWidget(self.detailed_table)

        self.short_circuit_label = QLabel("Результат по КЗ: ...")
        layout.addWidget(self.short_circuit_label)

        control_layout = QHBoxLayout()

        self.reset_button = QPushButton("Сбросить результаты")
        self.reset_button.clicked.connect(self.confirm_reset)
        control_layout.addWidget(self.reset_button)

        self.report_button = QPushButton("Сохранить результаты")
        self.report_button.clicked.connect(self.save_report_as_pdf)
        self.report_button.setEnabled(False)
        control_layout.addWidget(self.report_button)

        layout.addLayout(control_layout)

        self.device_status = QLabel("Устройство не подключено")
        layout.addWidget(self.device_status)

    def init_settings_tab(self):
        self.setWindowTitle("Управление пользователями")
        layout = QVBoxLayout(self.settings_tab)

        layout.addWidget(QLabel("Список пользователей:"))

        self.user_list = QListWidget()
        for uid, info in self.users.items():
            full_display = f"{uid} — {info.get('lastname', '')} {info.get('firstname', '')}"
            self.user_list.addItem(full_display)
        layout.addWidget(self.user_list)

        self.add_user_button = QPushButton("Добавить пользователя")
        self.add_user_button.clicked.connect(self.add_user)
        layout.addWidget(self.add_user_button)

        self.delete_user_button = QPushButton("Удалить пользователя")
        self.delete_user_button.clicked.connect(self.delete_user)
        layout.addWidget(self.delete_user_button)

        layout.addStretch()

        layout.addSpacing(20)
        layout.addWidget(QLabel("<b>Испытательный участок</b>"))

        test_area_layout = QHBoxLayout()
        self.test_area_label = QLabel(f"Название: {self.test_area_name}")
        self.rename_area_button = QPushButton("Изменить название")
        self.rename_area_button.clicked.connect(self.rename_test_area)

        test_area_layout.addWidget(self.test_area_label)
        test_area_layout.addWidget(self.rename_area_button)

        layout.addLayout(test_area_layout)
    
    def rename_test_area(self):
        new_name, ok = QInputDialog.getText(self, "Изменить название участка", "Введите новое название:")
        if ok and new_name.strip():
            self.test_area_name = new_name.strip()
            self.test_area_label.setText(f"Название: {self.test_area_name}")

    def init_reports_tab(self):
        layout = QVBoxLayout(self.reports_tab)

        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по заводскому номеру")
        filter_layout.addWidget(self.search_input)

        self.date_filter_edit = QDateEdit()
        self.date_filter_edit.setCalendarPopup(True)
        self.date_filter_edit.setDate(QDate.currentDate())
        self.date_filter_edit.dateChanged.connect(self.update_date_filter)
        filter_layout.addWidget(self.date_filter_edit)

        self.clear_filters_button = QPushButton("Сбросить фильтры")
        self.clear_filters_button.clicked.connect(self.clear_filters)
        filter_layout.addWidget(self.clear_filters_button)

        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.update_report_list)
        filter_layout.addWidget(self.search_button)

        layout.addLayout(filter_layout)

        self.report_list = QListWidget()
        layout.addWidget(QLabel("Сохраненные отчеты:"))
        layout.addWidget(self.report_list)

        self.update_report_list()

    def update_date_filter(self):
        self.selected_date = self.date_filter_edit.date().toString("yyyyMMdd")

    def clear_filters(self):
        self.search_input.clear()
        self.selected_date = ""
        self.date_filter_edit.setDate(QDate.currentDate())
        self.update_report_list()

    def update_report_list(self):
        self.report_list.clear()
        report_dir = "reports"
        os.makedirs(report_dir, exist_ok=True)
        serial_filter = self.search_input.text().strip().lower()
        date_filter = getattr(self, "selected_date", "")

        for file in sorted(os.listdir(report_dir)):
            if file.endswith(".pdf"):
                match_serial = serial_filter in file.lower() if serial_filter else True
                match_date = date_filter in file if date_filter else True
                if match_serial and match_date:
                    self.report_list.addItem(QListWidgetItem(file))

    def load_users(self):
        if os.path.exists("users.json"):
            with open("users.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                self.test_area_name = data.get("test_area_name", "Участок 1")
                return data.get("users", {})
        self.test_area_name = "Участок 1"
        return {}


    def save_users(self):
        data = {
            "users": self.users,
            "test_area_name": self.test_area_name
        }
        with open("users.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


    def show_login_dialog(self):
        self.users = self.load_users()
        login = LoginDialog(self.users)
        if login.exec() == QDialog.Accepted:
            self.current_user = login.selected_user
            self.user_label.setText(f"Пользователь: {self.current_user}")

            for i in range(self.user_list.count()):
                item = self.user_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == self.current_user:
                    self.user_list.setCurrentItem(item)
                    break

            is_admin = self.users[self.current_user]["role"] == "admin"
            if not is_admin:
                if self.tabs.indexOf(self.settings_tab) != -1:
                    self.tabs.removeTab(self.tabs.indexOf(self.settings_tab))
            elif self.tabs.indexOf(self.settings_tab) == -1:
                self.tabs.insertTab(1, self.settings_tab, "Настройки")

    def change_user(self, name):
        self.current_user = name
        self.user_label.setText(f"Пользователь: {name}")
        is_admin = self.users[name]["role"] == "admin"

        if not is_admin and self.tabs.indexOf(self.settings_tab) != -1:
            self.tabs.removeTab(self.tabs.indexOf(self.settings_tab))
        elif is_admin and self.tabs.indexOf(self.settings_tab) == -1:
            self.tabs.insertTab(1, self.settings_tab, "Настройки")

    def log_event(self, message):
        logging.info(message)

    def log_user_action(self, action):
        self.log_event(f"Пользователь {self.current_user} {action}")

    def add_user(self):
        dialog = AddUserDialog(self)
        if dialog.exec():
            user_data = dialog.get_user_data()
            user_id = user_data["user_id"]

            if user_id in self.users:
                QMessageBox.information(self, "Информация", "Пользователь уже существует.")
                return

            self.users[user_id] = {
                "lastname": user_data["lastname"],
                "firstname": user_data["firstname"],
                "middlename": user_data["middlename"],
                "password": hash_password(user_data["password"]),
                "role": "operator"
            }

            display_name = f"{user_id} — {user_data['lastname']} {user_data['firstname']}"
            self.user_list.addItem(display_name)
            self.log_user_action(f"добавил пользователя {user_id}")
            self.save_users()

    def delete_user(self):
        item = self.user_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Ошибка", "Выберите пользователя для удаления.")
            return

        display_text = item.text()
        uid = display_text.split(" — ")[0]

        if uid == "Default":
            QMessageBox.warning(self, "Ошибка", "Нельзя удалить пользователя по умолчанию.")
            return

        confirm = QMessageBox.question(self, "Подтверждение", f"Удалить пользователя {uid}?", QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            del self.users[uid]
            row = self.user_list.row(item)
            self.user_list.takeItem(row)
            self.log_user_action(f"удалил пользователя {uid}")
            self.save_users()
            QMessageBox.information(self, "Удалено", f"Пользователь {uid} удален.")

    def confirm_reset(self):
        if QMessageBox.question(self, "Подтверждение", "Вы уверены, что хотите сбросить результаты?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for i in range(8):
                for j in range(4):
                    self.detailed_table.setItem(i, j, QTableWidgetItem(""))
            self.short_circuit_label.setText("Результат по КЗ: ...")
            self.results_received = False
            self.serial.enabled = True
            self.status_label.setText("Ожидание результатов испытаний...")

    def on_device_connected(self, success):
        if success:
            self.device_connected = True
            self.status_label.setText("Ожидание результатов испытаний...")
            self.device_status.setText("Устройство подключено")
            self.report_button.setEnabled(True)

    def on_test_received(self, timestamp, duration):
        if not self.device_connected or self.results_received:
            return
        for i in range(8):
            for j in range(4):
                val = random.choice(["+", "-"])
                item = QTableWidgetItem(val)
                font = item.font()
                font.setPointSize(16)
                item.setFont(font)
                if val == "-":
                    item.setForeground(Qt.red)
                item.setTextAlignment(Qt.AlignCenter)
                self.detailed_table.setItem(i, j, item)

        result = random.choice([
            "СКУ ЛИАБ сработало по короткому замыканию",
            "СКУ ЛИАБ не сработало по короткому замыканию",
            "Порог по КЗ не достигнут"
        ])
        self.short_circuit_label.setText("Результат по КЗ: " + result)
        self.results_received = True
        self.serial.enabled = False
        self.status_label.setText("Результаты получены")

    def save_report_as_pdf(self):
        from reportlab.lib.utils import simpleSplit
        system_name, ok1 = QInputDialog.getText(self, "Название системы контроля", "Введите название системы контроля:")
        if not ok1 or not system_name.strip():
            QMessageBox.warning(self, "Ошибка", "Название системы контроля обязательно.")
            return

        # Запрос заводского номера
        serial_number, ok2 = QInputDialog.getText(self, "Заводской номер", "Введите заводской номер:")
        if not ok2 or not serial_number.strip():
            QMessageBox.warning(self, "Ошибка", "Заводской номер обязателен.")
            return
        
        system_name = system_name.strip().replace(" ", "_")
        serial_number = serial_number.strip().replace(" ", "_")

        os.makedirs("reports", exist_ok=True)
        filename = f"reports/report_{system_name}_{serial_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        # Регистрация шрифтов
        pdfmetrics.registerFont(TTFont("TimesNewRoman", "C:/Windows/Fonts/times.ttf"))
        pdfmetrics.registerFont(TTFont("TimesNewRoman-Bold", "C:/Windows/Fonts/timesbd.ttf"))

        c = canvas.Canvas(filename, pagesize=A4)
        width, height = A4
        min_y_margin = 50

        def ensure_y_space(c, y, decrement=20, font="TimesNewRoman", size=12):
            if y - decrement < min_y_margin:
                c.showPage()
                c.setFont(font, size)
                return height - 50
            return y

        # Заголовки
        c.setFont("TimesNewRoman-Bold", 12)
        y = height - 140
        c.drawCentredString(width / 2, height - 50, "ПРОТОКОЛ")
        c.setFont("TimesNewRoman", 12)
        c.drawCentredString(width / 2, height - 70, "Проверки соответствия системы контроля литий-ионной аккумуляторной батареи")
        c.drawCentredString(width / 2, height - 85, "функциональным требованиям")

        c.drawString(50, height - 110, "№ 1246")
        date_str = datetime.now().strftime("«%d» %B %Y г.")
        c.drawRightString(width - 50, height - 110, date_str)


        bms_model = self.bms_model if hasattr(self, "bms_model") else "BMS_ABC123"
        test_area = self.test_area_name if hasattr(self, "test_area_name") else "испытательный участок ООО «__________»"
        date_str = datetime.now().strftime("«%d» %B %Y г.")

        # Раздел 1
        c.drawString(50, y, "1.  Объект испытания: система контроля литий-ионной аккумуляторной батареи")
        y = ensure_y_space(c, y, 15)
        y -= 15
        c.setFont("TimesNewRoman-Bold", 12)
        c.drawString(70, y, bms_model)
        c.setFont("TimesNewRoman", 12)
        c.drawString(150, y, f"зав. № {serial_number.strip()}.")
        y = ensure_y_space(c, y, 25)
        y -= 25

        # Раздел 2
        c.drawString(50, y, "2.  Цель испытания:")
        y = ensure_y_space(c, y, 15)
        y -= 15
        lines = [
            "Проверка соответствия системы контроля литий-ионной аккумуляторной батареи",
            "функциональным требованиям",
            "по защите аккумуляторной батареи от перезаряда, переразряда, токов короткого замыкания.",
            "- отключение тока заряда при напряжении 4,25±0,05 В на любом из аккумуляторов;",
            "- отключение тока разряда при напряжении 2,85±0,05 В на любом из аккумуляторов;",
            "- отключение при превышении тока разряда свыше 50 А."
        ]
        for line in lines:
            y = ensure_y_space(c, y, 15)
            c.drawString(70, y, line)
            y -= 15
        y = ensure_y_space(c, y, 10)
        y -= 10

        # Раздел 3
        c.drawString(50, y, f"3.  Дата проведения испытания: {date_str}")
        y = ensure_y_space(c, y, 20)
        y -= 20

        # Раздел 4
        c.drawString(50, y, f"4.  Место проведения испытания: {test_area}.")
        y = ensure_y_space(c, y, 30)
        y -= 30

        # Раздел 5: Результаты испытания
        y = ensure_y_space(c, y, 20)
        c.drawString(50, y, "5.  Результаты испытания:")
        y -= 20

        y, has_negative_result = self.draw_results_table(c, y)

        # Проверка отключения по превышению тока
        sc_text = self.short_circuit_label.text().lower()
        if "сработало по короткому замыканию" in sc_text:
            discharge_status = "выполнено"
        else:
            discharge_status = "не выполнено"
            has_negative_result = True  # если не сработало — это тоже негативный результат

        y -= 5
        y = ensure_y_space(c, y, 20)
        c.drawString(50, y, f"Отключение разряда по превышению тока 50 А: {discharge_status}")
        y -= 30

        # Раздел 6: Заключение
        text_width_limit = width - 100  # 50 отступ слева и справа

        conclusion_text = (
            f"6. Заключение\n"
            f"Система контроля литий-ионной аккумуляторной батареи {system_name} "
            f"зав. № {serial_number} прошла проверку на соответствие функциональным требованиям по "
            f"защите аккумуляторной батареи от перезаряда, переразряда, токов короткого замыкания "
            f"с {'отрицательным' if has_negative_result else 'положительным'} результатом и "
            f"{'не ' if has_negative_result else ''}пригодна к использованию по назначению."
        )

        lines = simpleSplit(conclusion_text, "TimesNewRoman", 12, text_width_limit)

        c.setFont("TimesNewRoman", 12)

        for line in lines:
            if y < min_y_margin + 20:  
                c.showPage()
                c.setFont("TimesNewRoman", 12)
                y = height - 50
            c.drawString(50, y, line)
            y -= 15 

        # Получение ФИО
        user_info = self.users.get(self.current_user, {})
        lastname = user_info.get("lastname", "")
        firstname = user_info.get("firstname", "")
        middlename = user_info.get("middlename", "")
        fio_short = f"{lastname} {firstname[:1]}.{middlename[:1]}."

        # Подпись
        if y < min_y_margin + 60:
            c.showPage()
            c.setFont("TimesNewRoman", 12)
            y = height - 50

        y -= 20
        c.drawString(50, y, "Испытание проводил:")

        y -= 20
        c.drawString(50, y, "Инженер:")
        c.drawRightString(width - 50, y, fio_short)

        y -= 20
        c.drawString(50, y, "Контролер ОТК:")
        c.drawRightString(width - 50, y, "Финогенова Е.С.")

        c.save()

        # Хеш-файл
        file_hash = self.calculate_file_hash(filename)
        log_event(f"Report: {filename}, hash: {file_hash}")

        self.update_report_list()
        self.confirm_reset()

    def draw_results_table(self, c, start_y):
        from reportlab.platypus import Table, TableStyle
        from reportlab.lib import colors
        from reportlab.lib.units import mm

        width, height = A4

        data = [
            ["№ канала",
                "Работа при\nнапряжении\nниже 4,2 В",
                "Отключение при\nнапряжении\nвыше 4,3 В",
                "Работа при\nнапряжении\nвыше 2,9 В",
                "Отключение при\nнапряжении\nниже 2,8 В"]
        ]

        has_negative_result = False

        for i in range(self.detailed_table.rowCount()):
            row = [str(i + 1)]
            for j in range(self.detailed_table.columnCount()):
                item = self.detailed_table.item(i, j)
                value = item.text() if item else ""
                if value == "-":
                    has_negative_result = True
                row.append(value)
            data.append(row)

        table = Table(data, colWidths=[20*mm, 40*mm, 40*mm, 40*mm, 40*mm])
        table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'TimesNewRoman', 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ]))

        table.wrapOn(c, 50, start_y)
        table.drawOn(c, 50, start_y - table._height)

        return start_y - table._height - 40, has_negative_result

    def calculate_file_hash(self, filepath):
        sha256 = hashlib.sha256()
        with open(filepath, "rb") as f:
            while chunk := f.read(8192):
                sha256.update(chunk)
        return sha256.hexdigest()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())