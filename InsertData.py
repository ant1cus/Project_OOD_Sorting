import datetime
import psutil
import json
import pathlib
import queue
import sys
import os

import Main
import logging

from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo, QDir
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from Default import DefaultWindow
from Check import check_data
from Insert import InsertTableData
from Report import GetReport


class MainWindow(QMainWindow, Main.Ui_MainWindow):  # Главное окно

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        self.queue = queue.Queue(maxsize=1)
        self.default_path = pathlib.Path.cwd()
        filename = str(datetime.date.today()) + '_logs.log'
        os.makedirs(pathlib.Path('logs'), exist_ok=True)
        filemode = 'a' if pathlib.Path('logs', filename).is_file() else 'w'
        logging.basicConfig(filename=pathlib.Path('logs', filename),
                            level=logging.DEBUG,
                            filemode=filemode,
                            format="%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
        if pathlib.Path.exists(pathlib.Path(pathlib.Path.cwd(), 'report')) is False:
            os.mkdir(pathlib.Path(pathlib.Path.cwd(), 'report'))
        self.pushButton_open_data_file.clicked.connect((lambda: self.browse(self.lineEdit_path_data_file)))
        self.pushButton_open_start_folder.clicked.connect((lambda: self.browse(self.lineEdit_path_start_folder)))
        self.pushButton_open_finish_folder.clicked.connect((lambda: self.browse(self.lineEdit_path_finish_folder)))
        self.pushButton_start.clicked.connect(self.insert_data)
        self.pushButton_open_report.clicked.connect(self.report)
        self.action_settings_default.triggered.connect(self.default_settings)
        self.list = {'path-path_data_file': ['Путь к файлу выгрузки', self.lineEdit_path_data_file],
                     'path-path_start_folder': ['Путь к документам', self.lineEdit_path_start_folder],
                     'path-path_finish_folder': ['Путь к конечной папке', self.lineEdit_path_finish_folder],
                     'path-font_size': ['Размер шрифта', self.lineEdit_size]
                     }
        try:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                dict_load = json.load(f)
                self.data = dict_load['widget_settings']
        except FileNotFoundError:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "w", encoding='utf-8-sig') as f:
                data_insert = {"widget_settings": {}}
                json.dump(data_insert, f, ensure_ascii=False, sort_keys=True, indent=4)
                self.data = {}
        self.default_data(self.data)

    def browse(self, line_edit):  # Для кнопки открыть
        if 'folder' in self.sender().objectName():
            directory = QFileDialog.getExistingDirectory(self, "Открыть папку", QDir.currentPath())
        else:
            directory = QFileDialog.getOpenFileName(self, "Открыть файл", QDir.currentPath())
        if directory and isinstance(directory, tuple):
            if directory[0]:
                line_edit.setText(directory[0])
        elif directory and isinstance(directory, str):
            line_edit.setText(directory)

    def report(self):
        sending_data = {'logging': logging, 'queue': self.queue, 'default_path': self.default_path}
        self.thread = GetReport(sending_data)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def insert_data(self):
        sending_data = check_data(self.lineEdit_path_data_file, self.lineEdit_path_start_folder,
                                  self.lineEdit_path_finish_folder, self.lineEdit_size)
        if isinstance(sending_data, list):
            self.on_message_changed(sending_data[0], sending_data[1])
            return
        # Если всё прошло запускаем поток
        sending_data['logging'], sending_data['queue'] = logging, self.queue
        sending_data['default_path'] = self.default_path
        self.thread = InsertTableData(sending_data)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def errors(self):
        text = self.queue.get_nowait()
        self.on_message_changed('Внимание!', 'Ошибки в загруженных данных:\n' + '\n'.join(text['errors']))

    def on_message_changed(self, title, description):
        if title == 'УПС!':
            QMessageBox.critical(self, title, description)
        elif title == 'Внимание!':
            QMessageBox.warning(self, title, description)
        elif title == 'Вопрос?':
            self.statusBar().clearMessage()
            ans = QMessageBox.question(self, title, description,
                                       QMessageBox.Yes | QMessageBox.No,
                                       QMessageBox.Yes)
            if ans == QMessageBox.Yes:
                self.thread.queue.put(True)
            else:
                self.thread.queue.put(False)
            self.thread.event.set()
        elif title == 'Пауза':
            self.statusBar().clearMessage()
            ans = QMessageBox.question(self, title, description, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ans == QMessageBox.No:
                self.thread.queue.put(True)
            else:
                self.thread.queue.put(False)
            self.thread.event.set()

    def default_settings(self):  # Запускаем окно с настройками по умолчанию.
        self.close()
        window_add = DefaultWindow(self, self.default_path, self.list)
        window_add.show()

    def default_data(self, incoming_data):
        for element in self.list:
            if element in incoming_data:
                if 'checkBox' in element or 'groupBox' in element:
                    self.list[element][1].setChecked(True) if incoming_data[element] \
                        else self.list[element][1].setChecked(False)
                elif 'radioButton' in element:
                    for radio, button in zip(incoming_data[element], self.list[element][1]):
                        if radio:
                            button.setChecked(True)
                        else:
                            button.setAutoExclusive(False)
                            button.setChecked(False)
                        button.setAutoExclusive(True)
                else:  # Если любой другой элемент
                    self.list[element][1].setText(incoming_data[element])  # Помещаем значение

    def show_mess(self, value):  # Вывод значения в статус бар
        self.statusBar().showMessage(value)

    def closeEvent(self, event):
        if os.listdir(pathlib.Path(self.default_path, 'report')):
            for file in os.listdir(pathlib.Path(self.default_path, 'report')):
                try:
                    if '~' not in file:
                        os.remove(pathlib.Path(self.default_path, 'report', file))
                except PermissionError:
                    self.on_message_changed('Внимание!', f'Сохраните у себя или закройте временный файл {file}'
                                                         f' с отчётом, документ будет удалён при закрытии')
                    event.ignore()
                    return
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    translator = QTranslator(app)
    locale = QLocale.system().name()
    path = QLibraryInfo.location(QLibraryInfo.TranslationsPath)
    translator.load('qtbase_%s' % locale.partition('_')[0], path)
    app.installTranslator(translator)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
