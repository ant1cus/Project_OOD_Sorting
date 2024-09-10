import pathlib

from PyQt5.QtCore import QSize, QEvent
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtWidgets import QWidget, QButtonGroup, QPushButton, QSizePolicy, QLineEdit, QMessageBox

import exception


class ExceptionWindow(QWidget, exception.Ui_Form):

    def __init__(self, parent, exception_gadget):
        super().__init__()
        self.setupUi(self)
        self.parent = parent
        self.exception_gadget = exception_gadget
        self.default_path = pathlib.Path.cwd()  # Путь для файла настроек
        self.numAddWidget = 0
        self.device_group = []
        self.gadget = {}
        self.button = {}
        self.buttongroup_del = QButtonGroup()
        # Коннекты
        self.buttongroup_del.buttonClicked[int].connect(self.del_button_clicked)
        self.pushButton_add.clicked.connect(self.add_widget)
        self.pushButton_save.clicked.connect(self.save_exception)
        for key in self.exception_gadget:
            self.numAddWidget += 1
            self.add_widget(self.exception_gadget[key])

    def add_widget(self, gadget=False):
        self.device_group.insert(self.numAddWidget, 1)  # Добавляем в список
        self.gadget[self.numAddWidget] = QLineEdit()  # Создаем лайн для имени
        if gadget:
            self.gadget[self.numAddWidget].setText(gadget)
        self.gadget[self.numAddWidget].setPlaceholderText('Напишите название техники для исключения')
        self.gadget[self.numAddWidget].setFont(QFont("Times", 10, QFont.Light))  # Размер шрифта
        self.gadget[self.numAddWidget].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
        self.gadget[self.numAddWidget].setFixedSize(16777215, 35)  # Размеры вручную
        self.gridLayout_gadget_name.addWidget(self.gadget[self.numAddWidget], self.numAddWidget, 0)  # Добавляем
        self.button[self.numAddWidget] = QPushButton()  # Создаем кнопку с подписью удалить
        self.button[self.numAddWidget].installEventFilter(self)
        self.button[self.numAddWidget].setToolTip('Удалить элемент')
        self.button[self.numAddWidget].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
        self.button[self.numAddWidget].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
        self.button[self.numAddWidget].setStyleSheet('''QPushButton {
                                                        border: none;
                                                        margin: 0px;
                                                        padding: 0px;
                                                    }''')
        self.button[self.numAddWidget].setIcon(QIcon(str(pathlib.Path(self.default_path, 'icons', 'del_button.png'))))
        self.button[self.numAddWidget].setIconSize(QSize(30, 30))
        self.button[self.numAddWidget].setFixedSize(35, 35)  # Размеры вручную
        self.buttongroup_del.addButton(self.button[self.numAddWidget], self.numAddWidget)  # Добавляем в группу
        self.gridLayout_del_gadget.addWidget(self.button[self.numAddWidget], self.numAddWidget, 0)  # Добавляем в фрейм
        self.numAddWidget += 1

    def del_button_clicked(self, number):  # Если нажмем кнопку удалить
        self.button[number].removeEventFilter(self)
        self.gadget[number].deleteLater()
        self.button[number].deleteLater()
        del self.gadget[number]
        del self.button[number]

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Enter:
            obj.setIconSize(QSize(35, 35))  # Размеры вручную
        elif event.type() == QEvent.Leave:
            obj.setIconSize(QSize(30, 30))  # Размеры вручную
        return QWidget.eventFilter(self, obj, event)

    def save_exception(self):
        self.exception_gadget.clear()
        for enum, line in enumerate(self.gadget):
            if self.gadget[line].text():
                self.exception_gadget[enum] = self.gadget[line].text()
        self.parent.rewrite_exception(self.exception_gadget)
        QMessageBox.warning(self, 'Сохранение исключений', 'Текущая конфигурация настроек исключений была сохранена!')

    def closeEvent(self, event):
        self.close()
