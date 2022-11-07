from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtWidgets import QMainWindow

import sys # Только для доступа к аргументам командной строки
# https://habr.com/ru/post/456534/
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('C:/Users/User/YandexDisk/Научная работа/Пары/Энергетика/Проектная деятельность/mainwindow.ui', self)
        self.pushButton.setText('Check text')

        self.pushButton.clicked.connect(self.the_button_was_clicked)  # Remember to pass the definition/method, not the return value!

    def the_button_was_clicked(self):
        print("Clicked!")
        print(self.lineEdit.text())

app = QApplication(sys.argv)

# Создаём виджет Qt — окно.
window = MainWindow()
window.show()  # Важно: окно по умолчанию скрыто.

# Запускаем цикл событий.
app.exec()


#w1=MainWindow



print("Check")