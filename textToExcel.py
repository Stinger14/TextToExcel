from os import listdir
from os.path import isfile, join
import xlwt
import xlrd
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, \
    QInputDialog, QLineEdit, QHBoxLayout, QVBoxLayout, QMessageBox, QFileDialog
from PyQt5.QtGui import QIcon, QFont
from subprocess import call

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.title = "Text to Excel"
        self.top = 30
        self.left = 30
        self.width = 550
        self.height = 450
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.l1 = QLabel(self)
        self.l1.setText("Especifique la ruta al directorio: ")
        self.l1.setFont(QFont('Sans Serif', 12))

        self.b1 = QPushButton("Convertir")
        self.b2 = QPushButton("Salir")
        self.b3 = QPushButton("Ruta...")
        self.le = QLineEdit()
        self.fd = QLineEdit()
        self.fd = QFileDialog()

        h_box = QHBoxLayout()
        h_box.addStretch()
        h_box.addWidget(self.b1)
        h_box.addStretch()
        h_box.addWidget(self.b2)
        h_box.addStretch()

        v_box = QVBoxLayout()
        v_box.addWidget(self.l1)
        v_box.addStretch()
        v_box.addWidget(self.le)
        v_box.addStretch()
        v_box.addWidget(self.b3)
        v_box.addLayout(h_box)
        v_box.addStretch()

        v_box2 = QVBoxLayout()
        v_box2.addWidget(self.b3)
        v_box2.addStretch()

        self.setLayout(v_box)

        self.b1.clicked.connect(self.btn_clicked)
        self.b2.clicked.connect(self.btn_clicked)
        self.b3.clicked.connect(self.btn_clicked)

        self.show()

    def _openFile(self):
        path = str(QFileDialog.getExistingDirectory())
        self.le.setText('{}'.format(path))

    def btn_clicked(self):
        sender = self.sender()
        if sender.text() == 'Convertir' and self.le.text() != '':
            self.window(self.le.text())
            QMessageBox.about(self, "Notificación", "Conversión de archivos "
                                                    "completada")
            self.le.clear()

        elif sender.text() == 'Salir':
            sys.exit(0)

        elif sender.text() == 'Ruta...':
            self._openFile()



    def window(self, p):
        self.path = p
        textfiles = [ join(self.path,f) for f in listdir(self.path) \
                      if isfile(join(self.path,f)) and '.txt' in  f]

        def is_number(s):
            try:
                float(s)
                return True
            except ValueError:
                return False

        style = xlwt.XFStyle()
        style.num_format_str = '#,###0.00'

        try:
            for textfile in textfiles:
                with open(textfile, 'r+') as f:
                    row_list = []

                    for row in f:
                        # remove blank lines
                        if row.rstrip():
                            row_list.append(row.split('|'))

                    column_list = zip(*row_list)
                    workbook = xlwt.Workbook()
                    worksheet = workbook.add_sheet('Sheet1')
                    i = 0
                    for column in column_list:
                        for item in range(len(column)):
                            value = column[item].strip()
                            if is_number(value):
                                worksheet.write(item, i, float(value), style=style)
                            else:
                                worksheet.write(item, i, value)
                        i += 1
                    workbook.save(textfile.replace('.txt', '.xls'))
        except PermissionError:
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Window()
    sys.exit(app.exec_())