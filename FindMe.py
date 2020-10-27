import os
import psutil
import PyQt5.QtWidgets as Qtw
from PyQt5.QtGui import QCursor
import PyQt5.QtCore as QtCore
import docx
from functools import partial


class MainWindow(Qtw.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Find Me')
        self.setLayout(Qtw.QVBoxLayout())
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Find Me')
        self.setFixedSize(480, 480)
        self.show()

        container = Qtw.QWidget()
        container.setLayout(Qtw.QGridLayout())

        # Combo for disk selection
        label1 = Qtw.QLabel()
        label1.setText("Select a disk to search through:")
        combo = Qtw.QComboBox()
        disks = self.get_disks()
        for d in disks:
            combo.addItem(d)

        # Textbox for string input
        label2 = Qtw.QLabel()
        label2.setText("Enter the UPI to look for:")
        text_box = Qtw.QLineEdit()

        # Text area for found files
        label3 = Qtw.QLabel()
        label3.setText("Found files:")
        self.list_of_files = Qtw.QListWidget()

        # Button to do the search
        button = Qtw.QPushButton('Find')
        button.clicked.connect(partial(self.get_list_of_files, text_box, combo))

        # Adding elements to container
        container.layout().addWidget(label1, 0, 0, 1, 10)
        container.layout().addWidget(combo, 1, 2, 1, 8)
        container.layout().addWidget(label2, 2, 0, 1, 10)
        container.layout().addWidget(text_box, 3, 2, 1, 8)
        container.layout().addWidget(button, 4, 4, 1, 2)

        # Adding elements to window
        self.layout().addWidget(container)
        self.layout().addWidget(label3)
        self.layout().addWidget(self.list_of_files)

    def get_list_of_files(self, textbox, combo):  # Finds files and displays them.
        self.list_of_files.clear()
        self.setCursor(QCursor(QtCore.Qt.WaitCursor))
        upi = str(textbox.text())  # Text to look for
        selected_disk = combo.currentText()  # Chosen disk to search through

        os.chdir(selected_disk)  # working directory changed to selected disk

        list_doc_of_files = []  # will hold doc files

        for dirpath, subdirs, files in os.walk(selected_disk):  # loop through the files find doc files and store them
            for x in files:
                if x.endswith(".docx"):
                    list_doc_of_files.append(os.path.join(dirpath, x))

        found_files = []  # will hold files containing the string

        for file in list_doc_of_files:
            if "node_modules" in file:
                continue
            doc = docx.Document(file)
            for i in doc.paragraphs:
                if upi in i.text:
                    found_files.append(file)
                    break

        if len(found_files) == 0:  # No found files
            self.popup()
        self.list_of_files.addItems(found_files)
        self.setCursor(QCursor(QtCore.Qt.CustomCursor))

    def get_disks(self):  # gets all the disk partitions of the system
        partitions = psutil.disk_partitions()
        disks = []
        for p in partitions:
            disks.append(p.device)

        return disks

    def popup(self):  # shown when the search finishes with no file found
        msg = Qtw.QMessageBox()
        msg.setWindowTitle("Alert")
        msg.setText("No files found!")

        self.layout().addWidget(msg)


app = Qtw.QApplication([])
mw = MainWindow()
app.setStyle(Qtw.QStyleFactory.create('Fusion'))
app.exec_()
