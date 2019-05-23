import sys
import os
import shutil
import datetime
from xlutils.copy import copy
from xlrd import *

from PyQt5.QtGui import QValidator
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

dic_month = {"Mois":"","Janvier":"01","Février":"02","Mars":"03","Avril":"04","Mai":"05","Juin":"06","Juillet":"07","Aout":"08","Septembre":"09","Octobre":"10","Novembre":"11","Décembre":"12"}

class Widget(QWidget):
    def __init__(self, *args, **kwargs):
        QWidget.__init__(self, *args, **kwargs)
        vbox = QVBoxLayout(self)
        hlay = QHBoxLayout(self)
        hlay2 = QHBoxLayout(self)
        self.lineEdit = QLineEdit()
        self.lineEdit.setPlaceholderText("Rechercher")
        self.cb_month = QComboBox()
        self.cb_month.addItems(["Mois","Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Aout","Septembre","Octobre","Novembre","Décembre"])
        self.cb_year = QComboBox()
        date = datetime.datetime.now()
        self.cb_year.addItems(["Année"]+[str(i) for i in range(2007,date.year+1)])
        self.btn1 = QPushButton("Créer un chantier")

        vbox.addWidget(self.btn1)
        hlay2.addWidget(self.lineEdit)
        hlay2.addWidget(self.cb_month)
        hlay2.addWidget(self.cb_year)
        vbox.addLayout(hlay2)
        vbox.addLayout(hlay)
        listeVB = QVBoxLayout(self)
        headerLay = QHBoxLayout(self)

        self.name_header = QLabel("Dossier : ")
        self.name_header.setAlignment(Qt.AlignRight)
        self.name_header.setStyleSheet("font-size: 30px;")
        self.name_header.resize(100,100)
        self.header = QLabel("")
        self.header.setStyleSheet("font-size: 30px;")
        headerLay.addWidget(self.name_header)
        headerLay.addWidget(self.header)
        listeVB.addLayout(headerLay)
        self.treeview = QListView()
        self.listview = QListView()
        listeVB.addWidget(self.listview)
        hlay.addWidget(self.treeview)
        hlay.addLayout(listeVB)


        self.path = os.path.abspath("C:/Users/flo12/PycharmProjects/classeur_de_chantier/Chantier")
        self.model = QStringListModel()
        self.fichier = ""
        self.month = "Mois"
        self.year = ""
        self.update_view()

        self.fileModel = QFileSystemModel()

        self.treeview.setModel(self.model)
        self.treeview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.listview.setModel(self.fileModel)

        self.treeview.clicked.connect(self.on_clicked)
        self.treeview.doubleClicked.connect(self.double_clicked)
        self.listview.doubleClicked.connect(self.click)
        self.lineEdit.textChanged.connect(self.entrPress)
        self.cb_month.currentIndexChanged.connect(self.select_month)
        self.cb_year.currentIndexChanged.connect(self.select_year)
        self.btn1.clicked.connect(self.create_dir)

    def on_clicked(self, index):
        self.listview.setRootIndex(self.fileModel.setRootPath(self.liste_dir[index.row()].absolutePath()))
        self.header.setText(self.liste_dir[index.row()].dirName())

    def double_clicked(self, index):
        os.startfile(self.liste_dir[index.row()].absolutePath())

    def click(self,index):
        path = self.fileModel.fileInfo(index).absoluteFilePath()
        os.startfile(path)

    def find_all(self,name, month, year, path):
        result = []
        for dirs in os.listdir(path):
            if str(name).lower() in str(dirs).lower():
                dir = os.path.join(path,dirs)
                directory = QDir(dir)
                onlyfiles = [f for f in os.listdir(dir) if os.path.isfile(os.path.join(dir, f))]
                statinfo = os.path.getctime(os.path.join(dir,onlyfiles[0]))
                date = datetime.datetime.fromtimestamp(statinfo).strftime('%Y-%m-%d %H:%M:%S')
                if dic_month[month] in date[5:7]:
                    if year in date[:4] or year == "Année":
                        result.append(directory)
        return result

    def entrPress(self):
        self.fichier = self.lineEdit.text()
        self.update_view()

    def select_month(self):
        self.month = self.cb_month.currentText()
        self.update_view()

    def select_year(self):
        self.year = self.cb_year.currentText()
        self.update_view()


    def create_dir(self):
        self.d = QDialog()
        self.d.resize(200,240)
        self.line = QLineEdit(self.d)
        self.nom = QLineEdit(self.d)
        self.prenom = QLineEdit(self.d)
        self.adresse = QLineEdit(self.d)
        self.line.resize(100,20)
        self.nom.resize(100,20)
        self.prenom.resize(100,20)
        self.adresse.resize(100,20)
        self.line.move(50,0)
        self.nom.move(50,30)
        self.prenom.move(50,60)
        self.adresse.move(50,90)
        self.line.setPlaceholderText("Nom du chantier")
        self.nom.setPlaceholderText("Nom du client")
        self.prenom.setPlaceholderText("Prenom du client")
        self.adresse.setPlaceholderText("Adresse du client")
        self.line.textChanged.connect(self.disableButton)
        self.radio_devis = QCheckBox("Devis", self.d)
        self.radio_devis.setChecked(True)
        self.radio_devis.move(20,120)
        self.radio_barreau = QCheckBox("Barreau", self.d)
        self.radio_barreau.setChecked(True)
        self.radio_barreau.move(20,160)
        self.b1 = QPushButton("ok", self.d)
        self.b1.setEnabled(False)
        self.b1.resize(100,40)
        self.b1.move(50, 180)
        self.b1.clicked.connect(self.create_file)
        self.d.setWindowTitle("Creation")
        self.d.setWindowModality(Qt.ApplicationModal)
        self.d.exec_()

    def disableButton(self):
        if len(self.line.text()) > 0:
            self.b1.setEnabled(True)

    def create_file(self):
        self.d.done(1)
        text=self.line.text()
        path = os.path.join(self.path, str(text))
        try:
            os.makedirs(path)
            shutil.copyfile('source/modele_plan.dwg', os.path.join(path,"plan_"+str(text)+".dwg"))
            if self.radio_devis.isChecked():
                shutil.copyfile('source/modele_espacement.xls', os.path.join(path, "calcul_espacement_" + str(text) + ".xls"))
            if self.radio_barreau.isChecked():
                xls_path = os.path.join(path, "devis_" + str(text) + ".xls")
                shutil.copyfile('source/modele_devis.xls',xls_path)
                name="toto"
                if self.nom.text():
                    self.insert_in_xls(xls_path, self.nom)

        except OSError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Le dossier existe déja")
            msg.setWindowTitle("Erreur")
            retval = msg.exec_()
        self.update_view()
        self.listview.setRootIndex(self.fileModel.setRootPath(path))
        self.header.setText(text)


    def update_view(self):
        self.liste_dir = self.find_all(self.fichier, self.month, self.year, self.path)
        liste = []
        for i in self.liste_dir:
            liste.append(i.dirName())
        self.model.removeRows(0, self.model.rowCount())
        self.model.setStringList(liste)

    def insert_in_xls(self,path,name):
        w = copy(open_workbook(path))
        w.get_sheet(1).write(11, 6, name)
        w.save(path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Widget()
    w.showMaximized()
    sys.exit(app.exec_())