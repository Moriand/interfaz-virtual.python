import sys
import csv
from pandas.io.parsers import read_csv
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextDocument, QTextCursor, QTextTableFormat, QFont
from PyQt5.QtPrintSupport import QPrintPreviewDialog, QPrinter
from PyQt5.QtWidgets import (QMainWindow, QApplication, QFileDialog, QTableWidget, QTableWidgetItem, QFrame, QMessageBox)
import untitled_ui
import resource_rc


class Table(QTableWidget):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setFrameShape(QFrame.StyledPanel)
        self.setFrameShadow(QFrame.Sunken)
        self.setAlternatingRowColors(True)
        self.setGridStyle(Qt.CustomDashLine)
        self.setRowCount(10)
        self.setColumnCount(10)

        # etc.

    def copy(self):
        self.copied_cells = sorted(self.selectedIndexes())
        # print("presionado Ctrl+C")

    def paste(self):
        if self.copied_cells:
            r = self.currentRow() - self.copied_cells[0].row()
            c = self.currentColumn() - self.copied_cells[0].column()
            # print("presionado Ctrl+V")
            # print(r)
            for cell in self.copied_cells:
                self.setItem(cell.row() + r, cell.column() + c, QTableWidgetItem(cell.data()))


class Ui_MainWindow(QMainWindow, untitled_ui.Ui_MainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)

        self.actionNew.setShortcut("Ctrl+N")
        self.actionOpen.setShortcut("Ctrl+O")
        self.actionSave.setShortcut("Ctrl+S")
        self.actionPaste.setShortcut("Ctrl+V")
        self.actionCopy.setShortcut("Ctrl+C")
        self.actionPrint.setShortcut("Ctrl+P")

        self.actionNew.triggered.connect(self.on_new)
        self.actionOpen.triggered.connect(self.On_open_file)
        self.actionSave.triggered.connect(self.On_save_file)
        self.actionPaste.triggered.connect(self.On_paste)
        self.actionCopy.triggered.connect(self.On_copy)
        self.actionPrint.triggered.connect(self.on_printpreview)

    def on_new(self):
        print("nueva pestaña")
        self.table = Table()
        self.tabWidget.addTab(self.table, "* Tabla Nueva *")

    def On_open_file(self, e):
        path = QFileDialog.getOpenFileName(self, "Abrir fichero", "", "File ( *.csv *.txt *.md);; All files (*.*)")
        print(path[0])
        if path:
            df = read_csv(path[0], header=None, delimiter=',', keep_default_na=False, error_bad_lines=False)
            header = df.iloc[0]
            ret = QMessageBox.question(self,"CSV visor", "Desea usar esto como encabezado?\n\n"+ str(header.values), QMessageBox.Yes | QMessageBox.No, defaultButton= QMessageBox.Yes)
            if ret == QMessageBox.Yes:
                df = df[1:]
                self.table.setColumnCount(len(df.columns))
                self.table.setRowCount(len(df.index))

                for i in range(len(df.index)):
                    for j in range(len(df.columns)):
                        self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

                for j in range(len(df.columns)):
                    m = QTableWidgetItem(header[j])
                    self.table.setHorizontalHeaderItem(j, m)
            else:
                self.table.setColumnCount(len(df.columns))
                self.table.setRowCount(len(df.index))

                for i in range(len(df.index)):
                    for j in range(len(df.columns)):
                        self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))
                for j in range(self.table.columnCount()):
                    m = QTableWidgetItem(str(j))
                    self.table.setHorizontalHeaderItem(j, m)
        # self.table.selectRow(0)
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        print("Abrir Archivo")

    def On_save_file(self, e):
        path = QFileDialog.getSaveFileName(self, "Guardar como", "", "File csv ( *.csv *.txt)")
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        if path:
            with open(path[0], 'w') as stream:
                print("guardando", path)
                writer = csv.writer(stream, delimiter=',')
                for row in range(self.table.rowCount()):
                    rowdata = []
                    for column in range(self.table.columnCount()):
                        item = self.table.item(row, column)
                        if item is not  None:
                            rowdata.append(item.text())
                        else:
                            rowdata.append('')
                    writer.writerow(rowdata)


    def On_paste(self, e):
        w = self.tabWidget.currentWidget()
        if hasattr(w, "paste") and callable(w.paste):
            w.paste()
            print(w.paste())
        print("pegado de texto")

    def On_copy(self):
        w = self.tabWidget.currentWidget()
        if hasattr(w, "copy") and callable(w.copy):
            w.copy()

    def on_printpreview(self):
        if self.table.rowCount() == 0:
            print("No hay datos")
        else:
            dlg = QPrintPreviewDialog()
            dlg.setFixedSize(1000, 700)
            dlg.paintRequested.connect(self.paint_resquest)
            dlg.exec_()
            print("ha sido cerrado")

    def paint_resquest(self, printer):
        for row in range(self.table.rowCount()):
            for column in range(self.table.columnCount()):
                myitem = self.table.item(row, column)
                if myitem is None:
                    item = QTableWidgetItem("")
                    self.table.setItem(row, column, item)
        printer.setDocName("Lista")
        document = QTextDocument()
        cursor = QTextCursor(document)
        model = self.table.model()
        tableFormat = QTextTableFormat()
        tableFormat.setBorder(0.3)
        tableFormat.setBorderStyle(3);
        tableFormat.setTopMargin(0);
        tableFormat.setCellPadding(6)
        table = cursor.insertTable(model.rowCount(), model.columnCount(), tableFormat)
        for row in range(table.rows()):
            for column in range(table.columns()):
                cursor.insertText(self.table.item(row, column).text())
                cursor.movePosition(QTextCursor.NextCell)
        document.print_(printer)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = Ui_MainWindow()
    frame.show()
    sys.exit(app.exec_())