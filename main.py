import datetime
import sys

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (QWidget, QToolTip,
                             QPushButton, QApplication, QTableWidget, QAbstractItemView, QLabel, QTableWidgetItem,
                             QGridLayout, QMessageBox)
from PyQt5.QtCore import Qt
import program


class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.table = QTableWidget(5, 8, self)
        self.label = QLabel(self)
        self.btn_1 = QPushButton(self)
        self.btn_2 = QPushButton(self)
        self.btn_3 = QPushButton(self)
        self.btn_4 = QPushButton(self)
        self.grid = QGridLayout(self)
        self.count = 0
        self.excel_content = program.parse_excel("config_info.xlsx")
        if not self.excel_content:
            QMessageBox.about(self, "提示", "当前程序路径下，没有找到配置表格")
            sys.exit()
        self.excel_row = len(self.excel_content)
        self.initUI()



    def initUI(self):
        QToolTip.setFont(QFont('SansSerif', 10))
        # self.table.setFixedSize(800, 400)
        # self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # self.table.resizeColumnsToContents(TableWidget)
        # QTableWidget.resizeColumnsToContents(self.table)
        # QTableWidget.resizeRowsToContents(self.table)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setHorizontalHeaderLabels(['编号', '启动程序地址', '配置文件地址', '数据库地址', ' 批号 ', '       统计       ', '是否在线', '   PID   '])
        self.setGeometry(300, 300, 1200, 600)
        self.set_table()
        QTableWidget.resizeColumnsToContents(self.table)
        QTableWidget.resizeRowsToContents(self.table)
        self.table.itemClicked.connect(self.row_changed)
        self.table.itemDoubleClicked.connect(self.dou_clik)
        self.label.move(650, 30)
        self.label.setText("启动统计后显示")
        self.label.setFixedSize(100, 30)

        self.btn_1.setText("选中行统计")
        self.btn_1.move(650, 80)
        self.btn_1.clicked.connect(self.btn_1_click)
        self.btn_1.setFixedSize(100, 30)

        self.btn_3.setText("所有行统计")
        self.btn_3.move(650, 130)
        self.btn_3.clicked.connect(self.btn_3_click)
        self.btn_3.setFixedSize(100, 30)

        self.btn_2.setText("检查是否在线")
        self.btn_2.move(650, 180)
        self.btn_2.clicked.connect(self.btn_2_click)
        self.btn_2.setFixedSize(100, 30)

        self.btn_4.setText("刷新程序配置")
        self.btn_4.move(650, 230)
        self.btn_4.clicked.connect(self.btn_4_click)
        self.btn_4.setFixedSize(100, 30)

        self.grid.addWidget(self.table, 0, 0, 5, 1)
        self.grid.addWidget(self.btn_1, 0, 1)
        self.grid.addWidget(self.btn_2, 2, 1)

        self.grid.addWidget(self.btn_3, 1, 1)
        self.grid.addWidget(self.btn_4, 3, 1)
        # self.grid.addWidget(self.label,2,1)
        self.setWindowTitle("领取宝图业务控制中心")
        self.table.verticalHeader().setVisible(False)
        self.show()

    def row_changed(self):
        # self.label.setText("clecl")
        row = self.table.selectedItems()[0].row()
        # self.label.setText(str(row))

    def dou_clik(self):
        row = self.table.selectedItems()[0].row()
        item_exe = self.table.item(row, 1).text()
        item_config = self.table.item(row, 2).text()
        process_pid = program.start_prgram(item_exe, item_config)
        row = self.table.selectedItems()[0].row()
        self.table.item(row, 7).setText(str(process_pid))
        QTableWidget.resizeColumnsToContents(self.table)
        QTableWidget.resizeRowsToContents(self.table)

    def set_table(self):
        self.count += 1
        self.table.setRowCount(self.excel_row)
        for row in range(self.excel_row):
            item_data = QTableWidgetItem(self.excel_content[row][0])
            self.table.setItem(row, 0, item_data)

            item_data = QTableWidgetItem(self.excel_content[row][1])
            self.table.setItem(row, 1, item_data)

            database_content = program.get_db_pi_path(self.excel_content[row][2])
            if not database_content:
                QMessageBox.about(self, "提示", self.excel_content[row][2] + " 没有找到该文件")
                if self.count == 1:
                    sys.exit()
                else:
                    return False

            item_data = QTableWidgetItem(self.excel_content[row][2])
            self.table.setItem(row, 2, item_data)

            item_data = QTableWidgetItem(database_content[0])
            self.table.setItem(row, 3, item_data)

            item_data = QTableWidgetItem(database_content[1])
            item_data.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 4, item_data)

            item_data = QTableWidgetItem("")
            item_data.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 5, item_data)

            item_data = QTableWidgetItem("")
            item_data.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 6, item_data)

            item_data = QTableWidgetItem("")
            item_data.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 7, item_data)
        return True

    def update_sum(self):
        sum_gold = 0
        sum_silver = 0
        for row in range(self.excel_row):
            gold_silver = self.table.item(row, 5).text()
            if gold_silver != '':
                one_gold = int((gold_silver.split('\n')[0]).split("：")[1])
                one_silver = int((gold_silver.split('\n')[1]).split("：")[1])
                sum_gold += one_gold
                sum_silver += one_silver
        return [sum_gold, sum_silver]

    def btn_1_click(self):
        # self.label.setText("btn_1")
        if self.table.selectedItems() == []:
            QMessageBox.about(self, "提示", "请选择一行")
            return
        row = self.table.selectedItems()[0].row()
        process_path = self.table.item(row, 1).text()
        one_gold_silver = program.income_sum(process_path)

        if not one_gold_silver:
            log_path = process_path.split("BaoTu.exe")[0] + "Log\\" + datetime.date.today().strftime(
                '%Y_%m_%d') + ".log"
            QMessageBox.about(self, "提示", log_path + " 没有找到该文件")
            return

        self.table.item(row, 5).setText(
            "金豆：" + str(one_gold_silver[0]) + '\n' + "银豆：" + str(one_gold_silver[1]))
        QTableWidget.resizeColumnsToContents(self.table)
        QTableWidget.resizeRowsToContents(self.table)

        sum_gold_silve = self.update_sum()
        self.label.setText("金豆：" + str(sum_gold_silve[0]) + '\n' + "银豆：" + str(sum_gold_silve[1]))

    def btn_2_click(self):
        # self.label.setText("btn_2")
        for row in range(self.excel_row):

            pid_str = self.table.item(row, 7).text()
            if pid_str != '':
                process_exist = program.pid_is_exist(int(self.table.item(row, 7).text()))
                if process_exist:
                    self.table.item(row, 6).setText("是")
                else:
                    self.table.item(row, 6).setText("否")

    def btn_3_click(self):
        for row in range(self.excel_row):
            process_path = self.table.item(row, 1).text()
            one_gold_silver = program.income_sum(process_path)

            if not one_gold_silver:
                log_path = process_path.split("BaoTu.exe")[0] + "Log\\" + datetime.date.today().strftime(
                    '%Y_%m_%d') + ".log"
                QMessageBox.about(self, "提示", log_path + " 没有找到该文件")
                return

            self.table.item(row, 5).setText(
                "金豆：" + str(one_gold_silver[0]) + '\n' + "银豆：" + str(one_gold_silver[1]))
            QTableWidget.resizeColumnsToContents(self.table)
            QTableWidget.resizeRowsToContents(self.table)

            sum_gold_silve = self.update_sum()
            self.label.setText("金豆：" + str(sum_gold_silve[0]) + '\n' + "银豆：" + str(sum_gold_silve[1]))

    def btn_4_click(self):
        self.excel_content = program.parse_excel("config_info.xlsx")
        self.excel_row = len(self.excel_content)
        if self.set_table():
            QMessageBox.about(self, "提示", "刷新成功")

    def resizeEvent(self, Event):
        self.label.move(self.width() - 100, self.height() - 100)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
