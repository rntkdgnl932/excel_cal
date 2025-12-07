# excel_ui.py
# 탭 통합 메인 UI: 기존 엑셀 계산기 + 네이버/쿠팡 송장 읽기

import sys
from PyQt5 import QtWidgets

from excel_cal_ui import ExcelCalWindow      # 기존 부가세/3종 엑셀 UI
from read_excel import ReadInvoiceWidget     # 새로 만들 송장 읽기 탭


class MainTabbedWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("하비 브라운 엑셀 도구 모음")
        self.resize(1200, 800)

        # 중앙에 탭 위젯 배치
        tabs = QtWidgets.QTabWidget(self)
        self.setCentralWidget(tabs)

        # 1. 기존 부가세/3종 엑셀 생성기 탭
        self.cal_window = ExcelCalWindow()   # QMainWindow 이지만 QWidget처럼 탭에 넣어도 동작함
        tabs.addTab(self.cal_window, "부가세 계산 / 3종 엑셀")

        # 2. 네이버/쿠팡 송장 엑셀 읽기 탭
        self.read_invoice_widget = ReadInvoiceWidget(self)
        tabs.addTab(self.read_invoice_widget, "네이버·쿠팡 송장 엑셀 읽기")


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainTabbedWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
