import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import QTimer
import mysql.connector
import time

SLEEP_TIME = 3.6

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Kiwoom Login
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")

        # OpenAPI+ Event
        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)

        self.setWindowTitle("PyStock")
        self.setGeometry(300, 300, 300, 150)

        label = QLabel('종목코드: ', self)
        label.move(20, 20)

        self.code_edit = QLineEdit(self)
        self.code_edit.move(80, 20)
        self.code_edit.setText("039490")

        btn1 = QPushButton("조회", self)
        btn1.move(190, 20)
        btn1.clicked.connect(self.btn1_clicked)

        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(10, 60, 280, 80)
        self.text_edit.setEnabled(False)

        # DB 연결 및 테이블 생성
        self.conn = mysql.connector.connect(
            host='localhost',  # MySQL 서버 호스트
            user='root',  # MySQL 사용자 이름
            password='1029384756',  # MySQL 비밀번호
            database='jusicool'  # 사용할 데이터베이스 이름
        )
        self.cursor = self.conn.cursor()
        self.create_table()

        # Timer 설정
        self.stock_codes = ["005930", "035420", "035720", "051910", "207940",  # 예시로 5개 추가, 필요한 주식 코드 추가 필요
                            "000660", "005380", "068270", "005935", "017670",
                            "096770", "034730", "000270", "012330", "028260",
                            "051900", "015760", "105560", "055550", "032830",
                            "066570", "018260", "003550", "036570", "000810",
                            "066570", "036490", "051600", "011170", "003410",
                            "034220", "011070", "035720", "009150", "018880",
                            "005940", "035720", "010950", "024110", "090430",
                            "016360", "025560", "012800", "009830", "034730",
                            "010120", "008600", "004170", "010620", "012630"]

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.request_stock_data)

    def create_table(self):
        # 테이블 초기화
        self.cursor.execute('''DROP TABLE IF EXISTS stock_data''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS stock_data
                             (code VARCHAR(20) PRIMARY KEY, 
                             name VARCHAR(50), 
                             volume BIGINT,
                             price_5min BIGINT, 
                             price_10min BIGINT,
                             price_24h BIGINT,
                             price_1week BIGINT,
                             price_1month BIGINT)''')

    def event_connect(self, err_code):
        if err_code == 0:
            self.text_edit.append("로그인 성공")
            self.get_stock_list()

    def get_stock_list(self):
        # 타이머 시작 (3.6초 간격으로 요청)
        time.sleep(SLEEP_TIME)

    def request_stock_data(self):
        if not self.stock_codes:
            self.timer.stop()
            return

        self.code = self.stock_codes.pop(0)
        self.request_price("5MIN")
        # time.sleep(SLEEP_TIME)
        self.request_price("10MIN")
        # time.sleep(SLEEP_TIME)
        self.request_price("24H")
        # time.sleep(SLEEP_TIME)
        self.request_price("1WEEK")
        # time.sleep(SLEEP_TIME)
        self.request_price("1MONTH")
        # time.sleep(SLEEP_TIME)

    def request_price(self, interval):
        if interval == "5MIN":
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.code)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10080_req_5min", "opt10080", 0, "0101")
        elif interval == "10MIN":
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.code)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10080_req_10min", "opt10080", 0, "0101")
        elif interval == "24H":
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.code)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10081_req_24h", "opt10081", 0, "0101")
        elif interval == "1WEEK":
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.code)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10082_req_1week", "opt10082", 0, "0101")
        elif interval == "1MONTH":
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", self.code)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10083_req_1month", "opt10083", 0, "0101")

    def btn1_clicked(self):
        code = self.code_edit.text()
        self.text_edit.append("종목코드: " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")

    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "opt10001_req":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")
            volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "거래량")
            
            if name and volume:
                self.update_db_bulk(self.code, name.strip(), volume.strip())

        elif "opt10080" in rqname or "opt10081" in rqname or "opt10082" in rqname or "opt10083" in rqname:
            price_5min = None
            price_10min = None
            price_24h = None
            price_1week = None

            if "opt10080" in rqname:
                price_5min = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "현재가").strip()
            elif "opt10081" in rqname:
                price_10min = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "현재가").strip()
            elif "opt10082" in rqname:
                price_24h = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "현재가").strip()
            elif "opt10083" in rqname:
                price_1week = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "현재가").strip()

            self.update_db_bulk(self.code, None, None, price_5min, price_10min, price_24h, price_1week)

    def update_db_bulk(self, code, name=None, volume=None, price_5min=None, price_10min=None, price_24h=None, price_1week=None):
        try:
            if name is not None and volume is not None:
                self.cursor.execute("INSERT INTO stock_data (code, name, volume) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE name = VALUES(name), volume = VALUES(volume)", (code, name, volume))

            if any([price_5min, price_10min, price_24h, price_1week]):
                sql = "INSERT INTO stock_data (code, price_5min, price_10min, price_24h, price_1week) VALUES (%s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE price_5min = VALUES(price_5min), price_10min = VALUES(price_10min), price_24h = VALUES(price_24h), price_1week = VALUES(price_1week)"
                val = (code, price_5min, price_10min, price_24h, price_1week)
                self.cursor.execute(sql, val)

            self.conn.commit()
            self.text_edit.append("데이터베이스에 저장되었습니다.")

        except mysql.connector.Error as e:
            self.text_edit.append(f"Error: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
