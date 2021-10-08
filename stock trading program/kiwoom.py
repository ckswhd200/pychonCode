import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import sqlite3
import pandas as pd

TR_REQ_TIME_INTERVAL = 0.2  # 키움증권은 1초에 최대 5번의 tr 요청을 허용하므로 0.2초 대기


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.opw00018_output = {'single': [], 'multi': []}

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect)  # OnEventConnect 이벤트 발생 시 실행할 메서드와 연결
        self.OnReceiveTrData.connect(self._receive_tr_data)  # OnReceiveTrData 이벤트 발생시 호출되는 메서드
        self.OnReceiveChejanData.connect(self._receive_chejan_data)  # 주문 체결 시점에 키움서버에서 발생키시는 메서드

    def comm_connect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 전송
        self.login_event_loop = QEventLoop()  # 키움 서버에서 OnEventConnect 이벤트가 발생할 때까지 종료하지 않기 위해 이벤트루프
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()  # 연결됐으므로 루프 종료

    def set_input_value(self, id, value):  # 서버 통신 전 tran값을 입력
        self.dynamicCall("SetInputValue(QString, QString", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):  # 키움 서버로 tran 송신
        self.dynamicCall("CommRqData(QString, QString, int, QString", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()  # 서버로부터 통신이 올 때까지 대기하기 위해 이벤트루프 생성
        self.tr_event_loop.exec_()

    def _get_repeat_cnt(self, trcode, rqname):  # 요청한 데이터의 개수를 저장하기 위한 메서드
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def _comm_get_data(self, code, real_type, field_name, index, item_name):  # 데이터를 가져오기 위한 메서드
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString", code, real_type, field_name, index, item_name)
        return ret.strip()  # CommGetData 반환값 양쪽에 공백이 있기 때문에 제거

    def _opt10081(self, rqname, trcode):
        data_cnt = self._get_repeat_cnt(trcode, rqname)  # 데이터가 900개보다 작을 수 있으므로 개수 먼저 파악

        for i in range(data_cnt):
            date = self._comm_get_data(trcode, "", rqname, i, "일자")
            open = self._comm_get_data(trcode, "", rqname, i, "시가")
            high = self._comm_get_data(trcode, "", rqname, i, "고가")
            low = self._comm_get_data(trcode, "", rqname, i, "저가")
            close = self._comm_get_data(trcode, "", rqname, i, "현재가")
            volume = self._comm_get_data(trcode, "", rqname, i, "거래량")

            self.ohlcv['date'].append(date)
            self.ohlcv['open'].append(int(open))
            self.ohlcv['high'].append(int(high))
            self.ohlcv['low'].append(int(low))
            self.ohlcv['close'].append(int(close))
            self.ohlcv['volume'].append(int(volume))

    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2':  # TR 요청시 최대 900개의 데이터가 반환되고, opw00001 요청시 최대 20개 종목이 반환됨. next==2면 추가요청이 필요.
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == "opt10081_req":
            self._opt10081(rqname, trcode)
        elif rqname == "opw00001_req":
            self._opw00001(rqname, trcode)
        elif rqname == "opw00018_req":
            self._opw00018(rqname, trcode)

        try:
            self.tr_event_loop.exit()  # 이 메소드가 실행됐다는 건 OnReceiveTrData 이벤트가 정상적으로 실행된 것이므로 이벤트루프 종료
        except AttributeError:
            pass

    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')  # 코드가 123;456;789;처럼 세미콜론으로 구분된 문자열로 반환되므로 리스트로 변환
        return code_list[:-1]

    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_name

    def get_connect_state(self):
        ret = self.dynamicCall("GetConnectState()")
        return ret

    def send_order(self, rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no):
        self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                         [rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no])
        print("거래 성공")

    def get_chejan_data(self, fid):  # 체결잔고를 가져옴
        ret = self.dynamicCall("GetChejanData(int)", fid)
        return ret

    def _receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(gubun)
        print(self.get_chejan_data(9203))  # 주문번호
        print(self.get_chejan_data(302))  # 종목명
        print(self.get_chejan_data(900))  # 주문수량
        print(self.get_chejan_data(901))  # 주문가격

    def get_login_info(self, tag):
        ret = self.dynamicCall("GetLoginInfo(QString)", tag)
        return ret



    def _opw00001(self, rqname, trcode):
        d2_deposit = self._comm_get_data(trcode, "", rqname, 0, "d+2추정예수금")
        self.d2_deposit = Kiwoom.change_format(d2_deposit)

    def _opw00018(self, rqname, trcode):
        # 싱글 데이터 부분
        total_purchase_price = self._comm_get_data(trcode, "", rqname, 0, "총매입금액")
        total_eval_price = self._comm_get_data(trcode, "", rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, 0, "총평가손익금액")
        total_earning_rate = self._comm_get_data(trcode, "", rqname, 0, "총수익률(%)")
        estimated_deposit = self._comm_get_data(trcode, "", rqname, 0, "추정예탁자산")

        total_earning_rate = Kiwoom.change_format(total_earning_rate)

        if self.get_server_gubun():
            total_earning_rate = float(total_earning_rate) / 100
            total_earning_rate = str(total_earning_rate)

        self.opw00018_output['single'].append(Kiwoom.change_format(total_purchase_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_profit_loss_price))
        self.opw00018_output['single'].append(total_earning_rate)
        self.opw00018_output['single'].append(Kiwoom.change_format(estimated_deposit))

        print(Kiwoom.change_format(total_purchase_price))
        print(Kiwoom.change_format(total_eval_price))
        print(Kiwoom.change_format(total_eval_profit_loss_price))
        print(Kiwoom.change_format(total_earning_rate))
        print(Kiwoom.change_format(estimated_deposit))

        rows = self._get_repeat_cnt(trcode, rqname)
        for i in range(rows):
            name = self._comm_get_data(trcode, "", rqname, i, "종목명")
            quantity = self._comm_get_data(trcode, "", rqname, i, "보유수량")
            purchase_price = self._comm_get_data(trcode, "", rqname, i, "매입가")
            current_price = self._comm_get_data(trcode, "", rqname, i, "현재가")
            eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, i, "평가손익")
            earning_rate = self._comm_get_data(trcode, "", rqname, i, "수익률(%)")


            quantity = Kiwoom.change_format(quantity)
            purchase_price = Kiwoom.change_format(purchase_price)
            current_price = Kiwoom.change_format(current_price)
            eval_profit_loss_price = Kiwoom.change_format(eval_profit_loss_price)
            earning_rate = Kiwoom.change_format2(earning_rate)

            self.opw00018_output['multi'].append([name, quantity, purchase_price, current_price, eval_profit_loss_price, earning_rate])
            print("멀티데이터 : ", self.opw00018_output['multi'])


    def reset_opw00018_output(self):
        self.opw00018_output = {'single': [], 'multi': []}

    @staticmethod
    def change_format(data):
        strip_data = data.lstrip('-0')
        if strip_data == '':
            strip_data = '0'

        try:
            format_data = format(int(strip_data), ',d')
        except:
            format_data = format(float(strip_data))

        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    @staticmethod
    def change_format2(data):
        strip_data = data.lstrip('-0')

        if strip_data == '':
            strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    def get_server_gubun(self):  # 모의투자 서버에서 주는 데이터와, 실 투자 서버에서 주는 수익률 데이터의 형식이 다른 경우 구분해주는 메서드
        ret = self.dynamicCall("KOA_Function(QString, QString)", "GetServerGubun", "")
        return ret


if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect()

    print(1)

    account_number = kiwoom.get_login_info("ACCNO")
    account_number = account_number.split(';')[0]

    kiwoom.set_input_value("계좌번호", account_number)
    kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

    # 계좌번호 : 8009618711
