#!/usr/bin/env python
# coding: utf-8

### 네이버 메일 기반으로 작성되었으니 보내는 메일은 네이버 메일로 부탁드립니다.

import os
import re
import datetime
import zipfile
import shutil
import pickle
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.encoders import encode_base64
import pandas as pd
import matplotlib as mpl
import matplotlib.font_manager as fm
import openpyxl # 터미널에서 추가 install 필요 (openpyxl)
from win32com import client # 터미널에서 추가 install 필요 (pywin32)

def naver_mail(naver_id, naver_pw, toAddr, mail_header, txt, attached_file_list=[]):
    naver = smtplib.SMTP_SSL('smtp.naver.com', 465)
    naver.ehlo()
    naver.login(naver_id, naver_pw)
    msg = MIMEMultipart("mixed")
    msg.preamble = ''

    msg['Subject'] = Header(mail_header, 'utf8')
    msg['From'] = naver_id
    msg['To'] = toAddr

    msg.attach(MIMEText(txt, _charset="utf8"))

    if attached_file_list != []:
        msg = attach_files(msg, attached_file_list)

    naver.sendmail(naver_id, toAddr, msg.as_string())
    naver.quit()

def main_program():
    # 만약 이메일보관파일을 사용하신다면 아래 두 코드를 사용하시면 됩니다.
    #with open("my_email_acct.dat", "rb") as file:
    #    email = pickle.load(file)

    df = pd.read_excel("고객주문내역.xlsx", sheet_name="Sheet1")

    for i in range(len(df)):
        date = f"""{df.iloc[i]["주문일시"]}"""
        price = f"""{df.iloc[i]["제품가격"]}"""
        count = f"""{df.iloc[i]["주문수량"]}"""
        pay = int(price) * int(count)

        naver_id = "ysm00813@naver.com"  # 사용하시는 메일 주소를 넣어주시면 됩니다. 네이버 메일로 부탁드립니다.
        naver_pw = "mosquitto1!"  # 메일 주소의 비밀번호를 넣어주시면 됩니다. 비밀번호를 바로 입력해주세요.
        #naver_pw = email[naver_id] #만약 이메일보관파일을 사용하신다면 위 코드가 아닌 이 코드를 사용하시면 됩니다.

        toAddr = f"""{df.iloc[i]["이메일"]}"""
        mail_header = "[하나둘상사] 주문배송정보 확인메일"
        txt = f"""
    {df.iloc[i]["고객명"]} 고객님 안녕하십니까. 하나둘상사입니다.
    하나둘상사의 제품을 구입해주셔서 감사합니다.
    주문정보와 배송정보를 확인해주세요.

    주문번호 : {df.iloc[i]["주문번호"]}
    주문일시 : {date[0:4]}년 {date[5:7]}월 {date[8:10]}일 {date[11:13]}시 {date[14:16]}분
    제품명 : {df.iloc[i]["제품명"]}
    주문수량 : {count}
    제품가격 : {price}원
    총 주문가격 : {pay}원

    택배사 : {df.iloc[i]["택배사"]}
    송장번호 : {df.iloc[i]["송장번호"]}
    배송주소 : {df.iloc[i]["배송주소"]}

    문의사항이 있으시다면 연락 바랍니다.
    감사합니다.
    """

        # 아래는 파일 첨부 코드 입니다.
        # 엑셀 파일만 첨부, pdf 파일만 첨부, 엑셀과 pdf 파일 모두 첨부하는 코드 총 세 종류를 구현했습니다.
        # 어느 하나를 사용하실 때 나머지 코드는 주석으로 처리하셔야 됩니다.

        # 1. 엑셀 파일만 첨부 하실 경우 아래 한 줄 코드를 사용하시면 됩니다. (아래 여섯줄 주석처리 필요)
        #attached_file_list = [make_bill("고객용간이영수증.xlsx", "고객주문내역.xlsx", i)]


        # 2. pdf 파일만 첨부 하실 경우 아래 세 줄 코드를 사용하시면 됩니다. (위 한줄, 아래 세줄 주석처리)
        #a = make_bill("고객용간이영수증.xlsx", "고객주문내역.xlsx", i)
        #name = a.replace(".xlsx", "")
        #attached_file_list = [pdffile(name)]


        # 3. 아래 세 줄은 엑셀 파일과 pdf 파일 모두 첨부하는 코드입니다. (위의 네줄 주석처리)
        a = make_bill("고객용간이영수증.xlsx", "고객주문내역.xlsx", i)
        name = a.replace(".xlsx", "")
        attached_file_list = [a, pdffile(name)]


        naver_mail(naver_id, naver_pw, toAddr, mail_header, txt, attached_file_list)

def make_bill(file, user_db, num):
    df = pd.read_excel(user_db, sheet_name="Sheet1")
    price = int(f"""{df.iloc[num]["제품가격"]}""")
    count = int(f"""{df.iloc[num]["주문수량"]}""")
    date = f"""{df.iloc[num]["주문일시"]}"""
    pay = price * count
    dt = datetime.datetime.now()
    wb = openpyxl.load_workbook(file)
    sheet = wb["Sheet1"]
    sheet["A2"] = f"""{df.iloc[num]["고객명"]} 귀하"""
    sheet["A10"] = f"{dt.year}-{dt.month:02d}-{dt.day:02d}"
    sheet["E10"] = pay
    sheet["A13"] = date[5:7] + "월 " + date[8:10] + "일"
    sheet["C13"] = f"""{df.iloc[num]["제품명"]}"""
    sheet["F13"] = count
    sheet["H13"] = price
    sheet["J13"] = pay
    sheet["J23"] = pay
    bill_file = f"""고객영수증({df.iloc[num]["고객명"]}.{df.iloc[num]["주문번호"]})"""
    try:
        wb.save(f"{bill_file}.xlsx")
    except PermissionError:
        print("영수증 파일을 닫고 다시 실행해주세요")
    return f"{bill_file}.xlsx"


def attach_files(msg, attached_file_list):
    for file in attached_file_list:
        if os.path.isfile(file) != True:
            print(f"{file}이 없습니다.")
            break
        part = MIMEBase("application", "octet-stream", _charset="utf8")
        part.set_payload(open(file, "rb").read())
        encode_base64(part)
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=("utf8", "", os.path.basename(file))
        )
        msg.attach(part)
    return msg

def pdffile(file):
    xlApp = client.Dispatch("Excel.Application")

    books = xlApp.Workbooks.Open(os.path.abspath(rf'D:\Users\Storage\PycharmProjects\Ex\{file}.xlsx'))
    # 위 코드에서 엑셀 파일의 절대 경로를 따옴표 안에 바꿔서 넣어주시면 됩니다.
    # 엑셀 파일 이름은 file로 넘어오기 때문에 위 주소에서 file 전 까지만 넣어주셔야 합니다.

    ws = books.Worksheets[1]  # 안에 숫자는 변환할 엑셀 시트 순서 번호를 입력해주시면 됩니다. (0 ~ )
                              # "고객용강의영수증.xlsx"는 sheet1(두 번째 시트)를 사용했습니다.(1)
    ws.Visible = 1

    ws.ExportAsFixedFormat(0, os.path.abspath(rf'D:\Users\Storage\PycharmProjects\Ex\{file}.pdf'))
    # 위 코드에서 엑셀 파일의 절대 경로를 따옴표 안에 바꿔서 넣어주시면 됩니다.
    # 엑셀 파일 이름은 file로 넘어오기 때문에 위 주소에서 file 전 까지만 넣어주셔야 합니다.

    books.Save()
    xlApp.Quit()
    return f"{file}.pdf"

main_program()