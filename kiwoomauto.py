"""
키움 OpenApi+ 모듈을 업데이트하기 위해서 번개2를 실행했다가 종료시킨다.
스케줄러에 등록해서 사용한다.

author: 서경동
last edit: 2017. 01. 10.
"""


from pywinauto import application
from pywinauto import timings
import time
import os


# Account
account = []
with open("C:\\Users\\seoga\\PycharmProjects\\PyTrader\\account.txt", 'r') as f:
    account = f.readlines()

# 번개2 실행 및 로그인
app = application.Application()
app.start("C:\Kiwoom\KiwoomFlash2\khministarter.exe")

title = "번개 Login"
dlg = timings.WaitUntilPasses(20, 0.5, lambda: app.window_(title=title))

idForm = dlg.Edit0
idForm.SetFocus()
idForm.TypeKeys(account[0])

passForm = dlg.Edit2
passForm.SetFocus()
passForm.TypeKeys(account[1])

certForm = dlg.Edit3
certForm.SetFocus()
certForm.TypeKeys(account[2])

loginBtn = dlg.Button0
loginBtn.Click()

# 업데이트가 완료될 때 까지 대기
while True:
    time.sleep(5)
    with os.popen('tasklist /FI "IMAGENAME eq khmini.exe"') as f:
        lines = f.readlines()
        if len(lines) >= 3:
            break

# 번개2 종료
time.sleep(30)
os.system("taskkill /im khmini.exe")
