# -*- coding: utf-8 -*-
import sys, re, time, random, requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication, QWidget, QGridLayout, QLabel, QLineEdit, QComboBox,
    QGroupBox, QCheckBox, QPushButton, QTextEdit, QMessageBox, QStackedWidget
)

# ---------------- Selenium for calendar ----------------
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
CAL_URL = "https://www.nchu.edu.tw/calendar/"

def parse_year_month(text: str, last_year=None):
    text = re.sub(r"\s+", "", text)
    year = month = None
    m = re.search(r"(\d+)年(\d+)月", text)
    if m: return int(m.group(1)) + 1911, int(m.group(2))
    m = re.search(r"(\d+)年", text)
    if m: year = int(m.group(1)) + 1911
    m = re.search(r"(\d+)月", text)
    if m:
        month = int(m.group(1))
        if year is None and last_year is not None: year = last_year
    return year, month

def is_red_cell(td):
    s = (td.get("style") or "").lower()
    return ("#800000" in s) or ("#8b0000" in s) or ("background-color:maroon" in s)

def crawl_red_days_roc_strings():
    """return set like {'1150201','1150207', ...} (ROC yyyMMdd)"""
    opt = Options()
    opt.add_argument("--headless=new")
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-gpu")
    d = webdriver.Chrome(options=opt)
    try:
        d.get(CAL_URL)
        WebDriverWait(d, 12).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(1.2)
        soup = BeautifulSoup(d.page_source, "html.parser")

        last_year = current_month = None
        red = set()
        for td in soup.find_all("td"):
            if (td.get("id") or "").startswith("rowspan"):
                ym = td.get_text("", strip=True)
                y, m = parse_year_month(ym, last_year)
                if y is not None: last_year = y
                if m is not None: current_month = m
                continue
            if is_red_cell(td) and last_year and current_month:
                day = td.get_text(strip=True)
                if day.isdigit():
                    red.add(f"{last_year-1911:03d}{current_month:02d}{int(day):02d}")
        return red
    finally:
        d.quit()

# ---------------- worker ----------------
class SubmitWorker(QThread):
    status = pyqtSignal(str)
    done   = pyqtSignal(str)

    def __init__(self, username, password, mode_with_time, day_choice,
                 begin_time, end_time, hours, contents, parent=None):
        super().__init__(parent)
        self.username=username; self.password=password
        self.mode_with_time=mode_with_time
        self.day_choice=day_choice
        self.begin_time=begin_time; self.end_time=end_time; self.hours=hours
        self.contents=[c for c in contents if c]

    def log(self, s): self.status.emit(s)

    def run(self):
        try:
            self.log("抓取行事曆紅底日期…")
            red_days = crawl_red_days_roc_strings()
            self.log(f"紅底日期數：{len(red_days)}")

            s = requests.Session()
            login_url = "https://psf.nchu.edu.tw/punch/login_chk.jsp"
            r = s.post(login_url, data={"txtLoginID":self.username,"txtLoginPWD":self.password}, timeout=20)
            if "/Menu.jsp" not in r.text:
                self.done.emit("登入失敗：請檢查帳號或密碼。"); return

            today = datetime.today()
            first = today.replace(day=1)
            dates = []
            for i in range((today-first).days+1):
                d = first + timedelta(days=i)
                if d.weekday()<5:
                    roc=f"{d.year-1911:03d}{d.strftime('%m%d')}"
                    if roc not in red_days: dates.append(roc)
            if not dates:
                self.done.emit("沒有可用的平日（非紅底）。"); return

            if self.day_choice=="全部平日":
                selected = dates[:]
            else:
                n=int(self.day_choice)
                if n>len(dates):
                    self.done.emit(f"可選天數不足（可選 {len(dates)} 天，要 {n} 天）。"); return
                selected = random.sample(dates, n)

            if not self.contents:
                self.done.emit("請至少勾選一個工作內容。"); return

            if self.mode_with_time:
                list_url="https://psf.nchu.edu.tw/punch/PunchList.jsp"
                add_url ="https://psf.nchu.edu.tw/punch/PunchListS.jsp"
            else:
                list_url="https://psf.nchu.edu.tw/punch/PunchList_A.jsp"
                add_url ="https://psf.nchu.edu.tw/punch/PunchListS_A.jsp"

            page = s.get(list_url, timeout=20)
            soup = BeautifulSoup(page.content,"html.parser")
            sel = soup.find("select", {"name":"schno"})
            if not sel or not sel.find("option"):
                self.done.emit("找不到 schno 下拉選單。"); return
            schno = sel.find("option").get("value")

            ok=fail=0
            for roc_date in sorted(selected):
                work = random.choice(self.contents)
                if self.mode_with_time:
                    if not (self.begin_time and self.end_time and self.hours):
                        self.done.emit("請輸入完整時間欄位（begintime/endtime/hours）。"); return
                    payload = {"date":roc_date,"begtime":self.begin_time,"endtime":self.end_time,
                               "hours":self.hours,"work":"行政事務","schno":schno,"hidACT":"add"}
                else:
                    payload = {"date":roc_date,"work":work,"schno":schno,"hidACT":"add"}

                self.log(f"提交：{roc_date}（{payload['work']}）")
                resp = s.post(add_url, data=payload, timeout=20)
                if "ERROR:null" in resp.text: ok+=1
                else: fail+=1

            self.done.emit(f"完成：成功 {ok}，失敗 {fail}。")
        except Exception as e:
            self.done.emit(f"例外：{e}")

# ---------------- main window ----------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("學習日誌自動化工具（Qt 兩鍵切換版，含時間=行政事務固定）")
        #self.resize(780, 480)

        self.mode_with_time = True

        g = QGridLayout(self); r=0
        # === 第 0 列：模式切換按鈕 ===
        self.btn_simple = QPushButton("學習日誌")
        self.btn_time   = QPushButton("出勤紀錄表")
        g.addWidget(self.btn_simple, 0, 0, 1, 1, alignment=Qt.AlignLeft)
        g.addWidget(self.btn_time,   0, 1, 1, 1, alignment=Qt.AlignLeft)
        self.btn_simple.clicked.connect(lambda: self.set_mode(False))
        self.btn_time.clicked.connect(lambda: self.set_mode(True))

        # === 第 1 列：帳號 ===
        r = 1
        g.addWidget(QLabel("帳號(身份證字號):"), r,0, alignment=Qt.AlignRight)
        self.ed_user = QLineEdit(); self.ed_user.setFixedWidth(240)
        g.addWidget(self.ed_user, r,1, alignment=Qt.AlignLeft)

        r+=1
        g.addWidget(QLabel("密碼(民國出生年月日):"), r,0, alignment=Qt.AlignRight)
        self.ed_pass = QLineEdit(); self.ed_pass.setEchoMode(QLineEdit.Password); self.ed_pass.setFixedWidth(240)
        g.addWidget(self.ed_pass, r,1, alignment=Qt.AlignLeft)

        # Stack pages
        r+=1
        self.stack = QStackedWidget()
        g.addWidget(self.stack, r,0, 1, 4)

        # 含時間版
        self.page_time = QWidget(); gt = QGridLayout(self.page_time); pr=0
        gt.addWidget(QLabel("開始時間 (HHmm):"), pr,0, alignment=Qt.AlignRight)
        self.ed_begin = QLineEdit("0830"); self.ed_begin.setFixedWidth(240)
        gt.addWidget(self.ed_begin, pr,1, alignment=Qt.AlignLeft)

        pr+=1
        gt.addWidget(QLabel("結束時間 (HHmm):"), pr,0, alignment=Qt.AlignRight)
        self.ed_end = QLineEdit("1730"); self.ed_end.setFixedWidth(240)
        gt.addWidget(self.ed_end, pr,1, alignment=Qt.AlignLeft)

        pr+=1
        gt.addWidget(QLabel("總時數 (hours):"), pr,0, alignment=Qt.AlignRight)
        self.ed_hours = QLineEdit("8"); self.ed_hours.setFixedWidth(240)
        gt.addWidget(self.ed_hours, pr,1, alignment=Qt.AlignLeft)

        pr+=1
        gt.addWidget(QLabel("選擇天數:"), pr,0, alignment=Qt.AlignRight)
        self.cb_days_A = QComboBox(); self.cb_days_A.addItems([str(i) for i in range(1,11)]+["全部平日"])
        self.cb_days_A.setCurrentIndex(0); self.cb_days_A.setFixedWidth(240)
        gt.addWidget(self.cb_days_A, pr,1, alignment=Qt.AlignLeft)

        pr+=1
        gt.addWidget(QLabel("工作內容:"), pr,0, alignment=Qt.AlignRight)
        gbA = QGroupBox(); gbAl = QGridLayout(gbA)
        self.ck_admin_fixed = QCheckBox("行政事務")
        self.ck_admin_fixed.setChecked(True)
        self.ck_admin_fixed.setEnabled(False)
        gbAl.addWidget(self.ck_admin_fixed, 0, 0)
        gt.addWidget(gbA, pr,1,1,2)

        # 簡易版
        self.page_simple = QWidget(); gs = QGridLayout(self.page_simple); pr=0
        gs.addWidget(QLabel("選擇天數:"), pr,0, alignment=Qt.AlignRight)
        self.cb_days_B = QComboBox(); self.cb_days_B.addItems([str(i) for i in range(1,11)]+["全部平日"])
        self.cb_days_B.setCurrentIndex(0); self.cb_days_B.setFixedWidth(240)
        gs.addWidget(self.cb_days_B, pr,1, alignment=Qt.AlignLeft)

        pr+=1
        gs.addWidget(QLabel("工作內容:"), pr,0, alignment=Qt.AlignRight)
        gbB = QGroupBox(); gbBl = QGridLayout(gbB)
        self.ck_read_B=QCheckBox("閱讀文獻"); self.ck_buy_B=QCheckBox("購買材料")
        self.ck_exp_B =QCheckBox("實驗實做"); self.ck_admin_B=QCheckBox("行政事務")
        for i,w in enumerate([self.ck_read_B,self.ck_buy_B,self.ck_exp_B,self.ck_admin_B]): gbBl.addWidget(w,0,i)
        self.ck_admin_B.hide()  
        self.ck_read_B.setChecked(True)
        gs.addWidget(gbB, pr,1,1,2)

        self.stack.addWidget(self.page_time)   # index 0
        self.stack.addWidget(self.page_simple) # index 1
        self.stack.setCurrentIndex(0)

        # 狀態 & 提交
        r+=1
        self.txt = QTextEdit(); self.txt.setReadOnly(True); self.txt.setFixedHeight(160)
        g.addWidget(self.txt, r,0,1,4)
        r+=1
        self.btn_go = QPushButton("提交")
        g.addWidget(self.btn_go, r,0,1,4, alignment=Qt.AlignHCenter)
        self.btn_go.clicked.connect(self.on_submit)

        self.worker=None

    def set_mode(self, with_time: bool):
        self.mode_with_time = with_time
        self.stack.setCurrentIndex(0 if with_time else 1)

    def log(self, s): self.txt.append(s)

    def on_submit(self):
        username=self.ed_user.text().strip()
        password=self.ed_pass.text().strip()
        if not username or not password:
            QMessageBox.warning(self,"提醒","請輸入帳號與密碼。"); return

        if self.mode_with_time:
            day_choice = self.cb_days_A.currentText()
            contents   = ["行政事務"]   # 固定
            begin=self.ed_begin.text().strip(); end=self.ed_end.text().strip(); hours=self.ed_hours.text().strip()
        else:
            day_choice = self.cb_days_B.currentText()
            contents   = [w.text() for w in [self.ck_read_B,self.ck_buy_B,self.ck_exp_B,self.ck_admin_B] if w.isChecked()]
            begin=end=hours=""

        self.btn_go.setEnabled(False); self.txt.clear(); self.log("開始執行…")
        self.worker = SubmitWorker(username,password,self.mode_with_time,day_choice,begin,end,hours,contents)
        self.worker.status.connect(self.log)
        self.worker.done.connect(self.on_done)
        self.worker.start()

    def on_done(self, summary:str):
        self.log(summary)
        self.btn_go.setEnabled(True)
        self.worker=None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec_())
