#!/usr/bin/env python
# coding: utf-8


import sys
import time
import tkinter as tk
from tkinter import ttk

import pyautogui as gui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

####################################################################################################
# setting
groupId = "********"  # アカウント管理グループID
organization = "**県**市"  # 組織名
rooms = 30  # 退出する部屋数
mail_domain = "@****.******.lg.jp"  # メールアドレスのドメイン
url_login = "https://tb.logochat.st-japan.asp.lgwan.jp/signin"  # ログインページのURL
# url_login = "https://logochat.jp/signin" #インターネット側はこちら
####################################################################################################


####################################################################################################
# GUI
def submit():
    global username, password, groupId, organization, rooms
    username = entry_username.get()
    password = entry_password.get()
    groupId = entry_groupId.get()
    organization = entry_organization.get()
    rooms = room_var.get()
    root.destroy()


def password_appear():
    if bool_check.get() == True:
        entry_password["show"] = ""
    else:
        entry_password["show"] = "*"


root = tk.Tk()
root.title("LoGoチャット自動退出")
root.geometry("630x300")

label_username = tk.Label(root, text="メールアドレスまたはログインID:")
label_username.grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
entry_username = tk.Entry(root, width=50, justify="center")
entry_username.insert(0, mail_domain)
entry_username.grid(row=0, column=1, padx=10, pady=5)

label_password = tk.Label(root, text="パスワード:")
label_password.grid(row=1, column=0, padx=10, pady=5, sticky=tk.E)
entry_password = tk.Entry(root, width=50, justify="center", show="*")
entry_password.grid(row=1, column=1, padx=10, pady=5)

bool_check = tk.BooleanVar()
bool_check.set(False)
check = tk.Checkbutton(
    variable=bool_check, command=password_appear, text="パスワードを表示"
)
check.grid(row=2, column=1, padx=10, pady=5)

label_groupId = tk.Label(root, text="アカウント管理グループID:")
label_groupId.grid(row=3, column=0, padx=10, pady=5, sticky=tk.E)
entry_groupId = tk.Entry(root, width=50, justify="center")
entry_groupId.insert(0, groupId)
entry_groupId.grid(row=3, column=1, padx=10, pady=5)

label_rooms = tk.Label(root, text="退出する部屋数:")
label_rooms.grid(row=4, column=0, padx=10, pady=5, sticky=tk.E)
room_var = tk.IntVar()
room_var.set(rooms)
rooms_menu = ttk.Combobox(
    root, textvariable=room_var, values=[i for i in range(10, 301, 10)], width=10
)
rooms_menu.grid(row=4, column=1, padx=10, pady=5, sticky=tk.W)

label_organization = tk.Label(root, text="組織名:")
label_organization.grid(row=5, column=0, padx=10, pady=5, sticky=tk.E)
entry_organization = tk.Entry(root, width=50, justify="center")
entry_organization.insert(0, "奈良県橿原市")
entry_organization.grid(row=5, column=1, padx=10, pady=5)

submit_button = tk.Button(root, text="処理開始", command=submit)
submit_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()


####################################################################################################
# ブラウザ操作

# GoogleChromeを起動
chrome_options = webdriver.ChromeOptions()
driver_path = "chromedriver.exe"
service = Service(executable_path=driver_path)
driver = webdriver.Chrome(service=service)
driver.implicitly_wait(20)

# ログインページにアクセス
driver.get(url_login)
driver.maximize_window()

# ログイン
elem_userid = driver.find_element(By.NAME, "email_or_signin_id")
elem_userpass = driver.find_element(By.NAME, "password")
elem_groupid = driver.find_element(By.NAME, "group_alias")
elem_login_button = driver.find_element(By.ID, "btn_signin")
elem_groupid.send_keys(groupId)
elem_userid.send_keys(username)
elem_userpass.send_keys(password)
elem_login_button.click()
time.sleep(4)

# 組織選択
elem_soshiki = driver.find_element(By.CLASS_NAME, "navbar-header")
elem_soshiki.click()
time.sleep(2)
elem_organization = driver.find_element(
    By.XPATH, f"//span[contains(text(),'{organization}')]"
)
elem_organization.click()
time.sleep(4)

# トーク一覧表示
elem_talksearch = driver.find_element(By.ID, "talk-searchbox")
elem_talksearch.click()

for i in range(4):
    gui.press("tab")

for i in range(25):
    time.sleep(0.5)
    gui.press("end")

talks = driver.find_element(By.ID, "talks")
talks = talks.find_elements(By.TAG_NAME, "li")

# トークIDを全取得
talkid = []
for talk in talks:
    talkid.append(talk.get_attribute("id"))
talkid.reverse()

# トークから退出
for x in range(int(rooms)):
    exit_room = driver.find_element(By.ID, talkid[x])
    exit_room.click()

    talk_menu = driver.find_element(By.XPATH, "//span[@aria-label='トークメニュー']")
    talk_menu.click()

    exit_btn = driver.find_element(By.XPATH, "//span[text()='トークから退出']")
    exit_btn.click()

    confirm_ok_btn = driver.find_element(By.ID, "confirm-ok")
    confirm_ok_btn.click()
    time.sleep(1)


driver.quit()

sys.exit(0)
