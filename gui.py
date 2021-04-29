#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import PySimpleGUI as sg
import browser_action as ba
import importlib
importlib.reload(ba)

#excelからリンク等を読み込み()
S=ba.excel_read(2,2)

#  セクション1 - オプションの設定と標準レイアウト
sg.theme('Dark Blue 3')

# idとパスはおまけ
layout = [
    [sg.Text("対象ページよりリンク取得")],
    [sg.Text("link", size=(15, 1)), sg.InputText(S[0])],
    [sg.Text("ログインid", size=(15, 1)), sg.InputText(S[1])],
    [sg.Text("パスワード", size=(15, 1)), sg.InputText(S[2], password_char="●")],
    [sg.Text("抜き出し条件", size=(15, 1)), sg.InputText(S[3])],
    [sg.Text("書き込み開始行", size=(15, 1)), sg.InputText(S[4])],
    [sg.Text("書き込み開始列", size=(15, 1)), sg.InputText(S[5])],
    [sg.Submit(button_text="実行",key="p1")],
    [sg.Text("リンク集を用いてデータ収集")],
    [sg.Text("読み込み開始行", size=(15, 1)), sg.InputText(S[4])],
    [sg.Text("読み込み開始列", size=(15, 1)), sg.InputText(S[5])],
    [sg.Text("抜き出し条件", size=(15, 1)), sg.InputText(S[6])],
    [sg.Text("書き込み開始行", size=(15, 1)), sg.InputText(S[4])],
    [sg.Text("書き込み開始列", size=(15, 1)), sg.InputText(int(S[5])+1)],
    [sg.Submit(button_text="実行",key="p2")]
]

# セクション 2 - ウィンドウの生成
window = sg.Window("スクレイピングテスト", layout)

# セクション 3 - イベントループ
while True:
    event, values = window.read()

    if event is None:
        print("exit")
        break

    if event == "p1":
        links=ba.get_link(values[0],values[3])
        ba.excel_links_write(links,values[4],values[5])
        print("リンク収集完了")
        
    if event == "p2":
        links2=ba.excel_read(values[6],values[7])
        cnt=0
        for link in links2:
            items=ba.get_item(link,values[8])
            ba.excel_items_write(items,values[9],values[10])
        
        
        print("データ収集完了")
# セクション 4 - ウィンドウの破棄と終了
window.close()


# In[ ]:





# In[ ]:




