import win32com.client as win32
import datetime
from datetime import timedelta


# --------------------------------------------------------------------   
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = #'要寄信的人或信箱'
mail.CC = #'要副本的人'
mail.Subject = #'信件標題'
mail.Body = #'信件內文'
mail.Attachments.Add(#'要附的檔案')
mail.Send() #寄送信件 

# 加入排程的話 就能自動發信了