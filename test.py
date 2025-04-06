#	python test.py ical_setting.xlsx ical_TEST.xlsx

#	ical.pywにする時、一番下の二行をコメントアウトする！
#	webbrowser.open('http://saclaopr19.spring8.or.jp/~lognote/calendar/gantt-group-tasks-together.html')
#	break

# Formatter     Shift+Alt+F

import requests
from requests.exceptions import Timeout
import re
import pandas as pd
import sys
from matplotlib.dates import DateFormatter
from icalendar import Calendar, Event

import datetime

import plotly.figure_factory as ff
import plotly
import random
import os
import shutil

import webbrowser
import time

""" Japanese"""
import locale
dt = datetime.datetime(2018, 1, 1)
print(locale.getlocale(locale.LC_TIME))
print(dt.strftime('%A, %a, %B, %b'))
locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
print(locale.getlocale(locale.LC_TIME))
print(dt.strftime('%A, %a, %B, %b'))
"""------"""



args = sys.argv
print("arg1:" + args[1])
config_file_setting = args[1]
config_file_sig = args[2]

df_set = pd.read_excel(config_file_setting,
                       sheet_name="setting", header=None, index_col=0)
# print(df_set)
df_sig = pd.read_excel(config_file_sig, sheet_name="sig")
# print(df_sig)


"""  -------------------------------------------------------------------------------------  """


def get_acc_sync(url):

    # print(url)
    try:
        res = requests.get(url, timeout=(30.0, 30.0))
    except Exception as e:
        #print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!@get_acc_sync	' + url)
        print(e.args)
        return ''
    else:
        res.raise_for_status()
        return res.text


class SigInfo:
    def __init__(self):
        self.srv = ''
        self.url = ''
        self.sname = ''
        self.sid = 0
        self.sta = ''
        self.sto = ''
        self.time = ''
        self.val = ''
        self.sortedval = []
        self.rave = []
        self.rave_sigma = []
        self.d = {}
        self.t = {}
        self.mu = 0
        self.icaldata = ''
        self.sigma = 0


sig = [SigInfo() for _ in range(len(df_sig))]


tmp_summary_before = "test"

JST = datetime.timezone(datetime.timedelta(hours=+9), 'JST')

now = datetime.datetime.now()
dt1 = now + datetime.timedelta(days=-3)
dt2 = now + datetime.timedelta(days=23)

list_dt = []
list_dt.append(now)
list_dt.append(dt1)
list_dt.append(dt2)


while True:



    df = []
    colors = {}
    
    

    for n, s in enumerate(sig, 0):
        s.icaldata = get_acc_sync(str(df_sig.loc[n]['url']))
        # print(s.icaldata)
        cal = Calendar.from_ical(s.icaldata)
        m = 0
        for ev in cal.walk():

            try:
                start_dt_datetime = datetime.datetime.strptime(
                    str(ev.decoded("dtstart")), '%Y-%m-%d %H:%M:%S+09:00')
            except Exception as e:
                # print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                continue
            else:
                if (start_dt_datetime - now).days < -60:
                    continue
#		        else:
#			        print('(start_dt_datetime - now).days	' + str((start_dt_datetime - now).days))

            if ev.name == 'VEVENT':
                start_dt = ev.decoded("dtstart")
                end_dt = ev.decoded("dtend")
                try:
                    summary = ev['summary']
                except Exception as e:
                    print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                else:
                    d = {}
                    d["Task"] = str(df_sig.loc[n]['label'])
                    d["Start"] = start_dt
                    d["Finish"] = end_dt

                    tmp_summary = str(summary).replace(' ', '')

                    df.append(d)

                    print('start_dt	' + str(start_dt))
                    print('end_dt	' + str(end_dt))

                    tmp_summary = re.sub("（.+?）", "", tmp_summary)  # カッコで囲まれた部分を消す

                    tmp_summary = tmp_summary.rstrip('<br>')
                    tmp_summary = tmp_summary.replace("/30Hz", "")
                    tmp_summary = tmp_summary.replace("/60Hz", "")
                    tmp_summary = tmp_summary.replace("SEED", "<i>SEED</i>")

                    if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                        print("NOW", tmp_summary)
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: -2px -2px 1px #000, 2px 2px 1px #000, -2px 2px 1px #000, 2px -2px 1px #000;">' + tmp_summary + '</span>'
                    else:
                        print("Not NOW", tmp_summary)                        
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 0px 0px 2px #000">' + tmp_summary + '</span>'

                    Row = tmp_summary.count('<br>')+1  # 行数

                    if "BL" in tmp_summary:
                        print("DUMMY :	"+tmp_summary)
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 0px 0px 2px #000">' + tmp_summary + '</span>'
                    elif "加速器調整" in tmp_summary:
                        charsize = 21
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 0px 0px 2px #000">' + tmp_summary + '</span>'
                    elif str(df_sig.loc[n]['label']) == "運":
                        tmp_summary = tmp_summary.replace("・", "/")
                        charsize = 27
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: -2px -2px 1px #000, 2px 2px 1px #000, -2px 2px 1px #000, 2px -2px 1px #000;">' + tmp_summary + '</span>'
                    elif str(df_sig.loc[n]['label']) == "リング":
                        tmp_summary = tmp_summary.replace("(Ring)", "")
                        tmp_summary = tmp_summary.replace("変更", "変更<br>")
                        charsize = 15
                    else:  # User
                        print("DUMMY :	"+tmp_summary)
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 0px 0px 2px #000">' + tmp_summary + '</span>'

                    print(str(start_dt) + " ~ " +
                          str(end_dt) + " :	" + tmp_summary)
                    d["Resource"] = tmp_summary
                    d["Complete"] = n  # str(summary)


                    if "BL-study" in tmp_summary:
                        colors[tmp_summary] = '#%02X%02X%02X' % (random.randint(50, 50), random.randint(10, 10), 255)
                    elif "BL調整" in tmp_summary:
                        colors[tmp_summary] = '#%02X%02X%02X' % (random.randint(50, 50), random.randint(50, 50), 255)
#                        colors[tmp_summary] = '#%02X%02X%02X' % (random.randint(50, 50), random.randint(120, 120), 255)
                    elif "加速器調整" in tmp_summary:
                        colors[tmp_summary] = '#%02X%02X%02X' % (130, 130, 130)  # '#%02X%02X%02X' % (100, 100, 100)
                    elif str(df_sig.loc[n]['label']) == "運":
                        colors[tmp_summary] = '#%02X%02X%02X' % (0,0,0)
                    elif str(df_sig.loc[n]['label']) == "リング":
                        colors[tmp_summary] = '#%02X%02X%02X' % (130,130,130)
                    else:  # User
                        # colors[tmp_summary] = '#%02X%02X%02X' % (255, random.randint(0, 10), random.randint(0, 10))
                        colors[tmp_summary] = '#%02X%02X%02X' % (
                            205, random.randint(1, 1), random.randint(7, 7))

                    da = {}

                    if str(df_sig.loc[n]['label']) == "リング":
                        da['x'] = start_dt + datetime.timedelta(weeks=0, days=0, hours=8, minutes=0, seconds=0, milliseconds=0, microseconds=0)

                    da['y'] = float(df_sig.loc[n]['annote_y'])


                    da['text'] = tmp_summary
                    if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                        print("NOW")

                    tmp_summary_before = tmp_summary

#		            print("-------------------------------------------" + summary)
#		            print("-------------------------------------------" + colors[summary])
            m += 1
