#	python ical.py ical_setting.xlsx ical.xlsx

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
#import time
#from datetime import datetime
#from datetime import datetime, timedelta, timezone

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

while True:

    now = datetime.datetime.now()
#    sta = now + datetime.timedelta(days=-5)
#    sto = now + datetime.timedelta(days=21)
    sta = now + datetime.timedelta(days=-3)
    sto = now + datetime.timedelta(days=23)

    df = []
    annots = []
    colors = {}

    first_flg = 0
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
                    #	                summary = ev['summary'].encode('utf-8')
                    summary = ev['summary']
#		            description =  ev['description']
#		            description =  ev.decoded("description")
                except Exception as e:
                    print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                else:
                    #		            print(str(start_dt) + " ~ " + str(end_dt) + " :	" + str(summary))

                    #		            diff = (now - start_dt).total_seconds()

                    #		            if (now - start_dt).total_seconds() > 0 & (now - end_dt).total_seconds() < 0:
                    #		            	print("NOW")
                    #		            time.sleep(13)

                    d = {}
                    d["Task"] = str(df_sig.loc[n]['label'])
#		            d["Task"] = str(summary)
                    d["Start"] = start_dt
                    d["Finish"] = end_dt

                    tmp_summary = str(summary).replace(' ', '')

                    df.append(d)
                    charsize = 20
                    onerowhour = 12  # 　1行の時間巾　文字サイズcharsizeを20とすると12時間（1シフト分）くらい　ブラウザで見た感じ
                    Hdt_N = ((end_dt - start_dt).total_seconds() / 3600) / onerowhour
#		            if Hdt_N < 1:	#  12時間（1シフト分）より短い期間だったら文字サイズを小さくする
#		                charsize=  charsize * Hdt_N

                    print('start_dt	' + str(start_dt))
                    print('end_dt	' + str(end_dt))
                    print('Hdt_N	' + str(Hdt_N))


                    Mojisu = 17  # ＊文字以上なら改行する　Default

                    if Hdt_N !=0:
                        Mojisu = Mojisu/Hdt_N  # 文字が小さかったら、より長い文字数を納められるので
                    else:
                        Hdt_N =1

                    if "Seed" in tmp_summary:
                        print("SEED")
                        tmp_summary += "SEED"

                    tmp_summary = re.sub(
                        "（.+?）", "", tmp_summary)  # カッコで囲まれた部分を消す
                    if len(tmp_summary) > Mojisu:  # ＊文字以上なら改行する
                        tmp_summary = tmp_summary.replace("BL-study", "BL-study<br>")
                        tmp_summary = tmp_summary.replace("BLstudy", "BLstudy<br>")
                        tmp_summary = tmp_summary.replace("G", "G<br>")
                        tmp_summary = tmp_summary.replace("BL調整", "BL調整<br>")

                    tmp_summary = tmp_summary.rstrip('<br>')
                    tmp_summary = tmp_summary.replace("/30Hz", "")
                    tmp_summary = tmp_summary.replace("/60Hz", "")
                    tmp_summary = tmp_summary.replace("SEED", "<i>SEED</i>")

                      

                    """
                    if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                        print("NOW")
#                        tmp_summary = '<b><em>' + tmp_summary + '</em></b>'
#                        tmp_summary = '<span style="font-family:游明朝 Medium;"><em>' + tmp_summary + '</em></span>'
#                        tmp_summary = '<span style="color: #000;text-shadow:1px 1px 0 #FFF, -1px -1px 0 #FFF, -1px 1px 0 #FFF, 1px -1px 0 #FFF, 0px 1px 0 #FFF,  0-1px 0 #FFF, -1px 0 0 #FFF, 1px 0 0 #FFF;">' + tmp_summary + '</span>'                        
#                        tmp_summary = '<span style="color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-shadow:1px 1px 0 #FFF, -1px -1px 0 #FFF, -1px 1px 0 #FFF, 1px -1px 0 #FFF, 0px 1px 0 #FFF,  0-1px 0 #FFF, -1px 0 0 #FFF, 1px 0 0 #FFF;">' + tmp_summary + '</span>'                        
#                        tmp_summary = '<span style="color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-shadow:2px 2px 0 #FFF;">' + tmp_summary + '</span>'                        
#                        tmp_summary = '<span style="color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-shadow: 2px 2px 0 #111;;">' + tmp_summary + '</span>'                  
#                        tmp_summary = '<span style="color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-shadow: 3px 2px 1px rgba(20,20,20,20.3);">' + tmp_summary + '</span>'                  
#                        tmp_summary = '<span style="color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-shadow: 2px 2px 2px #eee, 0px -1px 2px #555;">' + tmp_summary + '</span>'  
#                        tmp_summary = '<b><span style="animation: blinkEffect 1s ease infinite;     color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: -2px -2px 3px #eee, 2px 2px 1px #111, -2px 2px 3px #777, 2px -2px 3px #777;">' + tmp_summary + '</span></b>'  
#                        tmp_summary = '<b><span style="animation: blinkEffect 1s ease infinite;     color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: -2px -2px 3px #fff, 2px 2px 1px #111, -2px 2px 3px #aaa, 2px -2px 3px #aaa;">' + tmp_summary + '</span></b>'  
                        tmp_summary = '<b><span style="animation: blinkEffect 1s ease infinite;     color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 2px 2px 0px #111;">' + tmp_summary + '</span></b>'  
                    else:
                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 0px 0px 2px #000">' + tmp_summary + '</span>'
#                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 1px 1px 0px #111">' + tmp_summary + '</span>'
                    """


                    if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                        print("NOW")
                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: -2px -2px 1px #000, 2px 2px 1px #000, -2px 2px 1px #000, 2px -2px 1px #000;">' + tmp_summary + '</span>'
                    else:
                        tmp_summary = '<span style="font-family:游明朝 Medium; color: ' + str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]) + ';text-decoration: blink;      text-shadow: 0px 0px 2px #000">' + tmp_summary + '</span>'






                    # print(tmp_summary)

                    Row = tmp_summary.count('<br>')+1  # 行数

                    if Hdt_N/Row < 1:  # 12時間（1シフト分）より短い期間だったら文字サイズを小さくする
                        charsize = charsize * Hdt_N/Row
                        tmp_summary = '<b>' + tmp_summary + '</b>'
                    if charsize < 1:
                        charsize = 1

                    """
		            if len(tmp_summary) > 30:	#100文字以上なら文字を小さく
		                charsize=  charsize * 0.5
		                tmp_summary = '<b>' + tmp_summary + '</b>'
		            """

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
#		            da['x'] = start_dt + (end_dt - start_dt)/2
                    if Hdt_N/Row < 1:
                        da['x'] = start_dt + datetime.timedelta(weeks=0, days=0, hours=3*(Row/Hdt_N), minutes=0, seconds=0, milliseconds=0, microseconds=0)
                    else:
                        da['x'] = start_dt + (  (end_dt - start_dt)/2 )
#MOTO                        da['x'] = start_dt + datetime.timedelta(weeks=0, days=0, hours=5*(Row), minutes=0, seconds=0, milliseconds=0, microseconds=0)

                    if str(df_sig.loc[n]['label']) == "リング":
                        da['x'] = start_dt + datetime.timedelta(weeks=0, days=0, hours=8, minutes=0, seconds=0, milliseconds=0, microseconds=0)

                    da['y'] = float(df_sig.loc[n]['annote_y'])


                    try:
                        description = ev['description']
                        tmp_summary = "♦" + tmp_summary   #"<em>★</em>" + tmp_summary
                    except Exception as e:
                        print('No descripton!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')

                    da['text'] = tmp_summary
# DAME	            da['bbox'] = dict(boxstyle="rarrow,pad=0.3", fc="cyan", ec="b", lw=2)
                    da['showarrow'] = False
                    da['textangle'] = -90
#                    da['font'] = dict(size=charsize, family='serif', color=str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]))
                    da['font'] = dict(size=charsize, family='游明朝', color=str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]))
                    if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                        print("NOW")
                        da['textangle'] = -100

#		            if str(df_sig.loc[n]['label'])=="運":
#		                da['textangle'] = -90#-120
#		                da['font'] = dict(size=27,family='serif',color=str(str(df_sig.loc[n]['annote_color']).replace("1","").strip().splitlines()[0]))
#		            family	[ 'serif' | 'sans-serif' | 'cursive' | 'fantasy' | 'monospace'

                    tmp_summary_before = tmp_summary

                    annots.append(da)

                    da = {}
                    try:
                        description = ev['description']
                    except Exception as e:
                        print('No descripton!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                    else:
                        #print('descripton OK	')
                        da['x'] = start_dt + \
                            (end_dt - start_dt) - (end_dt - start_dt)/4
                        da['y'] = float(df_sig.loc[n]['annote_y'])
                        da['text'] = "<i>" + str(description) + "</i>"
                        da['showarrow'] = False  # True
                        da['textangle'] = -90
                        da['font'] = dict(color=str(
                            str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]))
                        # annots.append(da)

#		            print("-------------------------------------------" + summary)
#		            print("-------------------------------------------" + colors[summary])
            m += 1
#		    print("m = -------------------------------------------" + str(m))

##############################
            da = {}
            da['x'] = now + datetime.timedelta(days=-3)
            da['y'] = float(df_sig.loc[n]['annote_y'])
            da['text'] = str(df_sig.loc[n]['label'])
            da['showarrow'] = False  # True
            da['textangle'] = -90
            da['bgcolor'] = "#000000"
            da['font'] = dict(size=37, family='serif', color=str(str(df_sig.loc[n]['label_color']).replace("1", "").strip().splitlines()[0]))
            annots.append(da)

#            da = {}
#            da['x'] = now + datetime.timedelta(days=-3.0)
#            da['y'] = -0.7
#            da['text'] = '♦印は詳細アリ'
 #           da['showarrow'] = False  # True
 #           da['textangle'] = -90
 #           da['font'] = dict(size=17, family='serif', color=str(str(df_sig.loc[n]['label_color']).replace("1", "").strip().splitlines()[0]))
 #           annots.append(da)

            """
            da = {}
            da['x'] = now + datetime.timedelta(days=-0.023)
            da['y'] = -0.7
            da['text'] = ">" #"<em>></em>"
            da['showarrow'] = False  # True
            da['textangle'] = -90
            da['font'] = dict(size=45, family='serif', color="yellow")
            annots.append(da)

            da = {}
            da['x'] = now + datetime.timedelta(days=-0.023)
            da['y'] = 3.7
            da['text'] = ">"#"<em><</em>"
            da['showarrow'] = False  # True
            da['textangle'] = -90
            da['font'] = dict(size=45, family='serif', color="yellow")
            annots.append(da)
            """
            da = {}
            da['x'] = now + datetime.timedelta(days=0.2)
            da['y'] = 3.0
            da['text'] = now.strftime('%m/%d %H:%M')
            da['showarrow'] = False  # True
            da['textangle'] = -90
            da['font'] = dict(size=8, family='serif', color="black")
            annots.append(da)

            da = {}
            da['x'] = now + datetime.timedelta(days=0.15)
            da['y'] = 0.7
#            da['text'] = '<span style="opacity: 0.8;">‣‣‣‣‣‣‣..............................................................................</span>'
#            da['text'] = '<span style="opacity: 0.8;">‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣>‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣>‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣</span>'

#            da['text'] = '<span style="opacity: 0.8;">‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣-・・・・・・・・・・・・・・・・・・・・・・・・・</span>'
#            da['text'] = '<span style="opacity: 0.8;">‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣ ・・・・・・・・・・・・・・・・・・・・・・・・・</span>'
#            da['text'] = '<span style="opacity: 0.8;">‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣- ・・・・・・・・・・・・・・・・・・・・・・・・・</span>'

#            da['text'] = '<span style="opacity: 0.8;"> > ‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣- ・・・・・・・・・・・・・・・・・・・・・・・・・</span>'

            da['text'] = '<span style="opacity: 0.8;">⋆ ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆ </span>'
#            da['text'] = '<span style="opacity: 0.8;">　　　　　　　　　　　　　本日 ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  </span>'
#            da['text'] = '<span style="opacity: 0.8;">||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||</span>'
#            da['text'] = '<span style="opacity: 0.8;">| | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | </span>'
#            da['text'] = '<span style="opacity: 0.8;">　　　　　　　　　　　　　☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  </span>'
#            da['text'] = '<span style="opacity: 0.8;">　　　　　　　　　　　　　★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ★ ☆ ★ ☆ ★ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ☆ ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  ⋆  </span>'


#            da['text'] = '<span style="opacity: 0.8;"> > ‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣- ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼</span>'
#            da['text'] = '<span style="font-size : 8pt";">▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼  ▼ </span>'

#            da['text'] = '<span style="opacity: 0.8;"> > ‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣                                              </span>'


#            da['text'] = '‣ ‣ ‣ ‣ ‣ ‣ ‣ ‣               '
#            str_tmp = dt.strftime('%a')
#            print('str_tmp =        ' + str_tmp)
#            da['text'] = '‣ ‣ ‣ ‣ ‣ ‣ ‣ ' + str_tmp + '               ' 

#dt.strftime('%A, %a, %B, %b')

#            da['text'] = '<span style="opacity: 0.8;">‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣‣-</span>'
            da['showarrow'] =  True  #False
            da['textangle'] = -90
#            da['font'] = dict(size=31, family='monospace', color="yellow")
            da['font'] = dict(size=20, family='monospace', color="yellow")
#            da['font'] = dict(size=15, family='monospace', color="yellow")
            annots.append(da)


        """
		d = {'Task':str(df_sig.loc[n]['label']), 'Start':sta, 'Finish':sto, 'Resource':'Marker'}
		df.append(d)
		if first_flg==0:
			colors['Marker'] = '#%02X%02X%02X' % (255,0,n)
			first_flg=1					
		"""
        print("-------------------------------------------")

#		print(df)
#		print("-------------------------------------------")
#		print(colors)


#	fig = ff.create_gantt-group-tasks-together(df, colors=colors, index_col='Resource', title='Schedule',
#                      show_colorbar=False, bar_width=0.495, width=1300, height=600, showgrid_x=True, showgrid_y=False, group_tasks=True)
    fig = ff.create_gantt(df, colors=colors, index_col='Resource', title='Schedule',
                          show_colorbar=False, bar_width=0.495, width=1550, height=850, showgrid_x=True, showgrid_y=False, group_tasks=True)

#	fig = ff.create_gantt(df, colors=colors, index_col='Resource', title='Schedule',
#                      show_colorbar=False, bar_width=0.5, width=1500, showgrid_x=True, showgrid_y=False, group_tasks=True)

#	print(annots)
    fig['layout']['annotations'] = annots

#	OK
    fig['layout'].update(xaxis=dict(tickformat="%_m/%-d %a", tick0='2022-7-01 10:00:00',
                         tickmode='linear', dtick=24 * 60 * 60 * 1000, tickcolor="gray", tickwidth=0.1))

    fig.update_xaxes(
        showgrid=True,
        tickangle=270,
        ticks="inside",  # ticks="outside",
        tickson="boundaries",
        tickwidth=0.0001,
        tickcolor='dimgrey',
        ticklen=1120,
        tickfont=dict(size=30),
        # rangeslider_visible=True
    )

#	OK?
    """
	fig.update_layout(xaxis={'domain': [0, 1],
                             'mirror': True,
                             'showgrid': True,
                             'showline': True,
                             'zeroline': False,
                             'showticklabels': True,
                             'ticks':""})
	"""


#	shiftNum = 0 - now.weekday() 	#月曜日が0で日曜日が6	0は目標とする曜日でMondya曜日の意味。
    shiftNum = 1 - now.weekday()  # 月曜日が0で日曜日が6	1は目標とする曜日でTuesday曜日の意味。
    print('now.weekday()   ' + str(now.weekday()))
    print('shiftNum   ' + str(shiftNum))
    shiftNum = shiftNum+7 if shiftNum < 0 else shiftNum
    print('shiftNum   ' + str(shiftNum))
    print('Next Monday  ' + str(shiftNum))
    next = now+datetime.timedelta(weeks=0, days=shiftNum, hours=0,
                                  minutes=0, seconds=0, milliseconds=0, microseconds=0)
    print('next   ' + str(next))
    next = datetime.datetime(next.year, next.month, next.day, 10, 0, 0)
    print('next   ' + str(next))

    fig.update_layout(shapes=[
#        dict(type='line', yref='paper', y0=-1, y1=1, xref='x', x0=now, x1=now,
#             fillcolor="black", opacity=0.5, line=dict(color="yellow", width=1, dash="solid")),

#        dict(type='line', yref='paper', y0=0, y1=1, xref='x', x0=now, x1=now,
#             fillcolor="greenyellow", opacity=0.5, line=dict(color="yellow", width=5, dash="dot")),

        dict(type='line', yref='paper', y0=-0.01, y1=1.01, xref='x', x0=next+datetime.timedelta(weeks=0, days=-7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), x1=next +
             datetime.timedelta(weeks=0, days=-7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), fillcolor="greenyellow", opacity=1.0, line=dict(color="yellow", width=3, dash="solid")),
        dict(type='line', yref='paper', y0=-0.01, y1=1.01, xref='x', x0=next, x1=next,
             fillcolor="greenyellow", opacity=1.0, line=dict(color="yellow", width=3, dash="solid")),
        dict(type='line', yref='paper', y0=-0.01, y1=1.01, xref='x', x0=next+datetime.timedelta(weeks=0, days=7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), x1=next +
             datetime.timedelta(weeks=0, days=7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), fillcolor="greenyellow", opacity=1.0, line=dict(color="yellow", width=3, dash="solid")),
        dict(type='line', yref='paper', y0=-0.01, y1=1.01, xref='x', x0=next+datetime.timedelta(weeks=0, days=14, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), x1=next +
             datetime.timedelta(weeks=0, days=14, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), fillcolor="greenyellow", opacity=1.0, line=dict(color="yellow", width=3, dash="solid")),

        #	    dict(type= 'line', yref= 'paper', y0= 0, y1= 1, xref= 'x', x0= next+datetime.timedelta(weeks=0, days=-7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), x1= next+datetime.timedelta(weeks=0, days=-7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), fillcolor="greenyellow", opacity=1.0, line=dict(color="gray", width=1, dash="dot")),
        #	    dict(type= 'line', yref= 'paper', y0= 0, y1= 1, xref= 'x', x0= next, x1= next, fillcolor="greenyellow", opacity=1.0, line=dict(color="gray", width=1, dash="dot")),
        #	    dict(type= 'line', yref= 'paper', y0= 0, y1= 1, xref= 'x', x0= next+datetime.timedelta(weeks=0, days=7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), x1= next+datetime.timedelta(weeks=0, days=7, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), fillcolor="greenyellow", opacity=1.0, line=dict(color="gray", width=1, dash="dot")),
        #	    dict(type= 'line', yref= 'paper', y0= 0, y1= 1, xref= 'x', x0= next+datetime.timedelta(weeks=0, days=14, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), x1= next+datetime.timedelta(weeks=0, days=14, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0), fillcolor="greenyellow", opacity=1.0, line=dict(color="gray", width=1, dash="dot")),

        #	    dict(type= 'line', yref= 'paper', y0= 0, y1= 1, xref= 'x', x0= now + datetime.timedelta(days=7), x1= now + datetime.timedelta(days=7), fillcolor="gray", opacity=0.5 )
    ],
        #	template='plotly_dark',
        margin=dict(r=1, t=1, b=10, l=1)
    )

    fig.update_xaxes(range=[sta, sto])
    fig.update_yaxes(range=[-0.7, 3.7])


#	fig['layout'].update( xaxis = dict( tickformat="%d %B(%a)", tickmode = 'linear', dtick = 24 * 60 * 60 * 1000 ))
#	fig['layout'].update( xaxis = dict( tickformat="%m/%d", tickmode = 'linear', dtick = 604800000 ) )

#	fig['layout'].update(autosize=True)
#	fig['layout'].update(autosize=False, margin=go.Margin(l=0, b=100), xaxis=dict(tickformat="%d-%m-%Y", autotick=False, tick0=-259200000, dtick=604800000))
#	fig['layout'].update(autosize=False, margin=go.Margin(l=0, r=0, b=50))

    """
	axes = plt.gcf().get_axes()
	for axis in axes:
		plt.axes(axis)
		print('### Updated	###  '  + str(axis))
	"""
    plotly.offline.plot(
        fig, filename='gantt-group-tasks-together.html', auto_open=False)

    print('### Updated END	###  ')

    src = 'C:\me\ical_to_ganto\gantt-group-tasks-together.html'
    if os.path.isfile(src):
        print('Sonzai')
#		copy = 'C:/me/test.html'
        copy = '//saclaoprfs01.spring8.or.jp/log_note/calendar/gantt-group-tasks-together.html'
        shutil.copyfile(src, copy)
#        webbrowser.open('http://saclaopr19.spring8.or.jp/~lognote/calendar/gantt-group-tasks-together.html')
    else:
        print('Not Sonzai')

    print("---------------------------")
    print(df_set.loc['interval'][1])
    print("---------------------------")
    time.sleep(int(df_set.loc['interval'][1]))
#    time.sleep(df_set.loc['interval'][1].astype(int))
#    break


"""MEMO	Plotly
fig.add_annotation(
        x=2,
        y=5,
        xref="x",
        yref="y",
        text="max=5",
        showarrow=True,
        font=dict(
            family="Courier New, monospace",
            size=16,
            color="#ffffff"
            ),
        align="center",
        arrowhead=2,
        arrowsize=1,
        arrowwidth=2,
        arrowcolor="#636363",
        ax=20,
        ay=-30,
        bordercolor="#c7c7c7",
        borderwidth=2,
        borderpad=4,
        bgcolor="#ff7f0e",
        opacity=0.8
        )

"""
