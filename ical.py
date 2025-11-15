#	python ical.py ical_setting.xlsx ical_SHISETUCHOUSEI.xlsx
#   
#   -u オプションを付けると「運転集計用に表示する範囲をユニットの開始終了にした」　が、ローカルに置いたHTMLファイルではブラウザ上でjavascriptを実行してくれる拡張機能「Tampermonky」が動いてくれないので、画像にしてから回転させる処理を入れた。
#   
#	ical.pywにする時、一番下の二行をコメントアウトする！

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
import argparse

import pytz
def safe_strptime(str_dt):
    """
    日時（タイムゾーン付き）または日付のみの文字列をdatetime型に安全に変換する。
    日付のみの場合、時刻は 00:00:00、タイムゾーンは JST (+09:00) を設定する。
    """
    str_dt = str(str_dt)
    
    # タイムゾーンの設定
    tokyo_tz = pytz.timezone('Asia/Tokyo')
    
    # 1. タイムゾーン付きの日時フォーマットで試行
    format_full = '%Y-%m-%d %H:%M:%S%z'
    try:
        # 成功した場合、既存のタイムゾーン情報を持つ datetime オブジェクトを返す
        dt_object = datetime.datetime.strptime(str_dt, format_full)
        return dt_object
    
    except ValueError:
        # 2. 失敗した場合、日付のみのフォーマットで試行
        format_date_only = '%Y-%m-%d'
        try:
            # 日付のみとしてパース。時刻は自動的に 00:00:00 になる (ここが要求通り)
            dt_object = datetime.datetime.strptime(str_dt, format_date_only)
            
            # JST (+09:00) のタイムゾーン情報を付与
            dt_object_tz = tokyo_tz.localize(dt_object)
            
            return dt_object_tz
        
        except ValueError as e:
            # 3. どちらのフォーマットでも失敗した場合
            print(f"エラー: '{str_dt}' は指定されたどのフォーマットにも一致しません。")
            raise e


""" Japanese"""
import locale
dt = datetime.datetime(2018, 1, 1)
print(locale.getlocale(locale.LC_TIME))
print(dt.strftime('%A, %a, %B, %b'))
locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
print(locale.getlocale(locale.LC_TIME))
print(dt.strftime('%A, %a, %B, %b'))



##################################################
parser = argparse.ArgumentParser(description='ファイル処理のプログラム。')
# 2. 引数の追加
# 位置引数 (必須)
parser.add_argument('config_file_setting', 
                    help='入力として使用するファイルパスを指定します。')
parser.add_argument('config_file_sig', 
                    help='入力として使用するファイルパスを指定します。')
parser.add_argument('-v', '--verbose', 
                    action='store_true', 
                    help='詳細な処理情報を出力します。')
parser.add_argument('-u', '--unten', 
                    action='store_true', 
                    help='運転集計用に、出力するユニットの期間を表示します。')
parser.add_argument('--limit', 
                    type=int, 
                    default=10, 
                    help='テスト')
args = parser.parse_args()
if args.verbose:
    print("✅ 詳細モード (verbose) が有効です。")
else:
    print("❌ 標準モードで実行します。")
if args.unten:
    print("✅ 運転集計モード (unten) が有効です。")
else:
    print("❌ 標準モードで実行します。")    
print(f"📘 入力ファイル1: {args.config_file_setting}")
print(f"📘 入力ファイル2: {args.config_file_sig}")
print(f"🔢 処理制限数: {args.limit}")
##################################################

df_set = pd.read_excel(args.config_file_setting,
                       sheet_name="setting", header=None, index_col=0)
# print(df_set)
df_sig = pd.read_excel(args.config_file_sig, sheet_name="sig")
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

    if args.unten:
        print("✅ untenモードが有効です。")
        with open(r"C:\me\unten\OperationSummary\dt_beg.txt", mode='r', encoding="UTF-8") as f:
            buff_dt_beg = f.read()
        with open(r"C:\me\unten\OperationSummary\dt_end.txt", mode='r', encoding="UTF-8") as f:
            buff_dt_end = f.read()
        sta = datetime.datetime.strptime(buff_dt_beg, "%Y/%m/%d %H:%M")
        sta = sta +  datetime.timedelta(days=-1)        
        sto = datetime.datetime.strptime(buff_dt_end, "%Y/%m/%d %H:%M")
        sto = sto +  datetime.timedelta(days=2)
    else:
        print("❌ 標準モードで実行します。")
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
#                start_dt_datetime = datetime.datetime.strptime(
#                    str(ev.decoded("dtstart")), '%Y-%m-%d %H:%M:%S+09:00')
                start_dt_datetime = safe_strptime(ev.decoded("dtstart")).replace(tzinfo=None) # replace(tzinfo=None) でタイムゾーン情報を削除
            except Exception as e:
                print('Exception@A  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
                continue
            
#                print(f"変換された日時: {start_dt_datetime} ({type(start_dt_datetime)})")
#                print(f"sta: {sta} ({type(sta)})")
            print(f"正常にパースされた日時: {start_dt_datetime}")
            if (start_dt_datetime - sta).days < -30:#  余裕をもってstaの30日前より前のイベントから表示する
                continue

            if ev.name == 'VEVENT':
#                start_dt = ev.decoded("dtstart")
#                end_dt = ev.decoded("dtend")
                start_dt = safe_strptime(ev.decoded("dtstart")).replace(tzinfo=None) # replace(tzinfo=None) でタイムゾーン情報を削除
                end_dt = safe_strptime(ev.decoded("dtend")).replace(tzinfo=None) # replace(tzinfo=None) でタイムゾーン情報を削除

                print(f"DEBUG:  start_dt: {start_dt} ({type(start_dt)})")
                try:
                    #	                summary = ev['summary'].encode('utf-8')
                    summary = ev['summary']
#		            description =  ev['description']
#		            description =  ev.decoded("description")
                except Exception as e:
                    print('Exception@B  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
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

#                    print('start_dt	' + str(start_dt))
#                    print('end_dt	' + str(end_dt))
 #                   print('Hdt_N	' + str(Hdt_N))

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

#                    if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                    if (now - start_dt).total_seconds() > 0 and (now - end_dt).total_seconds() < 0:
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

                    if "BL" in tmp_summary:
                        print("", end="")
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

                    print(str(start_dt) + " ~ " + str(end_dt) + " :	" + tmp_summary)
                    d["Resource"] = tmp_summary
                    d["Complete"] = n  # str(summary)

                    if str(df_sig.loc[n]['label']) == "運":
                        colors[tmp_summary] = '#%02X%02X%02X' % (0,0,0) # 運は表示されない。ical.xlsxの下(SCSS+)の方から順に表示され、ギリギリ施設調整が見える
                    elif str(df_sig.loc[n]['label']) == "リング":
                        colors[tmp_summary] =   '#%02X%02X%02X' % (130,130,130)
                    elif str(df_sig.loc[n]['label']) == "施設調整":
                        colors[tmp_summary] = '#%02X%02X%02X' % (200,127,80) 
                    elif "BL-study" in tmp_summary:
                        colors[tmp_summary] = '#%02X%02X%02X' % (random.randint(50, 50), random.randint(10, 10), 255)
                    elif "BL調整" in tmp_summary:
                        colors[tmp_summary] = '#%02X%02X%02X' % (random.randint(50, 50), random.randint(50, 50), 255)
                    elif "加速器調整" in tmp_summary:
                        colors[tmp_summary] = '#%02X%02X%02X' % (130, 130, 130)
                    else:  # User
                        colors[tmp_summary] = '#%02X%02X%02X' % (
                            205, random.randint(1, 1), random.randint(7, 7))
                                            
                    da = {} # tmp_summary を表示する位置を微調整
                    if Hdt_N/Row < 1:
                        da['x'] = start_dt + datetime.timedelta(weeks=0, days=0, hours=3*(Row/Hdt_N), minutes=0, seconds=0, milliseconds=0, microseconds=0)
                    else:
                        da['x'] = start_dt + (  (end_dt - start_dt)/2 )

                    if str(df_sig.loc[n]['label']) == "リング":
                        da['x'] = start_dt + datetime.timedelta(weeks=0, days=0, hours=8, minutes=0, seconds=0, milliseconds=0, microseconds=0)

                    da['y'] = float(df_sig.loc[n]['annote_y'])


                    try:
                        description = ev['description']
                        tmp_summary = "♦" + tmp_summary   #"<em>★</em>" + tmp_summary
                    except Exception as e:
                        print('', end="")

                    da['text'] = tmp_summary
# DAME	            da['bbox'] = dict(boxstyle="rarrow,pad=0.3", fc="cyan", ec="b", lw=2)
                    da['showarrow'] = False
                    da['textangle'] = -90
#                    da['font'] = dict(size=charsize, family='serif', color=str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]))
                    da['font'] = dict(size=charsize, family='游明朝', color=str(str(df_sig.loc[n]['annote_color']).replace("1", "").strip().splitlines()[0]))
                    #if (now.astimezone(JST) - start_dt).total_seconds() > 0 and (now.astimezone(JST) - end_dt).total_seconds() < 0:
                    if (now - start_dt).total_seconds() > 0 and (now - end_dt).total_seconds() < 0:
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
                        print('', end="")
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
#            print("m = -------------------------------------------" + str(m))

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
#        if n==1: os._exit(0)
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


    shiftNum = -1 - now.weekday() 	#月曜日が0で日曜日が6	0は目標とする曜日でMondya曜日の意味。 なぜか-1しないといけない。。
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

#===  一週間おきに黄色い線を付ける  ===================================================
    line_style = dict(color="yellow", width=3, dash="solid")
    shape_base = dict(
        type='line', 
        yref='paper', 
        y0=-0.01, 
        y1=1.01, 
        xref='x', 
        fillcolor="greenyellow", 
        opacity=1.0, 
        line=line_style
    )
    # 0日後から70日後まで（7日刻み）のtimedeltaを作成
    day_offsets = range(-1000, 1000, 7) 
    shapes_list = [
        dict(
            shape_base, 
            x0=next + datetime.timedelta(days=offset), 
            x1=next + datetime.timedelta(days=offset)
        )
        for offset in day_offsets
    ]
    fig.update_layout(
        shapes=shapes_list,
        margin=dict(r=1, t=1, b=10, l=1)
    )
#======================================================

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

    
    if args.unten:
        print('<<< 画像表示中...    ', end="")
        import plotly.io as pio # plotly.ioモジュールをインポート   回転させたいがブラウザだと難しいので一旦画像にしてPILで回転させる
        from PIL import Image
        output_image_path = 'gantt_chart.png'
        pio.write_image(fig, output_image_path, format='png', scale=1) # scale解像度
        try:
            with Image.open(output_image_path) as img:
                rotated_img = img.transpose(Image.ROTATE_270)
                rotated_img.show()
        except FileNotFoundError:
            print(f"エラー: '{output_image_path}' が見つかりません。Plotlyでの画像生成が成功したか確認してください。")
        except Exception as e:
            print(f"画像の回転中にエラーが発生しました: {e}")
        print(' 完了 >>>')
        os._exit(0)

    plotly.offline.plot(
        fig, filename='gantt-group-tasks-together.html', auto_open=False)
    print('### Updated END	###  ')

    src='C:\me\ical_to_ganto\gantt-group-tasks-together.html'
    if os.path.isfile(src):
        print('Sonzai')
        copy = '//saclaoprfs01.spring8.or.jp/log_note/calendar/gantt-group-tasks-together.html'                    
        try:
            shutil.copyfile(src, copy) #  //saclaoprfs01.spring8.or.jp　に繋がらないと落ちるのエラー処理入れ サーバーsaclaoprfs01.spring8.or.jpへは書き込み権限のあるユーザーでログインしてる必要がある
            print("ログサーバーへコピーが完了しました。")
            """
            try:
                browser = webbrowser.get('C:/Program Files/Google/Chrome/Application/chrome.exe %s')
                browser.open('http://saclaopr19.spring8.or.jp/~lognote/calendar/gantt-group-tasks-together.html', new=2) # new=2 は新しいタブまたはウィンドウで開くことを意味します
            except webbrowser.Error:
                print("Chromeブラウザが見つかりませんでした。デフォルトブラウザで開きます。")
                webbrowser.open('gantt-group-tasks-together.html')
            """          
        except Exception as e:
            print(f"予期しないエラーが発生しました: {e}")
            print("たぶんログサーバーにアクセスできない。DOSで叩いてみて下さい「net use \\saclaoprfs01.spring8.or.jp /user:log_user4 ses@sacla5712」")
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
