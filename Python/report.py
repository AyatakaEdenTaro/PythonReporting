import os
import sys
import configparser
import glob
import csv
from pptx import Presentation

# 変数定義
PYTHON_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_INI_PATH = os.path.join(PYTHON_DIR,"config.ini")

# iniファイル存在チェック
if not os.path.exists(CONFIG_INI_PATH):
    print("【ERROR】code00001")
    print("iniファイルが存在しません。")
    print("report.pyと同じ階層にconfig.iniを作成してください。")
    sys.exit()

# iniファイル読み取り
try:
    config_ini = configparser.ConfigParser()
    config_ini.read(CONFIG_INI_PATH, encoding="utf-8")

    REPLACE_SETTINGS_FOLDER = config_ini["REPORT"]["ReplaceSettingsFolder"]
except KeyError as e:
    print("【ERROR】code00002")
    print("iniファイルの中身が存在しません。")
    print("config.iniを以下のように作成してください。")
    print("1.セクション[""REPORT""]を追加します。")
    print("2.パラメータReplaceSettingsFolderにPowerPoint画像置換設定フォルダを指定します。")
    print("画像置換設定フォルダのファイルは再帰的に取得されます。")
    print("csvのファイル以外は無視されます。")
    sys.exit()
except configparser.ParsingError as e:
    print("【ERROR】code00003")
    print("iniファイルの書き方が間違っています。")
    print("config.iniを以下のように設定してください。")
    print("1.セクション[""REPORT""]を追加します。")
    print("2.パラメータReplaceSettingsFolderにPowerPoint画像置換設定フォルダを指定します。")
    print("画像置換設定フォルダのファイルは再帰的に取得されます。")
    print("csvのファイル以外は無視されます。")
    sys.exit()

# ファイル一覧取得(再帰的)
files_path = glob.glob(os.path.join(REPLACE_SETTINGS_FOLDER,"**"), recursive=True)

print("以下のPowerPointに画像を適用しました。")

for file_path in files_path:
    # 拡張子チェック
    if not os.path.splitext(file_path)[1] == ".csv":
        continue

    # 画像ファイル名からPowerPointファイル名取得処理追加

    # 画像ファイル一覧の取得
    replace_list = []
    with open(file_path,'r',encoding='UTF-8',newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',', quotechar='"')
        header = next(reader)

        for row in reader:
            replace_list.append(row)

    print(replace_list)






"""
report_img_path = "./after/report.png"

# PowerPointファイル読み込み
prs = Presentation("./レポート.pptx")

# 画像の追加と配置元シェイプの削除
for i, sld in enumerate(prs.slides, start=1):
    for shape in sld.shapes:
        if(shape.text=="ここに画像"):
            sld.shapes.add_picture(report_img_path,
            shape.left,
            shape.top,
            width=shape.width,
            height=shape.height)
            sp = shape.element
            sp.getparent().remove(sp)

prs.save("./コピーレポート.pptx")
"""