import os
import sys
import configparser
import glob
import shutil
import cv2

# 変数定義
PYTHON_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_INI_PATH = os.path.join(PYTHON_DIR,"config.ini")

# iniファイル存在チェック
if not os.path.exists(CONFIG_INI_PATH):
    print("【ERROR】code00001")
    print("iniファイルが存在しません。")
    print("image.pyと同じ階層にconfig.iniを作成してください。")
    sys.exit()

# iniファイル読み取り
try:
    config_ini = configparser.ConfigParser()
    config_ini.read(CONFIG_INI_PATH, encoding="utf-8")

    BEFORE_FOLDER = config_ini["IMAGE"]["BeforeFolder"]
    AFTER_FOLDER = config_ini["IMAGE"]["AfterFolder"]
    EXTENTION_LIST = config_ini["IMAGE"]["FileExtension"].split()
except KeyError as e:
    print("【ERROR】code00002")
    print("iniファイルの中身が存在しません。")
    print("config.iniを以下のように作成してください。")
    print("1.セクション[""IMAGE""]を追加します。")
    print("2.パラメータBeforeFolderにトリミング前の画像フォルダを指定します。")
    print("画像フォルダのファイルは再帰的に取得されます。")
    print("3.パラメータAfterFolderにトリミング後の画像フォルダを指定します。")
    print("トリミング前の画像フォルダにフォルダが存在する場合は、フォルダを新たに作成します。")
    print("4.パラメータFileExtensionにトリミングする画像ファイルの拡張子を空白区切りで指定します。")
    print("OpenCVで処理できない拡張子を指定した場合はエラーが発生します。")
    sys.exit()
except configparser.ParsingError as e:
    print("【ERROR】code00003")
    print("iniファイルの書き方が間違っています。")
    print("config.iniを以下のように設定してください。")
    print("1.セクション[""IMAGE""]を追加します。")
    print("2.パラメータBeforeFolderにトリミング前の画像フォルダを指定します。")
    print("画像フォルダのファイルは再帰的に取得されます。")
    print("3.パラメータAfterFolderにトリミング後の画像フォルダを指定します。")
    print("トリミング前の画像フォルダにフォルダが存在する場合は、フォルダを新たに作成します。")
    print("4.パラメータFileExtensionにトリミングする画像ファイルの拡張子を空白区切りで指定します。")
    print("OpenCVで処理できない拡張子を指定した場合はエラーが発生します。")
    sys.exit()

# ファイル一覧取得(再帰的)
files_path = glob.glob(os.path.join(BEFORE_FOLDER,"**"), recursive=True)

print("画像をトリミングします。")

for file_path in files_path:
    # 拡張子チェック
    if not os.path.splitext(file_path)[1] in EXTENTION_LIST:
        continue

    before_img_path = file_path
    after_img_dir = os.path.dirname(file_path).replace(BEFORE_FOLDER,AFTER_FOLDER)
    after_img_path = os.path.join(after_img_dir,os.path.basename(file_path))
    print("(適用元)"+before_img_path)
    print("(適用先)"+after_img_path)

    # フォルダ作成(AfterにBeforeのフォルダが存在しない場合)
    if not os.path.exists(after_img_dir):
        os.makedirs(after_img_dir, exist_ok=True)

    # ファイルコピー
    try:
        shutil.copyfile(before_img_path, after_img_path)
    except shutil.SameFileError:
        print("【ERROR】code00004")
        print("iniファイルのパラメータBeforeFolder、AfterFolderは末尾にスラッシュ/が入っているとエラーが発生します。")
        print("(エラー例)./Image/PowerAutomate/")
        print("以下のように修正してください。")
        print("(修正例)./Image/PowerAutomate")
        sys.exit()

    # 画像トリミング開始
    trim_img = cv2.imread(after_img_path)
    check_img = cv2.imread(after_img_path, cv2.IMREAD_GRAYSCALE)

    get_trim_x = check_img.shape[1]
    get_trim_y = check_img.shape[0]

    for x in range(check_img.shape[1]):
        if check_img[0][x] == 0:
            get_trim_x = x - 2
            break

    for y in range(check_img.shape[0]):
        if check_img[y][0] == 0:
            get_trim_y = y
            break

    trim_img = trim_img[0 : get_trim_y,0 : get_trim_x]
    cv2.imwrite(after_img_path, trim_img)
