# C互換ライブラリ呼び出しモジュール
# 必要なAPI (kernel32)をロード
from ctypes import windll, Structure, byref, wintypes


# GetStdHandleの詳細
# https://docs.microsoft.com/ja-jp/windows/console/getstdhandle
# 標準出力を取得
class cutil:
    stdout_handle = windll.kernel32.GetStdHandle(-11)
    GetConsoleInfo = windll.kernel32.GetConsoleScreenBufferInfo
    SetConsoleAttribute = windll.kernel32.SetConsoleTextAttribute

    # 構造体を定義
    # 返される情報は以下を参照
    # https://docs.microsoft.com/ja-jp/windows/console/console-screen-buffer-info-str
    # dwSize:コンソール画面のサイズ
    # dwCursorPosition:カーソルの列座標と行座標
    # wAttributes:文字の属性
    # https://docs.microsoft.com/ja-jp/windows/console/console-screen-buffers#character-attributes
    # srWindow:コンソールの左上と右下の座標
    # dwMaximumWindowSize:最大サイズ時の座標
    class console_screen_buffer_info(Structure):
        _fields_ = [
            ("dwSize", wintypes._COORD),
            ("dwCursorPosition", wintypes._COORD),
            ("wAttributes", wintypes.WORD),
            ("srWindow", wintypes.SMALL_RECT),
            ("dwMaximumWindowSize", wintypes._COORD),
        ]


class display:
    def __init__(self):
        # 現状のコンソール色設定を取得
        self.info_ = cutil.console_screen_buffer_info()
        cutil.GetConsoleInfo(cutil.stdout_handle, byref(self.info_))

    def error(self, msg_list):
        # 文字色を赤に変更
        # FOREGROUND_RED 0x0004
        # FOREGROUND_INTENSITY 0x0008
        cutil.SetConsoleAttribute(
            cutil.stdout_handle, 0x0008 | 0x0004)

        for msg in msg_list:
            print(msg)

        # もとの色設定に戻す
        cutil.SetConsoleAttribute(
            cutil.stdout_handle, self.info_.wAttributes)
