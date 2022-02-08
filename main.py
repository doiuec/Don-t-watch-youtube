import time
import win32com.client
import win32con
import ctypes
from plyer import notification

'''
ctypes で C の dll ライブラリを呼び出し，各関数を実装してみる
なお， cdll の呼び出し規約は cdeclで， windll の呼び出し規約は stdcalなので，
スタックを関数がクリーンアップするか，呼び出し元がするかに
注意する必要がある（たぶん）
今回は WindowsAPI なので stdcall
'''

def _get_running_window() -> list:
    # 存在するウィンドウを列挙する EnumWindows
    # 見つかったウィンドウがコールバック関数で返される
    EnumWindows = ctypes.windll.user32.EnumWindows

    ## l.50 EnumWindows(EnumWindowsProc(callback), None)
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                        ctypes.POINTER(ctypes.c_int),
                                        ctypes.POINTER(ctypes.c_int))

    # ウィンドウの名前の長さを取得する．
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW

    # ウィンドウの名前を取得する．
    GetWindowText = ctypes.windll.user32.GetWindowTextW

    # そのウィンドウが見えるか見えないか
    # 単にウィンドウの視点からの確認
    # chromeでウィンドウを開いていても，見えないタブは見えないことになる
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible

    # 現在実行中のプロセスをここに入れる
    # ここに "YouTube" があったらアウト
    titles = []

    def callback(hwnd, lParam):
        if IsWindowVisible:
            length = GetWindowTextLength(hwnd)
            buffer = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buffer, length + 1)
            titles.append(buffer.value)
            return True
    EnumWindows(EnumWindowsProc(callback), None)
    
    # 集合にして冗長な情報を削る
    titles = set(titles)
    # print(*titles, sep="\n")
    return list(titles)

def is_youtube_open() -> bool:
    process_list = _get_running_window()
    reject = ["youtube", "YouTube", "Youtube"]
    for process in process_list:
        for reject_ in reject:
            if reject_ in process:
                return True
    return False

def mouse_move_close(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def open_toast( flag ):
        notification.notify(
        title="警告",
        message="youtube を終了します",
        app_name="python",
        app_icon="warnning.ico",
        timeout=1)

def Cortana( speech ):
    sapi = win32com.client.Dispatch("SAPI.SpVoice")
    cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
    cat.SetID(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False)
    v = [t for t in cat.EnumerateTokens() if t.GetAttribute("Name") == "Microsoft Sayaka"]

    if v:
        oldv = sapi.Voice
        sapi.Voice = v[0]
        sapi.Speak( speech )
        sapi.Voice = oldv

def main():
    print("監視中...")
    while(True):
        if is_youtube_open():
            time.sleep(1)
            mouse_move_close(1900,10)
            open_toast("warnning")
            Cortana( "勉強しろ" )


if __name__ == '__main__':
    main()
