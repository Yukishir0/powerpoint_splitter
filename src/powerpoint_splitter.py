# ■前提
#   1.実行には PowerPoint がインストールされている必要がある（Microsoft 365 でもOK）
#   2.実行には Python がインストールされている必要がある ( 動作はVer.3.10.6で確認 )
#   3.初回のみライブラリをインストールしておく（ コマンド実行「 pip install pywin32 」）
# ■使い方
#   1 実行するとファイル選択ダイアログが出る
#   2 .pptx ファイルを選択
#   3 そのファイルと同じフォルダに split_slides フォルダが作成される
#   4 各スライドが slide_1.pptx, slide_2.pptx ... と保存される

import os
import pythoncom
from tkinter import Tk, filedialog
import win32com.client

def select_pptx_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        filetypes=[("PowerPoint Files", "*.pptx")],
        title="分割したい PowerPoint ファイルを選んでください"
    )
    return file_path

def split_pptx_using_com(input_file):
    pythoncom.CoInitialize()
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # 非表示にするとエラーになるので True のまま

    # 絶対パスを使用（日本語やスペース対応）
    input_file = os.path.abspath(input_file)

    if not os.path.exists(input_file):
        print("× ファイルが存在しません:", input_file)
        return

    try:
        presentation = powerpoint.Presentations.Open(input_file, WithWindow=True)
    except Exception as e:
        print("× PowerPointでファイルを開けませんでした。エラー内容:")
        print(e)
        return

    output_dir = os.path.join(os.path.dirname(input_file), "split_slides")
    os.makedirs(output_dir, exist_ok=True)

    slide_count = presentation.Slides.Count

    for i in range(1, slide_count + 1):
        new_pres = powerpoint.Presentations.Add()
        presentation.Slides(i).Copy()
        new_pres.Slides.Paste(1)
        output_path = os.path.join(output_dir, f"slide_{i}.pptx")
        new_pres.SaveAs(output_path)
        new_pres.Close()
        print(f"✅ 保存完了: {output_path}")

    presentation.Close()
    powerpoint.Quit()
    print("◎ すべてのスライドの保存が完了しました！")

if __name__ == "__main__":
    pptx_file = select_pptx_file()
    if pptx_file:
        split_pptx_using_com(pptx_file)
    else:
        print("△! ファイルが選択されませんでした。")
