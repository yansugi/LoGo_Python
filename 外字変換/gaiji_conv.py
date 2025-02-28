#%%
import os
import tkinter as tk
from tkinter import ttk  # ttkモジュールをインポート
from tkinter import StringVar, messagebox

import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD

# 外字ディクショナリを作成
rename_df = pd.read_csv("gaiji.csv", encoding="utf-8")
rename_dict = dict(zip(rename_df["外字"], rename_df["平字"]))  # 外字と平字の辞書を作成

# 選択したカラム名を格納するグローバル変数
selected_column = None
input_file_path = None  # 入力ファイルのパスを格納する変数


def drop(event):
    global input_file_path
    input_file_path = event.data.strip("{}")  # ドロップされたパスの前後の波括弧を削除
    file_extension = os.path.splitext(input_file_path)[1].lower()  # 拡張子を取得
    global name_df
    try:
        if file_extension == ".csv":
            name_df = pd.read_csv(input_file_path)
            print(f"{input_file_path} をCSVファイルとして読み込みました。")
        elif file_extension in (".xls", ".xlsx"):
            name_df = pd.read_excel(input_file_path)
            print(f"{input_file_path} をExcelファイルとして読み込みました。")
        else:
            messagebox.showerror("エラー", "ファイルの種類を確認してください")
            return

        # ドラッグ＆ドロップのウィンドウのサイズと位置を取得
        width = root.winfo_width()
        height = root.winfo_height()
        x = root.winfo_x()
        y = root.winfo_y()

        # ドラッグ＆ドロップのウィンドウを閉じる
        root.destroy()

        # カラム名を取得し、プルダウンメニューを表示
        show_column_selection(width, height, x, y)
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
        messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {e}")
    except Exception as e:
        messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")


def show_column_selection(width, height, x, y):
    # 新しいウィンドウを作成
    column_window = tk.Tk()
    column_window.title("列の選択")

    # 新しいウィンドウのサイズと位置を設定
    column_window.geometry(f"{width}x{height}+{x}+{y}")
    column_window.configure(background="lightgrey")  # 背景色をlightgreyに設定

    # ラベルを追加
    label = tk.Label(
        column_window,
        text="変換する列を選んでください",
        bg="lightgrey",  # ラベルの背景色を設定
    )
    label.grid(
        row=0, column=0, columnspan=2, pady=(10, 0), sticky="nsew"
    )  # 上部に余白を追加し、中央に配置

    # カラム名を取得
    columns = name_df.columns.tolist()

    # Comboboxの選択肢を作成
    selected_var = StringVar(column_window)

    # Comboboxを作成
    combobox = ttk.Combobox(column_window, textvariable=selected_var, values=columns)
    combobox.set(columns[0])  # デフォルトの選択肢
    combobox.grid(
        row=1, column=0, padx=(10, 0), pady=10, sticky="ew"
    )  # 左側に配置し、中央に拡張

    # 選択したカラム名を表示するボタンを作成
    select_button = tk.Button(
        column_window,
        text="選択",
        command=lambda: select_column(selected_var, column_window),
        bg="#191970",  # ボタンの背景色を設定
        fg="white",  # ボタンの文字色を白に設定
    )
    select_button.grid(row=1, column=1, padx=(5, 10), pady=10)  # ボタンを右隣に配置

    # カラム名を選択する関数を作成
    def select_column(selected_var, window):
        global selected_column
        selected_column = selected_var.get()  # 選択したカラム名を格納
        print(f"選択したカラム名: {selected_column}")

        # 外字を平字に変換する関数
        def convert_gaiji(name):
            for gaiji, heiji in rename_dict.items():
                name = name.replace(gaiji, heiji)  # 外字を平字に置き換え
            return name

        name_df[selected_column] = name_df[selected_column].apply(convert_gaiji)

        # 変換後に外字を含むかどうかをチェック
        def has_external_chars(text):
            for char in text:
                code_point = ord(char)
                # 私用領域（U+E000 - U+F8FF）やそれ以外の範囲をチェック
                if (
                    code_point >= 0xE000 and code_point <= 0xF8FF
                ) or code_point > 0xFFFF:
                    return True
            return

        # 使用例
        name_df["外字判定"] = name_df[selected_column].apply(has_external_chars)

        # 新しいファイル名を生成
        base_name = os.path.basename(input_file_path)
        new_file_name = f"（外字変換）{base_name}"
        new_file_path = os.path.join(os.path.dirname(input_file_path), new_file_name)

        # データを新しいファイルに保存
        if input_file_path.endswith(".csv"):
            name_df.to_csv(new_file_path, index=False)
        else:
            name_df.to_excel(new_file_path, index=False)

        messagebox.showinfo("ファイル保存", f"{new_file_name} として保存しました。")
        window.destroy()  # ウィンドウを閉じる

    # カラム名の選択ウィンドウのカラムの幅を設定
    column_window.grid_columnconfigure(0, weight=1)  # Comboboxの列にスペースを与える
    column_window.grid_columnconfigure(1, weight=0)  # ボタンの列にはスペースを与えない

    # メインループの開始
    column_window.mainloop()


# メインウィンドウの作成
root = TkinterDnD.Tk()
root.title("外字変換くん")
root.configure(background="#2B2B2B")

# ドロップターゲットの作成
drop_target = tk.Label(
    root,
    text="ここにcsvファイルまたはエクセルファイルをドロップしてください",
    bg="lightgray",
    width=50,
    height=10,
)
drop_target.pack(padx=10, pady=10)

# ドロップイベントのバインド
drop_target.drop_target_register(DND_FILES)
drop_target.dnd_bind("<<Drop>>", drop)

# メインループの開始
root.mainloop()

# %%
