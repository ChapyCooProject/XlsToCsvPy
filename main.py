import sys
import os
import json
import pandas as pd
from tkinter import Tk, filedialog, messagebox, simpledialog
from pprint import pformat

import tempfile
import msoffcrypto
import pathlib

# 実行ファイルが存在する絶対パス
def get_current_path(exe_path):
    if getattr(sys, "frozen", False):
        # The application is frozen
        currentPath =  os.path.dirname(exe_path)
        # messagebox.showinfo("frozen", "currentPath:\n" + currentPath)
    else:
        # The application is not frozen
        # Change this bit to match where you store your data files:
        currentPath =  os.path.dirname(exe_path)
        if currentPath == "":
            currentPath =  os.path.dirname(__file__)
        # messagebox.showinfo("not frozen", "currentPath:\n" + currentPath)
    return currentPath

# Excelファイルの選択
def select_excel_file(exe_path):
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")], initialdir = get_current_path(exe_path))
    return file_path

# パスワード入力
def get_excel_password(excel_file_path):
    root = Tk()
    # root.withdraw()
    root.attributes("-topmost", True)
    root.geometry("0x0")
    root.overrideredirect(True)

    sw = root.winfo_screenwidth() 
    sh = root.winfo_screenheight()
    w = root.winfo_width()+500 # simpledialogをなるべく画面中央に寄せるための処置
    h = root.winfo_height()+250 # simpledialogをなるべく画面中央に寄せるための処置
    root.geometry('{}x{}+{}+{}'.format(0, 0, int((sw-w)/2), int((sh-h)/2)))

    root.lift()
    root.focus_force()

    file_name = os.path.basename(excel_file_path)
    password = simpledialog.askstring("パスワード入力（"+file_name+"）", "Excelファイルのパスワードを入力してください。", parent=root)

    return password

# JSONファイルのパスを取得
def get_json_file_path(excel_file_path):
    file_name = os.path.basename(excel_file_path)
    file_name = file_name.replace(".xlsx", ".json")
    file_name = file_name.replace(".xls", ".json")
    file_path = os.path.join(os.path.dirname(excel_file_path), file_name)
    return file_path

# CSVファイルの保存先
def get_csv_file_path(excel_file_path):
    file_name = os.path.basename(excel_file_path)
    file_name = file_name.replace(".xlsx", ".csv")
    file_name = file_name.replace(".xls", ".csv")
    file_path = os.path.join(os.path.dirname(excel_file_path), file_name)
    return file_path

# ExcelファイルをCSVに変換
def convert_excel_to_csv(excel_file, csv_file, conf):

    try:
        
        tmp_file = csv_file.replace(".csv", ".tmp")

        if conf:
            
            if conf["excel_password"]:
                # パスワードを解除してテンポラリファイルを作成
                with open(excel_file, "rb") as f:
                    with tempfile.TemporaryFile() as tf:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=conf["excel_password"])
                        office_file.decrypt(tf)
                        excel_data = pd.read_excel(tf, sheet_name = conf["sheet_name"], dtype = object)
            else:
                # Excelファイル読み込み
                excel_data = pd.read_excel(excel_file, sheet_name = conf["sheet_name"], dtype = object)

            # CSVファイル出力用として空のDataFrameを定義
            csv_data = pd.DataFrame()
        
            for excel_col, col_info in conf["column_mapping"].items():
                if col_info["data_type"] == "整数":
                    csv_data[col_info["csv_col_name"]] = excel_data.iloc[:, int(excel_col)].astype("int64")
                    # csv_data[col_info["csv_col_name"]] = excel_data.iloc[:, int(excel_col)].astype(str).replace("nan", "")
                elif col_info["data_type"] == "小数":
                    csv_data[col_info["csv_col_name"]] = excel_data.iloc[:, int(excel_col)].astype("float64")
                elif col_info["data_type"] == "文字列":
                    csv_data[col_info["csv_col_name"]] = excel_data.iloc[:, int(excel_col)].astype(str).replace("nan", "")
                elif col_info["data_type"] == "日付":
                    csv_data[col_info["csv_col_name"]] = excel_data.iloc[:, int(excel_col)].dt.strftime(col_info["fmt"])
                elif col_info["data_type"] == "時刻":
                    csv_data[col_info["csv_col_name"]] = pd.to_datetime(excel_data.iloc[:, int(excel_col)],format=col_info["fmt_from"])
                    csv_data[col_info["csv_col_name"]] = csv_data[col_info["csv_col_name"]].dt.strftime(col_info["fmt_to"])
                else:
                    csv_data[col_info["csv_col_name"]] = excel_data.iloc[:, int(excel_col)]
        
            csv_data.to_csv(tmp_file, 
                            header = conf["has_header"], 
                            index = conf["has_index"],
                            encoding = conf["encoding"],
                            sep = conf["sep_char"],
                            mode = conf["write_mode"],
                            quoting = conf["quoting"]) 
        else:

            exlpass = get_excel_password(excel_file)

            if exlpass:
                # パスワードを解除してテンポラリファイルを作成
                with open(excel_file, "rb") as f:
                    with tempfile.TemporaryFile() as tf:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=exlpass)
                        office_file.decrypt(tf)
                        excel_data = pd.read_excel(tf)
            else:
                # Excelファイル読み込み
                excel_data = pd.read_excel(excel_file)

            # CSVファイル出力
            excel_data.to_csv(tmp_file,
                              header = True, # ヘッダあり
                              index = False, # 行番号なし
                              # encoding = "utf_8",
                              # sep = ",",
                              # mode = "w",
                              quoting = 2) 

        # 【write_mode】
        #   w:新規（既存は上書き）
        #   x:新規（既存は上書き不可）
        #   a:追記
        # 【quoting】
        #   0:QUOTE_MINIMAL（区切り文字、クォーテーション、改行など特別な文字を含むフィールドのみクォートする。）
        #   1:QUOTE_ALL（全てのフィールドをクォート。）
        #   2:QUOTE_NONNUMERIC（全ての非数値フィールドをクォート。）
        #   3:QUOTE_NONE（全てのフィールドをクォートしない。値に含まれる区切り文字は設定されているエスケープ文字でエスケープされる。）

        #print(pformat(csv_data))

        #出力したファイルの拡張子を「tmp」から「csv」へ変更
        os.renames(tmp_file, csv_file)

    except Exception as e:
        messagebox.showinfo("Excel To CSV", pformat(e))


if __name__ == "__main__":
    
    # 変換元となるExcelファイルを取得
    excel_file = ""
    if len(sys.argv) == 2:
        # コマンドライン引数もしくはドラッグアンドドロップで取得できた場合
        excel_file = sys.argv[1]
    else:
        # 直接起動した場合 -> ダイアログを表示
        excel_file = select_excel_file(sys.argv[0])

    if excel_file:

        # 拡張子を小文字に統一
        sf = pathlib.PurePath(excel_file).suffix
        excel_file = excel_file.replace(sf, sf.lower())

        # 設定ファイルの読み込み
        configJson = ""
        try:
            jsonPath = get_json_file_path(excel_file)
            with open(jsonPath, "r", encoding="utf-8") as config_file:
                configJson = json.load(config_file)
        except FileNotFoundError as e:
            configJson = ""
            print(pformat(e))
            # messagebox.showinfo("Excel To CSV", "設定ファイル（config.json）が見つかりませんでした。\n\n" + pformat(e))
        except Exception as e:
            messagebox.showinfo("Excel To CSV", pformat(e))
            sys.exit()
    
        # 設定ファイルの読み込み
        csv_file = get_csv_file_path(excel_file)
        if csv_file:
            convert_excel_to_csv(excel_file, csv_file, configJson)
            print("変換が完了しました。")
        else:
            print("CSVファイルの保存先が指定されていません。")
    else:

        # 変換元となるExcelファイルを取得できなかった時は終了
        print("Excelファイルが選択されていません。")

