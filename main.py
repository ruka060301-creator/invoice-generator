"""
請求書自動作成アプリ - GUI

CSV・テンプレート・出力先を選び、対象年月を入力して請求書を作成する。
データ処理は csv_loader / excel_writer に任せ、この層は画面だけを担当する。
"""

import datetime
import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, ttk

import csv_loader
import excel_writer

# 前回選んだパスを保存しておくファイル(このスクリプトと同じ場所に作られる)
SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "settings.json")


class InvoiceApp:
    """請求書作成アプリのメイン画面。"""

    def __init__(self, root):
        self.root = root
        self.root.title("請求書自動作成")
        self.root.geometry("620x420")
        self.root.resizable(False, False)

        # 入力欄と結びつける変数
        self.csv_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.year = tk.StringVar()
        self.month = tk.StringVar()

        self._build_widgets()
        self._load_settings()

    # -- 画面の組み立て ---------------------------------------------------

    def _build_widgets(self):
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text="請求書自動作成",
            font=("Meiryo UI", 14, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 12))

        # ファイル選択欄を3つ並べる
        self._add_path_row(frame, 1, "ねっぱんCSV", self.csv_path,
                           self._choose_csv)
        self._add_path_row(frame, 2, "テンプレート", self.template_path,
                           self._choose_template)
        self._add_path_row(frame, 3, "出力先フォルダ", self.output_dir,
                           self._choose_output_dir)

        # 対象年月
        period = ttk.Frame(frame)
        period.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(12, 4))
        ttk.Label(period, text="対象年月").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Entry(period, textvariable=self.year, width=6,
                  justify=tk.CENTER).pack(side=tk.LEFT)
        ttk.Label(period, text="年").pack(side=tk.LEFT, padx=(4, 12))
        ttk.Entry(period, textvariable=self.month, width=4,
                  justify=tk.CENTER).pack(side=tk.LEFT)
        ttk.Label(period, text="月").pack(side=tk.LEFT, padx=(4, 0))

        # 実行ボタン
        self.run_button = ttk.Button(
            frame, text="請求書を作成", command=self._on_run
        )
        self.run_button.grid(row=5, column=0, columnspan=3, pady=14)

        # 結果表示欄
        ttk.Label(frame, text="実行結果").grid(row=6, column=0, sticky=tk.W)
        self.log = tk.Text(frame, height=8, width=72, state=tk.DISABLED,
                           wrap=tk.WORD, font=("Meiryo UI", 9))
        self.log.grid(row=7, column=0, columnspan=3, pady=(4, 0))

    def _add_path_row(self, parent, row, label, variable, command):
        """「ラベル + 入力欄 + 選択ボタン」の1行を作る。"""
        ttk.Label(parent, text=label, width=13).grid(
            row=row, column=0, sticky=tk.W, pady=3
        )
        ttk.Entry(parent, textvariable=variable, width=52).grid(
            row=row, column=1, pady=3
        )
        ttk.Button(parent, text="選択...", command=command).grid(
            row=row, column=2, padx=(8, 0), pady=3
        )

    # -- ファイル選択 -----------------------------------------------------

    def _choose_csv(self):
        path = filedialog.askopenfilename(
            title="ねっぱんCSVを選択",
            filetypes=[("CSVファイル", "*.csv *.CSV"), ("すべて", "*.*")],
        )
        if path:
            self.csv_path.set(path)

    def _choose_template(self):
        path = filedialog.askopenfilename(
            title="請求書テンプレートを選択",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべて", "*.*")],
        )
        if path:
            self.template_path.set(path)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.output_dir.set(path)

    # -- 実行 -------------------------------------------------------------

    def _on_run(self):
        """作成ボタンが押されたときの処理。"""
        self._clear_log()

        error = self._validate()
        if error:
            self._log(error)
            return

        # 処理中は二重実行を防ぐためボタンを無効化する
        self.run_button.config(state=tk.DISABLED)
        self._log("処理を開始しました...")

        # 画面が固まらないよう、別スレッドで実行する
        thread = threading.Thread(target=self._run_task, daemon=True)
        thread.start()

    def _validate(self):
        """入力内容を検証する。問題があればエラーメッセージを返す。"""
        if not self.csv_path.get():
            return "ねっぱんCSVを選択してください。"
        if not self.template_path.get():
            return "テンプレートを選択してください。"
        if not self.output_dir.get():
            return "出力先フォルダを選択してください。"

        try:
            year = int(self.year.get())
            month = int(self.month.get())
        except ValueError:
            return "対象年月は数字で入力してください。"

        if not 2000 <= year <= 2100:
            return "対象年が正しくありません。"
        if not 1 <= month <= 12:
            return "対象月は 1〜12 で入力してください。"
        return None

    def _run_task(self):
        """実際の作成処理。別スレッドで動く。"""
        year = int(self.year.get())
        month = int(self.month.get())

        try:
            df = csv_loader.load_reservations(self.csv_path.get(), year, month)
            self._log(f"対象データ: {len(df)} 件")

            output_path = excel_writer.create_invoice(
                self.template_path.get(), df, year, month, self.output_dir.get()
            )
            self._log("")
            self._log("請求書を作成しました。")
            self._log(output_path)
            self._log("")
            self._log("※ Excelで開いて金額が空欄の場合は Ctrl + Alt + F9 を"
                      "押すと再計算されます。")
            self._save_settings()

        except (csv_loader.CsvLoadError, excel_writer.ExcelWriteError) as e:
            self._log("")
            self._log(f"エラー: {e}")
        except Exception as e:
            self._log("")
            self._log(f"予期しないエラーが発生しました: {e}")
        finally:
            self.run_button.config(state=tk.NORMAL)

    # -- 結果表示 ---------------------------------------------------------

    def _log(self, message):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _clear_log(self):
        self.log.config(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.config(state=tk.DISABLED)

    # -- 設定の保存 / 復元 -------------------------------------------------

    def _save_settings(self):
        """次回起動時のためにパスを保存する。失敗しても処理は続行する。"""
        settings = {
            "csv_path": self.csv_path.get(),
            "template_path": self.template_path.get(),
            "output_dir": self.output_dir.get(),
        }
        try:
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def _load_settings(self):
        """前回のパスを復元する。年月は「先月」を初期値にする。"""
        try:
            with open(SETTINGS_PATH, encoding="utf-8") as f:
                settings = json.load(f)
            self.csv_path.set(settings.get("csv_path", ""))
            self.template_path.set(settings.get("template_path", ""))
            self.output_dir.set(settings.get("output_dir", ""))
        except (OSError, json.JSONDecodeError):
            pass

        # 請求書は締めた後に作るので、初期値は「先月」が実用的
        today = datetime.date.today()
        last_month = today.replace(day=1) - datetime.timedelta(days=1)
        self.year.set(str(last_month.year))
        self.month.set(str(last_month.month))


def main():
    root = tk.Tk()
    InvoiceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()