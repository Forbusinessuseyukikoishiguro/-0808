#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
転職活動メール自動生成ツール

【機能】
- 企業名・宛名を入力してメール自動生成
- 書類確認・面談・面接の返信テンプレート
- 日程調整の自動化
- プレビュー機能付き
- クリップボードコピー対応

【必要なライブラリ】
標準ライブラリのみで動作

作成者: [あなたの名前]
更新日: 2024-08-08
目的: 転職活動の効率化・時間短縮
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import datetime
import re
from typing import Dict, List


class JobHuntingEmailTool:
    """転職活動メール自動生成ツール"""

    def __init__(self, root):
        self.root = root
        self.root.title("転職活動メール自動生成ツール")
        self.root.geometry("900x700")

        # 変数の初期化
        self.company_name = tk.StringVar()
        self.contact_person = tk.StringVar()
        self.your_name = tk.StringVar(value="石黒")  # デフォルト名前
        self.email_template = tk.StringVar(value="document_confirm")
        self.date1 = tk.StringVar()
        self.date2 = tk.StringVar()
        self.date3 = tk.StringVar()
        self.time1 = tk.StringVar(value="10:00")
        self.time2 = tk.StringVar(value="14:00")
        self.time3 = tk.StringVar(value="16:00")
        self.additional_info = tk.StringVar()

        # メールテンプレート定義
        self.templates = self.define_templates()

        self.setup_ui()
        self.generate_email()  # 初期表示

    def define_templates(self) -> Dict[str, Dict]:
        """メールテンプレートを定義"""
        return {
            "document_confirm": {
                "name": "書類選考通過確認",
                "subject": "書類選考通過のご連絡について",
                "template": """採用ご担当者様
いつもお世話になっております。
{your_name}と申します。

この度は、書類選考通過のご連絡をいただき、誠にありがとうございます。
{company_name}様でのお仕事に大変興味を持っており、次の選考に進ませていただけることを嬉しく思っております。

次回の面接につきまして、以下の日程でご都合はいかがでしょうか。

【候補日程】
第1希望：{date1} {time1}〜
第2希望：{date2} {time2}〜  
第3希望：{date3} {time3}〜

上記以外でも、平日{additional_info}であれば調整可能です。
ご都合をお聞かせいただければと思います。

何かご不明な点がございましたら、お気軽にお声かけください。
お忙しい中恐れ入りますが、よろしくお願いいたします。""",
            },
            "casual_interview": {
                "name": "カジュアル面談日程調整",
                "subject": "カジュアル面談の日程について",
                "template": """採用ご担当者様
いつもお世話になっております。
{your_name}と申します。

この度は、カジュアル面談の機会をいただき、ありがとうございます。
{company_name}様について詳しくお話を伺えることを楽しみにしております。

面談の日程につきまして、以下でご都合はいかがでしょうか。

【候補日程】
第1希望：{date1} {time1}〜
第2希望：{date2} {time2}〜
第3希望：{date3} {time3}〜

面談形式につきましては、{additional_info}で対応可能です。
ご都合に合わせて調整いたします。

当日お話しさせていただきたい内容：
・{company_name}様の事業内容について
・募集ポジションの詳細について  
・今後のキャリアパスについて

お忙しい中恐れ入りますが、ご確認のほどよろしくお願いいたします。""",
            },
            "first_interview": {
                "name": "一次面接日程調整",
                "subject": "一次面接の日程調整について",
                "template": """採用ご担当者様　
いつもお世話になっております。
{your_name}と申します。

この度は、一次面接の機会をいただき、誠にありがとうございます。
{company_name}様とお話しさせていただけることを大変楽しみにしております。

面接の日程につきまして、以下の候補日でご調整いただけますでしょうか。

【候補日程】
第1希望：{date1} {time1}〜
第2希望：{date2} {time2}〜
第3希望：{date3} {time3}〜

{additional_info}

面接に際して事前に準備すべき資料や、
当日お持ちすべきものがございましたらお教えください。

お忙しい中恐れ入りますが、何卒よろしくお願いいたします。""",
            },
            "interview_thanks": {
                "name": "面接後お礼メール",
                "subject": "面接のお礼",
                "template": """採用ご担当者様　
いつもお世話になっております。
{your_name}と申します。

本日は貴重なお時間をいただき、面接の機会をくださり誠にありがとうございました。

{contact_person}様をはじめ、面接官の皆様から{company_name}様の事業内容や
募集ポジションについて詳しくお話を伺うことができ、
より一層{company_name}様でお仕事をさせていただきたいという気持ちが強くなりました。

特に{additional_info}についてのお話が印象的で、
自分の経験を活かして貢献できるのではないかと感じております。

次回の選考結果を心よりお待ちしております。

改めまして、本日は貴重なお時間をいただき、ありがとうございました。
今後ともどうぞよろしくお願いいたします。""",
            },
            "schedule_change": {
                "name": "日程変更依頼",
                "subject": "面接日程変更のお願い",
                "template": """いつもお世話になっております。
{your_name}と申します。

{date1}に予定しておりました面接の件でご連絡いたします。

大変申し訳ございませんが、急用により{additional_info}ため、
面接の日程を変更していただくことは可能でしょうか。

【変更希望日程】
第1希望：{date2} {time2}〜
第2希望：{date3} {time3}〜

採用ご担当者様　直前のご連絡となり、大変ご迷惑をおかけして申し訳ございません。
ご都合を確認いただき、再調整していただけますでしょうか。

お忙しい中恐れ入りますが、よろしくお願いいたします。""",
            },
            "decline_politely": {
                "name": "丁重な辞退連絡",
                "subject": "選考辞退のご連絡",
                "template": """いつもお世話になっております。
{your_name}と申します。

この度は、{company_name}様の選考において貴重なお時間をいただき、
誠にありがとうございました。

大変申し上げにくいのですが、{additional_info}により、
今回の選考を辞退させていただきたく、ご連絡いたします。

{contact_person}様には親身にご対応いただき、
{company_name}様の魅力的な事業内容についても詳しく教えていただき、
心より感謝しております。

このような結果となり、大変申し訳ございません。

また機会がございましたら、その際はよろしくお願いいたします。
末筆ながら、{company_name}様のますますのご発展を心よりお祈りしております。""",
            },
            "question_inquiry": {
                "name": "質問・問い合わせ",
                "subject": "募集要項についてのご質問",
                "template": """採用ご担当者様
いつもお世話になっております。
{your_name}と申します。

{company_name}様の求人について、いくつかご質問がございます。
お忙しい中恐れ入りますが、ご回答いただけますでしょうか。

【ご質問内容】
{additional_info}

上記についてお教えいただけますと幸いです。

応募に際して他に準備すべき書類や、
確認しておくべき事項がございましたら併せてお教えください。

お忙しい中恐れ入りますが、何卒よろしくお願いいたします。""",
            },
        }

    def setup_ui(self):
        """UIセットアップ"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 基本情報エリア
        self.create_basic_info_area(main_frame)

        # テンプレート選択エリア
        self.create_template_area(main_frame)

        # 日程入力エリア
        self.create_schedule_area(main_frame)

        # プレビューエリア
        self.create_preview_area(main_frame)

        # ボタンエリア
        self.create_button_area(main_frame)

        # グリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

    def create_basic_info_area(self, parent):
        """基本情報入力エリア作成"""
        basic_frame = ttk.LabelFrame(parent, text="基本情報", padding="5")
        basic_frame.grid(
            row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10)
        )
        basic_frame.columnconfigure(1, weight=1)

        # 企業名
        ttk.Label(basic_frame, text="企業名:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 5)
        )
        ttk.Entry(basic_frame, textvariable=self.company_name, width=40).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10)
        )

        # 担当者名
        ttk.Label(basic_frame, text="担当者名:").grid(
            row=0, column=2, sticky=tk.W, padx=(10, 5)
        )
        ttk.Entry(basic_frame, textvariable=self.contact_person, width=25).grid(
            row=0, column=3, sticky=(tk.W, tk.E)
        )

        # あなたの名前
        ttk.Label(basic_frame, text="あなたの名前:").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0)
        )
        ttk.Entry(basic_frame, textvariable=self.your_name, width=25).grid(
            row=1, column=1, sticky=tk.W, pady=(5, 0)
        )

    def create_template_area(self, parent):
        """テンプレート選択エリア作成"""
        template_frame = ttk.LabelFrame(parent, text="メールテンプレート", padding="5")
        template_frame.grid(
            row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10)
        )

        # テンプレート選択
        ttk.Label(template_frame, text="テンプレート:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10)
        )

        template_combo = ttk.Combobox(
            template_frame, textvariable=self.email_template, width=30
        )
        template_combo["values"] = [
            f"{key}: {value['name']}" for key, value in self.templates.items()
        ]
        template_combo.grid(row=0, column=1, sticky=tk.W)
        template_combo.bind("<<ComboboxSelected>>", self.on_template_changed)

        # 追加情報
        ttk.Label(template_frame, text="追加情報:").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0)
        )
        ttk.Entry(template_frame, textvariable=self.additional_info, width=60).grid(
            row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0)
        )

    def create_schedule_area(self, parent):
        """日程入力エリア作成"""
        schedule_frame = ttk.LabelFrame(parent, text="日程設定", padding="5")
        schedule_frame.grid(
            row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10)
        )

        # 日程入力
        for i, (date_var, time_var) in enumerate(
            [
                (self.date1, self.time1),
                (self.date2, self.time2),
                (self.date3, self.time3),
            ],
            1,
        ):
            ttk.Label(schedule_frame, text=f"第{i}希望:").grid(
                row=i - 1, column=0, sticky=tk.W, padx=(0, 5)
            )

            # 日付入力
            ttk.Entry(schedule_frame, textvariable=date_var, width=20).grid(
                row=i - 1, column=1, padx=(0, 5)
            )
            ttk.Label(schedule_frame, text="時刻:").grid(
                row=i - 1, column=2, sticky=tk.W, padx=(10, 5)
            )
            ttk.Entry(schedule_frame, textvariable=time_var, width=10).grid(
                row=i - 1, column=3, padx=(0, 10)
            )

            # 今日から数日後の日付を自動設定
            future_date = datetime.date.today() + datetime.timedelta(days=i + 2)
            if not date_var.get():
                date_var.set(future_date.strftime("%Y年%m月%d日(%a)"))

        # 日程入力ヘルプ
        help_text = "例: 2024年8月15日(木) または 8/15(木)"
        ttk.Label(schedule_frame, text=help_text, font=("", 8)).grid(
            row=3, column=1, sticky=tk.W, pady=(5, 0)
        )

    def create_preview_area(self, parent):
        """プレビューエリア作成"""
        preview_frame = ttk.LabelFrame(parent, text="メールプレビュー", padding="5")
        preview_frame.grid(
            row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)
        )
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)

        # 件名表示
        self.subject_label = ttk.Label(
            preview_frame, text="件名: ", font=("", 9, "bold")
        )
        self.subject_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        # メール本文
        self.preview_text = scrolledtext.ScrolledText(
            preview_frame, height=15, width=80, wrap=tk.WORD
        )
        self.preview_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def create_button_area(self, parent):
        """ボタンエリア作成"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(0, 10))

        ttk.Button(
            button_frame, text="🔄 プレビュー更新", command=self.generate_email
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            button_frame,
            text="📋 クリップボードコピー",
            command=self.copy_to_clipboard,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            button_frame, text="🗂️ テンプレート管理", command=self.manage_templates
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="💾 設定保存", command=self.save_settings).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(button_frame, text="🧹 クリア", command=self.clear_all).pack(
            side=tk.LEFT
        )

    def on_template_changed(self, event=None):
        """テンプレート変更時の処理"""
        self.generate_email()

    def generate_email(self):
        """メール生成"""
        try:
            # テンプレート取得
            template_key = (
                self.email_template.get().split(":")[0]
                if ":" in self.email_template.get()
                else self.email_template.get()
            )
            if template_key not in self.templates:
                template_key = "document_confirm"

            template_info = self.templates[template_key]

            # 変数の準備
            variables = {
                "company_name": self.company_name.get() or "[企業名]",
                "contact_person": self.contact_person.get() or "[担当者名]",
                "your_name": self.your_name.get() or "[あなたの名前]",
                "date1": self.date1.get() or "[第1希望日]",
                "date2": self.date2.get() or "[第2希望日]",
                "date3": self.date3.get() or "[第3希望日]",
                "time1": self.time1.get() or "[時刻1]",
                "time2": self.time2.get() or "[時刻2]",
                "time3": self.time3.get() or "[時刻3]",
                "additional_info": self.additional_info.get() or "[追加情報]",
            }

            # メール生成
            subject = template_info["subject"].format(**variables)
            body = template_info["template"].format(**variables)

            # プレビュー更新
            self.subject_label.config(text=f"件名: {subject}")
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, body)

        except Exception as e:
            messagebox.showerror("エラー", f"メール生成エラー: {e}")

    def copy_to_clipboard(self):
        """クリップボードにコピー"""
        try:
            # 件名と本文を結合
            subject = self.subject_label.cget("text").replace("件名: ", "")
            body = self.preview_text.get(1.0, tk.END).strip()

            full_email = f"件名: {subject}\n\n{body}"

            # クリップボードにコピー
            self.root.clipboard_clear()
            self.root.clipboard_append(full_email)

            messagebox.showinfo(
                "完了",
                "メール内容をクリップボードにコピーしました！\nメールソフトにペーストして使用してください。",
            )

        except Exception as e:
            messagebox.showerror("エラー", f"コピーエラー: {e}")

    def manage_templates(self):
        """テンプレート管理ウィンドウ"""
        template_window = tk.Toplevel(self.root)
        template_window.title("テンプレート管理")
        template_window.geometry("600x400")

        # テンプレート一覧
        frame = ttk.Frame(template_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="利用可能なテンプレート:", font=("", 11, "bold")).pack(
            anchor=tk.W, pady=(0, 10)
        )

        for key, template in self.templates.items():
            template_frame = ttk.LabelFrame(frame, text=template["name"], padding="5")
            template_frame.pack(fill=tk.X, pady=(0, 10))

            ttk.Label(template_frame, text=f"用途: {template['name']}").pack(
                anchor=tk.W
            )
            ttk.Label(
                template_frame, text=f"件名: {template['subject']}", font=("", 8)
            ).pack(anchor=tk.W)

            # 選択ボタン
            ttk.Button(
                template_frame,
                text="このテンプレートを使用",
                command=lambda k=key: self.select_template(k, template_window),
            ).pack(anchor=tk.E, pady=(5, 0))

    def select_template(self, template_key, window):
        """テンプレート選択"""
        self.email_template.set(
            f"{template_key}: {self.templates[template_key]['name']}"
        )
        self.generate_email()
        window.destroy()

    def save_settings(self):
        """設定保存"""
        try:
            settings = {
                "company_name": self.company_name.get(),
                "contact_person": self.contact_person.get(),
                "your_name": self.your_name.get(),
                "additional_info": self.additional_info.get(),
            }

            # 簡単な設定ファイル保存（実際のアプリではJSONファイルなどに保存）
            messagebox.showinfo(
                "完了", "設定を保存しました！\n（この機能は今後の拡張で実装予定）"
            )

        except Exception as e:
            messagebox.showerror("エラー", f"保存エラー: {e}")

    def clear_all(self):
        """全項目クリア"""
        self.company_name.set("")
        self.contact_person.set("")
        self.date1.set("")
        self.date2.set("")
        self.date3.set("")
        self.additional_info.set("")
        self.generate_email()

    def auto_fill_sample(self):
        """サンプルデータで自動入力"""
        self.company_name.set("株式会社○○")
        self.contact_person.set("人事ご担当者様")
        self.additional_info.set("オンライン・オフィス")

        # 今日から3日後、4日後、5日後の日付を設定
        for i, date_var in enumerate([self.date1, self.date2, self.date3], 3):
            future_date = datetime.date.today() + datetime.timedelta(days=i)
            weekday = ["月", "火", "水", "木", "金", "土", "日"][future_date.weekday()]
            date_var.set(f"{future_date.strftime('%Y年%m月%d日')}({weekday})")

        self.generate_email()


def create_quick_email_window():
    """クイック作成ウィンドウ"""
    quick_window = tk.Toplevel()
    quick_window.title("クイックメール作成")
    quick_window.geometry("500x300")

    frame = ttk.Frame(quick_window, padding="20")
    frame.pack(fill=tk.BOTH, expand=True)

    # 簡単入力
    ttk.Label(frame, text="企業名:", font=("", 10, "bold")).grid(
        row=0, column=0, sticky=tk.W, pady=5
    )
    company_entry = ttk.Entry(frame, width=30)
    company_entry.grid(row=0, column=1, padx=(10, 0), pady=5)

    ttk.Label(frame, text="メールタイプ:", font=("", 10, "bold")).grid(
        row=1, column=0, sticky=tk.W, pady=5
    )
    type_combo = ttk.Combobox(
        frame, values=["書類確認", "面談調整", "面接調整", "お礼"], width=27
    )
    type_combo.grid(row=1, column=1, padx=(10, 0), pady=5)
    type_combo.set("書類確認")

    def quick_generate():
        company = company_entry.get()
        mail_type = type_combo.get()

        if company and mail_type:
            # メインウィンドウに値を設定（実装時はメインクラスのインスタンスにアクセス）
            messagebox.showinfo(
                "完了", f"{company}様への{mail_type}メールを生成しました！"
            )
            quick_window.destroy()
        else:
            messagebox.showwarning("警告", "企業名とメールタイプを入力してください。")

    ttk.Button(
        frame, text="メール生成", command=quick_generate, style="Accent.TButton"
    ).grid(row=2, column=1, pady=20)


def main():
    """メイン関数"""
    root = tk.Tk()
    app = JobHuntingEmailTool(root)

    # スタイル設定
    style = ttk.Style()
    try:
        style.configure("Accent.TButton", foreground="white", background="blue")
    except:
        pass

    # メニューバー追加
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    # ファイルメニュー
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="ファイル", menu=file_menu)
    file_menu.add_command(label="サンプル入力", command=app.auto_fill_sample)
    file_menu.add_separator()
    file_menu.add_command(label="終了", command=root.quit)

    # ヘルプメニュー
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="ヘルプ", menu=help_menu)
    help_menu.add_command(
        label="使い方",
        command=lambda: messagebox.showinfo(
            "使い方",
            "1. 企業名・担当者名を入力\n2. テンプレートを選択\n3. 日程を入力\n4. プレビュー確認\n5. クリップボードにコピー",
        ),
    )

    # 初期状態でサンプルデータを入力
    app.auto_fill_sample()

    root.mainloop()


if __name__ == "__main__":
    main()
