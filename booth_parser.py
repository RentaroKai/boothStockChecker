import csv
import re
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText

def parse_products(text: str):
    """
    与えられたテキストから
    「基礎商品名, 商品名, 価格, 在庫, 販売数, 売上金額, 発送待ちメール数」
    を抽出し、辞書のリストとして返す
    """
    # 公開ステータス系キーワード
    base_name_triggers = {"公開中", "非公開", "下書き"}

    # 無視してよいキーワード（ノイズになりそうなもの）
    # ※ ここから "未発送" は除外し、フローの中で手動で処理する。
    ignore_keywords = {
        "商品管理", "商品登録商品リスト管理", "すべて下書き公開中非公開",
        "編集する", "支払待ち –", "支払済み –"
        # "未発送" は除外 → 処理の中で明示的にスキップさせる
    }

    # 改行で分割し、前後の空白を除去
    raw_lines = [line.strip() for line in text.splitlines()]
    # 空行や無視キーワードだけの行は除外
    lines = [line for line in raw_lines if line and line not in ignore_keywords]

    products = []

    current_base_name = None       # 基礎商品名
    current_name = None            # バリエーション or 商品名
    shipment_wait_count = None     # 発送待ちメール数
    price = None
    stock = None
    sold = None
    revenue = None

    # 値の次の行を取得するためのフラグ
    next_is_price = False
    next_is_stock = False
    next_is_sold = False
    next_is_revenue = False

    # 金額用正規表現 ("¥ 2,500" などを抜き出して数値化する)
    yen_pattern = re.compile(r"¥\s*([0-9,]+)")

    # 処理を楽にするため、最後の行を番兵的に追加
    lines.append("_END_OF_TEXT_")

    i = 0
    while i < len(lines):
        line = lines[i]

        # ----------------------------------------
        # 1) 基礎商品名を取得
        #   「公開中」「非公開」「下書き」などを見つけたら、その次の行を base_name に。
        if line in base_name_triggers:
            if i + 1 < len(lines):
                potential_name = lines[i + 1]
                if potential_name == "_END_OF_TEXT_":
                    break
                current_base_name = potential_name
            i += 2
            # 新たな基礎商品になったタイミングで発送待ちメール数などをクリア
            shipment_wait_count = None
            current_name = None
            price = None
            stock = None
            sold = None
            revenue = None
            continue

        # ----------------------------------------
        # 2) 「発送待ちメール数」の検出
        #
        #    パターン: "0" (数値) の次の行が "未発送" → これは発送待ちメール数とする
        #    例: 
        #       0
        #       未発送
        #
        if line.isdigit() and (i + 1 < len(lines)) and (lines[i+1] == "未発送"):
            # これは発送待ちメール数
            shipment_wait_count = int(line)
            # さらに次行の「未発送」は商品名としないのでスキップ
            i += 2
            continue

        # ----------------------------------------
        # 3) もし行が「未発送」だったら商品名ではないのでスキップ
        #    (上の処理で数字+未発送のペアを先に処理するが、
        #     数字がなかった場合でも「未発送」単独で行が来る場合はスキップ)
        if line == "未発送":
            i += 1
            continue

        # ----------------------------------------
        # 4) 「価格」「在庫」「販売数」「売上金額」のキーワードを見つけたらフラグを立てる
        if line == "価格":
            next_is_price = True
            i += 1
            continue
        elif line == "在庫":
            next_is_stock = True
            i += 1
            continue
        elif line == "販売数":
            next_is_sold = True
            i += 1
            continue
        elif line == "売上金額":
            next_is_revenue = True
            i += 1
            continue

        # ----------------------------------------
        # 5) フラグが立っていれば、現在の line を値とみなす
        if next_is_price:
            match = yen_pattern.search(line)
            if match:
                price = int(match.group(1).replace(",", ""))
            else:
                # "¥"なしの場合に一応対応
                try:
                    price = int(line.replace(",", ""))
                except ValueError:
                    price = None
            next_is_price = False
            i += 1
            continue

        if next_is_stock:
            try:
                stock = int(line)
            except ValueError:
                stock = None
            next_is_stock = False
            i += 1
            continue

        if next_is_sold:
            try:
                sold = int(line)
            except ValueError:
                sold = None
            next_is_sold = False
            i += 1
            continue

        if next_is_revenue:
            match = yen_pattern.search(line)
            if match:
                revenue = int(match.group(1).replace(",", ""))
            else:
                # "¥"なしの場合に対応
                try:
                    revenue = int(line.replace(",", ""))
                except ValueError:
                    revenue = None
            next_is_revenue = False

            # 売上金額まで取れたら → 1レコード確定
            if current_name is None:
                # バリエーション名が無ければ、基礎商品名 = 商品名
                current_name = current_base_name

            products.append({
                "基礎商品名": current_base_name,
                "商品名": current_name,
                "価格": price,
                "在庫": stock,
                "販売数": sold,
                "売上金額": revenue,
                "発送待ちメール数": shipment_wait_count,  # ← ここに記録
            })
            # 次の商品に備えてクリア
            current_name = None
            price = None
            stock = None
            sold = None
            revenue = None
            # 発送待ちメール数は、同一商品・バリエーション内で使い回す可能性があるため、
            # ここでクリアするかどうかは要件次第。 
            # もし毎バリエーションで別の数を記録したければクリアしておく。
            #
            # shipment_wait_count = None

            i += 1
            continue

        # ----------------------------------------
        # 6) ここに来たら、商品名（バリエーション名）かもしれない
        if line == "_END_OF_TEXT_":
            break

        # 基礎商品名がまだ無いときは何もしない
        if current_base_name is None:
            i += 1
            continue

        # 数字だけの行は「在庫」や「販売数」かもしれないが、
        # 価格などのフラグが立っていない状態なのでノイズ扱い
        if line.isdigit():
            i += 1
            continue

        # 商品名 (バリエーション名) とみなす
        current_name = line
        i += 1

    return products

class BoothParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Booth在庫チェッカー")
        self.root.geometry("800x600")

        # メインフレーム
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 説明ラベル
        ttk.Label(main_frame, text="BOOTHの商品管理ページの内容をコピー＆ペーストしてください：").pack(anchor=tk.W)

        # テキスト入力エリア（スクロール可能）
        self.text_input = ScrolledText(main_frame, height=20, width=80)
        self.text_input.pack(fill=tk.BOTH, expand=True, pady=10)

        # CSV出力ボタン
        ttk.Button(main_frame, text="CSVを作成", command=self.create_csv).pack(pady=10)

    def create_csv(self):
        input_text = self.text_input.get("1.0", tk.END)
        if not input_text.strip():
            messagebox.showwarning("警告", "テキストが入力されていません。")
            return

        try:
            products_info = parse_products(input_text)
            fieldnames = [
                "基礎商品名",
                "商品名",
                "価格",
                "在庫",
                "販売数",
                "売上金額",
                "発送待ちメール数"
            ]
            
            with open("output.csv", mode="w", encoding="utf-8", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for p in products_info:
                    writer.writerow(p)
            
            messagebox.showinfo("成功", "CSVファイルを作成しました。\nファイル名: output.csv")
        
        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの作成中にエラーが発生しました。\n{str(e)}")

def main():
    root = tk.Tk()
    app = BoothParserApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
