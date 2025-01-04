import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import re

def parse_booth_text(text: str):
    """
    貼り付けられたBOOTHの長文テキストから必要項目を抽出。
    バリエーションがない商品も拾えるように修正。
    """
    lines = text.splitlines()

    # このリストに「商品ブロック」を順番にためこんでいく
    products_blocks = []

    current_block = []
    for line in lines:
        line = line.strip()
        if line in ("公開中", "非公開", "下書き"):
            # 次の商品が始まる合図なので、前のブロックをまとめて保存
            if current_block:
                products_blocks.append(current_block)
            # 新しいブロックをスタート
            current_block = [line]
        else:
            # それ以外はとりあえず今のブロックに詰め込む
            current_block.append(line)

    # ループを抜けたあと、最後のブロックも追加
    if current_block:
        products_blocks.append(current_block)

    results = []
    
    # いま products_blocks は「[公開中, 商品名, URL?, ...]」のように
    # 一商品ぶんずつのリストになっているはず
    for block in products_blocks:
        # まずは公開ステータス、商品名、URL を取得
        status, product_name, product_url, rest_lines = parse_product_header(block)
        
        # その後の行にバリエーションがあるかどうかをチェックして取得
        variations = parse_variations(rest_lines)
        
        if variations:
            # バリエーションごとのレコードを全部まとめる
            for var in variations:
                record = {
                    "公開ステータス": status,
                    "パーツ名": product_name,
                    "URL": product_url,
                    "バリエーション名": var["バリエーション名"],
                    "支払い待ち": var["支払い待ち"],
                    "未発送": var["未発送"],
                    "価格": var["価格"],
                    "在庫": var["在庫"],
                    "販売数": var["販売数"],
                    "売上金額": var["売上金額"],
                    "入荷待ちメール数": var["入荷待ちメール数"],
                }
                results.append(record)
        else:
            # バリエーションが見つからない場合は単一商品として1つレコードを作る
            # 価格や在庫等は「rest_lines」をまるっと探してみる
            single_data = parse_single_product(rest_lines)
            record = {
                "公開ステータス": status,
                "パーツ名": product_name,
                "URL": product_url,
                "バリエーション名": "",
                "支払い待ち": single_data["支払い待ち"],
                "未発送": single_data["未発送"],
                "価格": single_data["価格"],
                "在庫": single_data["在庫"],
                "販売数": single_data["販売数"],
                "売上金額": single_data["売上金額"],
                "入荷待ちメール数": single_data["入荷待ちメール数"],
            }
            results.append(record)

    return results


def parse_product_header(block: list):
    """
    block の最初の方から
      - 公開ステータス
      - 商品名
      - URL(あれば)
    を取得して残りを返す
    """
    # 例: block = ["公開中", "つちのこ祭り", "https://xxx", "2", "未発送", ...]
    status = block[0]
    product_name = ""
    product_url = ""
    
    # block[1] が商品名という前提で雑に書く
    if len(block) >= 2:
        product_name = block[1]

    # block[2] が URL かもしれない
    # "http" を含んでいればURL扱い
    idx = 2
    if len(block) > 2 and "http" in block[2]:
        product_url = block[2]
        idx = 3

    rest_lines = block[idx:]
    return status, product_name, product_url, rest_lines


def parse_variations(lines: list):
    """
    block の残り行からバリエーション行を探して抽出する。
    見つからなかったら [] を返す。
    """
    results = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # バリエーションっぽい行かどうかチェック
        if any(keyword in line for keyword in [
            # 体のパーツ
            "肩","体", "腕", "脚", "胸", "腿", "腹", "腰", "首", "手首", "股関節", "足首",
            # スキンカラー・色
            "風", "白", "黒", "ピンク", "モカ", "グレー", "クリアー", "黄", 
            # サイズ
            "LL", " L ", " M ", " S ", " (大) "," (小) ",
        ]):
            variation_name = line
            pay_wait = 0
            not_shipped = 0
            price_str = ""
            stock_str = ""
            backorder_count = ""
            sales_count = ""
            sales_amount = ""

            # ざっくり近所の行を見て "支払い待ち" "未発送" をカウント
            look_ahead = 1
            while (i+look_ahead) < len(lines):
                check_line = lines[i+look_ahead]
                if "支払待ち" in check_line:
                    pay_wait += 1
                if "未発送" in check_line:
                    # 未発送の行の前の行を確認
                    if i+look_ahead-1 >= 0:
                        prev_line = lines[i+look_ahead-1].strip()
                        if prev_line.isdigit():  # 数字のみの行なら入荷待ちメール数
                            backorder_count = prev_line
                    not_shipped += 1
                # 価格 or 在庫 のキーワードが出始めたらストップ
                if "価格" in check_line or "在庫" in check_line:
                    break
                look_ahead += 1

            # 価格・在庫・販売数・売上金額を探す
            for j in range(1, 20):  # 探索範囲を広げる
                if i + j >= len(lines):
                    break
                check_line = lines[i+j].strip()
                if check_line == "価格":
                    # 次の行が価格
                    if (i+j+1) < len(lines):
                        p_line = lines[i+j+1].strip()
                        if p_line.startswith("¥"):
                            p_line = p_line.replace("¥", "").strip()
                        price_str = p_line
                elif check_line == "在庫":
                    if (i+j+1) < len(lines):
                        s_line = lines[i+j+1].strip()
                        stock_str = s_line
                elif check_line == "販売数":
                    if (i+j+1) < len(lines):
                        sales_line = lines[i+j+1].strip()
                        sales_count = sales_line
                elif check_line == "売上金額":
                    if (i+j+1) < len(lines):
                        amount_line = lines[i+j+1].strip()
                        if amount_line.startswith("¥"):
                            amount_line = amount_line.replace("¥", "").strip()
                        sales_amount = amount_line

            results.append({
                "バリエーション名": variation_name,
                "支払い待ち": pay_wait,
                "未発送": not_shipped,
                "価格": price_str,
                "在庫": stock_str,
                "販売数": sales_count,
                "売上金額": sales_amount,
                "入荷待ちメール数": backorder_count
            })
        i += 1

    return results


def parse_single_product(lines: list):
    """
    バリエーション行が1つも無い場合に、
    単一商品の「支払い待ち/未発送の有無」「価格」「在庫」をざっくり拾う。
    """
    pay_wait = 0
    not_shipped = 0
    price_str = ""
    stock_str = ""
    backorder_count = ""
    sales_count = ""
    sales_amount = ""

    # とりあえず全部の行を走査して拾う
    for i in range(len(lines)):
        check_line = lines[i].strip()  # strip()を追加
        if "支払待ち" in check_line:
            pay_wait += 1
        if "未発送" in check_line:
            # 未発送の行の前の行を確認
            if i > 0:
                prev_line = lines[i-1].strip()
                if prev_line.isdigit():  # 数字のみの行なら入荷待ちメール数
                    backorder_count = prev_line
            not_shipped += 1

        if check_line == "価格":
            if i+1 < len(lines):
                p_line = lines[i+1].strip()
                if p_line.startswith("¥"):
                    p_line = p_line.replace("¥", "").strip()
                price_str = p_line
        elif check_line == "在庫":
            if i+1 < len(lines):
                stock_str = lines[i+1].strip()
        elif check_line == "販売数":
            if i+1 < len(lines):
                sales_count = lines[i+1].strip()
        elif check_line == "売上金額":
            if i+1 < len(lines):
                amount_line = lines[i+1].strip()
                if amount_line.startswith("¥"):
                    amount_line = amount_line.replace("¥", "").strip()
                sales_amount = amount_line

    return {
        "支払い待ち": pay_wait,
        "未発送": not_shipped,
        "価格": price_str,
        "在庫": stock_str,
        "販売数": sales_count,
        "売上金額": sales_amount,
        "入荷待ちメール数": backorder_count,
    }


def convert_to_csv():
    text = text_box.get("1.0", tk.END)
    if not text.strip():
        messagebox.showwarning("警告", "テキストが空です。貼り付けてから実行してください。")
        return
    
    parsed_data = parse_booth_text(text)
    if not parsed_data:
        messagebox.showinfo("情報", "パース結果が0件でした。パターンが合わない可能性があります。")
        return
    
    # URLが空の行を除外
    filtered_data = [row for row in parsed_data if row["URL"]]
    
    if not filtered_data:
        messagebox.showinfo("情報", "URLを含む有効な商品データが見つかりませんでした。")
        return
    
    # 除外された行数を計算
    removed_count = len(parsed_data) - len(filtered_data)
    
    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
    )
    if not file_path:
        return
    
    fieldnames = [
        "公開ステータス",
        "パーツ名",
        "URL",
        "バリエーション名",
        "支払い待ち",
        "未発送",
        "価格",
        "在庫",
        "販売数",
        "売上金額",
        "入荷待ちメール数"
    ]
    
    with open(file_path, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in filtered_data:  # filtered_dataを使用
            writer.writerow(row)
    
    # 入荷待ちメールがある商品をチェック
    waiting_products = []
    for row in filtered_data:  # filtered_dataを使用
        if (row["公開ステータス"] == "公開中" and 
            row["入荷待ちメール数"] and 
            row["入荷待ちメール数"].isdigit() and 
            int(row["入荷待ちメール数"]) > 0):
            waiting_products.append(row["パーツ名"])
    
    # メッセージを作成
    message = f"CSVファイルを出力しました:\n{file_path}"
    if removed_count > 0:
        message += f"\n\nURL未設定の{removed_count}件の行を除外しました。"
    if waiting_products:
        # 重複を除去して一意な商品名のリストにする
        unique_waiting_products = list(set(waiting_products))
        message += "\n\n以下の商品に入荷待ちメールが来ています：\n"
        message += "\n".join(f"・{product}" for product in unique_waiting_products)
    
    messagebox.showinfo("完了", message)


def main():
    global text_box
    
    root = tk.Tk()
    root.title("BOOTHテキスト→CSV変換ツール（在庫と価格強化版）")
    
    lbl = tk.Label(root, text="ここにBOOTHの商品管理ページからコピーした文章を貼り付けてください")
    lbl.pack(anchor="w", padx=5, pady=5)
    
    text_box = tk.Text(root, wrap="word", width=100, height=30)
    text_box.pack(padx=5, pady=5)
    
    btn_frame = tk.Frame(root)
    btn_frame.pack(padx=5, pady=5, anchor="e")
    
    convert_btn = tk.Button(btn_frame, text="CSVに変換", command=convert_to_csv)
    convert_btn.pack(side="right", padx=5)
    
    root.mainloop()


if __name__ == "__main__":
    main()
