#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
index.html のデザインはそのままに、
1) <!-- BEGIN: DATA_TABLE --> ～ <!-- END: DATA_TABLE --> の中だけを data.xlsx 由来のHTML(カテゴリ毎のカード)に差し替え
2) <!-- BEGIN: LAST_UPDATED --> ～ <!-- END: LAST_UPDATED --> の中だけを現在日時に差し替え
"""
import os, sys, re, datetime as dt
import pandas as pd
import html as htmlmod

BEGIN_TABLE = "<!-- BEGIN: DATA_TABLE -->"
END_TABLE   = "<!-- END: DATA_TABLE -->"
BEGIN_DATE  = "<!-- BEGIN: LAST_UPDATED -->"
END_DATE    = "<!-- END: LAST_UPDATED -->"

# 列名の候補（柔軟に対応）
COL_LABELS = {
    'カテゴリ': ['カテゴリ','カテゴリー','category','Category'],
    '氏名': ['氏名','名前','name','Name'],
    '会社名': ['会社名','会社','企業','company','Company'],
    '役職': ['役職','肩書','title','Title','役割'],
    '住所': ['住所','所在地','address','Address'],
    '電話番号': ['電話番号','電話','Tel','TEL','tel'],
    'FAX番号': ['FAX番号','FAX','Fax','fax'],
    '携帯番号': ['携帯番号','携帯','Mobile','mobile','Phone2'],
    'アドレス': ['アドレス','メール','Email','email','Mail','mail']
}

def detect_columns(df):
    mapping = {}
    cols_lower = {c.lower(): c for c in df.columns}
    for key, candidates in COL_LABELS.items():
        found = None
        for cand in candidates:
            # 完全一致（大文字小文字を無視）
            for col in df.columns:
                if col.lower() == cand.lower():
                    found = col; break
            if found: break
            # 部分一致（例：電話番号(会社) など）
            for col in df.columns:
                if cand.lower() in col.lower():
                    found = col; break
            if found: break
        mapping[key] = found
    return mapping


def escape_text(s):
    return htmlmod.escape(str(s)) if s is not None else ""


def build_cards_block(df):
    mapping = detect_columns(df)
    # 欠損列に備え、存在しない場合は空列を作る
    for key, col in mapping.items():
        if col is None:
            df[key] = ""
        else:
            df[key] = df[col]
    df = df.fillna("")

    # カテゴリ毎にグループ化
    cat_series = df['カテゴリ'] if 'カテゴリ' in df.columns else df.get(mapping.get('カテゴリ'), "")
    categories = []
    if 'カテゴリ' in df.columns:
        categories = [c for c in df['カテゴリ'].unique() if str(c).strip() != ""]
    elif mapping.get('カテゴリ') and mapping['カテゴリ'] in df.columns:
        categories = [c for c in df[mapping['カテゴリ']].unique() if str(c).strip() != ""]
    else:
        categories = ['すべて']
        df['カテゴリ'] = 'すべて'

    # カードHTMLを構築
    out = []
    for cat in categories:
        out.append(f"<div class='category' data-category='{escape_text(cat)}'>")
        out.append(f"<h2>{escape_text(cat)}</h2>")
        out.append("<div class='contacts active'>")
        sub = df[df['カテゴリ'] == cat]
        for _, row in sub.iterrows():
            name = escape_text(row.get('氏名', ''))
            company = escape_text(row.get('会社名', ''))
            title = escape_text(row.get('役職', ''))
            addr = escape_text(row.get('住所', ''))
            tel  = str(row.get('電話番号', '')).strip()
            fax  = escape_text(row.get('FAX番号', ''))
            mob  = str(row.get('携帯番号', '')).strip()
            mail = str(row.get('アドレス', '')).strip()

            # 1件分のカード
            card = ["<div class='contact-item'>"]
            if name: card.append(f"<strong>{name}</strong>")
            if company: card.append(f"<span>会社名：{company}</span>")
            if title: card.append(f"<span>役職：{title}</span>")
            if addr: card.append(f"<span>住所：{addr}</span>")
            if tel:
                t = htmlmod.escape(tel)
                card.append(f"<span>電話：<a href='tel:{t}'>{t}</a></span>")
            if fax: card.append(f"<span>FAX：{fax}</span>")
            if mob:
                m = htmlmod.escape(mob)
                card.append(f"<span>携帯：<a href='tel:{m}'>{m}</a></span>")
            if mail:
                mm = htmlmod.escape(mail)
                card.append(f"<span>メール：<a href='mailto:{mm}'>{mm}</a></span>")
            card.append("</div>")
            out.append("".join(card))
        out.append("</div>")  # contacts
        out.append("</div>")  # category
    return "\n".join(out)


def replace_block(original_html, new_block, begin_tag, end_tag):
    b = original_html.find(begin_tag)
    e = original_html.find(end_tag)
    if b == -1 or e == -1 or e < b:
        print(f"[ERROR] タグが見つかりません: {begin_tag} / {end_tag}")
        sys.exit(1)
    head = original_html[: b + len(begin_tag)]
    tail = original_html[e:]
    middle = "\n" + new_block + "\n"
    return head + middle + tail


def main():
    repo_root  = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(repo_root, "data.xlsx")
    index_path = os.path.join(repo_root, "index.html")

    if not os.path.exists(index_path):
        print("[ERROR] index.html が見つかりません:", index_path); sys.exit(1)
    if not os.path.exists(excel_path):
        print("[ERROR] data.xlsx が見つかりません:", excel_path); sys.exit(1)

    # data.xlsx 読み込み（先頭シート）
    try:
        df = pd.read_excel(excel_path, sheet_name=0, engine="openpyxl")
    except Exception as e:
        print("[ERROR] Excel 読み込みに失敗:", e); sys.exit(1)

    html = open(index_path, "r", encoding="utf-8").read()

    # --- 1) DATA_TABLE 差し替え ---
    new_html_block = build_cards_block(df)
    html = replace_block(html, new_html_block, BEGIN_TABLE, END_TABLE)

    # --- 2) LAST_UPDATED 差し替え ---
    now_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    updated_text = f"<div class=\"meta\">最終更新: {now_str}（data.xlsx から生成）</div>"
    html = replace_block(html, updated_text, BEGIN_DATE, END_DATE)

    # 書き出し
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(html)
    print("[OK] index.html を更新しました。")

if __name__ == "__main__":
    main()
