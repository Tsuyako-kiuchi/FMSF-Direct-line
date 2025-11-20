import pandas as pd

def normalize_category(cat):
    if pd.isna(cat):
        return "その他業者"
    if '発注者' in cat: return '発注者'
    if 'コンサル' in cat: return 'コンサル'
    if '設計' in cat: return '設計監理'
    if '別途' in cat: return '別途業者'
    if '施工図' in cat: return '施工図'
    if '躯体' in cat: return '躯体工事'
    if '仕上げ' in cat: return '仕上げ工事'
    if '職員' in cat: return '職員'
    return 'その他業者'

def generate_html():
    # ✅ 最初のシートを自動読み込み（シート名エラー防止）
    df = pd.read_excel("data.xlsx", engine="openpyxl")
    df = df.dropna(subset=["会社名", "氏名"])
    df['カテゴリ'] = df['カテゴリ'].apply(normalize_category)

    category_order = ['発注者','コンサル','設計監理','別途業者','施工図','躯体工事','仕上げ工事','その他業者','職員']
    df['カテゴリ'] = pd.Categorical(df['カテゴリ'], categories=category_order, ordered=True)
    df = df.sort_values('カテゴリ')

    html = """<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>電話帳</title>
<style>
body { font-family: Arial, sans-serif; margin: 0; padding: 0; background-color: #f9f9f9; }
header { background-color: #333; color: white; padding: 10px; text-align: center; font-size: 1.5em; }
.search-bar { position: sticky; top: 0; background: #333; padding: 10px; display: flex; justify-content: center; flex-wrap: wrap; }
.search-bar input { width: 70%; padding: 10px; font-size: 1em; margin-bottom: 5px; }
.search-bar button { padding: 10px; margin-left: 5px; background: #555; color: white; border: none; cursor: pointer; font-size: 1em; }
.category { margin: 10px; background: #fff; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
.category h2 { margin: 0; padding: 15px; background: #eee; cursor: pointer; font-size: 1.2em; }
.contacts { display: none; padding: 10px; }
.contact-item { margin-bottom: 10px; padding: 8px; border-bottom: 1px solid #ddd; font-size: 1em; }
.contact-item strong { display: block; font-size: 1.1em; }
.contact-item span { display: block; margin: 2px 0; }
.contact-item a { color: #004080; text-decoration: none; }
#auth-screen { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: #f9f9f9; display: flex; flex-direction: column; justify-content: center; align-items: center; }
#auth-screen input { font-size: 1.2em; padding: 10px; margin: 10px; width: 150px; text-align: center; }
#auth-screen button { font-size: 1.2em; padding: 10px 20px; background: #004080; color: white; border: none; cursor: pointer; }
#auth-error { color: red; margin-top: 10px; }
#no-results { text-align: center; color: red; font-size: 1.2em; display: none; margin-top: 20px; }
/* ✅ スマホ対応 */
@media (max-width: 600px) {
    header { font-size: 1.2em; }
    .search-bar input { width: 100%; margin-bottom: 10px; }
    .search-bar button { width: 48%; margin: 5px 1%; }
    .category h2 { font-size: 1.1em; padding: 10px; }
    .contact-item { font-size: 0.9em; }
}
</style></head><body>
<div id='auth-screen'><h2>PINコードを入力してください</h2><input type='password' id='pin' maxlength='4' placeholder='0000'><button onclick='checkPIN()'>認証</button><div id='auth-error'></div></div>
<header style='display:none;'>電話帳</header>
<div class='search-bar' style='display:none;'><input type='text' id='search' placeholder='検索（会社名・氏名）'><button onclick='performSearch()'>検索</button><button onclick='clearSearch()'>クリア</button></div>
<div id='no-results'>該当者なし</div>
<div id='content' style='display:none;'>"""

    for category, group in df.groupby('カテゴリ'):
        html += f"<div class='category'><h2 onclick='toggleCategory(this)'>{category}</h2><div class='contacts'>"
        for _, row in group.iterrows():
            html += "<div class='contact-item'>"
            html += f"<strong>{row['氏名']}</strong>"
            html += f"<span>会社名：{row['会社名']}</span>"
            if pd.notna(row['役職']): html += f"<span>役職：{row['役職']}</span>"
            if pd.notna(row['住所']): html += f"<span>住所：{row['住所']}</span>"
            if pd.notna(row['電話番号']): html += f"<span>電話：<a href='tel:{row['電話番号']}'>{row['電話番号']}</a></span>"
            if pd.notna(row['FAX番号']): html += f"<span>FAX：{row['FAX番号']}</span>"
            if pd.notna(row['携帯番号']): html += f"<span>携帯：<a href='tel:{row['携帯番号']}'>{row['携帯番号']}</a></span>"
            if pd.notna(row['アドレス']): html += f"<span>メール：<a href='mailto:{row['アドレス']}'>{row['アドレス']}</a></span>"
            html += "</div>"
        html += "</div></div>"

    html += """</div><script>
function checkPIN(){const pin=document.getElementById('pin').value;if(pin==='0000'){document.getElementById('auth-screen').style.display='none';document.querySelector('header').style.display='block';document.querySelector('.search-bar').style.display='flex';document.getElementById('content').style.display='block';}else{document.getElementById('auth-error').textContent='PINコードが違います';}}
function toggleCategory(header){const contacts=header.nextElementSibling;contacts.style.display=contacts.style.display==='block'?'none':'block';}
function performSearch(){const query=document.getElementById('search').value.toLowerCase();let anyMatch=false;document.querySelectorAll('.category').forEach(cat=>{let matchFound=false;const contacts=cat.querySelector('.contacts');cat.querySelectorAll('.contact-item').forEach(item=>{const text=item.innerText.toLowerCase();if(text.includes(query)){item.style.display='block';matchFound=true;anyMatch=true;}else{item.style.display='none';}});if(matchFound){cat.style.display='block';contacts.style.display='block';}else{cat.style.display='none';contacts.style.display='none';}});document.getElementById('no-results').style.display=anyMatch?'none':'block';}
function clearSearch(){document.getElementById('search').value='';document.getElementById('no-results').style.display='none';document.querySelectorAll('.category').forEach(cat=>{cat.style.display='block';cat.querySelector('.contacts').style.display='none';cat.querySelectorAll('.contact-item').forEach(item=>{item.style.display='block';});});}
document.getElementById('search').addEventListener('input',performSearch);
</script></body></html>"""

    with open("index.html","w",encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    generate_html()