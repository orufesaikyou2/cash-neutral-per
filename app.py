import streamlit as st
import yfinance as yf
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

st.set_page_config(
    page_title="キャッシュニュートラルPER 計算ツール",
    layout="centered"
)

KABUTAN_BASE = "https://kabutan.jp"


# ===== データ取得・計算 =====

def get_value(series, *keys):
    for key in keys:
        try:
            val = series.get(key)
            if val is not None and not pd.isna(val):
                return float(val)
        except Exception:
            continue
    return 0.0


def fetch_yahoo_data(ticker_symbol):
    """Yahoo Finance から財務データを取得する"""
    try:
        ticker = yf.Ticker(ticker_symbol)
        info = ticker.info

        stock_price = info.get("currentPrice") or info.get("regularMarketPrice")
        if not stock_price:
            return {"success": False, "error": "株価データを取得できませんでした。銘柄コードを確認してください。"}

        shares_issued = info.get("sharesOutstanding")
        if not shares_issued:
            return {"success": False, "error": "株式数データを取得できませんでした。"}

        balance_sheet = ticker.balance_sheet
        if balance_sheet is None or balance_sheet.empty:
            return {"success": False, "error": "貸借対照表データを取得できませんでした。"}

        latest_bs = balance_sheet.iloc[:, 0]
        bs_date = balance_sheet.columns[0].strftime("%Y年%m月%d日")

        treasury_shares = get_value(latest_bs, "Treasury Shares Number", "TreasurySharesNumber")
        shares_outstanding = shares_issued - treasury_shares

        current_assets = get_value(latest_bs, "Current Assets", "CurrentAssets", "TotalCurrentAssets")
        investment_securities = get_value(
            latest_bs,
            "Available For Sale Securities", "AvailableForSaleSecurities",
            "Long Term Investments", "LongTermInvestments",
            "Investments And Advances", "InvestmentsAndAdvances",
            "Other Investments",
        )
        total_liabilities = get_value(
            latest_bs,
            "Total Liabilities Net Minority Interest", "TotalLiabilitiesNetMinorityInterest",
            "Total Liab", "TotalLiab", "Total Liabilities",
        )

        financials = ticker.financials
        if financials is None or financials.empty:
            return {"success": False, "error": "損益計算書データを取得できませんでした。"}

        latest_fin = financials.iloc[:, 0]
        fin_date = financials.columns[0].strftime("%Y年%m月%d日")

        net_income = get_value(
            latest_fin,
            "Net Income Common Stockholders", "NetIncomeCommonStockholders",
            "Net Income", "NetIncome",
        )

        return {
            "success": True,
            "company_name": info.get("longName", ticker_symbol),
            "currency": info.get("currency", "JPY"),
            "stock_price": stock_price,
            "shares_issued": shares_issued,
            "treasury_shares": treasury_shares,
            "shares_outstanding": shares_outstanding,
            "bs_date": bs_date,
            "fin_date": fin_date,
            "current_assets": current_assets,
            "investment_securities": investment_securities,
            "total_liabilities": total_liabilities,
            "net_income": net_income,
        }
    except Exception as e:
        return {"success": False, "error": f"データ取得中にエラーが発生しました: {str(e)}"}


def fetch_kessan_auto(stock_code):
    """
    kabutan.jp（株探）から来期予想当期純利益と期中平均株式数を取得する。

    期中平均株式数は決算短信に記載の数値を直接取得できないため、
    実績の「当期純利益 ÷ 1株当たり当期純利益(EPS)」で算出する。
    """
    code = re.sub(r"[^0-9]", "", stock_code)
    url = f"{KABUTAN_BASE}/stock/finance?code={code}"

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "ja,en-US;q=0.9",
        "Referer": "https://kabutan.jp/",
    }

    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        return {"success": False, "error": f"kabutan.jp へのアクセスに失敗しました: {e}"}

    result = {"success": False, "avg_shares": None, "forecast_net_income_m": None}

    def to_float(s):
        cleaned = re.sub(r"[^\d\.]", "", str(s).replace(",", "").replace("，", ""))
        try:
            return float(cleaned) if cleaned else None
        except ValueError:
            return None

    # kabutan の年次業績テーブルは id="finance_box"
    table = soup.find("table", id="finance_box")
    if table is None:
        # IDで見つからない場合は全テーブルから「最終益」を含むものを探す
        for t in soup.find_all("table"):
            if "最終益" in t.get_text():
                table = t
                break

    if table is None:
        return {"success": False, "error": "kabutan.jp から業績テーブルが見つかりませんでした。手動で入力してください。"}

    rows = table.find_all("tr")
    if not rows:
        return {"success": False, "error": "kabutan.jp のテーブルが空です。手動で入力してください。"}

    # ヘッダー行から列インデックスを特定
    header_cells = rows[0].find_all(["th", "td"])
    headers_text = [c.get_text(strip=True) for c in header_cells]

    profit_idx, eps_idx = None, None
    for i, h in enumerate(headers_text):
        if "最終益" in h:
            profit_idx = i
        if "1株益" in h:
            eps_idx = i

    if profit_idx is None or eps_idx is None:
        return {"success": False, "error": f"kabutan.jp の列が見つかりませんでした（列名: {headers_text}）。手動で入力してください。"}

    latest_actual_found = False

    for row in rows[1:]:
        cells = row.find_all(["td", "th"])
        if len(cells) <= max(profit_idx, eps_idx):
            continue

        period = cells[0].get_text(strip=True)
        profit = to_float(cells[profit_idx].get_text())
        eps = to_float(cells[eps_idx].get_text())

        # 「予」を含む行 = 会社予想（「実」「連」は除く）
        is_forecast = "予" in period and "実" not in period

        if is_forecast and profit is not None and result["forecast_net_income_m"] is None:
            result["forecast_net_income_m"] = int(profit)

        elif not is_forecast and not latest_actual_found and profit and eps and eps > 0:
            # 期中平均株式数 = 当期純利益(百万円) × 1,000,000 ÷ EPS(円)
            result["avg_shares"] = int(profit * 1_000_000 / eps)
            latest_actual_found = True

    if result["forecast_net_income_m"] or result["avg_shares"]:
        result["success"] = True
    else:
        return {"success": False, "error": "kabutan.jp から数値を読み取れませんでした。手動で入力してください。"}

    return result


def format_number(num, currency="JPY"):
    symbol = "¥" if currency == "JPY" else "$"
    abs_num = abs(num)
    sign = "-" if num < 0 else ""
    if abs_num >= 1e12:
        return f"{sign}{symbol}{abs_num/1e12:.2f}兆"
    elif abs_num >= 1e8:
        return f"{sign}{symbol}{abs_num/1e8:.2f}億"
    elif abs_num >= 1e4:
        return f"{sign}{symbol}{abs_num/1e4:.2f}万"
    else:
        return f"{sign}{symbol}{abs_num:,.0f}"


# ===== UI =====
st.title("キャッシュニュートラルPER 計算ツール")

with st.expander("使い方"):
    st.markdown(
        """
- **日本株**：証券コード + `.T` を入力　例）`7203.T`（トヨタ）、`3139.T`（ラクト・ジャパン）
- **米国株**：ティッカーシンボルをそのまま入力　例）`AAPL`
- 「データ取得」→ 「決算短信から自動取得」の順に押すと、来期予想利益と期中平均株式数が自動で入ります。
- 期中平均株式数は決算短信の数値と若干異なる場合があります。その場合は手動で修正してください。
"""
    )

# セッション状態の初期化
for key, default in [
    ("yahoo_result", None),
    ("avg_shares", None),
    ("forecast_net_income_m", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ---- ステップ1：Yahoo Finance からデータ取得 ----
col1, col2 = st.columns([4, 1])
with col1:
    ticker_input = st.text_input(
        "銘柄コード",
        placeholder="例: 3139.T　または　AAPL",
        label_visibility="collapsed",
    )
with col2:
    fetch_btn = st.button("データ取得", use_container_width=True, type="primary")

if fetch_btn:
    if not ticker_input.strip():
        st.warning("銘柄コードを入力してください。")
    else:
        with st.spinner("Yahoo Finance からデータを取得中..."):
            r = fetch_yahoo_data(ticker_input.strip().upper())
        if r["success"]:
            st.session_state.yahoo_result = r
            # 初期値をセット（まだ決算短信取得前）
            st.session_state.avg_shares = int(r["shares_outstanding"])
            st.session_state.forecast_net_income_m = (
                int(round(r["net_income"] / 1_000_000)) if r["net_income"] else 0
            )
        else:
            st.session_state.yahoo_result = None
            st.error(f"エラー: {r['error']}")
            st.info("日本株は末尾に `.T` をつけてください（例: `3139.T`）。")

# ---- ステップ2：決算短信から自動取得 + 手動修正 + 計算 ----
if st.session_state.yahoo_result:
    r = st.session_state.yahoo_result
    currency = r["currency"]
    symbol = "¥" if currency == "JPY" else "$"

    st.success(f"**{r['company_name']}** のデータを取得しました")
    st.caption(
        f"貸借対照表：{r['bs_date']}期末　|　損益計算書：{r['fin_date']}期末　（Yahoo Finance より）"
    )
    st.markdown("---")

    # 決算短信自動取得ボタン
    kessan_btn = st.button("決算短信から自動取得（株探）", type="secondary")
    if kessan_btn:
        with st.spinner("kabutan.jp（株探）から業績データを取得中..."):
            kr = fetch_kessan_auto(ticker_input.strip().upper())

        if kr["success"]:
            msgs = []
            if kr["avg_shares"]:
                st.session_state.avg_shares = kr["avg_shares"]
                msgs.append(f"期中平均株式数：**{kr['avg_shares']:,} 株**")
            if kr["forecast_net_income_m"]:
                st.session_state.forecast_net_income_m = kr["forecast_net_income_m"]
                msgs.append(f"来期予想当期純利益：**{kr['forecast_net_income_m']:,} 百万円**")

            if msgs:
                st.success("決算短信から取得しました：" + "　|　".join(msgs))
            else:
                st.warning(
                    "PDFは取得できましたが数値の自動読み取りに失敗しました。"
                    "下の入力欄に手動で入力してください。"
                )
        else:
            st.error(f"決算短信の取得に失敗しました：{kr['error']}")
            st.info("下の入力欄に決算短信の数値を手動で入力してください。")

    st.markdown("---")

    # 手動修正フィールド（期中平均株式数のみ）
    st.subheader("数値の確認・修正")
    st.markdown("期中平均株式数が決算短信と異なる場合はここで修正してください。")

    override_shares = st.number_input(
        "期中平均株式数（株）",
        min_value=1,
        value=st.session_state.avg_shares or 1,
        step=1,
        format="%d",
        help="決算短信「発行済株式数」欄の③期中平均株式数",
    )

    calc_btn = st.button("計算する", type="primary")

    if calc_btn:
        override_net_income = (st.session_state.forecast_net_income_m or 0) * 1_000_000

        if override_net_income == 0:
            st.error("当期純利益が取得できていません。先に「決算短信から自動取得」ボタンを押してください。")
        else:
            market_cap = r["stock_price"] * override_shares
            cash_adjustment = (
                r["current_assets"]
                + r["investment_securities"] * 0.7
                - r["total_liabilities"]
            )
            adjusted_market_cap = market_cap - cash_adjustment
            cash_neutral_per = adjusted_market_cap / override_net_income
            normal_per = market_cap / override_net_income

            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("キャッシュニュートラルPER", f"{cash_neutral_per:.1f} 倍")
            with c2:
                st.metric("通常PER（参考）", f"{normal_per:.1f} 倍")
            with c3:
                price_str = (
                    f"¥{r['stock_price']:,.0f}"
                    if currency == "JPY"
                    else f"${r['stock_price']:,.2f}"
                )
                st.metric("株価", price_str)

            st.markdown("---")
            st.subheader("計算の内訳")

            inv_adj = r["investment_securities"] * 0.7
            rows = [
                ("株価", f"{symbol}{r['stock_price']:,.0f}"),
                ("期中平均株式数", f"{override_shares:,} 株"),
                ("時価総額　(A)", format_number(market_cap, currency)),
                ("", ""),
                ("流動資産　(B)", format_number(r["current_assets"], currency)),
                ("投資有価証券　(C)", format_number(r["investment_securities"], currency)),
                ("投資有価証券 × 70%　(C×0.7)", format_number(inv_adj, currency)),
                ("総負債　(D)", format_number(r["total_liabilities"], currency)),
                ("調整額　(B + C×0.7 − D)", format_number(cash_adjustment, currency)),
                ("", ""),
                ("調整後時価総額　(A − 調整額)", format_number(adjusted_market_cap, currency)),
                ("当期純利益（来期通期予想）", format_number(override_net_income, currency)),
                ("★ キャッシュニュートラルPER", f"{cash_neutral_per:.2f} 倍"),
            ]

            df = pd.DataFrame(rows, columns=["項目", "値"])
            st.table(df)

            with st.expander("計算式について"):
                st.markdown(
                    """
**計算式：**
```
キャッシュニュートラルPER =
  { 株価 × 期中平均株式数 − (流動資産 + 投資有価証券×70% − 総負債) }
  ÷ 当期純利益（来期通期予想）
```

| 項目 | 出典 |
|------|------|
| 株価 | Yahoo Finance（リアルタイム） |
| 期中平均株式数 | kabutan.jp の実績値から算出（当期純利益 ÷ EPS） |
| 流動資産・投資有価証券・総負債 | Yahoo Finance（最新期末） |
| 当期純利益 | kabutan.jp の来期通期会社予想（親会社株主帰属） |
"""
                )

            if r["investment_securities"] == 0:
                st.info(
                    "投資有価証券のデータが自動取得できなかったため0円としています。"
                    "有価証券報告書等でご確認ください。"
                )
