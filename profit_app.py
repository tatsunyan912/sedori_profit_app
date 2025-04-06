import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="せどり実質利益計算アプリ", layout="centered")

st.title("せどり実質利益計算アプリ（ポイント率対応）")

# 入力欄
with st.form(key="profit_form"):
    st.subheader("商品情報の入力")
    item_name = st.text_input("商品名", value="")

    selling_price = st.number_input("販売価格（円）", min_value=0, value=0, step=100)
    cost_price = st.number_input("仕入れ価格（円）", min_value=0, value=0, step=100)
    shipping_cost = st.number_input("送料（円）", min_value=0, value=215, step=100)

    platform_fee_percent = st.number_input("販売手数料（%）", min_value=0.0, max_value=100.0, value=10.0, step=0.1)
    shop_point_percent = st.number_input("店舗ポイント還元率（%）", min_value=0.0, max_value=100.0, value=10.0, step=0.1)
    card_point_percent = st.number_input("クレカポイント還元率（%）", min_value=0.0, max_value=100.0, value=7.0, step=0.1)

    submitted = st.form_submit_button("計算して保存")

# データ保存リスト（セッションに保存）
if "saved_data" not in st.session_state:
    st.session_state.saved_data = []

if submitted:
    fee = selling_price * (platform_fee_percent / 100)
    total_cost = cost_price + shipping_cost
    cash_profit = selling_price - fee - total_cost
    point_profit = cost_price * ((shop_point_percent + card_point_percent) / 100)
    total_profit = cash_profit + point_profit

    # 利益率の計算
    cash_profit_rate = (cash_profit / total_cost * 100) if total_cost > 0 else 0
    total_profit_rate = (total_profit / total_cost * 100) if total_cost > 0 else 0

    # 結果表示
    st.subheader("計算結果")
    st.write(f"**現金利益：{cash_profit:,.0f} 円**")
    st.write(f"ポイント利益：{point_profit:,.0f} 円")
    st.write(f"総合利益：{total_profit:,.0f} 円")
    st.write(f"現金利益率：{cash_profit_rate:.2f}%")
    st.write(f"総合利益率（ポイント含む）：{total_profit_rate:.2f}%")

    # 保存データに追加
    st.session_state.saved_data.append({
        "入力日": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "商品名": item_name,
        "販売価格": selling_price,
        "仕入れ価格": cost_price,
        "送料": shipping_cost,
        "販売手数料（%）": platform_fee_percent,
        "ポイント合計（%）": shop_point_percent + card_point_percent,
        "現金利益": cash_profit,
        "ポイント利益": point_profit,
        "総合利益": total_profit,
        "現金利益率（%）": cash_profit_rate,
        "総合利益率（%）": total_profit_rate
    })

# データ表示
if st.session_state.saved_data:
    st.subheader("保存済みデータ一覧")
    df = pd.DataFrame(st.session_state.saved_data)
    st.dataframe(df, use_container_width=True)

    # Excel出力
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="利益データ")
        workbook = writer.book
        worksheet = writer.sheets["利益データ"]
        last_row = len(df) + 1
        worksheet.write(f"L{last_row + 1}", "現金利益合計")
        worksheet.write_formula(f"M{last_row + 1}", f"=SUM(H2:H{last_row})")

    st.download_button(
        label="エクセルでダウンロード",
        data=output.getvalue(),
        file_name="sedori_profit_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )