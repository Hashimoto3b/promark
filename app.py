import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def process_data(store_df, ad_dfs):
    store_df["日付"] = pd.to_datetime(store_df["日付"], errors="coerce")
    ad_df = pd.concat(ad_dfs, ignore_index=True)
    ad_df["日付"] = pd.to_datetime(ad_df.get("日") or ad_df.get("日付"), errors="coerce")
    merged = pd.merge(ad_df, store_df, on="日付", how="outer")

    # KPI算出
    merged["ROAS"] = merged["売上（円）"] / merged["Cost"]
    merged["CPA"] = merged["Cost"] / merged["CV"]
    merged["LTV"] = merged["売上（円）"] / merged["CV"]
    merged["ROI"] = (merged["売上（円）"] - merged["Cost"]) / merged["Cost"]

    BENCHMARKS = {"ROAS": 1.2, "CPA": 3000, "LTV": 6000, "ROI": 0.1}
    comments = []
    roas_avg = merged["ROAS"].mean(skipna=True)
    cpa_avg = merged["CPA"].mean(skipna=True)
    ltv_avg = merged["LTV"].mean(skipna=True)
    roi_avg = merged["ROI"].mean(skipna=True)

    if roas_avg < BENCHMARKS["ROAS"]:
        comments.append("ROASが業界平均を下回っています。ターゲティングや訴求強化を推奨します。")
    else:
        comments.append("ROASは業界平均以上です。現状の施策を維持・拡大を検討ください。")

    if cpa_avg > BENCHMARKS["CPA"]:
        comments.append("CPAが高めです。クリエイティブやLP改善を推奨します。")
    else:
        comments.append("CPAは業界平均以下で良好です。現状維持で効率化を。")

    if ltv_avg < BENCHMARKS["LTV"]:
        comments.append("LTVが低めです。リピート促進やクロスセルを強化しましょう。")
    else:
        comments.append("LTVは良好です。維持施策を継続しましょう。")

    if roi_avg < BENCHMARKS["ROI"]:
        comments.append("ROIが低く、投資回収が不十分です。抜本的な施策見直しを推奨します。")
    else:
        comments.append("ROIは業界平均以上です。現状施策を拡大可能です。")

    # Excel 出力
    wb = Workbook()
    ws = wb.active
    ws.title = "KPIレポート"
    for row in dataframe_to_rows(merged, index=False, header=True):
        ws.append(row)

    cws = wb.create_sheet("改善コメント")
    for c in comments:
        cws.append([c])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Webマーケ分析アプリ (ベンチマーク比較付き)")
    st.write("来店データとMETA広告データをアップロードしてください。")

    store_file = st.file_uploader("来店データファイル (Excel)", type="xlsx")
    ad_file = st.file_uploader("META広告データファイル (Excel)", type="xlsx")

    if store_file and ad_file:
        store_df = pd.read_excel(store_file)
        ad_sheets = pd.read_excel(ad_file, sheet_name=None)
        ad_dfs = list(ad_sheets.values())

        st.success("データを読み込みました。KPIを計算中です…")
        excel_output = process_data(store_df, ad_dfs)

        st.download_button("分析レポートをダウンロード", data=excel_output, file_name="マーケ分析レポート.xlsx")

if __name__ == "__main__":
    main()
