import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt  # 只留著以後備用
from mpl_toolkits.mplot3d import Axes3D  # 目前沒用到，但保留
import plotly.express as px

# ==========================
# 基本設定（依照你的 Excel 模板）
# ==========================

SHEET_DETAIL = "細部點座標"
SHEET_CONTROL = "控制點 (ControlPoints)"  # 如果工作表叫「控制點」，改成 "控制點"

COL_POINT = "點號"
COL_N = "N座標"
COL_E = "E座標"
COL_H = "H座標"


# ==========================
# 工具函式：讀取 Excel
# ==========================

def load_points(xls, sheet_name: str) -> pd.DataFrame:
    """
    從指定工作表讀取三維座標資料
    xls 可以是上傳的檔案物件（streamlit file_uploader 給的）
    """
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # 檢查欄位是否存在
    for col in [COL_POINT, COL_N, COL_E, COL_H]:
        if col not in df.columns:
            raise KeyError(f"在工作表「{sheet_name}」找不到欄位：{col}")

    return df  # 不在這裡 dropna，畫圖前再處理


# ==========================
# 繪圖：平面圖 (N–E) - 使用 plotly，可放大
# ==========================

def plot_plan_interactive(detail_df: pd.DataFrame,
                          control_df: pd.DataFrame | None = None,
                          show_labels: bool = True):
    """平面 N–E 圖（plotly 版，可放大）"""

    # 只取有 N/E 的點
    detail_valid = detail_df.dropna(subset=[COL_N, COL_E]) if detail_df is not None else pd.DataFrame()
    control_valid = control_df.dropna(subset=[COL_N, COL_E]) if (control_df is not None and not control_df.empty) else pd.DataFrame()

    # 組合兩種點成一個 DataFrame，方便 plotly 上色
    frames = []
    if not detail_valid.empty:
        df_d = detail_valid.copy()
        df_d["點類型"] = "細部點"
        frames.append(df_d)
    if not control_valid.empty:
        df_c = control_valid.copy()
        df_c["點類型"] = "控制點"
        frames.append(df_c)

    if not frames:
        return None

    all_points = pd.concat(frames, ignore_index=True)

    # hover 資訊
    hover_data = {
        COL_POINT: True,
        COL_N: True,
        COL_E: True,
        COL_H: True,
        "點類型": True,
    }

    fig = px.scatter(
        all_points,
        x=COL_E,
        y=COL_N,
        color="點類型",
        hover_name=COL_POINT,
        hover_data=hover_data,
        symbol="點類型",
    )

    fig.update_layout(
        title="平面圖：細部點 + 控制點（可用滑鼠/手指框選放大）",
        xaxis_title="E (m)",
        yaxis_title="N (m)",
        yaxis_scaleanchor="x",  # 保持比例 1:1
        legend_title="點類型",
        height=600,
    )

    # 如果不要在圖上顯示標籤，只保留 hover
    if not show_labels:
        return fig

    # 顯示固定標籤（在點旁邊印點號）
    fig.update_traces(
        text=all_points[COL_POINT],
        textposition="top center",
        textfont=dict(size=9),
        mode="markers+text",
    )

    return fig


# ==========================
# 繪圖：三維圖 (E–N–H) - 使用 plotly，可放大旋轉
# ==========================

def plot_3d_interactive(detail_df: pd.DataFrame,
                        control_df: pd.DataFrame | None = None):
    """三維圖：細部點 + 控制點（plotly 版，可旋轉、放大）"""

    detail_valid = detail_df.dropna(subset=[COL_N, COL_E, COL_H]) if detail_df is not None else pd.DataFrame()
    control_valid = control_df.dropna(subset=[COL_N, COL_E, COL_H]) if (control_df is not None and not control_df.empty) else pd.DataFrame()

    frames = []
    if not detail_valid.empty:
        df_d = detail_valid.copy()
        df_d["點類型"] = "細部點"
        frames.append(df_d)
    if not control_valid.empty:
        df_c = control_valid.copy()
        df_c["點類型"] = "控制點"
        frames.append(df_c)

    if not frames:
        return None

    all_points = pd.concat(frames, ignore_index=True)

    hover_data = {
        COL_POINT: True,
        COL_N: True,
        COL_E: True,
        COL_H: True,
        "點類型": True,
    }

    fig = px.scatter_3d(
        all_points,
        x=COL_E,
        y=COL_N,
        z=COL_H,
        color="點類型",
        hover_name=COL_POINT,
        hover_data=hover_data,
        symbol="點類型",
    )

    fig.update_layout(
        title="三維圖：細部點 + 控制點（可拖曳旋轉 / 滾輪放大）",
        scene=dict(
            xaxis_title="E (m)",
            yaxis_title="N (m)",
            zaxis_title="H (m)",
        ),
        legend_title="點類型",
        height=650,
    )

    return fig


# ==========================
# Streamlit App：測量可視化助手
# ==========================

def main():
    st.set_page_config(page_title="測量可視化助手", layout="wide")

    st.title("📐 測量可視化助手")
    st.caption("使用你的 Excel 計算模板，自動繪製可放大、可旋轉的平面與三維座標圖")

    # --- 模板下載 ---
    st.subheader("下載 Excel 計算模板")
    try:
        with open("calculation template.xlsx", "rb") as f:
            st.download_button(
                label="📥 點我下載計算模板",
                data=f,
                file_name="calculation_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except FileNotFoundError:
        st.warning("⚠ 找不到 calculation template.xlsx，請確認檔案有放在與 app.py 同一資料夾。")

    st.markdown("---")

    # --- 上傳 Excel ---
    st.subheader("上傳計算成果 Excel 檔")
    uploaded_file = st.file_uploader(
        "請上傳依照『計算模板』填好的 .xlsx 檔案",
        type=["xlsx"]
    )

    show_labels = st.checkbox("平面圖顯示點號標籤", value=True)

    if uploaded_file is None:
        st.info("請先上傳 Excel 檔案後再進行繪圖。")
        return

    # --- 讀取細部點 ---
    try:
        detail_df = load_points(uploaded_file, SHEET_DETAIL)
    except Exception as e:
        st.error(f"讀取細部點座標失敗：{e}")
        return

    # --- 讀取控制點（可選） ---
    try:
        control_df = load_points(uploaded_file, SHEET_CONTROL)
    except Exception:
        control_df = pd.DataFrame()
        st.warning("⚠ 未找到控制點工作表或欄位，將只顯示細部點。")

    # --- 顯示資料表 ---
    st.subheader("細部點座標表")
    st.dataframe(detail_df, use_container_width=True)

    if not control_df.empty:
        st.subheader("控制點座標表")
        st.dataframe(control_df, use_container_width=True)

    st.markdown("---")

    # --- 繪圖（左右兩欄，使用 plotly_chart，可以放大） ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("平面圖 (N–E)")
        fig_plan = plot_plan_interactive(detail_df, control_df, show_labels=show_labels)
        if fig_plan is None:
            st.warning("沒有有效的細部點 / 控制點可以繪製平面圖。請確認 N/E 座標有計算完成。")
        else:
            st.plotly_chart(fig_plan, use_container_width=True)

    with col2:
        st.subheader("三維圖 (E–N–H)")
        fig_3d = plot_3d_interactive(detail_df, control_df)
        if fig_3d is None:
            st.warning("沒有有效的細部點 / 控制點可以繪製三維圖。請確認 N/E/H 座標有計算完成。")
        else:
            st.plotly_chart(fig_3d, use_container_width=True)


if __name__ == "__main__":
    main()


