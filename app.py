import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # 啟用 3D 投影用的

# === 固定設定：依照你的計算模板 ===

SHEET_DETAIL = "細部點座標"
SHEET_CONTROL = "控制點 (ControlPoints)"  # 如果你後來改成「控制點」，這裡就改 "控制點"

COL_POINT = "點號"
COL_N = "N座標"
COL_E = "E座標"
COL_H = "H座標"


# === 工具函式：讀取上傳的 Excel ===

def load_points(xls, sheet_name: str) -> pd.DataFrame:
    """從指定工作表讀取三維座標資料（使用上傳的 Excel 檔）"""
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # 檢查欄位是否存在
    for col in [COL_POINT, COL_N, COL_E, COL_H]:
        if col not in df.columns:
            raise KeyError(f"在工作表「{sheet_name}」找不到欄位：{col}")

    df_clean = df.dropna(subset=[COL_N, COL_E, COL_H])
    return df_clean


def set_equal_3d_axes(ax, x, y, z):
    """讓 3D 圖比例一致"""
    x_min, x_max = x.min(), x.max()
    y_min, y_max = y.min(), y.max()
    z_min, z_max = z.min(), z.max()

    max_range = max(x_max - x_min, y_max - y_min, z_max - z_min) / 2.0

    x_mid = (x_max + x_min) / 2.0
    y_mid = (y_max + y_min) / 2.0
    z_mid = (z_max + z_min) / 2.0

    ax.set_xlim(x_mid - max_range, x_mid + max_range)
    ax.set_ylim(y_mid - max_range, y_mid + max_range)
    ax.set_zlim(z_mid - max_range, z_mid + max_range)


# === 畫平面圖 (N-E) ===

def plot_plan(detail_df: pd.DataFrame, control_df: pd.DataFrame | None = None, show_labels: bool = True):
    fig, ax = plt.subplots()

    # 細部點
    if not detail_df.empty:
        x = detail_df[COL_E]
        y = detail_df[COL_N]
        labels = detail_df[COL_POINT].astype(str)

        ax.scatter(x, y, s=10, marker="o", label="細部點")
        if show_labels:
            for xi, yi, label in zip(x, y, labels):
                ax.text(xi, yi, label, fontsize=6)

    # 控制點
    if control_df is not None and not control_df.empty:
        x = control_df[COL_E]
        y = control_df[COL_N]
        labels = control_df[COL_POINT].astype(str)

        ax.scatter(x, y, s=40, marker="^", label="控制點")
        if show_labels:
            for xi, yi, label in zip(x, y, labels):
                ax.text(xi, yi, label, fontsize=7, fontweight="bold")

    ax.set_xlabel("E (m)")
    ax.set_ylabel("N (m)")
    ax.set_aspect("equal", adjustable="box")
    ax.set_title("平面圖：細部點 + 控制點")
    ax.legend()

    return fig


# === 畫 3D 圖 (E, N, H) ===

def plot_3d(detail_df: pd.DataFrame, control_df: pd.DataFrame | None = None, show_labels: bool = False):
    fig = plt.figure()
    ax = fig.add_subplot(111, projection="3d")

    xs, ys, zs = [], [], []

    # 細部點
    if not detail_df.empty:
        x = detail_df[COL_E]
        y = detail_df[COL_N]
        z = detail_df[COL_H]
        labels = detail_df[COL_POINT].astype(str)

        ax.scatter(x, y, z, s=10, marker="o", label="細部點")
        if show_labels:
            for xi, yi, zi, label in zip(x, y, z, labels):
                ax.text(xi, yi, zi, label, fontsize=6)

        xs.append(x)
        ys.append(y)
        zs.append(z)

    # 控制點
    if control_df is not None and not control_df.empty:
        x = control_df[COL_E]
        y = control_df[COL_N]
        z = control_df[COL_H]
        labels = control_df[COL_POINT].astype(str)

        ax.scatter(x, y, z, s=40, marker="^", label="控制點")
        if show_labels:
            for xi, yi, zi, label in zip(x, y, z, labels):
                ax.text(xi, yi, zi, label, fontsize=7, fontweight="bold")

        xs.append(x)
        ys.append(y)
        zs.append(z)

    ax.set_xlabel("E (m)")
    ax.set_ylabel("N (m)")
    ax.set_zlabel("H (m)")
    ax.set_title("三維圖：細部點 + 控制點")
    ax.legend()

    if xs:
        x_all = pd.concat(xs)
        y_all = pd.concat(ys)
        z_all = pd.concat(zs)
        set_equal_3d_axes(ax, x_all, y_all, z_all)

    return fig


# === Streamlit App：測量可視化助手 ===

def main():
    st.set_page_config(page_title="測量可視化助手", layout="wide")

    st.title("📐 測量可視化助手")
    st.subheader("下載 Excel 計算模板")

with open("calculation template.xlsx", "rb") as f:
    st.download_button(
        label="📥 點我下載計算模板",
        data=f,
        file_name="calculation_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption("使用你的 Excel 計算模板，自動繪製平面與三維座標圖")

    uploaded_file = st.file_uploader(
        "請上傳使用『計算模板』填好的 Excel 檔 (.xlsx)",
        type=["xlsx"],
    )

    show_labels = st.checkbox("顯示點號標籤", value=True)

    if uploaded_file is None:
        st.info("請先上傳一個 Excel 檔案。")
        return

    try:
        # 讀兩個工作表
        detail_df = load_points(uploaded_file, SHEET_DETAIL)
    except Exception as e:
        st.error(f"讀取細部點座標失敗：{e}")
        return

    # 控制點可選
    try:
        control_df = load_points(uploaded_file, SHEET_CONTROL)
    except Exception:
        control_df = pd.DataFrame()
        st.warning("找不到控制點工作表或欄位，將只顯示細部點。")

    # 顯示資料表
    st.subheader("細部點座標表")
    st.dataframe(detail_df)

    if not control_df.empty:
        st.subheader("控制點座標表")
        st.dataframe(control_df)

    # 繪圖（左右兩欄）
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("平面圖 (N–E)")
        fig_plan = plot_plan(detail_df, control_df, show_labels=show_labels)
        st.pyplot(fig_plan)

    with col2:
        st.subheader("三維圖 (E–N–H)")
        fig_3d = plot_3d(detail_df, control_df, show_labels=False)
        st.pyplot(fig_3d)


if __name__ == "__main__":
    main()

