import re

import streamlit as st
import pandas as pd
import plotly.express as px

# ==========================
# 基本設定（依照你的 Excel 模板）
# ==========================

SHEET_DETAIL = "細部點座標"
SHEET_CONTROL = "控制點 (ControlPoints)"  # 如果工作表叫「控制點」，就改成 "控制點"

COL_POINT = "點號"
COL_N = "N座標"
COL_E = "E座標"
COL_H = "H座標"


# ==========================
# 工具函式：讀取 Excel
# ==========================

def load_points(xls, sheet_name: str) -> pd.DataFrame:
    """
    從指定工作表讀取三維座標資料。
    xls 可以是上傳的檔案物件（streamlit file_uploader 給的）。
    """
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # 檢查欄位是否存在
    for col in [COL_POINT, COL_N, COL_E, COL_H]:
        if col not in df.columns:
            raise KeyError(f"在工作表「{sheet_name}」找不到欄位：{col}")

    return df  # 不在這裡 dropna，畫圖前再處理


# ==========================
# 細部點分類：依點號判斷點類型 & 標籤
# ==========================

def classify_detail_points(detail_df: pd.DataFrame) -> pd.DataFrame:
    """
    根據點號內容分類細部點：
    S -> 補點（深藍）
    B -> 建物（淺藍）
    R -> 道路（淺灰）
    L -> 路燈（黃色）
    T -> 樹木（深綠）
    F -> 花圃（淺綠）
    O -> 其他（淺紫）
    其餘 -> 細部點（預設）
    """
    if detail_df is None or detail_df.empty:
        return detail_df

    df = detail_df.copy()
    df["點類型"] = "[細部點]"
    pt_str = df[COL_POINT].astype(str)

    # 依序分類，只有目前還是 [細部點] 的才覆蓋
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("S", case=False, na=False), "點類型"] = "[補點]"
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("B", case=False, na=False), "點類型"] = "[建物]"
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("R", case=False, na=False), "點類型"] = "[道路]"
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("L", case=False, na=False), "點類型"] = "[路燈]"
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("T", case=False, na=False), "點類型"] = "[樹木]"
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("F", case=False, na=False), "點類型"] = "[花圃]"
    mask = df["點類型"] == "[細部點]"
    df.loc[mask & pt_str.str.contains("O", case=False, na=False), "點類型"] = "[其他]"

    return df


# ==========================
# 命名工具：從 B 點推算下一個編號，不重複
# ==========================

def infer_naming_style_and_next_indices(base_name: str,
                                        all_names: pd.Series,
                                        c: int):
    """
    從 B 點點號推斷命名風格：
      - T-1, T-2 -> 產生 T-3, T-4...
      - T1, T2   -> 產生 T3, T4...
    從 all_names 中找出同風格的最大編號，然後連續往後 C 個，保證不重複。

    回傳: (style, prefix, [index1, index2, ...])
        style: 'hyphen' 或 'plain'
        prefix: 例如 'T'
    """
    name = str(base_name)

    # 嘗試 hyphen 風格: PREFIX-N
    m_hyphen = re.match(r"^(.*?)-(\d+)$", name)
    m_plain = re.match(r"^(.*?)(\d+)$", name)

    style = None
    prefix = None

    if m_hyphen:
        style = "hyphen"
        prefix = m_hyphen.group(1)
    elif m_plain:
        style = "plain"
        prefix = m_plain.group(1)
    else:
        # 沒有數字，預設用 plain 風格，從 1 開始
        style = "plain"
        prefix = name

    all_names_str = all_names.astype(str)

    # 找全部相同風格的現有編號
    existing_indices = []

    if style == "hyphen":
        pattern = re.compile(rf"^{re.escape(prefix)}-(\d+)$")
        for s in all_names_str:
            m = pattern.match(s)
            if m:
                existing_indices.append(int(m.group(1)))
    else:  # plain
        pattern = re.compile(rf"^{re.escape(prefix)}(\d+)$")
        for s in all_names_str:
            m = pattern.match(s)
            if m:
                existing_indices.append(int(m.group(1)))

    max_idx = max(existing_indices) if existing_indices else 0
    used_names = set(all_names_str)

    indices = []
    cur = max_idx
    while len(indices) < c:
        cur += 1
        candidate = f"{prefix}-{cur}" if style == "hyphen" else f"{prefix}{cur}"
        if candidate in used_names:
            # 理論上不會常發生，但還是保險一下
            continue
        indices.append(cur)
        used_names.add(candidate)

    return style, prefix, indices


# ==========================
# 支距法：產生新點（繼承 A/B 類型 & 顏色）
# ==========================

def generate_offset_points(all_points: pd.DataFrame,
                           point_a: str,
                           point_b: str,
                           k: float,
                           c: int) -> pd.DataFrame:
    """
    支距法：
    - 從 A、B 兩點，沿著 AB 方向，自 B 起每次 K 倍 AB 向量，重複 C 次。
    - 新點點號依據 B 點命名風格，延續編號，不與任何既有點號重複。
    - 新點的「點類型」：
        若 A、B 類型相同 -> 使用該類型；
        若不同 -> 使用 B 的類型。
    """

    # 確保有 "點類型" 欄位（控制點和細部點都應該已設定）
    if "點類型" not in all_points.columns:
        all_points = all_points.copy()
        all_points["點類型"] = "[細部點]"

    row_a = all_points[all_points[COL_POINT] == point_a]
    row_b = all_points[all_points[COL_POINT] == point_b]

    if row_a.empty or row_b.empty:
        raise ValueError("找不到指定的點 A 或點 B")

    Na, Ea, Ha = float(row_a[COL_N].iloc[0]), float(row_a[COL_E].iloc[0]), float(row_a[COL_H].iloc[0])
    Nb, Eb, Hb = float(row_b[COL_N].iloc[0]), float(row_b[COL_E].iloc[0]), float(row_b[COL_H].iloc[0])

    dN = Nb - Na
    dE = Eb - Ea
    dH = Hb - Ha

    type_a = row_a["點類型"].iloc[0]
    type_b = row_b["點類型"].iloc[0]
    if type_a == type_b:
        new_type = type_a
    else:
        # 若 A、B 類型不同，以 B 為主
        new_type = type_b

    base_name = str(row_b[COL_POINT].iloc[0])
    style, prefix, indices = infer_naming_style_and_next_indices(
        base_name,
        all_points[COL_POINT],
        c
    )

    records = []

    for idx in indices:
        # 注意：這裡 factor 依「第幾個新點」排，跟 idx 數字無關
        factor = k * (len(records) + 1)
        Ni = Nb + factor * dN
        Ei = Eb + factor * dE
        Hi = Hb + factor * dH

        if style == "hyphen":
            pt_name = f"{prefix}-{idx}"
        else:
            pt_name = f"{prefix}{idx}"

        records.append({
            COL_POINT: pt_name,
            COL_N: Ni,
            COL_E: Ei,
            COL_H: Hi,
            "點類型": new_type,
        })

    return pd.DataFrame.from_records(records)


# ==========================
# 繪圖：平面圖 (N–E) - plotly，可放大
# ==========================

def plot_plan_interactive(detail_df: pd.DataFrame,
                          control_df: pd.DataFrame | None = None,
                          offset_df: pd.DataFrame | None = None,
                          show_labels: bool = True):
    """平面 N–E 圖（plotly 版，可放大）"""

    # 細部點分類 + 過濾有效
    if detail_df is not None and not detail_df.empty:
        detail_df = classify_detail_points(detail_df)
        detail_valid = detail_df.dropna(subset=[COL_N, COL_E])
    else:
        detail_valid = pd.DataFrame()

    # 控制點：標記類型
    if control_df is not None and not control_df.empty:
        control_valid = control_df.dropna(subset=[COL_N, COL_E]).copy()
        control_valid["點類型"] = "[控制點]"
    else:
        control_valid = pd.DataFrame()

    # 支距點（已含點類型）
    if offset_df is not None and not offset_df.empty:
        offset_valid = offset_df.dropna(subset=[COL_N, COL_E]).copy()
    else:
        offset_valid = pd.DataFrame()

    frames = []
    if not detail_valid.empty:
        frames.append(detail_valid)
    if not control_valid.empty:
        frames.append(control_valid)
    if not offset_valid.empty:
        frames.append(offset_valid)

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

    # 顏色與符號對照
    color_map = {
        "[控制點]": "#ff8800",  # 橘色
        "[補點]": "#003f7f",   # 深藍
        "[建物]": "#4fa3ff",   # 淺藍
        "[道路]": "#c0c0c0",   # 淺灰
        "[路燈]": "#ffd447",   # 黃
        "[樹木]": "#006400",   # 深綠
        "[花圃]": "#7ed957",   # 淺綠
        "[其他]": "#c792ea",   # 淺紫
        "[細部點]": "#888888", # 未分類細部點
    }

    symbol_map = {
        "[控制點]": "triangle-up",  # 橘色三角形
        "[補點]": "circle",
        "[建物]": "circle",
        "[道路]": "circle",
        "[路燈]": "circle",
        "[樹木]": "circle",
        "[花圃]": "circle",
        "[其他]": "circle",
        "[細部點]": "circle",
    }

    fig = px.scatter(
        all_points,
        x=COL_E,
        y=COL_N,
        color="點類型",
        symbol="點類型",
        hover_name=COL_POINT,
        hover_data=hover_data,
        text=COL_POINT,              # 🔹 每個點顯示自己點號
        color_discrete_map=color_map,
        symbol_map=symbol_map,
    )

    fig.update_layout(
        title="平面圖：控制點 + 細部點 + 支距點（可縮放拖曳）",
        xaxis_title="E (m)",
        yaxis_title="N (m)",
        yaxis_scaleanchor="x",  # 保持比例 1:1
        legend_title="點類型",
        height=600,
    )

    if show_labels:
        fig.update_traces(
            textposition="top center",
            textfont=dict(size=9),
            mode="markers+text",
        )
    else:
        # 不顯示文字只保留點
        fig.update_traces(
            text=None,
            mode="markers",
        )

    return fig


# ==========================
# 繪圖：三維圖 (E–N–H) - plotly，可旋轉
# ==========================

def plot_3d_interactive(detail_df: pd.DataFrame,
                        control_df: pd.DataFrame | None = None,
                        offset_df: pd.DataFrame | None = None):
    """三維圖：控制點 + 細部點 + 支距點（plotly，可旋轉、放大）"""

    if detail_df is not None and not detail_df.empty:
        detail_df = classify_detail_points(detail_df)
        detail_valid = detail_df.dropna(subset=[COL_N, COL_E, COL_H])
    else:
        detail_valid = pd.DataFrame()

    if control_df is not None and not control_df.empty:
        control_valid = control_df.dropna(subset=[COL_N, COL_E, COL_H]).copy()
        control_valid["點類型"] = "[控制點]"
    else:
        control_valid = pd.DataFrame()

    if offset_df is not None and not offset_df.empty:
        offset_valid = offset_df.dropna(subset=[COL_N, COL_E, COL_H]).copy()
    else:
        offset_valid = pd.DataFrame()

    frames = []
    if not detail_valid.empty:
        frames.append(detail_valid)
    if not control_valid.empty:
        frames.append(control_valid)
    if not offset_valid.empty:
        frames.append(offset_valid)

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

    color_map = {
        "[控制點]": "#ff8800",
        "[補點]": "#003f7f",
        "[建物]": "#4fa3ff",
        "[道路]": "#c0c0c0",
        "[路燈]": "#ffd447",
        "[樹木]": "#006400",
        "[花圃]": "#7ed957",
        "[其他]": "#c792ea",
        "[細部點]": "#888888",
    }

    symbol_map = {
        "[控制點]": "triangle-up",
        "[補點]": "circle",
        "[建物]": "circle",
        "[道路]": "circle",
        "[路燈]": "circle",
        "[樹木]": "circle",
        "[花圃]": "circle",
        "[其他]": "circle",
        "[細部點]": "circle",
    }

    fig = px.scatter_3d(
        all_points,
        x=COL_E,
        y=COL_N,
        z=COL_H,
        color="點類型",
        symbol="點類型",
        hover_name=COL_POINT,
        hover_data=hover_data,
        color_discrete_map=color_map,
        symbol_map=symbol_map,
    )

    # 3D 互動設定：
    # - camera.up = Z 軸朝上
    # - dragmode = "turntable"：類似「Z 軸始終向上旋轉」的模式
    fig.update_layout(
        title="三維圖：控制點 + 細部點 + 支距點（可旋轉 / 縮放）",
        scene=dict(
            xaxis_title="E (m)",
            yaxis_title="N (m)",
            zaxis_title="H (m)",
            aspectmode="data",
            camera=dict(up=dict(x=0, y=0, z=1)),
            dragmode="turntable",
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

    if "offset_points" not in st.session_state:
        st.session_state["offset_points"] = pd.DataFrame(
            columns=[COL_POINT, COL_N, COL_E, COL_H, "點類型"]
        )

    st.title("📐 測量可視化助手")
    st.caption("使用 Excel 計算模板，自動繪製可放大、可旋轉的平面與三維座標圖（含支距法）")

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
        detail_df_raw = load_points(uploaded_file, SHEET_DETAIL)
    except Exception as e:
        st.error(f"讀取細部點座標失敗：{e}")
        return

    # --- 讀取控制點（可選） ---
    try:
        control_df_raw = load_points(uploaded_file, SHEET_CONTROL)
    except Exception:
        control_df_raw = pd.DataFrame()
        st.warning("⚠ 未找到控制點工作表或欄位，將只顯示細部點。")

    # --- 顯示資料表 ---
    st.subheader("細部點座標表")
    st.dataframe(detail_df_raw, use_container_width=True)

    if not control_df_raw.empty:
        st.subheader("控制點座標表")
        st.dataframe(control_df_raw, use_container_width=True)

    # --- 準備給支距法用的「已分類所有點」 ---
    detail_classified = classify_detail_points(detail_df_raw) if not detail_df_raw.empty else pd.DataFrame()
    if not control_df_raw.empty:
        control_classified = control_df_raw.copy()
        control_classified["點類型"] = "[控制點]"
    else:
        control_classified = pd.DataFrame()

    existing_offset = st.session_state["offset_points"]
    if not detail_classified.empty or not control_classified.empty or not existing_offset.empty:
        all_points_for_offset = pd.concat(
            [df for df in [detail_classified, control_classified, existing_offset] if not df.empty],
            ignore_index=True
        )
    else:
        all_points_for_offset = pd.DataFrame()

    st.markdown("---")
    st.subheader("支距法產生新點")

    # 支距法：目前依「細部點座標」的點號做 A、B 選擇
    point_choices = detail_df_raw[COL_POINT].astype(str).tolist()

    if len(point_choices) < 2:
        st.info("細部點少於兩點，無法執行支距法。")
        offset_df = st.session_state["offset_points"]
    else:
        col_a, col_b = st.columns(2)
        with col_a:
            point_a = st.selectbox("起點 A", point_choices, key="offset_A")
        with col_b:
            point_b = st.selectbox("終點 B", point_choices, key="offset_B")

        col_k, col_c = st.columns(2)
        with col_k:
            k = st.number_input("K 倍距離", min_value=0.0, value=1.0, step=0.1)
        with col_c:
            c = st.number_input("C 次（要生成幾個點）", min_value=1, max_value=100, value=3, step=1)

        if st.button("執行支距法並產生新點"):
            try:
                if all_points_for_offset.empty:
                    st.error("目前沒有可用的點資料供支距法使用。")
                    offset_df = st.session_state["offset_points"]
                else:
                    new_offset = generate_offset_points(all_points_for_offset, point_a, point_b, k, c)
                    # 新產生的支距點與既有支距點合併，避免覆蓋
                    offset_df = pd.concat(
                        [existing_offset, new_offset],
                        ignore_index=True
                    )
                    st.session_state["offset_points"] = offset_df
                    st.success(f"已從 {point_a} → {point_b} 方向產生 {len(new_offset)} 個支距點。")
            except Exception as e:
                st.error(f"支距法計算失敗：{e}")
                offset_df = st.session_state["offset_points"]
        else:
            offset_df = st.session_state["offset_points"]

    if not st.session_state["offset_points"].empty:
        st.write("目前所有支距法產生的點：")
        st.dataframe(st.session_state["offset_points"], use_container_width=True)

    st.markdown("---")

    # --- 繪圖（左右兩欄，使用 plotly_chart，可以放大） ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("平面圖 (N–E)")
        fig_plan = plot_plan_interactive(
            detail_df_raw,
            control_df_raw,
            offset_df=st.session_state["offset_points"],
            show_labels=show_labels,
        )
        if fig_plan is None:
            st.warning("沒有有效的細部點 / 控制點可以繪製平面圖。請確認 N/E 座標有計算完成。")
        else:
            st.plotly_chart(fig_plan, use_container_width=True)

    with col2:
        st.subheader("三維圖 (E–N–H)")
        fig_3d = plot_3d_interactive(
            detail_df_raw,
            control_df_raw,
            offset_df=st.session_state["offset_points"],
        )
        if fig_3d is None:
            st.warning("沒有有效的細部點 / 控制點可以繪製三維圖。請確認 N/E/H 座標有計算完成。")
        else:
            st.plotly_chart(fig_3d, use_container_width=True)
            st.caption("滑鼠拖曳旋轉、滾輪縮放。預設為 Z 軸朝上的旋轉模式（turntable）。")


if __name__ == "__main__":
    main()
