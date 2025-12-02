import re
import math
from typing import Dict, Tuple, Optional, List

import streamlit as st
import pandas as pd
import plotly.express as px


# ==========================
# 基本設定：欄位名稱
# ==========================

COL_POINT = "點號"
COL_N = "N座標"
COL_E = "E座標"
COL_H = "H座標"


# ==========================
# 自動偵測工作表
# ==========================

def auto_detect_sheets(xls_file) -> Tuple[pd.DataFrame, Optional[pd.DataFrame], str, Optional[str]]:
    """
    自動偵測上傳的 Excel 裡：
      - 哪一張是「細部點」工作表
      - 哪一張是「控制點」工作表（可有可無）

    規則：
      1) 只考慮同時擁有 COL_POINT, COL_N, COL_E, COL_H 四欄的工作表
      2) 工作表名稱包含「細部 / detail」優先當細部點
         名稱包含「控制 / control」優先當控制點
      3) 若還是不明，第一個符合條件的當細部點，第二個當控制點（如果有）

    回傳：(detail_df, control_df_or_None, detail_name, control_name_or_None)
    """
    xls = pd.ExcelFile(xls_file)
    candidates: Dict[str, pd.DataFrame] = {}

    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        if all(c in df.columns for c in [COL_POINT, COL_N, COL_E, COL_H]):
            candidates[name] = df

    if not candidates:
        raise ValueError("找不到同時包含「點號 / N座標 / E座標 / H座標」欄位的工作表。")

    detail_name = None
    control_name = None

    # 優先依名稱判斷
    for name in candidates.keys():
        lname = name.lower()
        if detail_name is None and ("細部" in name or "detail" in lname):
            detail_name = name
        if control_name is None and ("控制" in name or "control" in lname):
            control_name = name

    # 仍未決定時，用順序填補
    names_list = list(candidates.keys())
    if detail_name is None:
        detail_name = names_list[0]
    if control_name is None and len(names_list) >= 2:
        if names_list[1] != detail_name:
            control_name = names_list[1]

    detail_df = candidates[detail_name]
    control_df = candidates[control_name] if control_name is not None else None

    return detail_df, control_df, detail_name, control_name


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
# 命名工具：從起始點推算下一個編號，不重複
# ==========================

def infer_naming_style_and_next_indices(base_name: str,
                                        all_names: pd.Series,
                                        c: int):
    """
    從起始點點號推斷命名風格：
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
        pattern = re.compile(r"^" + re.escape(prefix) + r"-(\d+)$")
        for s in all_names_str:
            m = pattern.match(s)
            if m:
                existing_indices.append(int(m.group(1)))
    else:  # plain
        pattern = re.compile(r"^" + re.escape(prefix) + r"(\d+)$")
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
            continue
        indices.append(cur)
        used_names.add(candidate)

    return style, prefix, indices


# ==========================
# 支距法：以兩點距離 + NESW 方向
# ==========================

def compute_distance(all_points: pd.DataFrame, p1: str, p2: str) -> float:
    row1 = all_points[all_points[COL_POINT] == p1]
    row2 = all_points[all_points[COL_POINT] == p2]
    if row1.empty or row2.empty:
        raise ValueError("找不到距離基準點。")

    N1, E1 = float(row1[COL_N].iloc[0]), float(row1[COL_E].iloc[0])
    N2, E2 = float(row2[COL_N].iloc[0]), float(row2[COL_E].iloc[0])
    dN = N2 - N1
    dE = E2 - E1
    return math.sqrt(dN ** 2 + dE ** 2)


def generate_offset_points_directional(all_points: pd.DataFrame,
                                       dist_p1: str,
                                       dist_p2: str,
                                       start_point: str,
                                       direction: str,
                                       k: float,
                                       c: int) -> pd.DataFrame:
    """
    新版支距法：
      1) 先選兩點 dist_p1, dist_p2 計算距離 D
      2) 選起始點 start_point
      3) 選方向 direction ∈ {N, E, S, W}
      4) 設定 K 倍距離、C 次
         每一新點與前一點距離 = D * K，方向為 NESW

    新點的點號：
      - 依起始點 start_point 的命名風格（T-1/T1）往後編
      - 不與任何既有點號重複

    新點的點類型：
      - 與起始點相同（顏色和標籤一致）
    """

    if "點類型" not in all_points.columns:
        all_points = all_points.copy()
        all_points["點類型"] = "[細部點]"

    # 距離 D
    D = compute_distance(all_points, dist_p1, dist_p2)

    # 起始點資訊
    row_s = all_points[all_points[COL_POINT] == start_point]
    if row_s.empty:
        raise ValueError("找不到起始點。")

    Ns, Es, Hs = float(row_s[COL_N].iloc[0]), float(row_s[COL_E].iloc[0]), float(row_s[COL_H].iloc[0])
    start_type = row_s["點類型"].iloc[0]
    base_name = str(row_s[COL_POINT].iloc[0])

    style, prefix, indices = infer_naming_style_and_next_indices(
        base_name,
        all_points[COL_POINT],
        c
    )

    # 方向單位向量（只考慮平面 N, E）
    dir_map = {
        "N": (1.0, 0.0),
        "S": (-1.0, 0.0),
        "E": (0.0, 1.0),
        "W": (0.0, -1.0),
    }
    if direction not in dir_map:
        raise ValueError("方向必須為 N、E、S 或 W。")

    uN, uE = dir_map[direction]

    records = []
    cur_N, cur_E, cur_H = Ns, Es, Hs

    for idx in indices:
        step = D * k  # 每一段的長度
        cur_N += uN * step
        cur_E += uE * step
        cur_H = Hs  # 預設高度不變

        if style == "hyphen":
            pt_name = f"{prefix}-{idx}"
        else:
            pt_name = f"{prefix}{idx}"

        records.append({
            COL_POINT: pt_name,
            COL_N: cur_N,
            COL_E: cur_E,
            COL_H: cur_H,
            "點類型": start_type,
        })

    return pd.DataFrame.from_records(records)


# ==========================
# 繪圖：平面圖 (N–E) - plotly，可放大
# ==========================

def plot_plan_interactive(detail_df: pd.DataFrame,
                          control_df: Optional[pd.DataFrame],
                          offset_df: Optional[pd.DataFrame],
                          show_labels: bool,
                          allowed_types: List[str]):
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

    if allowed_types:
        all_points = all_points[all_points["點類型"].isin(allowed_types)]
        if all_points.empty:
            return None

    hover_data = {
        COL_POINT: True,
        COL_N: True,
        COL_E: True,
        COL_H: True,
        "點類型": True,
    }

    # 顏色與符號對照（2D）
    base_color_map = {
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

    base_symbol_map = {
        "[控制點]": "triangle-up",  # 2D 可以用三角形
        "[補點]": "circle",
        "[建物]": "circle",
        "[道路]": "circle",
        "[路燈]": "circle",
        "[樹木]": "circle",
        "[花圃]": "circle",
        "[其他]": "circle",
        "[細部點]": "circle",
    }

    used_types = all_points["點類型"].astype(str).unique().tolist()
    color_map = {t: base_color_map.get(t, "#000000") for t in used_types}
    symbol_map = {t: base_symbol_map.get(t, "circle") for t in used_types}

    fig = px.scatter(
        all_points,
        x=COL_E,
        y=COL_N,
        color="點類型",
        symbol="點類型",
        hover_name=COL_POINT,
        hover_data=hover_data,
        text=COL_POINT,
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
        fig.update_traces(
            text=None,
            mode="markers",
        )

    return fig


# ==========================
# 繪圖：三維圖 (E–N–H) - plotly，可旋轉
# ==========================

def plot_3d_interactive(detail_df: pd.DataFrame,
                        control_df: Optional[pd.DataFrame],
                        offset_df: Optional[pd.DataFrame],
                        allowed_types: List[str]):
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

    if allowed_types:
        all_points = all_points[all_points["點類型"].isin(allowed_types)]
        if all_points.empty:
            return None

    hover_data = {
        COL_POINT: True,
        COL_N: True,
        COL_E: True,
        COL_H: True,
        "點類型": True,
    }

    base_color_map = {
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

    # 3D 的 symbol 只能用這幾種：circle, circle-open, cross,
    # diamond, diamond-open, square, square-open, x
    base_symbol_map = {
        "[控制點]": "square-open",  # 3D 用方框代替三角形
        "[補點]": "circle",
        "[建物]": "circle",
        "[道路]": "circle",
        "[路燈]": "circle",
        "[樹木]": "circle",
        "[花圃]": "circle",
        "[其他]": "circle",
        "[細部點]": "circle",
    }

    used_types = all_points["點類型"].astype(str).unique().tolist()
    color_map = {t: base_color_map.get(t, "#000000") for t in used_types}
    symbol_map = {t: base_symbol_map.get(t, "circle") for t in used_types}

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

    # 3D 互動設定
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
# 匯出 Excel：把目前的細部點 + 控制點 + 支距點寫出去
# ==========================

def export_to_excel(detail_df: pd.DataFrame,
                    control_df: Optional[pd.DataFrame],
                    offset_df: Optional[pd.DataFrame]) -> bytes:
    """
    產生一份新的 Excel：
      - 工作表「細部點座標」：detail_df + offset_df（去掉 點類型 欄位）
      - 工作表「控制點」：control_df（若有，同樣去掉 點類型）
    回傳：Excel 檔案的位元組（給 st.download_button 用）
    """
    from io import BytesIO

    output = BytesIO()

    # 準備細部點
    detail_out = detail_df.copy()
    if "點類型" in detail_out.columns:
        detail_out = detail_out.drop(columns=["點類型"])

    # 支距點加入細部點
    if offset_df is not None and not offset_df.empty:
        offset_out = offset_df.copy()
        if "點類型" in offset_out.columns:
            offset_out = offset_out.drop(columns=["點類型"])
        detail_out = pd.concat([detail_out, offset_out], ignore_index=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_out.to_excel(writer, sheet_name="細部點座標", index=False)

        if control_df is not None and not control_df.empty:
            control_out = control_df.copy()
            if "點類型" in control_out.columns:
                control_out = control_out.drop(columns=["點類型"])
            control_out.to_excel(writer, sheet_name="控制點", index=False)

    output.seek(0)
    return output.getvalue()


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
    st.caption("使用 Excel 計算模板，自動繪製可放大、可旋轉的平面與三維座標圖（含新版支距法）")

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

    # --- 自動偵測工作表，取得細部點 & 控制點 ---
    try:
        detail_df_raw, control_df_raw, detail_name, control_name = auto_detect_sheets(uploaded_file)
        st.success(f"已偵測到細部點工作表：『{detail_name}』")
        if control_df_raw is not None and control_name is not None:
            st.info(f"已偵測到控制點工作表：『{control_name}』")
        else:
            st.warning("未偵測到控制點工作表，只使用一張工作表做細部點。")
    except Exception as e:
        st.error(f"偵測工作表失敗：{e}")
        return

    # --- 在網站上直接編輯 / 新增點 ---
    st.subheader("細部點座標表（可直接編輯 / 新增）")
    detail_df_edit = st.data_editor(
        detail_df_raw,
        num_rows="dynamic",
        use_container_width=True,
        key="detail_editor"
    )

    if control_df_raw is not None:
        st.subheader("控制點座標表（可直接編輯）")
        control_df_edit = st.data_editor(
            control_df_raw,
            num_rows="dynamic",
            use_container_width=True,
            key="control_editor"
        )
    else:
        control_df_edit = None

    # --- 準備支距法用的全點集合（已分類） ---
    detail_classified = classify_detail_points(detail_df_edit) if not detail_df_edit.empty else pd.DataFrame()
    if control_df_edit is not None and not control_df_edit.empty:
        control_classified = control_df_edit.copy()
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
    st.subheader("支距法產生新點（新版：兩點距離 + NESW 方向）")

    if all_points_for_offset.empty:
        st.info("目前沒有可用的點資料，請先在上方輸入或修改細部點 / 控制點座標。")
        offset_df = st.session_state["offset_points"]
    else:
        point_choices = all_points_for_offset[COL_POINT].astype(str).tolist()

        if len(point_choices) < 2:
            st.info("點位少於兩點，無法執行支距法。")
            offset_df = st.session_state["offset_points"]
        else:
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                dist_p1 = st.selectbox("距離基準點 1", point_choices, key="dist_p1")
            with col_p2:
                dist_p2 = st.selectbox("距離基準點 2", point_choices, key="dist_p2")

            # 顯示距離
            try:
                D_preview = compute_distance(all_points_for_offset, dist_p1, dist_p2)
                st.write(f"兩點距離 D = **{D_preview:.3f} m**")
            except Exception as e:
                st.error(f"距離計算錯誤：{e}")
                D_preview = None

            col_start, col_dir = st.columns(2)
            with col_start:
                start_point = st.selectbox("起始點", point_choices, key="start_point")
            with col_dir:
                direction = st.selectbox("方向（NESW）", ["N", "E", "S", "W"], key="direction")

            col_k, col_c = st.columns(2)
            with col_k:
                k = st.number_input("K 倍距離", min_value=0.0, value=1.0, step=0.1)
            with col_c:
                c = st.number_input("C 次（要生成幾個點）", min_value=1, max_value=100, value=3, step=1)

            if st.button("執行支距法並產生新點"):
                try:
                    new_offset = generate_offset_points_directional(
                        all_points_for_offset,
                        dist_p1,
                        dist_p2,
                        start_point,
                        direction,
                        k,
                        c
                    )
                    offset_df = pd.concat([existing_offset, new_offset], ignore_index=True)
                    st.session_state["offset_points"] = offset_df
                    st.success(
                        f"已從起始點 {start_point} 向 {direction} 方向，"
                        f"依距離({dist_p1}–{dist_p2}) × {k}，產生 {len(new_offset)} 個支距點。"
                    )
                except Exception as e:
                    st.error(f"支距法計算失敗：{e}")
                    offset_df = st.session_state["offset_points"]
            else:
                offset_df = st.session_state["offset_points"]

    if not st.session_state["offset_points"].empty:
        st.write("目前所有支距法產生的點：")
        st.data_editor(
            st.session_state["offset_points"],
            num_rows="dynamic",
            use_container_width=True,
            key="offset_editor"
        )

    st.markdown("---")

    # --- 標籤篩選：只顯示特定類型 ---
    all_types_set = set()
    if not detail_classified.empty:
        all_types_set.update(detail_classified["點類型"].unique().tolist())
    if not control_classified.empty:
        all_types_set.update(control_classified["點類型"].unique().tolist())
    if not existing_offset.empty:
        all_types_set.update(existing_offset["點類型"].unique().tolist())

    all_types_list = sorted(all_types_set)
    st.subheader("顯示的點類型篩選")
    if all_types_list:
        selected_types = st.multiselect(
            "選擇要顯示的點類型（留空 = 全部顯示）",
            options=all_types_list,
            default=all_types_list
        )
    else:
        selected_types = []

    st.markdown("---")

    # --- 繪圖（左右兩欄，使用 plotly_chart，開啟工具列下載按鈕） ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("平面圖 (N–E)")
        fig_plan = plot_plan_interactive(
            detail_df_edit,
            control_df_edit,
            offset_df=st.session_state["offset_points"],
            show_labels=show_labels,
            allowed_types=selected_types
        )

        if fig_plan is None:
            st.warning("沒有符合條件的點可以繪製平面圖。請確認 N/E 座標與標籤篩選。")
        else:
            st.plotly_chart(
                fig_plan,
                use_container_width=True,
                config={
                    "toImageButtonOptions": {
                        "format": "png",
                        "filename": "plan_view",
                        "scale": 2
                    }
                }
            )
            st.caption("💡 右上角工具列可使用「Download plot as png」下載平面圖。")

    with col2:
        st.subheader("三維圖 (E–N–H)")
        fig_3d = plot_3d_interactive(
            detail_df_edit,
            control_df_edit,
            offset_df=st.session_state["offset_points"],
            allowed_types=selected_types
        )

        if fig_3d is None:
            st.warning("沒有符合條件的點可以繪製三維圖。請確認 N/E/H 座標與標籤篩選。")
        else:
            st.plotly_chart(
                fig_3d,
                use_container_width=True,
                config={
                    "toImageButtonOptions": {
                        "format": "png",
                        "filename": "view3d",
                        "scale": 2
                    }
                }
            )
            st.caption("💡 右上角工具列可使用「Download plot as png」下載三維圖。")

    st.markdown("---")

    # --- 匯出 Excel（含目前所有修改 & 支距點） ---
    st.subheader("匯出目前成果為 Excel")
    if st.button("產生並顯示下載按鈕"):
        try:
            excel_bytes = export_to_excel(detail_df_edit, control_df_edit, st.session_state["offset_points"])
            st.download_button(
                label="📥 下載成果 Excel",
                data=excel_bytes,
                file_name="測量成果_含支距點.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"匯出 Excel 失敗：{e}")


if __name__ == "__main__":
    main()
