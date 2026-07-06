import io
import pandas as pd
import streamlit as st

# ============================================================
#  GIOVEN TRAVEL · Excel Toolkit
#  Design giovane, arioso, mobile-first
# ============================================================

st.set_page_config(
    page_title="GIOVEN TRAVEL",
    page_icon="✈️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ------------------------------------------------------------
#  STILE
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Nunito",
                     Roboto, "PingFang SC", "Microsoft YaHei", sans-serif;
    }
    [data-testid="stHeader"] { background: transparent; }
    #MainMenu, footer, [data-testid="stToolbar"] { visibility: hidden; }

    .block-container { max-width: 720px; padding-top: 1rem; padding-bottom: 4rem; }

    /* ---------- HERO ---------- */
    .hero {
        position: relative;
        background: linear-gradient(135deg,#a78bfa 0%,#8b5cf6 40%,#ec4899 100%);
        border-radius: 26px;
        padding: 26px 26px 24px;
        color:#fff;
        box-shadow: 0 18px 40px rgba(139,92,246,.32);
        margin-bottom: 26px;
        overflow: hidden;
    }
    .hero::after {
        content:""; position:absolute; right:-40px; top:-40px;
        width:160px; height:160px; border-radius:50%;
        background: rgba(255,255,255,.15);
    }
    .hero h1 { margin:0; font-size:1.9rem; font-weight:800; letter-spacing:.3px; }
    .hero p  { margin:6px 0 0; font-size:.95rem; opacity:.95; max-width:90%; }

    /* ---------- SEZIONI (leggere, senza scatole pesanti) ---------- */
    .sec { display:flex; align-items:center; gap:9px; margin: 26px 0 12px; }
    .sec .n {
        width:26px; height:26px; border-radius:9px;
        background:#ede9fe; color:#7c3aed; font-weight:800; font-size:.85rem;
        display:flex; align-items:center; justify-content:center;
    }
    .sec .t { font-size:1.05rem; font-weight:750; color:#1f2540; }
    .sec .t small { display:block; font-size:.72rem; color:#94a3b8; font-weight:500; }

    /* ---------- SELETTORE A CARTE (pulsanti modalità) ---------- */
    div[data-testid="column"] .stButton button {
        min-height: 88px; border-radius: 18px; font-weight: 700;
        font-size: 1rem; line-height:1.35; white-space: pre-line;
        transition: all .15s ease; padding: 10px;
    }
    /* inattivo */
    .stButton button[kind="secondary"] {
        background:#faf9ff; border:1.6px solid #ece7ff; color:#5b21b6;
    }
    .stButton button[kind="secondary"]:hover {
        border-color:#c4b5fd; background:#f5f2ff; transform: translateY(-2px);
    }
    /* attivo */
    .stButton button[kind="primary"] {
        background: linear-gradient(135deg,#8b5cf6,#7c3aed); color:#fff;
        border:none; box-shadow: 0 10px 24px rgba(124,58,237,.34);
        transform: translateY(-2px);
    }

    /* ---------- UPLOADER ---------- */
    [data-testid="stFileUploader"] section {
        border:2px dashed #c9bcff; border-radius:16px; background:#faf9ff; padding:14px;
    }
    [data-testid="stFileUploader"] section:hover { border-color:#8b5cf6; }

    /* ---------- STAT RISULTATO ---------- */
    .stat {
        display:flex; align-items:center; gap:14px;
        background: linear-gradient(135deg,#f5f3ff,#fce7f3);
        border:1px solid #eaddff; border-radius:18px; padding:16px 20px; margin-top:14px;
    }
    .stat .big { font-size:2.1rem; font-weight:800;
        background:linear-gradient(135deg,#7c3aed,#ec4899);
        -webkit-background-clip:text; -webkit-text-fill-color:transparent; line-height:1; }
    .stat .lbl { font-size:.9rem; color:#6b21a8; }
    .stat .lbl b { color:#4c1d95; }

    /* ---------- DOWNLOAD ---------- */
    .stDownloadButton button {
        background: linear-gradient(135deg,#7c3aed,#ec4899) !important;
        color:#fff !important; border:none !important; width:100%;
        border-radius:16px !important; font-weight:800 !important;
        padding:.8rem 1rem !important; font-size:1.02rem !important;
        box-shadow: 0 12px 26px rgba(124,58,237,.32) !important;
    }
    .stDownloadButton button:hover { filter:brightness(1.06); transform: translateY(-1px); }

    /* ---------- EXPANDER (opzioni) ---------- */
    [data-testid="stExpander"] { border:none !important; }
    [data-testid="stExpander"] summary { border-radius:12px; background:#f6f4ff; font-weight:600; }

    /* chip multiselect: angoli morbidi */
    [data-baseweb="tag"] { border-radius:9px !important; }

    /* ---------- RESPONSIVE ---------- */
    @media (max-width: 640px) {
        .block-container { padding-left:.8rem; padding-right:.8rem; }
        .hero { padding:22px 20px; border-radius:22px; }
        .hero h1 { font-size:1.55rem; }
        .hero p  { font-size:.86rem; max-width:100%; }
        div[data-testid="column"] .stButton button { min-height: 76px; font-size:.92rem; }
        .stat .big { font-size:1.8rem; }
    }
    /* ---------- DARK ---------- */
    @media (prefers-color-scheme: dark) {
        .sec .t { color:#f1f5f9; }
        .stat { background:linear-gradient(135deg,#2a2140,#3a1d33); border-color:#4c1d95; }
        .stat .lbl { color:#e9d5ff; } .stat .lbl b{ color:#fff; }
        .stButton button[kind="secondary"]{ background:#1e1b2e; border-color:#3b2f5e; color:#ddd6fe; }
        [data-testid="stFileUploader"] section{ background:#1a1730; border-color:#5b21b6; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
#  HERO
# ------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <h1>✈️ GIOVEN TRAVEL</h1>
        <p>上传名单 · 选日期 · 一键生成整洁的 Excel 报表</p>
    </div>
    """,
    unsafe_allow_html=True,
)


def sec(n, title, sub):
    st.markdown(
        f'<div class="sec"><div class="n">{n}</div>'
        f'<div class="t">{title}<small>{sub}</small></div></div>',
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------
#  ① Selettore modalità a carte
# ------------------------------------------------------------
if "mode" not in st.session_state:
    st.session_state.mode = "rooming"


def _pick(m):
    st.session_state.mode = m


sec("1", "选择报表类型", "Che tipo di elenco?")
c1, c2 = st.columns(2, gap="small")
c1.button(
    "🛏️\n排房表\nRooming list",
    key="m1",
    use_container_width=True,
    type="primary" if st.session_state.mode == "rooming" else "secondary",
    on_click=_pick,
    args=("rooming",),
)
c2.button(
    "👥\n团员名单\nPassenger list",
    key="m2",
    use_container_width=True,
    type="primary" if st.session_state.mode == "passenger" else "secondary",
    on_click=_pick,
    args=("passenger",),
)
modalita = st.session_state.mode

# ------------------------------------------------------------
#  ② Upload
# ------------------------------------------------------------
sec("2", "上传文件", "Carica l'Excel (.xlsx)")
file_caricato = st.file_uploader("上传", type=["xlsx"], label_visibility="collapsed")


if file_caricato is not None:
    try:
        df_origine = pd.read_excel(file_caricato)

        if modalita == "rooming":
            colonna_inizio, colonna_fine = "首晚入住日期", "末晚入住日期"
            prefisso_file, base_label = "在店客户名单", "首晚 / 末晚入住日期"
        else:
            colonna_inizio, colonna_fine = "上团日期", "下团日期"
            prefisso_file, base_label = "团队行程名单", "上团 / 下团日期"

        if colonna_inizio in df_origine.columns and colonna_fine in df_origine.columns:
            df_pulito = df_origine.copy()
            ignora = ["nan", "nat", "none", "", "null", "undefined", "-", "/"]
            for col in [colonna_inizio, colonna_fine]:
                df_pulito[col] = df_pulito[col].astype(str).str.strip()
                df_pulito = df_pulito[~df_pulito[col].str.lower().isin(ignora)]
            df_pulito[colonna_inizio] = pd.to_datetime(df_pulito[colonna_inizio], errors="coerce").dt.date
            df_pulito[colonna_fine] = pd.to_datetime(df_pulito[colonna_fine], errors="coerce").dt.date
            df_pulito = df_pulito.dropna(subset=[colonna_inizio, colonna_fine])

            if not df_pulito.empty:
                sec("3", "选择日期范围", f"基于 {base_label}")
                min_date = df_pulito[colonna_inizio].min()
                max_date = df_pulito[colonna_fine].max()
                intervallo = st.date_input(
                    "日期范围",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date,
                    label_visibility="collapsed",
                )

                if isinstance(intervallo, tuple) and len(intervallo) == 2:
                    d1, d2 = intervallo
                    df_filtrato = df_pulito[
                        (df_pulito[colonna_inizio] <= d2) & (df_pulito[colonna_fine] >= d1)
                    ].copy()
                    st.markdown(
                        f'<div class="stat"><div class="big">{len(df_filtrato)}</div>'
                        f'<div class="lbl">条数据符合条件<br><b>{d1}</b> → <b>{d2}</b></div></div>',
                        unsafe_allow_html=True,
                    )
                    nome_file_output = f"{prefisso_file}_{d1}_al_{d2}.xlsx"
                else:
                    df_filtrato = pd.DataFrame(columns=df_pulito.columns)
                    st.warning("请在日历中选择完整的结束日期。")
                    nome_file_output = f"{prefisso_file}.xlsx"
            else:
                st.error("⚠️ 日期列全部为空、错位或含乱码，无法筛选！")
                df_filtrato = pd.DataFrame()
                nome_file_output = "名单.xlsx"
        else:
            st.error(
                f"❌ 此模式需要 '{colonna_inizio}' 和 '{colonna_fine}' 两列，未找到。"
                "请检查表头或切换上方模式。"
            )
            df_filtrato = pd.DataFrame()
            nome_file_output = "名单.xlsx"

        # --- ④ Anteprima + download ---
        if not df_filtrato.empty:
            sec("4", "预览与下载", "Controlla e scarica")
            st.dataframe(df_filtrato, use_container_width=True, hide_index=True)

            colonne = list(df_filtrato.columns)
            with st.expander("⚙️ 更多选项：调整导出列的顺序"):
                colonne_sel = st.multiselect(
                    "选择并排序需要导出的列：", options=colonne, default=colonne
                )
            if not colonne_sel:
                colonne_sel = colonne

            df_ordinato = df_filtrato[colonne_sel].copy()
            for c in (colonna_inizio, colonna_fine):
                if c in df_ordinato.columns:
                    df_ordinato[c] = df_ordinato[c].astype(str)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_ordinato.to_excel(writer, index=False)

            st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
            st.download_button(
                label=f"📥 下载 {prefisso_file}",
                data=buffer.getvalue(),
                file_name=nome_file_output,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"运行过程中发生错误：{e}")
else:
    st.markdown(
        '<div style="text-align:center;color:#a5aec2;padding:22px 0 4px;font-size:.92rem;">'
        "👆 上传一个 Excel 文件即可开始 · Carica un file per iniziare</div>",
        unsafe_allow_html=True,
    )
