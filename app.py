"""
FT - Zona Sul | Painel Centralizado de Documentos
Auth via Google OAuth2 (fluxo manual — sem PKCE)
Restrição de domínio: apenas @rappi.com
Arquivos servidos a partir da pasta ./data/
"""

import os
from pathlib import Path
from urllib.parse import urlencode

import requests
import streamlit as st
import pandas as pd
import mammoth

# ─────────────────────────────────────────────
#  PAGE CONFIG  (primeira instrução obrigatória)
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="FT – Zona Sul | Documentos",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  CONSTANTES
# ─────────────────────────────────────────────
DATA_DIR       = Path(__file__).parent / "data"
ALLOWED_DOMAIN = "@rappi.com"
SUPPORTED_EXTS = {".xlsx", ".xls", ".docx"}

# ── Arquivo remoto: Consolidação de Integrações (Google Sheets) ───────────────
# O arquivo .xlsx de 37MB não fica no repositório; é lido direto do Google Sheets.
# Para funcionar, a planilha precisa ser compartilhada como "Qualquer pessoa com o link".
_CONSOLIDACAO_FILENAME = "Consolidação de Integrações Zona Sul.xlsx"
_GSHEET_ID             = "17TRmff2LDP2b5S8_qGKITE31bUN153uo"
# CSV export por aba — muito mais rápido que XLSX para leitura remota
_CONSOLIDACAO_CSV_URL  = (
    f"https://docs.google.com/spreadsheets/d/{_GSHEET_ID}"
    "/export?format=csv&sheet=Base+de+Integra%C3%A7%C3%B5es"
)
# XLSX export (usado para obter os nomes de todas as abas e abas secundárias)
_CONSOLIDACAO_XLSX_URL = (
    f"https://docs.google.com/spreadsheets/d/{_GSHEET_ID}/export?format=xlsx"
)

# Google OAuth endpoints
_AUTH_URI     = "https://accounts.google.com/o/oauth2/v2/auth"
_TOKEN_URI    = "https://oauth2.googleapis.com/token"
_USERINFO_URI = "https://www.googleapis.com/oauth2/v3/userinfo"
_SCOPES       = "openid email profile"

# ─────────────────────────────────────────────
#  ESTILOS GLOBAIS
# ─────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp {
        background: linear-gradient(135deg, #0f0f0f 0%, #1a1a1a 50%, #111111 100%);
        color: #e8e8e8;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1c1c1c 0%, #141414 100%);
        border-right: 1px solid #2a2a2a;
    }
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] label { color: #cccccc !important; font-size: 0.85rem; }

    .brand-block {
        display: flex; align-items: center; gap: 10px;
        padding: 1.2rem 0.5rem 0.8rem;
        border-bottom: 1px solid #2e2e2e; margin-bottom: 1rem;
    }
    .brand-dot {
        width: 36px; height: 36px;
        background: linear-gradient(135deg, #FF6B35, #FF4500);
        border-radius: 8px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem; font-weight: 700; color: white;
        box-shadow: 0 0 12px rgba(255,107,53,0.5); flex-shrink: 0;
    }
    .brand-text .title { font-size: 0.9rem; font-weight: 700; color: #FF6B35; }
    .brand-text .sub   { font-size: 0.7rem; color: #777; letter-spacing: 0.05em; }

    .user-card {
        background: linear-gradient(135deg, #1e1e1e, #252525);
        border: 1px solid #2e2e2e; border-radius: 12px;
        padding: 0.9rem; margin-bottom: 1rem;
        display: flex; align-items: center; gap: 10px;
    }
    .user-avatar {
        width: 38px; height: 38px; border-radius: 50%;
        background: linear-gradient(135deg, #FF6B35, #FF4500);
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem; font-weight: 700; color: white; flex-shrink: 0;
    }
    .user-info .name  { font-size: 0.82rem; font-weight: 600; color: #e8e8e8; }
    .user-info .email { font-size: 0.72rem; color: #888; }

    .stButton > button {
        background: linear-gradient(135deg, #FF6B35, #FF4500) !important;
        color: white !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
        font-size: 0.82rem !important; padding: 0.4rem 1rem !important;
        transition: all 0.2s ease !important; width: 100%;
    }
    .stButton > button:hover {
        opacity: 0.85 !important; transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(255,107,53,0.4) !important;
    }

    .stTabs [data-baseweb="tab-list"] { background: transparent; border-bottom: 2px solid #2a2a2a; }
    .stTabs [data-baseweb="tab"] { color: #888 !important; font-size: 0.82rem !important; font-weight: 500 !important; }
    .stTabs [aria-selected="true"] { color: #FF6B35 !important; border-bottom: 2px solid #FF6B35 !important; background: transparent !important; }

    .main-header { padding: 1.5rem 0 1rem; border-bottom: 1px solid #2a2a2a; margin-bottom: 1.5rem; }
    .main-header h1 { font-size: 1.6rem; font-weight: 700; color: #f0f0f0; margin: 0; }
    .main-header p  { font-size: 0.85rem; color: #666; margin: 0.2rem 0 0; }

    .welcome-card {
        background: linear-gradient(135deg, #1e1e1e, #1a1a1a);
        border: 1px solid #2e2e2e; border-left: 3px solid #FF6B35;
        border-radius: 12px; padding: 2.5rem; text-align: center; margin-top: 2rem;
    }
    .welcome-card h2 { color: #FF6B35; margin-bottom: 0.4rem; }
    .welcome-card p  { color: #777; font-size: 0.9rem; }

    .doc-container {
        background: #1a1a1a; border: 1px solid #2a2a2a;
        border-radius: 12px; padding: 2rem 2.5rem;
        line-height: 1.85; color: #d4d4d4; font-size: 0.95rem;
    }
    .doc-heading { color: #FF6B35; font-weight: 700; margin-top: 1.2rem; }
    .doc-para    { margin: 0.3rem 0; }

    .ext-badge {
        display: inline-block; background: rgba(255,107,53,0.15);
        border: 1px solid rgba(255,107,53,0.3); color: #FF6B35;
        border-radius: 6px; padding: 2px 10px;
        font-size: 0.75rem; font-weight: 600;
        margin-left: 8px; vertical-align: middle;
    }

    .access-denied {
        background: rgba(220,53,69,0.08); border: 1px solid rgba(220,53,69,0.3);
        border-radius: 12px; padding: 2rem; text-align: center;
    }
    .access-denied h2 { color: #dc3545; }
    .access-denied p  { color: #aaa; }

    .login-wrap {
        display: flex; flex-direction: column; align-items: center;
        justify-content: center; min-height: 80vh; gap: 1.5rem;
    }
    .login-card {
        background: linear-gradient(135deg, #1e1e1e, #1a1a1a);
        border: 1px solid #2e2e2e; border-radius: 16px;
        padding: 2.5rem 3rem; text-align: center; width: 100%; max-width: 420px;
    }
    .login-logo {
        width: 64px; height: 64px;
        background: linear-gradient(135deg, #FF6B35, #FF4500);
        border-radius: 16px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.8rem; font-weight: 800; color: white;
        margin: 0 auto 1.2rem;
        box-shadow: 0 0 24px rgba(255,107,53,0.5);
    }
    .login-title { font-size: 1.4rem; font-weight: 700; color: #f0f0f0; margin: 0 0 0.3rem; }
    .login-sub   { font-size: 0.85rem; color: #666; margin: 0 0 2rem; }
    .login-btn {
        display: inline-flex; align-items: center; gap: 0.6rem;
        background: linear-gradient(135deg, #FF6B35, #FF4500);
        color: white; text-decoration: none; border-radius: 10px;
        padding: 0.75rem 1.8rem; font-weight: 600; font-size: 0.9rem;
        box-shadow: 0 4px 16px rgba(255,107,53,0.4);
        transition: all 0.2s ease;
    }
    .login-btn:hover { opacity: 0.88; transform: translateY(-2px); }

    hr { border-color: #2a2a2a !important; }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #111; }
    ::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #FF6B35; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
#  OAUTH HELPERS  (fluxo manual, sem PKCE)
# ─────────────────────────────────────────────

def _get_secret(key: str) -> str:
    return st.secrets["google"][key]


def _get_redirect_uri() -> str:
    """Retorna localhost em dev, URL de produção no Streamlit Cloud.
    Usa múltiplos indicadores para detectar o ambiente com segurança.
    """
    is_cloud = any([
        os.environ.get("STREAMLIT_SHARING_MODE") == "streamlit-cloud",
        os.path.exists("/home/appuser"),          # container padrão do Streamlit Cloud
        os.environ.get("HOME", "").startswith("/home/appuser"),
        os.environ.get("USER") == "appuser",      # usuário padrão do Streamlit Cloud
    ])
    key = "redirect_uri_prod" if is_cloud else "redirect_uri_local"
    return st.secrets["google"][key]


def _build_auth_url() -> str:
    """Monta a URL de autorização do Google (sem PKCE)."""
    params = {
        "client_id":     _get_secret("client_id"),
        "redirect_uri":  _get_redirect_uri(),
        "response_type": "code",
        "scope":         _SCOPES,
        "access_type":   "online",
        "prompt":        "select_account",
    }
    return _AUTH_URI + "?" + urlencode(params)


def _exchange_code(code: str) -> dict:
    """Troca o authorization code por um access token."""
    resp = requests.post(
        _TOKEN_URI,
        data={
            "code":          code,
            "client_id":     _get_secret("client_id"),
            "client_secret": _get_secret("client_secret"),
            "redirect_uri":  _get_redirect_uri(),
            "grant_type":    "authorization_code",
        },
        timeout=10,
    )
    return resp.json()


def _fetch_user_info(access_token: str) -> dict:
    """Busca nome, email e foto do usuário autenticado."""
    resp = requests.get(
        _USERINFO_URI,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=10,
    )
    return resp.json()


def do_logout():
    for key in ("oauth_state", "access_token", "user_info", "connected"):
        st.session_state.pop(key, None)
    st.query_params.clear()
    st.rerun()


# ─────────────────────────────────────────────
#  FLUXO DE AUTENTICAÇÃO
# ─────────────────────────────────────────────

params = st.query_params

# ── Callback do Google (code presente na URL após redirect) ──
if "code" in params and not st.session_state.get("connected"):
    with st.spinner("Autenticando..."):
        token_data = _exchange_code(params["code"])

    if "error" in token_data:
        st.error(f"❌ Erro ao obter token: `{token_data.get('error_description', token_data['error'])}`")
        st.query_params.clear()
        st.stop()

    access_token = token_data.get("access_token", "")
    user_info    = _fetch_user_info(access_token)

    st.session_state["access_token"] = access_token
    st.session_state["user_info"]    = user_info
    st.session_state["connected"]    = True
    st.query_params.clear()
    st.rerun()

# ── Não autenticado: mostra tela de login ──
if not st.session_state.get("connected"):
    auth_url = _build_auth_url()

    st.markdown(
        f"""
        <div class="login-wrap">
          <div class="login-card">
            <div class="login-logo">FT</div>
            <div class="login-title">FT – Zona Sul</div>
            <div class="login-sub">Painel Centralizado de Documentos</div>
            <a class="login-btn" href="{auth_url}" target="_self">
              <svg width="18" height="18" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
                <path fill="#fff" d="M44.5 20H24v8.5h11.8C34.7 33.9 30 37 24 37c-7.2 0-13-5.8-13-13s5.8-13 13-13c3.1 0 5.9 1.1 8.1 2.9l6.4-6.4C34.6 4.1 29.6 2 24 2 11.8 2 2 11.8 2 24s9.8 22 22 22c11 0 21-8 21-22 0-1.3-.2-2.7-.5-4z"/>
              </svg>
              Entrar com Google
            </a>
            <p style="font-size:0.72rem; color:#444; margin-top:1.5rem;">
              Acesso restrito a colaboradores <strong style="color:#666;">@rappi.com</strong>
            </p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# ─────────────────────────────────────────────
#  VERIFICAÇÃO DE DOMÍNIO
# ─────────────────────────────────────────────

user_info  = st.session_state.get("user_info", {})
user_email = user_info.get("email", "")
user_name  = user_info.get("name", "")

if not user_email.lower().endswith(ALLOWED_DOMAIN):
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; justify-content:center; min-height:80vh;">
          <div class="access-denied">
            <div style="font-size:3rem;">🔒</div>
            <h2>Acesso Restrito</h2>
            <p>Este painel é exclusivo para colaboradores <strong>@rappi.com</strong>.<br/>
               O e-mail <code>{user_email}</code> não está autorizado.</p>
            <p style="margin-top:1rem; font-size:0.8rem; color:#555;">
              Entre em contato com o administrador se acredita que isso é um erro.
            </p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.sidebar:
        st.markdown(
            '<div class="brand-block">'
            '<div class="brand-dot">FT</div>'
            '<div class="brand-text"><div class="title">FT – Zona Sul</div>'
            '<div class="sub">PAINEL DE DOCUMENTOS</div></div></div>',
            unsafe_allow_html=True,
        )
        st.markdown("---")
        if st.button("🚪 Sair"):
            do_logout()
    st.stop()


# ─────────────────────────────────────────────
#  HELPERS DE UI
# ─────────────────────────────────────────────

def render_sidebar_brand():
    st.markdown(
        '<div class="brand-block">'
        '<div class="brand-dot">FT</div>'
        '<div class="brand-text">'
        '<div class="title">FT – Zona Sul</div>'
        '<div class="sub">PAINEL DE DOCUMENTOS</div>'
        '</div></div>',
        unsafe_allow_html=True,
    )


def render_user_card(name: str, email: str):
    picture = user_info.get("picture", "")
    if picture:
        avatar_html = f'<img src="{picture}" style="width:38px;height:38px;border-radius:50%;object-fit:cover;flex-shrink:0;">'
    else:
        initial = (name[0] if name else email[0]).upper()
        avatar_html = f'<div class="user-avatar">{initial}</div>'

    st.markdown(
        f'<div class="user-card">'
        f'{avatar_html}'
        f'<div class="user-info">'
        f'<div class="name">{name or "Usuário"}</div>'
        f'<div class="email">{email}</div>'
        f'</div></div>',
        unsafe_allow_html=True,
    )


def get_file_icon(ext: str) -> str:
    return {"xlsx": "📊", "xls": "📊", "docx": "📝"}.get(ext.lstrip("."), "📄")


def human_size(size_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} GB"


def list_data_files() -> list:
    if not DATA_DIR.exists():
        local_files = []
    else:
        local_files = [f for f in DATA_DIR.iterdir() if f.suffix.lower() in SUPPORTED_EXTS]

    # Adiciona entrada virtual para o arquivo remoto (não existe em disco)
    remote_virtual = DATA_DIR / _CONSOLIDACAO_FILENAME
    if not any(f.name == _CONSOLIDACAO_FILENAME for f in local_files):
        local_files.append(remote_virtual)

    # DOCX primeiro (0), depois outros (1). Alfabético dentro.
    return sorted(
        local_files,
        key=lambda f: (0 if f.suffix.lower() == ".docx" else 1, f.name.lower())
    )


def _is_remote_file(file_path: Path) -> bool:
    """Retorna True se o arquivo não existe em disco (é remoto)."""
    return not file_path.exists()


# ── Leituras cacheadas ─────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_integracoes_sample() -> pd.DataFrame:
    """Lê as 1.000 primeiras linhas via CSV (rápido, sem baixar o XLSX inteiro)."""
    return pd.read_csv(_CONSOLIDACAO_CSV_URL, nrows=1000)


@st.cache_data(show_spinner=False)
def _load_integracoes_filter_values() -> dict:
    """Lê apenas as 4 colunas de filtro via CSV para popular os selects."""
    FILTER_COLS = ["ID da Loja", "PRODUCT_ID", "SKU (ID)", "EAN"]
    df = pd.read_csv(_CONSOLIDACAO_CSV_URL, usecols=lambda c: c in FILTER_COLS)
    return {
        col: sorted(df[col].dropna().astype(str).unique().tolist())
        for col in FILTER_COLS
        if col in df.columns
    }


@st.cache_data(show_spinner=False)
def _load_integracoes_filtered(col: str, value: str) -> pd.DataFrame:
    """Carrega o CSV completo e aplica o filtro. Resultado fica em cache."""
    df = pd.read_csv(_CONSOLIDACAO_CSV_URL)
    return df[df[col].astype(str) == value]


@st.cache_data(show_spinner=False)
def _load_sheet_generic(file_path_str: str, sheet: str) -> pd.DataFrame:
    return pd.read_excel(file_path_str, sheet_name=sheet, engine="openpyxl")


@st.cache_data(show_spinner=False)
def _load_sheet_kpis_found_rate(file_path_str: str) -> pd.DataFrame:
    df = pd.read_excel(file_path_str, sheet_name="Found Rate", usecols="A:J", engine="openpyxl")
    if "TOTAL_ORDERS" in df.columns:
        df = df[df["TOTAL_ORDERS"].notna()]
        df = df[df["TOTAL_ORDERS"].astype(str).str.strip() != "#N/A"]
    return df


def render_xlsx(file_path: Path):
    is_remote = _is_remote_file(file_path)
    is_consolidacao = file_path.name == _CONSOLIDACAO_FILENAME

    try:
        if is_consolidacao:
            # Para o arquivo remoto, as abas são conhecidas (evita baixar XLSX só para metadados)
            sheet_names = ["Base de Integrações"]
        else:
            import openpyxl
            wb_meta = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet_names = wb_meta.sheetnames
            wb_meta.close()

        if not sheet_names:
            st.warning("⚠️ A planilha não contém nenhuma aba.")
            return

        tabs = st.tabs([f"📋 {s}" for s in sheet_names])
        fp_str = str(file_path)
        FILTER_COLS = ["ID da Loja", "PRODUCT_ID", "SKU (ID)", "EAN"]

        for tab, sheet in zip(tabs, sheet_names):
            with tab:
                try:
                    is_kpis_found_rate = (
                        file_path.name.lower() == "kpis das lojas.xlsx"
                        and sheet == "Found Rate"
                    )
                    is_integracoes_base = (
                        is_consolidacao and sheet == "Base de Integrações"
                    )

                    if is_kpis_found_rate:
                        df = _load_sheet_kpis_found_rate(fp_str)

                    elif is_integracoes_base:
                        st.markdown(
                            "<p style='font-size:0.82rem;color:#888;margin-bottom:0.3rem;'>"
                            "🔍 <strong>Filtrar base</strong> — escolha uma coluna e um valor "
                            "para carregar os dados filtrados. Sem filtro: exibe amostra de 1.000 linhas.</p>",
                            unsafe_allow_html=True,
                        )
                        c1, c2, c3 = st.columns([1, 2, 1])
                        with c1:
                            filtro_col = st.selectbox(
                                "Coluna",
                                options=[""] + FILTER_COLS,
                                format_func=lambda x: "Nenhum filtro" if x == "" else x,
                                key="integracoes_filter_col",
                                label_visibility="collapsed",
                            )
                        if filtro_col:
                            with st.spinner("Carregando valores únicos..."):
                                filter_vals = _load_integracoes_filter_values()
                            valores = filter_vals.get(filtro_col, [])
                            with c2:
                                filtro_val = st.selectbox(
                                    "Valor",
                                    options=[""] + valores,
                                    format_func=lambda x: "Selecione um valor..." if x == "" else x,
                                    key="integracoes_filter_val",
                                    label_visibility="collapsed",
                                )
                        else:
                            filtro_val = ""

                        if filtro_col and filtro_val:
                            with st.spinner(f"Filtrando por {filtro_col} = {filtro_val}..."):
                                df = _load_integracoes_filtered(filtro_col, filtro_val)
                            st.success(f"✅ {len(df):,} linhas com **{filtro_col}** = `{filtro_val}`")
                        else:
                            with st.spinner("Carregando amostra..."):
                                df = _load_integracoes_sample()
                            st.info("⚡ Amostra: primeiras 1.000 linhas. Use os filtros acima para buscar na base completa.")

                    else:
                        df = _load_sheet_generic(fp_str, sheet)

                    if df.empty:
                        st.info(f"A aba **{sheet}** está vazia.")
                    else:
                        st.caption(f"{len(df):,} linhas · {len(df.columns)} colunas")
                        st.dataframe(df, use_container_width=True, hide_index=False)
                except Exception as e:
                    st.error(f"Erro ao ler a aba **{sheet}**: {e}")
    except Exception as e:
        st.error(f"❌ Não foi possível abrir o arquivo.\n\n`{e}`")


def render_docx(file_path: Path):
    """Converte DOCX → HTML via mammoth (preserva tabelas, imagens, formatação)."""
    try:
        with open(file_path, "rb") as f:
            result = mammoth.convert_to_html(f)

        html_body = result.value
        if not html_body.strip():
            st.warning("⚠️ O documento parece estar vazio.")
            return

        # Avisos de conversão (elementos não suportados)
        if result.messages:
            with st.expander("⚠️ Avisos de conversão", expanded=False):
                for msg in result.messages:
                    st.caption(str(msg))

        # CSS inline para o conteúdo gerado pelo mammoth
        doc_css = """
        <style>
        /* Títulos — tamanhos explícitos para sobrepor o reset global do Streamlit */
        .doc-container h1 {
            color: #FF6B35 !important; font-weight: 800 !important;
            font-size: 2.6rem !important;
            margin-top: 1.8rem !important; margin-bottom: 0.6rem !important;
            border-bottom: 2px solid rgba(255,107,53,0.25);
            padding-bottom: 0.4rem;
        }
        .doc-container h2 {
            color: #FF6B35 !important; font-weight: 700 !important;
            font-size: 1.85rem !important;
            margin-top: 1.6rem !important; margin-bottom: 0.5rem !important;
        }
        .doc-container h3 {
            color: #e87d50 !important; font-weight: 700 !important;
            font-size: 1.45rem !important;
            margin-top: 1.4rem !important; margin-bottom: 0.4rem !important;
        }
        .doc-container h4 {
            color: #d4824a !important; font-weight: 600 !important;
            font-size: 1.25rem !important;
            margin-top: 1.2rem !important; margin-bottom: 0.35rem !important;
        }
        .doc-container h5, .doc-container h6 {
            color: #c07840 !important; font-weight: 600 !important;
            font-size: 1.15rem !important;
            margin-top: 1rem !important; margin-bottom: 0.3rem !important;
        }
        /* Corpo do texto */
        .doc-container p   { margin: 0.5rem 0; line-height: 1.95; font-size: 1.18rem !important; color: #dcdcdc !important; }
        .doc-container li  { margin: 0.35rem 0; line-height: 1.9; font-size: 1.18rem !important; color: #dcdcdc !important; }
        .doc-container strong { color: #ffffff !important; }
        .doc-container em  { color: #dddddd !important; font-style: italic; }
        .doc-container ul, .doc-container ol { padding-left: 1.6rem; margin: 0.6rem 0; }
        .doc-container img { max-width: 100%; border-radius: 8px; margin: 0.8rem 0;
                             border: 1px solid #2a2a2a; }
        .doc-container table {
            width: 100%; border-collapse: collapse;
            margin: 1.2rem 0; font-size: 1.12rem !important;
        }
        .doc-container th {
            background: rgba(255,107,53,0.15);
            color: #FF6B35; font-weight: 600;
            padding: 0.7rem 0.9rem;
            border: 1px solid #333; text-align: left;
        }
        .doc-container td {
            padding: 0.6rem 0.9rem;
            border: 1px solid #2a2a2a; color: #d4d4d4; font-size: 1.12rem !important;
        }
        .doc-container tr:nth-child(even) td { background: rgba(255,255,255,0.025); }
        .doc-container tr:hover td { background: rgba(255,107,53,0.05); }
        .doc-container a   { color: #FF6B35; }
        </style>
        """

        st.markdown(
            doc_css + f'<div class="doc-container">{html_body}</div>',
            unsafe_allow_html=True,
        )
    except Exception as e:
        st.error(f"❌ Não foi possível abrir o documento.\n\n`{e}`")


# ─────────────────────────────────────────────
#  LAYOUT PRINCIPAL (apenas @rappi.com)
# ─────────────────────────────────────────────

with st.sidebar:
    render_sidebar_brand()
    render_user_card(user_name, user_email)
    st.markdown("---")

    data_files = list_data_files()

    if not data_files:
        st.warning("Nenhum arquivo encontrado na pasta `data/`.")
        selected_file = None
    else:
        st.markdown(
            "<p style='font-size:0.75rem; font-weight:600; color:#777; "
            "letter-spacing:0.08em; margin-bottom:0.5rem;'>ARQUIVOS DISPONÍVEIS</p>",
            unsafe_allow_html=True,
        )
        file_names = [f.name for f in data_files]
        selected_name = st.selectbox(
            label="Selecione um arquivo",
            options=file_names,
            label_visibility="collapsed",
        )
        selected_file = DATA_DIR / selected_name

        ext  = selected_file.suffix.lower().lstrip(".")
        size = human_size(selected_file.stat().st_size) if selected_file.exists() else "Google Sheets"
        icon = get_file_icon(ext)
        st.markdown(
            f'<div style="background:#1e1e1e; border:1px solid #2a2a2a;'
            f'border-left:3px solid #FF6B35; border-radius:10px;'
            f'padding:0.7rem 0.9rem; margin-top:0.5rem;">'
            f'<div style="display:flex; align-items:center; gap:0.5rem;">'
            f'<span style="font-size:1.2rem;">{icon}</span>'
            f'<div><div style="font-size:0.75rem; font-weight:600; color:#e8e8e8;">{selected_name}</div>'
            f'<div style="font-size:0.65rem; color:#555; margin-top:2px;">{ext.upper()} · {size}</div>'
            f'</div></div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("---")
    if st.button("🚪 Logout"):
        do_logout()

# ── Área principal ────────────────────────────
st.markdown(
    '<div class="main-header">'
    '<h1>📄 Painel de Documentos</h1>'
    '<p>Visualize planilhas e documentos armazenados na pasta do projeto</p>'
    '</div>',
    unsafe_allow_html=True,
)

if not data_files or selected_file is None:
    st.markdown(
        '<div class="welcome-card">'
        '<div style="font-size:3rem;">📁</div>'
        '<h2>Pasta data/ vazia</h2>'
        '<p>Adicione arquivos <strong>.xlsx</strong> ou <strong>.docx</strong>'
        ' à pasta <code>data/</code> e recarregue a página.</p>'
        '</div>',
        unsafe_allow_html=True,
    )
else:
    ext      = selected_file.suffix.lower()
    file_sz  = human_size(selected_file.stat().st_size) if selected_file.exists() else "Google Sheets"

    st.markdown(
        f'<div style="display:flex; align-items:center; gap:0.6rem; margin-bottom:1.2rem;">'
        f'<span style="font-size:1.1rem; font-weight:600; color:#e8e8e8;">{selected_file.name}</span>'
        f'<span class="ext-badge">{ext.upper()}</span>'
        f'<span style="color:#555; font-size:0.78rem; margin-left:auto;">{file_sz}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if ext in (".xlsx", ".xls"):
        render_xlsx(selected_file)
    elif ext == ".docx":
        render_docx(selected_file)
    else:
        st.warning(f"Formato `{ext}` não suportado.")
