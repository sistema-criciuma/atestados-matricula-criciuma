import json
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

import base64
import gzip
from io import BytesIO

import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Tuple

import pandas as pd
import streamlit as st

from pdf_atestado import AtestadoData, generate_atestado_pdf, normalize_intlike, excel_serial_to_iso


BASE_DIR = Path(__file__).resolve().parent
USUARIOS_CSV = BASE_DIR / "usuarios.csv"
MATRICULAS_CSV = BASE_DIR / "matriculas_p_atestado.csv"
LOGO_PATH = BASE_DIR / "logo.png"  # opcional (se existir)
TZ = ZoneInfo("America/Sao_Paulo")


def read_csv_from_secret_gz_b64(secret_key: str, *, encoding: str, **read_csv_kwargs) -> pd.DataFrame:
    if secret_key not in st.secrets:
        raise KeyError(f"Secret não encontrado: {secret_key}")

    raw = str(st.secrets[secret_key])
    raw = re.sub(r"\s+", "", raw)

    data = gzip.decompress(base64.b64decode(raw))

    return pd.read_csv(BytesIO(data), encoding=encoding, **read_csv_kwargs)


def is_admin_user(usuario: str) -> bool:
    return (usuario or "").strip().upper() == "SME"


def safe_filename_name(nome: str) -> str:
    s = (nome or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))  # remove acentos
    s = s.replace("/", "-").replace("\\", "-")
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9_\-]", "", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "aluno"


@st.cache_data(show_spinner=False)
def load_usuarios() -> pd.DataFrame:
    if "USUARIOS_CSV_GZ_B64" in st.secrets:
        df = read_csv_from_secret_gz_b64(
            "USUARIOS_CSV_GZ_B64",
            encoding="latin-1",
            dtype=str,
            keep_default_na=False,
        )
    else:
        if not USUARIOS_CSV.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {USUARIOS_CSV}")
        df = pd.read_csv(USUARIOS_CSV, encoding="latin-1", dtype=str, keep_default_na=False)

    for col in ("Escola", "Usuario", "Senha"):
        if col not in df.columns:
            raise ValueError(
                f"usuarios.csv precisa ter as colunas: Escola, Usuario, Senha. Colunas atuais: {list(df.columns)}"
            )

    df["Escola"] = df["Escola"].astype(str).str.strip()
    df["Usuario"] = df["Usuario"].astype(str)
    df["Senha"] = df["Senha"].astype(str)
    return df


def auth_user(usuario: str, senha_digitada: str) -> Tuple[bool, dict]:
    dfu = load_usuarios()

    u_in = (usuario or "").strip()
    s_in = (senha_digitada or "").strip()

    if not u_in or not s_in:
        return False, {}

    row = dfu.loc[dfu["Usuario"].astype(str).str.strip() == u_in]
    if row.empty:
        return False, {}

    row = row.iloc[0].to_dict()
    senha_armazenada = str(row.get("Senha", "")).strip()

    if s_in != senha_armazenada:
        return False, {}

    admin = is_admin_user(u_in)

    return True, {
        "usuario": u_in,
        "escola": str(row.get("Escola", "")).strip(),
        "is_admin": admin,
    }


def using_matriculas_api() -> bool:
    return "MATRICULAS_API_URL" in st.secrets and "MATRICULAS_API_TOKEN" in st.secrets


@st.cache_data(show_spinner=False, ttl=300)
def api_get(op: str, params: dict | None = None) -> dict:
    url = str(st.secrets["MATRICULAS_API_URL"]).strip()
    token = str(st.secrets["MATRICULAS_API_TOKEN"]).strip()

    q = {"op": op, "token": token}
    if params:
        q.update(params)

    full_url = f"{url}?{urlencode(q)}"
    req = Request(full_url, headers={"User-Agent": "streamlit"})

    with urlopen(req, timeout=30) as resp:
        body = resp.read().decode("utf-8")

    data = json.loads(body)
    if not data.get("ok", False):
        raise RuntimeError(f"API retornou erro: {data}")
    return data


@st.cache_data(show_spinner=False, ttl=300)
def api_list_schools() -> list[str]:
    data = api_get("schools")
    return data.get("schools", [])


@st.cache_data(show_spinner=False, ttl=300)
def api_list_years(escola: str) -> list[str]:
    data = api_get("years", {"escola": (escola or "").strip().upper()})
    return data.get("years", [])


@st.cache_data(show_spinner=False, ttl=300)
def api_search(escola: str, ano: str, situacao: str, q: str, limit: int = 200) -> pd.DataFrame:
    data = api_get(
        "search",
        {
            "escola": (escola or "").strip().upper(),
            "ano": str(ano or "").strip(),
            "situacao": str(situacao or "").strip(),
            "q": str(q or "").strip(),
            "limit": int(limit),
        },
    )
    df = pd.DataFrame(data.get("rows", []))
    return prepare_matriculas_df(df)


@st.cache_data(show_spinner=False, ttl=300)
def api_student(escola: str, ano: str, id_norm: str) -> pd.DataFrame:
    data = api_get(
        "student",
        {
            "escola": (escola or "").strip().upper(),
            "ano": str(ano or "").strip(),
            "id_norm": str(id_norm or "").strip(),
        },
    )
    df = pd.DataFrame(data.get("rows", []))
    return prepare_matriculas_df(df)


def prepare_matriculas_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    required = [
        "ID Aluno", "Código INEP (Aluno)", "Nome", "Escola", "Turma", "Série",
        "Curso", "Data da Matrícula", "Ano", "Situação da Matrícula", "Turno", "Nome da mãe"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"API retornou dados sem colunas obrigatórias: {missing}")

    df["Escola_norm"] = df["Escola"].astype(str).str.strip().str.upper()
    df["Nome_norm"] = df["Nome"].astype(str).str.strip().str.upper()
    df["ID_norm"] = df["ID Aluno"].apply(normalize_intlike)
    df["INEP_norm"] = df["Código INEP (Aluno)"].apply(normalize_intlike)
    return df


@st.cache_data(show_spinner=False)
def load_matriculas() -> pd.DataFrame:
    if "MATRICULAS_CSV_GZ_B64" in st.secrets:
        df = read_csv_from_secret_gz_b64("MATRICULAS_CSV_GZ_B64", encoding="utf-8-sig")
    else:
        if not MATRICULAS_CSV.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {MATRICULAS_CSV}")
        df = pd.read_csv(MATRICULAS_CSV, encoding="utf-8-sig")

    required = [
        "ID Aluno", "Código INEP (Aluno)", "Nome", "Escola", "Turma", "Série",
        "Curso", "Data da Matrícula", "Ano", "Situação da Matrícula", "Turno", "Nome da mãe"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"matriculas_p_atestado.csv está sem colunas obrigatórias: {missing}")

    df["Escola_norm"] = df["Escola"].astype(str).str.strip().str.upper()
    df["Nome_norm"] = df["Nome"].astype(str).str.strip().str.upper()
    df["ID_norm"] = df["ID Aluno"].apply(normalize_intlike)
    df["INEP_norm"] = df["Código INEP (Aluno)"].apply(normalize_intlike)
    return df


def safe_first(series: pd.Series) -> str:
    for v in series.tolist():
        if pd.notna(v) and str(v).strip():
            return str(v).strip()
    return ""


def agg_unique(series: pd.Series) -> str:
    vals = [str(v).strip() for v in series.dropna().tolist() if str(v).strip()]
    uniq = sorted(set(vals))
    return ", ".join(uniq)


def pick_data_matricula(df_aluno: pd.DataFrame) -> str:
    raws = df_aluno["Data da Matrícula"].dropna().tolist()
    if not raws:
        return ""
    iso_dates = []
    for v in raws:
        iso = excel_serial_to_iso(v)
        if iso:
            iso_dates.append(iso)
    if not iso_dates:
        return str(raws[0])
    return min(iso_dates)


def build_atestado_data(df_aluno: pd.DataFrame, ano: Any, escola: str) -> AtestadoData:
    nome = safe_first(df_aluno["Nome"])
    nome_mae = safe_first(df_aluno["Nome da mãe"])

    return AtestadoData(
        ano=str(ano),
        escola=escola,
        nome=nome,
        nome_mae=nome_mae,
        turma=agg_unique(df_aluno["Turma"]),
        serie=agg_unique(df_aluno["Série"]),
        curso=agg_unique(df_aluno["Curso"]),
        turno=agg_unique(df_aluno["Turno"]),
        data_matricula=pick_data_matricula(df_aluno),
        id_aluno=safe_first(df_aluno["ID_norm"]),
        inep_aluno=safe_first(df_aluno["INEP_norm"]),
    )


def logout():
    for k in ("auth_ok", "usuario", "escola", "is_admin"):
        if k in st.session_state:
            del st.session_state[k]


def main():
    st.set_page_config(page_title="Atestado de Matrícula", layout="centered")
    st.title("Atestado de Matrícula")

    if "auth_ok" not in st.session_state:
        st.session_state["auth_ok"] = False

    if not st.session_state["auth_ok"]:
        st.subheader("Login")
        with st.form("login_form", clear_on_submit=False):
            usuario = st.text_input("Usuário", value="", autocomplete="username")
            senha = st.text_input("Senha", value="", type="password", autocomplete="current-password")
            ok = st.form_submit_button("Entrar")

        if ok:
            success, info = auth_user(usuario, senha)
            if success:
                st.session_state["auth_ok"] = True
                st.session_state["usuario"] = info["usuario"]
                st.session_state["escola"] = info["escola"]
                st.session_state["is_admin"] = bool(info.get("is_admin", False))
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")
        st.stop()

    usuario = st.session_state.get("usuario", "").strip()
    is_admin = bool(st.session_state.get("is_admin", False))

    st.write(f"Usuário: {usuario}")
    if is_admin:
        st.write("Perfil: admin (SME)")

    st.button("Sair", on_click=logout)

    use_api = using_matriculas_api()

    if use_api:
        if is_admin:
            escolas = api_list_schools()
            escola_sel = st.selectbox("Escola", options=escolas)
            escola = (escola_sel or "").strip()
        else:
            escola = st.session_state.get("escola", "").strip()
            st.write(f"Escola: {escola}")

        if not escola:
            st.warning("Nenhuma escola disponível.")
            st.stop()

        anos = api_list_years(escola)
        if not anos:
            st.warning("Não há anos disponíveis para esta escola.")
            st.stop()

        ano_sel = st.selectbox("Ano", options=anos, index=len(anos) - 1)

        sit_sel = st.text_input("Situação da matrícula (opcional)", value="Cursando").strip()
        if sit_sel.lower() in ("", "(todas)", "todas"):
            sit_sel = ""

        st.subheader("Buscar aluno")
        q = st.text_input("Digite nome, ID do aluno ou INEP", value="").strip().upper()

        df_busca = api_search(escola, ano_sel, sit_sel, q, limit=200)

    else:
        dfm = load_matriculas()

        if is_admin:
            escolas = sorted(dfm["Escola"].dropna().astype(str).str.strip().unique().tolist())
            escola_sel = st.selectbox("Escola", options=escolas)
            escola = (escola_sel or "").strip()
            dfm_escola = dfm.loc[dfm["Escola_norm"] == escola.upper()].copy()
        else:
            escola = st.session_state.get("escola", "").strip()
            st.write(f"Escola: {escola}")
            dfm_escola = dfm.loc[dfm["Escola_norm"] == escola.upper()].copy()

        if dfm_escola.empty:
            st.warning("Não há matrículas para esta escola na tabela.")
            st.stop()

        anos = sorted([str(a) for a in dfm_escola["Ano"].dropna().unique()])
        ano_sel = st.selectbox("Ano", options=anos, index=len(anos) - 1)

        df_ano = dfm_escola.loc[dfm_escola["Ano"].astype(str) == str(ano_sel)].copy()

        if "Situação da Matrícula" in df_ano.columns:
            situacoes = ["(todas)"] + sorted([str(x) for x in df_ano["Situação da Matrícula"].dropna().unique()])
            sit_sel = st.selectbox(
                "Situação da matrícula",
                options=situacoes,
                index=1 if "Cursando" in situacoes else 0,
            )
            if sit_sel != "(todas)":
                df_ano = df_ano.loc[df_ano["Situação da Matrícula"].astype(str) == sit_sel]

        st.subheader("Buscar aluno")
        q = st.text_input("Digite nome, ID do aluno ou INEP", value="").strip().upper()

        df_busca = df_ano
        if q:
            digits = "".join([c for c in q if c.isdigit()])
            mask_nome = df_ano["Nome_norm"].str.contains(q, na=False)
            mask_id = df_ano["ID_norm"].eq(digits) if digits else False
            mask_inep = df_ano["INEP_norm"].eq(digits) if digits else False
            df_busca = df_ano.loc[mask_nome | mask_id | mask_inep]

    if df_busca is None or df_busca.empty:
        st.info("Nenhum aluno encontrado com esse filtro.")
        st.stop()

    cols_show = ["Nome", "ID_norm", "INEP_norm", "Turma", "Turno", "Série", "Curso"]
    df_lista = (
        df_busca[cols_show]
        .drop_duplicates(subset=["ID_norm"])
        .sort_values(["Nome"])
        .reset_index(drop=True)
    )

    st.dataframe(df_lista.head(200), use_container_width=True, hide_index=True)

    options = df_lista.apply(
        lambda r: f"{r['Nome']} | ID {r['ID_norm']} | INEP {r['INEP_norm']} | {r['Turma']} | {r['Turno']}",
        axis=1,
    ).tolist()

    idx = st.selectbox("Selecionar aluno", options=list(range(len(options))), format_func=lambda i: options[i])
    sel_id = df_lista.loc[idx, "ID_norm"]

    if use_api:
        df_aluno = api_student(escola, ano_sel, sel_id)
    else:
        df_aluno = df_busca.loc[df_busca["ID_norm"] == sel_id].copy()

    if df_aluno is None or df_aluno.empty:
        st.warning("Não foi possível carregar os dados completos do aluno.")
        st.stop()

    st.subheader("Dados para o atestado")
    st.write(f"Aluno: {safe_first(df_aluno['Nome'])}")
    st.write(f"Mãe: {safe_first(df_aluno['Nome da mãe'])}")
    st.write(f"Turma(s): {agg_unique(df_aluno['Turma'])}")
    st.write(f"Série(s): {agg_unique(df_aluno['Série'])}")
    st.write(f"Curso(s): {agg_unique(df_aluno['Curso'])}")
    st.write(f"Turno(s): {agg_unique(df_aluno['Turno'])}")

    if st.button("Gerar PDF"):
        at = build_atestado_data(df_aluno, ano_sel, escola)
        emitted_dt = datetime.now(TZ)

        logo_path = str(LOGO_PATH) if LOGO_PATH.exists() else None
        pdf_bytes = generate_atestado_pdf(
            data=at,
            emitted_dt=emitted_dt,
            city_uf="Criciúma / SC",
            logo_path=logo_path,
            school_meta=None,
        )

        aluno_nome = safe_filename_name(at.nome)
        data_emissao = emitted_dt.strftime("%d-%m-%Y")
        filename = f"matricula_{aluno_nome}_{data_emissao}.pdf"

        st.success("PDF gerado.")
        st.download_button(
            label="Baixar PDF",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
        )


if __name__ == "__main__":
    main()

