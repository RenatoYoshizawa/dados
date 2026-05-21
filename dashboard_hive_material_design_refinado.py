import os
import time
import base64
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components


# =========================#
# CONFIGURAÇÕES
# =========================

# Repositório GitHub utilizado pelo RPA de monitoramento.
# Para repositório privado, cadastre TOKEN_GITHUB em: Streamlit Cloud > Settings > Secrets.
GITHUB_OWNER = "RenatoYoshizawa"
GITHUB_REPO = "Monitoramento"
GITHUB_BRANCH = "main"

GITHUB_ARQ_MONITORAMENTO = "monitoramento/Monitoramento.xlsx"
GITHUB_ARQ_CRITICAS = "monitoramento/Criticas.xlsx"
GITHUB_ARQ_HISTORICO = "monitoramento/Historico_Criticas.xlsx"
GITHUB_ARQ_MONITORAMENTO_HIST = f"historico/Monitoramento_{datetime.now().strftime('%Y_%m')}.csv"
GITHUB_ARQ_INCONSISTENCIAS_HIST = f"historico/Inconsistencias_{datetime.now().strftime('%Y_%m')}.csv"

CACHE_DIR = Path("/tmp/monitoramento_ecrv_cache")
CACHE_DIR.mkdir(parents=True, exist_ok=True)

ARQ_LOCAL_MONITORAMENTO = CACHE_DIR / "Monitoramento.xlsx"
ARQ_LOCAL_CRITICAS = CACHE_DIR / "Criticas.xlsx"
ARQ_LOCAL_HISTORICO = CACHE_DIR / "Historico_Criticas.xlsx"
ARQ_LOCAL_MONITORAMENTO_HIST = CACHE_DIR / "Monitoramento_Historico.csv"
ARQ_LOCAL_INCONSISTENCIAS_HIST = CACHE_DIR / "Inconsistencias_Historico.csv"
META_LOCAL = CACHE_DIR / "github_meta.json"

INTERVALO_VERIFICACAO_SEGUNDOS = 30


# =========================
# STREAMLIT
# =========================

st.set_page_config(
    page_title="Monitoramento e-CRV",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =========================
# CSS
# =========================

CSS = """
<style>
:root {
    --md-bg: #F6F8FC;
    --md-surface: #FFFFFF;
    --md-surface-container: #F8FAFD;
    --md-surface-variant: #F1F3F4;
    --md-primary: #1A73E8;
    --md-primary-dark: #174EA6;
    --md-text: #202124;
    --md-muted: #5F6368;
    --md-border: #E0E3EB;
    --md-green: #188038;
    --md-yellow: #F9AB00;
    --md-red: #D93025;
    --md-blue-soft: #E8F0FE;
    --md-shadow: 0 1px 2px rgba(60,64,67,.10), 0 2px 6px rgba(60,64,67,.08);
}

.stApp {
    background: var(--md-bg);
    color: var(--md-text);
    font-family: "Google Sans", "Roboto", Arial, sans-serif;
}

[data-testid="stHeader"] {
    background: transparent;
}

[data-testid="stToolbar"] {
    display: none;
}

/* REMOVE COMPLETAMENTE SIDEBAR NATIVA */
section[data-testid="stSidebar"] {
    display: none !important;
}

/* REMOVE BOTÃO NATIVO */
button[kind="header"] {
    display: none !important;
}

/* REMOVE ESPAÇO LATERAL */
[data-testid="stAppViewContainer"] {
    margin-left: 0 !important;
}

.block-container {
    padding-top: 1.4rem;
    padding-bottom: 1rem;
    padding-left: 5rem;
    max-width: 100%;
}

/* =========================
   MENU LATERAL HOVER
========================= */

.hover-menu {
    position: fixed;
    top: 0;
    left: 0;
    width: 58px;
    height: 100vh;
    background: #FFFFFF;
    border-right: 1px solid #E0E3EB;
    box-shadow: 0 1px 2px rgba(60,64,67,.10),
                0 2px 6px rgba(60,64,67,.08);
    z-index: 999999999;
    transition: width 0.25s ease;
    overflow: hidden;
    padding-top: 18px;
}

.hover-menu:hover {
    width: 280px;
}

.menu-icon {
    font-size: 24px;
    color: #5F6368;
    padding-left: 18px;
    margin-bottom: 28px;
}

.menu-title {
    opacity: 0;
    white-space: nowrap;
    font-size: 16px;
    font-weight: 700;
    color: #202124;
    padding: 0 18px 18px 18px;
    transition: opacity 0.2s ease;
}

.hover-menu:hover .menu-title {
    opacity: 1;
}

.menu-item {
    display: block;
    opacity: 0;
    white-space: nowrap;
    text-decoration: none !important;
    color: #202124 !important;
    font-size: 14px;
    font-weight: 600;
    padding: 13px 18px;
    transition: opacity 0.2s ease,
                background 0.2s ease;
}

.hover-menu:hover .menu-item {
    opacity: 1;
}

.menu-item:hover {
    background: #E8F0FE;
    color: #1A73E8 !important;
}

/* =========================
   RESTANTE CSS
========================= */

.hive-title {
    font-size: 30px;
    font-weight: 600;
    color: var(--md-text);
    margin-bottom: 2px;
    letter-spacing: -.2px;
}

.hive-subtitle {
    color: var(--md-muted);
    font-size: 13px;
    font-weight: 400;
    margin-top: 2px;
    margin-bottom: 18px;
}

.kpi-card {
    background: var(--md-surface);
    border: none;
    border-radius: 24px;
    padding: 20px;
    min-height: 148px;
    box-shadow: var(--md-shadow);
    margin-bottom: 18px;
}

.kpi-label {
    color: var(--md-muted);
    font-size: 13px;
    font-weight: 600;
    text-transform: uppercase;
}

.kpi-value {
    font-size: 32px;
    font-weight: 600;
    line-height: 1.08;
    margin-top: 8px;
}

.kpi-note {
    color: var(--md-muted);
    font-size: 11px;
    margin-top: 9px;
}

.panel {
    background: var(--md-surface);
    border-radius: 24px;
    padding: 20px 22px 16px 22px;
    box-shadow: var(--md-shadow);
    margin-bottom: 18px;
}

.panel-title {
    font-size: 15px;
    font-weight: 600;
    color: var(--md-text);
    margin-bottom: 10px;
}

.chart-scroll {
    overflow-x: auto;
    overflow-y: hidden;
    width: 100%;
    padding-bottom: 8px;
}

.chart-inner {
    min-width: 100%;
}

.dark-table {
    width: 100%;
    table-layout: auto;
    border-collapse: collapse;
    background: var(--md-surface);
    color: var(--md-text);
    font-size: 12px;
    border-radius: 18px;
    overflow: hidden;
    box-shadow: var(--md-shadow);
}

.dark-table thead th {
    background: var(--md-surface-variant);
    color: var(--md-text);
    text-align: left;
    padding: 10px;
    font-weight: 600;
    border-bottom: 1px solid var(--md-border);
    vertical-align: middle;
    height: auto !important;
    min-height: 0 !important;
    white-space: nowrap;
}

.dark-table tbody td {
    padding: 9px;
    border-bottom: 1px solid var(--md-border);
    white-space: normal;
    word-break: normal;
    overflow-wrap: break-word;
    vertical-align: top;
    height: auto !important;
}

/* Histórico de críticas: Data/Hora | Inconsistência | Total */
.dark-table.tabela-criticas {
    table-layout: fixed;
}

.dark-table.tabela-criticas th:nth-child(1),
.dark-table.tabela-criticas td:nth-child(1) {
    width: 180px;
    max-width: 180px;
    white-space: nowrap;
}

.dark-table.tabela-criticas th:nth-child(2),
.dark-table.tabela-criticas td:nth-child(2) {
    width: auto;
    white-space: normal;
    word-break: normal;
    overflow-wrap: break-word;
}

.dark-table.tabela-criticas th:nth-child(3),
.dark-table.tabela-criticas td:nth-child(3) {
    width: 140px;
    max-width: 140px;
    text-align: center;
    white-space: nowrap;
}

/* Histórico por serviço: Descrição | Total | % */
.dark-table.tabela-descricao {
    table-layout: fixed;
}

.dark-table.tabela-descricao th:nth-child(1),
.dark-table.tabela-descricao td:nth-child(1) {
    width: auto;
    white-space: normal;
    word-break: normal;
    overflow-wrap: break-word;
}

.dark-table.tabela-descricao th:nth-child(2),
.dark-table.tabela-descricao td:nth-child(2) {
    width: 120px;
    max-width: 120px;
    text-align: center;
    white-space: nowrap;
}

.dark-table.tabela-descricao th:nth-child(3),
.dark-table.tabela-descricao td:nth-child(3) {
    width: 120px;
    max-width: 120px;
    text-align: center;
    white-space: nowrap;
}

/* Scroll para tabela de monitoramento */
.table-scroll {
    max-height: 420px;
    overflow-y: auto;
    overflow-x: auto;
    border-radius: 18px;
    box-shadow: var(--md-shadow);
    margin-top: 6px;
    display: block;
    position: relative;
}

.table-scroll .dark-table {
    box-shadow: none !important;
    margin-bottom: 0 !important;
    border-radius: 0 !important;
    overflow: visible !important;
}

.table-scroll .dark-table thead {
    position: sticky;
    top: 0;
    z-index: 50;
}

.table-scroll .dark-table thead th {
    position: sticky !important;
    top: 0 !important;
    z-index: 60 !important;
    background: var(--md-surface-variant) !important;
}


.dark-table tbody tr:nth-child(even) {
    background: #FAFAFA;
}

.dark-table tbody tr:nth-child(odd) {
    background: var(--md-surface);
}

</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


# =========================
# MENU
# =========================

query_params = st.query_params

pagina = query_params.get(
    "pagina",
    "Monitoramento atual"
)

menu_html = """
<div class="hover-menu">

    <div class="menu-icon">
        ☰
    </div>

    <div class="menu-title">
        Monitoramento
    </div>

    <a class="menu-item"
       href="?pagina=Monitoramento atual">
       📊 Monitoramento atual
    </a>

    <a class="menu-item"
       href="?pagina=Histórico monitoramento">
       📁 Histórico monitoramento
    </a>

</div>
"""

st.html(menu_html)

# =========================
# GITHUB / CACHE LOCAL
# =========================

def obter_token_github() -> str:
    try:
        token = st.secrets.get("TOKEN_GITHUB", "")
        if token:
            return str(token).strip()
    except Exception:
        pass
    return os.getenv("TOKEN_GITHUB", "").strip()



def github_headers():
    token = obter_token_github()
    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "monitoramento-ecrv-streamlit",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


def carregar_meta_cache() -> dict:
    try:
        if META_LOCAL.exists():
            return json.loads(META_LOCAL.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def salvar_meta_cache(meta: dict):
    try:
        META_LOCAL.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def baixar_github_se_houver_alteracao(caminho_repo: str, destino: Path, obrigatorio: bool = False):
    """
    Consulta o GitHub pela API Contents.
    Se o SHA do arquivo não mudou, usa o arquivo local em cache.
    Se mudou, baixa o novo arquivo e atualiza o cache.
    """
    destino = Path(destino)
    meta = carregar_meta_cache()
    item_meta = meta.get(caminho_repo, {})

    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/{caminho_repo}"
    )

    resp = requests.get(
        api_url,
        headers=github_headers(),
        params={"ref": GITHUB_BRANCH},
        timeout=30,
    )

    if resp.status_code == 404:
        if obrigatorio:
            raise FileNotFoundError(f"Arquivo não encontrado no GitHub: {caminho_repo}")
        return None, {"status": "ausente", "path": caminho_repo}

    resp.raise_for_status()
    dados = resp.json()
    sha_remoto = dados.get("sha") or ""

    if destino.exists() and item_meta.get("sha") == sha_remoto:
        return destino, {
            "status": "cache",
            "path": caminho_repo,
            "sha": sha_remoto,
            "downloaded_at": item_meta.get("downloaded_at", ""),
        }

    conteudo_b64 = dados.get("content", "")
    encoding = dados.get("encoding", "")

    # Para arquivos maiores, a API Contents do GitHub pode não retornar
    # o campo "content" em base64. Nesses casos, faz fallback pelo download_url.
    if conteudo_b64 and encoding == "base64":
        conteudo = base64.b64decode(conteudo_b64)
    else:
        download_url = dados.get("download_url")
        if not download_url:
            raise RuntimeError(
                f"GitHub não retornou conteúdo base64 nem download_url válido para {caminho_repo}"
            )

        resp_download = requests.get(
            download_url,
            headers=github_headers(),
            timeout=60,
        )
        resp_download.raise_for_status()
        conteudo = resp_download.content

    tmp = destino.with_suffix(destino.suffix + ".tmp")
    tmp.write_bytes(conteudo)
    tmp.replace(destino)

    downloaded_at = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    meta[caminho_repo] = {
        "sha": sha_remoto,
        "downloaded_at": downloaded_at,
        "size": dados.get("size"),
    }
    salvar_meta_cache(meta)

    return destino, {
        "status": "baixado",
        "path": caminho_repo,
        "sha": sha_remoto,
        "downloaded_at": downloaded_at,
    }

# =========================
# FUNÇÕES AUXILIARES
# =========================

def fmt_num(valor):
    try:
        return f"{int(valor):,}".replace(",", ".")
    except Exception:
        return "0"


def arquivo_mtime(path: Path):
    if not path.exists():
        return None
    return os.path.getmtime(path)


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalizar_inconsistencia(texto):
    texto_original = str(texto or "").strip()

    # Extrai somente a mensagem quando vier em JSON
    try:
        obj = json.loads(texto_original)
        if isinstance(obj, dict):
            texto_original = obj.get("mensagem", texto_original)
    except Exception:
        pass

    texto = str(texto_original or "").upper().strip()

    # Remove acentos
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    # Remove códigos iniciais: 600-, 600, 015-, 904 etc.
    texto = re.sub(r"^\s*\d+\s*[-–]?\s*", "", texto)

    # Remove prefixos técnicos no início
    texto = re.sub(
        r"^(AQUI|TC02|TCO2|TC04|DHAB|EMP3|CAFI|EDUT|BLOQ|ERRO)\s*[-–]?\s*",
        "",
        texto
    )

    # Remove códigos alfanuméricos repetidos no início
    texto = re.sub(r"^[A-Z0-9]{3,6}\s+", "", texto)

    # remove símbolos soltos no início
    texto = re.sub(r"^[+]+\s*", "", texto)
    
    # remove números isolados no início
    texto = re.sub(r"^\d+\s+", "", texto)
    
    # remove combinações tipo "+ "
    texto = re.sub(r"^[+]\s*", "", texto)

    # Limpeza geral
    texto = re.sub(r"\.{2,}", " ", texto)
    texto = re.sub(r"[-_:;,]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    texto = re.sub(r"\s*\.+\s*$", "", texto).strip()

    # =========================
    # PADRONIZAÇÕES ESPECÍFICAS
    # =========================

    if "MULTA ATIVA NO RENAINF" in texto or "EXISTENCIA DE MULTA ATIVA NO RENAINF" in texto:
        return "MULTA ATIVA NO RENAINF"

    if "SOLICITACAO DE ESTAMPAGEM EM ABERTO" in texto:
        return "SOLICITACAO DE ESTAMPAGEM EM ABERTO"

    if "PENDENCIA DE EMPLACAMENTO" in texto and "PLACA MERCOSUL" in texto:
        return "PENDENCIA DE EMPLACAMENTO (PLACA MERCOSUL)"

    if "MUNICIPIO INVALIDO" in texto:
        return "MUNICIPIO INVALIDO"

    if "NUMERO INVALIDO" in texto and "MOTOR" not in texto:
        return "NUMERO INVALIDO"

    if "NUMERO DE MOTOR" in texto and ("CARACTERE INVALIDO" in texto or "INVALIDO" in texto):
        return "NUMERO DE MOTOR INVALIDO"

    if "NOTA FISCAL ELETRONICA NAO ENCONTRADA" in texto or "NAO AUTORIZADA" in texto:
        return "NOTA FISCAL ELETRONICA NAO ENCONTRADA/NAO AUTORIZADA"

    if "INTENCAO DE GRAVAME" in texto:
        return "NAO HA REGISTRO DE CONTRATO PARA INTENCAO DE GRAVAME"

    if "VEICULO COM DEBITO" in texto and "DETRAN" in texto:
        return "VEICULO COM DEBITO DETRAN"

    if "CEP INCOMPATIVEL" in texto:
        return "CEP INCOMPATIVEL COM MUNICIPIO"

    if "FICHA NAO CADASTRADA NO GEVER" in texto:
        return "FICHA NAO CADASTRADA NO GEVER"

    if "NAO CADASTRADA NO GEVER" in texto:
        return "FICHA NAO CADASTRADA NO GEVER"

    if "OCORRENCIA DE OBITO" in texto:
        return "PROPRIETARIO COM OCORRENCIA DE OBITO"

    if "BENEFICIO TRIBUTARIO" in texto:
        return "RESTRICAO DE BENEFICIO TRIBUTARIO"

    if "SEM COMUNICACAO COM O RENAVAM" in texto:
        return "SEM COMUNICACAO COM O RENAVAM"

    if "FINANCEIRA/SNG" in texto or "BAIXA DE GRAVAMES COMPETE A FINANCEIRA" in texto:
        return "GRAVAME COMPETE A FINANCEIRA/SNG"

    if "RETORNO INESPERADO NO ENVIO DA TRANSACAO" in texto:
        return "RETORNO INESPERADO NO ENVIO DA TRANSACAO"

    if "DO ESPELHO DO DOCUMENTO INVALIDO" in texto or "TIPO DO ESPELHO DO DOCUMENTO INVALIDO" in texto:
        return "ESPELHO DO DOCUMENTO INVALIDO"

    if "RG ORGAO E/OU UF NAO PERMITIDOS" in texto:
        return "RG/UF NAO PERMITIDOS PARA PESSOA JURIDICA"

    if "EXPEDIDOR R.G. INVALIDO" in texto or "ORGAO EXPEDIDOR R.G. INVALIDO" in texto:
        return "EXPEDIDOR RG INVALIDO"

    if "JA CADASTRADA ENTRE COM NOVAS INFORMACOES" in texto:
        return "PLACA JA CADASTRADA"

    if "MOTIVO DE EMISSAO INCOMPATIVEL" in texto:
        return "MOTIVO DE EMISSAO INCOMPATIVEL"

    if "RESTRICAO ADMINISTRATIVA" in texto:
        return "RESTRICAO ADMINISTRATIVA"

    if "RESTRICAO JUDICIAL" in texto or "RENAJUD" in texto:
        return "RESTRICAO JUDICIAL/RENAJUD"

    if "RESTRICAO FINANCEIRA" in texto or "ALIENACAO FIDUCIARIA" in texto:
        return "RESTRICAO FINANCEIRA"

    if "RESTRICAO TRIBUTARIA" in texto:
        return "RESTRICAO TRIBUTARIA"

    if "CHASSI" in texto and "INVALIDO" in texto:
        return "CHASSI INVALIDO"

    if "CHASSI" in texto and ("NAO ENCONTRADO" in texto or "INEXISTENTE" in texto):
        return "CHASSI NAO ENCONTRADO"

    if "PLACA" in texto and "INVALIDA" in texto:
        return "PLACA INVALIDA"

    if "PLACA" in texto and "NAO CADASTRADA" in texto:
        return "PLACA NAO CADASTRADA"

    return texto


def carregar_media_padrao() -> pd.DataFrame:
    dados = [
        ["06:40", 11, 65, 1, 4, 11, 0, 1],
        ["06:50", 13.75, 102, 2.5, 6.75, 6.75, 0.75, 1],
        ["07:00", 15.5, 130.75, 3.5, 14.5, 8, 0.75, 1],
        ["07:10", 20.75, 181, 4.5, 8, 12.75, 0.75, 1.25],
        ["07:20", 25.75, 235, 4.75, 9.75, 18.25, 0.5, 1.25],
        ["07:30", 33.67, 287, 7.33, 16.33, 26.67, 1.33, 1.33],
        ["07:40", 41.5, 398, 12.5, 18.75, 32, 0.5, 1.25],
        ["07:50", 52.5, 480.75, 13.25, 15.25, 40.5, 0.75, 1.25],
        ["08:00", 58.75, 561.5, 15.75, 44.25, 51.75, 5.5, 1.25],
        ["08:10", 77, 699.5, 24.5, 24.75, 64.25, 1.5, 1.5],
        ["08:20", 97, 816.75, 31.25, 18.5, 75.25, 1.25, 1.25],
        ["08:30", 120.25, 960.25, 44.75, 38.75, 91.75, 13.25, 1.25],
        ["08:40", 158.5, 1134, 60, 39, 112.25, 3.75, 2.5],
        ["08:50", 199.25, 1335.75, 71.75, 35.25, 132.75, 4.25, 2.5],
        ["09:00", 239.25, 1476.75, 92.5, 104.75, 148, 30, 2.75],
        ["09:10", 300.5, 1743.75, 135.75, 108.75, 173.5, 6.75, 3.75],
        ["09:20", 387.25, 2004.5, 150, 95, 201.25, 3.5, 4.25],
        ["09:30", 449, 2264.5, 176.75, 111.75, 234.75, 37.25, 4.75],
        ["09:40", 550.75, 2509.75, 222, 150, 263, 6.25, 5],
        ["09:50", 503.8, 2222.8, 197.2, 145.8, 240.8, 4.8, 4.4],
        ["10:00", 628.75, 2787.75, 277.25, 225.25, 316.25, 16.25, 5.5],
        ["10:10", 708, 3104.5, 278.5, 187.75, 331.75, 7, 6.25],
        ["10:20", 786.5, 3323, 333.25, 188, 361.75, 8.5, 8.25],
        ["10:30", 868.25, 3536.5, 380.5, 236.75, 390.5, 39.5, 8.75],
        ["10:40", 948.75, 3763.75, 423.25, 153.75, 421.25, 5.5, 10],
        ["10:50", 1039.75, 3935.75, 468.75, 170.75, 438.5, 19.5, 9.5],
        ["11:00", 1117, 4189, 523, 187.25, 489.5, 17.5, 11],
        ["11:10", 1206, 4465, 564.25, 134.5, 533.25, 10, 10.75],
        ["11:20", 1306.5, 4700.25, 604, 166, 584.75, 19.75, 12],
        ["11:30", 1354.5, 4991.75, 662.75, 180, 626.25, 32.5, 12.75],
        ["11:40", 1422.5, 5277.25, 691.75, 144.25, 654.25, 6.75, 13],
        ["11:50", 1530.5, 5510.25, 736.75, 172, 683.5, 16, 13.75],
        ["12:00", 1637, 5743, 779.5, 166.5, 715.5, 31.5, 15],
        ["12:10", 1726, 5981.5, 823.25, 160.75, 742.5, 14.75, 15.5],
        ["12:20", 1811.5, 6134.25, 842.75, 154.25, 786.75, 16.25, 16.25],
        ["12:30", 1875.75, 6306.25, 897.25, 145, 798.75, 41.75, 16.75],
        ["12:40", 2005.25, 6579.5, 911.75, 140.5, 853, 9, 17.5],
        ["12:50", 2099.5, 6688.25, 972, 130.5, 876, 19.5, 18],
        ["13:00", 2161.25, 6846, 1021.75, 167.75, 922, 15.75, 18.75],
        ["13:10", 2246.75, 7024.25, 1026.25, 155.5, 946.5, 10, 19.75],
        ["13:20", 2316, 7176.5, 1062.5, 152.25, 951.5, 14.5, 20.25],
        ["13:30", 2384.75, 7359, 1102, 126.25, 982.5, 28.25, 21],
        ["13:40", 2432.5, 7551, 1119, 107.25, 1012, 10.25, 21.75],
        ["13:50", 2507.25, 7624, 1169.25, 120.75, 1017.25, 24, 22.25],
        ["14:00", 2593.25, 7746, 1197, 128.5, 1060.5, 17, 23],
        ["14:10", 2657.25, 7837.5, 1196.5, 106.25, 1046.25, 9, 23.75],
        ["14:20", 2688.75, 8002.25, 1255.75, 89.5, 1076.75, 14.5, 24.5],
        ["14:30", 2789, 8260.5, 1286, 109, 1114.75, 48, 25.5],
        ["14:40", 2865, 8388.5, 1330.75, 91, 1126, 12.75, 25.75],
        ["14:50", 2929.5, 8541, 1346.75, 92.5, 1144, 17.25, 26],
        ["15:00", 3041.5, 8785.25, 1400.5, 106.75, 1198.75, 24, 27],
        ["15:10", 3101.75, 9052, 1461.5, 100, 1232, 11.25, 28],
        ["15:20", 3195.25, 9222.5, 1507.25, 91.75, 1258, 18.75, 29],
        ["15:30", 3316.75, 9624.5, 1513.75, 93.5, 1295.25, 33.25, 29.75],
        ["15:40", 3423.5, 9809.25, 1572.25, 66.25, 1323.75, 8.75, 31],
        ["15:50", 3526, 10088.75, 1628.75, 75, 1356.25, 21.75, 32],
        ["16:00", 3656, 10358.75, 1694, 113.75, 1387.25, 28.5, 32.75],
        ["16:10", 3771.25, 10639, 1784.25, 95.75, 1382.25, 15.75, 33],
        ["16:20", 3952.5, 10867.5, 1840.75, 66.5, 1421.75, 15.25, 34.5],
        ["16:30", 4200.5, 11422, 1926.5, 64.5, 1452.75, 57.25, 35.5],
        ["16:40", 4419.5, 11753.25, 1985.25, 51.75, 1490, 33.25, 36.25],
        ["16:50", 4553.5, 12002, 2074.75, 47, 1514.5, 24, 37.25],
        ["17:00", 4735, 12293.75, 2169.75, 54.5, 1538.25, 37.75, 38.25],
        ["17:10", 4896.25, 12541.5, 2268.25, 52, 1546, 23, 39],
        ["17:20", 5026.75, 12740.5, 2375.5, 45, 1568, 16.5, 39.75],
        ["17:30", 5167.75, 12937.5, 2476.75, 38.5, 1587.5, 23.25, 40.5],
        ["17:40", 5291.25, 13103, 2584, 31, 1599.75, 18.5, 41.25],
        ["17:50", 5400.25, 13245, 2663.5, 19.75, 1604.25, 14.75, 42],
        ["18:00", 5481, 13345.75, 2768.25, 18.5, 1624.75, 15.25, 43],
        ["18:10", 5552.75, 13446.25, 2835.75, 13.25, 1631.5, 11.25, 43.75],
        ["18:20", 5630, 13536, 2906.5, 10.5, 1633, 9, 44.5],
        ["18:30", 5678.25, 13587.25, 2965.5, 8, 1640.25, 5.5, 45.25],
        ["18:40", 5715.75, 13636.75, 2990.75, 4.75, 1646.25, 4.25, 45.75],
        ["18:50", 5757.25, 13686, 3023, 2.5, 1651, 4.25, 46.5],
        ["19:00", 5796.5, 13728.5, 3043.75, 4.25, 1658.75, 3, 47.25],
        ["19:10", 5829, 13775.25, 3069.75, 1.75, 1665, 2.25, 48],
        ["19:20", 5860.5, 13816.25, 3090.25, 1.25, 1668.75, 1.75, 48.75],
        ["19:30", 5883, 13860.75, 3108.75, 0.75, 1674.25, 1, 49.5],
        ["19:40", 5904.75, 13899.75, 3125.25, 0.5, 1681.25, 0.75, 50],
        ["19:50", 5922.25, 13939.25, 3142, 0.75, 1685.5, 0.75, 50.75],
        ["20:00", 5940.25, 13967.5, 3158.5, 1.5, 1690.25, 1.25, 51.5],
        ["20:10", 5958.25, 14001.25, 3174.5, 0.75, 1695, 0.5, 52],
        ["20:20", 5974, 14022.75, 3185.25, 0.75, 1698.75, 0.75, 52.5],
        ["20:30", 5986.75, 14043.5, 3197, 0.5, 1704.5, 0.5, 53],
        ["20:40", 6000.25, 14062.75, 3205.5, 0.5, 1706.75, 0.5, 53.75],
        ["20:50", 6015, 14080.75, 3213.25, 0.25, 1708, 1, 54.25],
        ["21:00", 6027.5, 14095, 3221.5, 0.5, 1710, 0.5, 54.75],
        ["21:10", 6041.5, 14114.25, 3231.5, 0.25, 1712.25, 0.25, 55.25],
        ["21:20", 6048.75, 14124.75, 3239.75, 0.25, 1713.25, 0.25, 55.75],
        ["21:30", 6055.25, 14134.5, 3247.5, 0.25, 1715.25, 0.25, 56.25],
        ["21:40", 6061.25, 14141, 3254.75, 0.25, 1716.5, 0.25, 56.75],
        ["21:50", 6065.75, 14149, 3261.25, 0.25, 1717.25, 0.25, 57.25],
    ]

    return pd.DataFrame(
        dados,
        columns=[
            "Intervalo_10min",
            "Automatizado",
            "Sucesso 2 e 3",
            "Sucesso 0km",
            "Fila 2 e 3",
            "Inconsistência 2 e 3",
            "Fila 0km",
            "Inconsistência 0km",
        ],
    )


@st.cache_data(show_spinner=False)
def carregar_excel(path_str: str, mtime: float):
    df = pd.read_excel(path_str, engine="openpyxl")
    df = normalizar_colunas(df)

    if "Horário" in df.columns:
        df["Horário"] = df["Horário"].astype(str).str.slice(0, 5)

    for col in df.columns:
        if col not in ("Horário", "Data/Hora", "Data"):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


@st.cache_data(show_spinner=False)
def carregar_excel_generico(path_str: str, mtime: float):
    df = pd.read_excel(path_str, engine="openpyxl")
    df = normalizar_colunas(df)
    return df


@st.cache_data(show_spinner=False)
def carregar_csv_historico(path_str: str, mtime: float):
    # Leitura robusta para CSV separado por ; ou ,
    try:
        df = pd.read_csv(path_str, sep=";", encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv(path_str, sep=",", encoding="utf-8-sig")

    df = normalizar_colunas(df)

    # Se o CSV entrou como uma única coluna contendo ;
    if len(df.columns) == 1 and ";" in str(df.columns[0]):
        df = pd.read_csv(path_str, sep=";", encoding="utf-8-sig")
        df = normalizar_colunas(df)

    df.columns = [str(c).strip() for c in df.columns]

    if "Data/Hora" in df.columns:
        df["Data/Hora"] = pd.to_datetime(
            df["Data/Hora"],
            dayfirst=True,
            errors="coerce",
            format="mixed",
        )
        df["Data"] = df["Data/Hora"].dt.date

    if "Horário" in df.columns:
        df["Horário"] = df["Horário"].astype(str).str.slice(0, 5)

    colunas_nao_numericas = {
        "Data/Hora",
        "Data",
        "Horário",
        "Tipo Histórico",
        "Tipo Historico",
        "Serviço",
        "Servico",
        "Inconsistência",
        "Inconsistencia",
        "Inconsistência nova no ciclo?",
        "STOP PROCESSO ativado?",
        "Serviços desligados",
        "Falhas ao desligar",
    }

    for col in df.columns:
        if col not in colunas_nao_numericas:
            convertido = pd.to_numeric(df[col], errors="coerce")
            if convertido.notna().any():
                df[col] = convertido.fillna(0)

    return df

def agrupar_monitoramento_por_horario(df_base: pd.DataFrame) -> pd.DataFrame:
    if df_base is None or df_base.empty or "Horário" not in df_base.columns:
        return df_base

    df_aux = df_base.copy()
    df_aux["Horário"] = df_aux["Horário"].astype(str).str.slice(0, 5)

    colunas_numericas = [
        c for c in df_aux.columns
        if c not in ("Data/Hora", "Data", "Horário")
        and pd.api.types.is_numeric_dtype(pd.to_numeric(df_aux[c], errors="coerce"))
    ]

    if not colunas_numericas:
        return df_aux.sort_values("Horário").reset_index(drop=True)

    agg = {c: "max" for c in colunas_numericas}
    if "Data/Hora" in df_aux.columns:
        agg["Data/Hora"] = "last"
    if "Data" in df_aux.columns:
        agg["Data"] = "last"

    return (
        df_aux
        .groupby("Horário", as_index=False)
        .agg(agg)
        .sort_values("Horário")
        .reset_index(drop=True)
    )


def media_coluna(df_media: pd.DataFrame, coluna: str, horario: str = None):
    """
    Retorna a média da coluna considerando o horário fechado de 10 em 10 minutos.

    Exemplo:
    11:42 -> usa 11:40
    14:18 -> usa 14:10

    Se não encontrar o horário, usa a média geral da coluna.
    """

    if df_media is None or df_media.empty or coluna not in df_media.columns:
        return None

    try:

        # =========================
        # MÉDIA POR HORÁRIO
        # =========================
        if horario and "Intervalo_10min" in df_media.columns:

            h = pd.to_datetime(str(horario), errors="coerce")

            if not pd.isna(h):

                intervalo = h.floor("10min").strftime("%H:%M")

                linha = df_media[
                    df_media["Intervalo_10min"].astype(str).str.strip() == intervalo
                ]

                if not linha.empty:

                    valor = pd.to_numeric(
                        linha.iloc[0][coluna],
                        errors="coerce"
                    )

                    if not pd.isna(valor):
                        return float(valor)

        # =========================
        # FALLBACK → MÉDIA GERAL
        # =========================
        s = pd.to_numeric(df_media[coluna], errors="coerce").dropna()

        if s.empty:
            return None

        return float(s.mean())

    except Exception:
        return None

def cor_saude(valor: int, media, tipo: str):
    """
    tipo='positivo': sucesso e automatizado
        verde: >= média
        amarelo: até 25% abaixo
        vermelho: abaixo de 25%

    tipo='negativo': fila e inconsistência
        verde: <= média
        amarelo: até 25% acima
        vermelho: acima de 25%
    """
    if media is None or media <= 0:
        return "#FFFFFF"

    valor = float(valor)

    if tipo == "positivo":
        if valor >= media:
            return "#188038"
        elif valor >= media * 0.75:
            return "#F9AB00"
        else:
            return "#D93025"

    if tipo == "negativo":
        if valor <= media:
            return "#188038"
        elif valor <= media * 1.25:
            return "#F9AB00"
        else:
            return "#D93025"

    return "#FFFFFF"

def cor_criticas_minuto(valor):
    valor = int(valor or 0)

    if valor == 0:
        return "#188038"  # verde
    elif 1 <= valor <= 5:
        return "#F9AB00"  # amarelo
    else:
        return "#D93025"  # vermelho

def obter_total_criticas_minuto(df_criticas: pd.DataFrame) -> int:
    if df_criticas is None or df_criticas.empty:
        return 0

    if "Tipo Linha" in df_criticas.columns:
        resumo = df_criticas[
            df_criticas["Tipo Linha"].astype(str).str.upper().str.strip() == "RESUMO"
        ]
        if not resumo.empty and "Total críticas no minuto" in resumo.columns:
            return int(pd.to_numeric(resumo.iloc[-1]["Total críticas no minuto"], errors="coerce") or 0)

    if "Total críticas no minuto" in df_criticas.columns:
        s = pd.to_numeric(df_criticas["Total críticas no minuto"], errors="coerce").dropna()
        if not s.empty:
            return int(s.iloc[-1])

    return 0


def _valor_sim(valor) -> bool:
    return str(valor or "").strip().upper() == "SIM"


def _to_int_safe(valor) -> int:
    try:
        if pd.isna(valor):
            return 0
    except Exception:
        pass
    try:
        return int(float(str(valor).replace(".", "").replace(",", ".").strip()))
    except Exception:
        return 0


def _coluna_existente(df: pd.DataFrame, alternativas: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    mapa = {str(c).strip().lower(): c for c in df.columns}
    for alt in alternativas:
        col = mapa.get(str(alt).strip().lower())
        if col is not None:
            return col
    return None


def _df_historico_full(df_hist: pd.DataFrame):
    """
    Retorna linhas de inconsistências por serviço,
    aceitando padrão antigo FULL e novo INCONSISTÊNCIA_DIA.
    """

    if df_hist is None or df_hist.empty:
        return pd.DataFrame()

    dfh = df_hist.copy()

    col_tipo = _coluna_existente(
        dfh,
        ["Tipo Histórico", "Tipo Historico"]
    )

    if col_tipo:

        tipo = (
            dfh[col_tipo]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        filtro = (
            tipo.str.contains("INCONSIST", na=False)
            &
            (
                tipo.str.contains("FULL", na=False)
                |
                tipo.str.contains("DIA", na=False)
            )
        )

        dfh = dfh[filtro].copy()

    col_servico = _coluna_existente(
        dfh,
        ["Serviço", "Servico", "Tipo Serviço", "Tipo Servico"]
    )

    col_inc = _coluna_existente(
        dfh,
        [
            "Inconsistência",
            "Inconsistencia",
            "Descrição Inconsistência",
            "Descricao Inconsistencia"
        ]
    )

    if not col_servico or not col_inc:
        return pd.DataFrame()

    return dfh


def _df_historico_full_ultimo_ciclo(df_hist: pd.DataFrame) -> pd.DataFrame:
    dfh = _df_historico_full(df_hist)
    if dfh.empty:
        return pd.DataFrame()

    col_data = _coluna_existente(dfh, ["Data/Hora", "Data Hora", "Data"])
    if not col_data:
        return dfh.tail(100).copy()

    datas = dfh[col_data].dropna().astype(str)
    if datas.empty:
        return dfh.tail(100).copy()

    ultima_data = datas.iloc[-1]
    return dfh[dfh[col_data].astype(str) == ultima_data].copy()


def _servico_para_chave(valor: str) -> str | None:

    txt = str(valor or "").lower()

    # 0KM
    if any(x in txt for x in [
        "primeiro registro",
        "primeiro registro do veículo",
        "primeiro registro do veiculo",
        "0km",
        "0 km",
        "pr"
    ]):
        return "0KM"

    # Transferência proprietário
    if any(x in txt for x in [
        "transferência proprietário",
        "transferencia proprietario",
        "proprietário",
        "proprietario",
        "tp"
    ]):
        return "Transferência 2"

    # Transferência município/estado
    if any(x in txt for x in [
        "transferência município",
        "transferencia municipio",
        "transferência estado",
        "transferencia estado",
        "município",
        "municipio",
        "estado",
        "tm",
        "te"
    ]):
        return "Transferência 3"

    return None

def _servicos_stop_sim(df_criticas: pd.DataFrame, df_hist: pd.DataFrame) -> set[str]:
    """
    Considera somente STOPs recentes (últimos 20 minutos)
    e identifica os serviços desligados.
    """

    fontes = []

    if df_criticas is not None and not df_criticas.empty:
        fontes.append(df_criticas)

    if df_hist is not None and not df_hist.empty:
        fontes.append(df_hist)

    registros = []

    for df0 in fontes:

        col_stop = _coluna_existente(df0, ["STOP PROCESSO ativado?"])
        col_data = _coluna_existente(df0, ["Data/Hora", "Data Hora", "Data"])
        col_servico = _coluna_existente(df0, ["Serviço", "Servico", "Tipo Serviço", "Tipo Servico"])
        col_desligados = _coluna_existente(df0, ["Serviços desligados", "Servicos desligados"])

        if not col_stop:
            continue

        tmp = df0.copy()

        # Apenas linhas com STOP = SIM
        tmp = tmp[
            tmp[col_stop].astype(str).str.upper().str.strip() == "SIM"
        ].copy()

        if tmp.empty:
            continue

        # Data do registro
        if col_data:
            tmp["_data"] = pd.to_datetime(
                tmp[col_data],
                dayfirst=True,
                errors="coerce"
            )
        else:
            tmp["_data"] = pd.Timestamp.now()

        for _, row in tmp.iterrows():

            textos = []

            if col_servico:
                textos.append(str(row.get(col_servico, "")))

            if col_desligados:
                textos.append(str(row.get(col_desligados, "")))

            combinado = " | ".join(textos).lower()

            chave = None

            if (
                "primeiro" in combinado
                or "0km" in combinado
                or "0 km" in combinado
            ):
                chave = "0KM"

            elif (
                "propriet" in combinado
            ):
                chave = "Transferência 2"

            elif (
                "estado" in combinado
                or "munic" in combinado
            ):
                chave = "Transferência 3"

            if chave:
                registros.append({
                    "servico": chave,
                    "data": row["_data"]
                })

    if not registros:
        return set()

    df_reg = pd.DataFrame(registros)

    # Mantém somente o STOP mais recente de cada serviço
    df_reg = (
        df_reg
        .sort_values("data")
        .drop_duplicates(subset=["servico"], keep="last")
    )

    return set(df_reg["servico"].tolist())

def _ultimo_stop_servico(df_criticas: pd.DataFrame, df_hist: pd.DataFrame, chave_servico: str):
    fontes = []

    if df_criticas is not None and not df_criticas.empty:
        fontes.append(df_criticas)

    if df_hist is not None and not df_hist.empty:
        fontes.append(df_hist)

    registros = []

    for df0 in fontes:
        col_stop = _coluna_existente(df0, ["STOP PROCESSO ativado?"])
        col_data = _coluna_existente(df0, ["Data/Hora", "Data Hora", "Data"])
        col_servico = _coluna_existente(df0, ["Serviço", "Servico", "Tipo Serviço", "Tipo Servico"])
        col_desligados = _coluna_existente(df0, ["Serviços desligados", "Servicos desligados"])

        if not col_stop or not col_data:
            continue

        tmp = df0[
            df0[col_stop].astype(str).str.upper().str.strip() == "SIM"
        ].copy()

        if tmp.empty:
            continue

        tmp["_data"] = pd.to_datetime(tmp[col_data], dayfirst=True, errors="coerce")

        for _, row in tmp.iterrows():
            textos = []

            if col_servico:
                textos.append(str(row.get(col_servico, "")))

            if col_desligados:
                textos.append(str(row.get(col_desligados, "")))

            combinado = " | ".join(textos)
            chave = _servico_para_chave(combinado)

            if chave == chave_servico and not pd.isna(row["_data"]):
                registros.append(row["_data"])

    if not registros:
        return None

    return max(registros)


def _existe_full_posterior_ao_stop(df_criticas: pd.DataFrame, df_hist: pd.DataFrame, chave_servico: str) -> bool:
    ultimo_stop = _ultimo_stop_servico(df_criticas, df_hist, chave_servico)

    if ultimo_stop is None:
        return False

    dfh = _df_historico_full(df_hist)

    if dfh.empty:
        return False

    col_data = _coluna_existente(dfh, ["Data/Hora", "Data Hora", "Data"])
    col_servico = _coluna_existente(dfh, ["Serviço", "Servico", "Tipo Serviço", "Tipo Servico"])

    if not col_data or not col_servico:
        return False

    tmp = dfh.copy()
    tmp["_data"] = pd.to_datetime(tmp[col_data], dayfirst=True, errors="coerce")
    tmp["_chave_servico"] = tmp[col_servico].apply(_servico_para_chave)

    tmp = tmp[
        (tmp["_chave_servico"] == chave_servico) &
        (tmp["_data"] > ultimo_stop)
    ].copy()

    return not tmp.empty


def status_robos(df_monitor: pd.DataFrame, df_criticas: pd.DataFrame, df_hist: pd.DataFrame = None):
    """
    Regra:
    - OFF: quando houver STOP PROCESSO ativado = SIM para o serviço.
    - ON: quando houver novo ciclo de monitoramento posterior ao STOP
      com aumento no contador de inconsistência do serviço.
    """

    status = {
        "0KM": "ON",
        "Transferência 2": "ON",
        "Transferência 3": "ON",
    }

    # Primeiro marca OFF pelos STOPs encontrados
    for servico in _servicos_stop_sim(df_criticas, df_hist):
        status[servico] = "OFF"

    if df_monitor is None or df_monitor.empty:
        return status

    dfm = normalizar_coluna_data_para_date(df_monitor)

    if "Data/Hora" not in dfm.columns:
        return status

    dfm["_data_hora_monitor"] = pd.to_datetime(
        dfm["Data/Hora"],
        dayfirst=True,
        errors="coerce",
        format="mixed",
    )

    def houve_aumento_apos_stop(chave_servico: str, coluna_incons: str) -> bool:
        ultimo_stop = _ultimo_stop_servico(df_criticas, df_hist, chave_servico)

        if ultimo_stop is None:
            return False

        if coluna_incons not in dfm.columns:
            return False

        tmp = dfm[dfm["_data_hora_monitor"] > ultimo_stop].copy()

        if len(tmp) < 2:
            return False

        tmp[coluna_incons] = pd.to_numeric(
            tmp[coluna_incons],
            errors="coerce"
        ).fillna(0)

        return tmp[coluna_incons].diff().fillna(0).gt(0).any()

    # 0KM volta ON se Inconsistência 0km aumentou após o STOP
    if status["0KM"] == "OFF":
        if houve_aumento_apos_stop("0KM", "Inconsistência 0km"):
            status["0KM"] = "ON"

    # Transferência 2 e 3 usam a mesma coluna agregada
    if status["Transferência 2"] == "OFF":
        if houve_aumento_apos_stop("Transferência 2", "Inconsistência 2 e 3"):
            status["Transferência 2"] = "ON"

    if status["Transferência 3"] == "OFF":
        if houve_aumento_apos_stop("Transferência 3", "Inconsistência 2 e 3"):
            status["Transferência 3"] = "ON"

    return status

def render_card(label, value, color, note="Último registro"):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="color:{color};">{fmt_num(value)}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_robos_card(status_dict, robo_monitoramento_online=True):
    status_0km = status_dict.get("0KM", "ON")
    status_t2 = status_dict.get("Transferência 2", "ON")
    status_t3 = status_dict.get("Transferência 3", "ON")
    status_monitoramento = "ON" if robo_monitoramento_online else "OFF"

    def bolinha(status):
        return "🟢" if status == "ON" else "🔴"

    def cor(status):
        return "#188038" if status == "ON" else "#D93025"

    html = f"""
    <div class="kpi-card">
        <div class="kpi-label">Robôs</div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-top:10px; align-items:start;">
            <div>
                <div style="font-size:13px; font-weight:600; line-height:1.45; margin-bottom:10px; white-space:nowrap;">
                    {bolinha(status_t2)} <span style="color:{cor(status_t2)}; font-weight:700;">{status_t2}</span> Transferência 2
                </div>
                <div style="font-size:13px; font-weight:600; line-height:1.45; white-space:nowrap;">
                    {bolinha(status_t3)} <span style="color:{cor(status_t3)}; font-weight:700;">{status_t3}</span> Transferência 3
                </div>
            </div>
            <div>
                <div style="font-size:13px; font-weight:600; line-height:1.45; margin-bottom:10px; white-space:nowrap;">
                    {bolinha(status_0km)} <span style="color:{cor(status_0km)}; font-weight:700;">{status_0km}</span> 0KM
                </div>
                <div style="font-size:13px; font-weight:600; line-height:1.45; white-space:nowrap;">
                    {bolinha(status_monitoramento)} <span style="color:{cor(status_monitoramento)}; font-weight:700;">{status_monitoramento}</span> e-CRV
                </div>
            </div>
        </div>
    </div>
    """

    st.markdown(html, unsafe_allow_html=True)

def fig_layout(fig, height=520):
    fig.update_layout(
        height=height,
        autosize=True,
        margin=dict(l=28, r=28, t=125, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#202124", family="Google Sans, Roboto, Arial"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.06,
            xanchor="left",
            x=0,
            font=dict(size=11, color="#5F6368"),
            bgcolor="rgba(255,255,255,0)",
        ),

        xaxis=dict(
            gridcolor="#E0E0E0",
            zerolinecolor="#E0E0E0",
            tickfont=dict(color="#5F6368"),
            tickangle=-45,
            rangeslider=dict(visible=True),
        ),

        yaxis=dict(
            gridcolor="#E0E0E0",
            zerolinecolor="#E0E0E0",
            tickfont=dict(color="#5F6368"),
        ),
    )
    return fig


def line_chart(df, cols, title):
    fig = go.Figure()

    x_labels = df["Horário"].astype(str).tolist() if "Horário" in df.columns else [str(i) for i in df.index]
    x_vals = list(range(len(x_labels)))

    for col in cols:
        if col not in df.columns:
            continue

        y_vals = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

        nome_lower = col.lower()
        if "sucesso" in nome_lower:
            cor = "#188038"
        elif "incons" in nome_lower:
            cor = "#D93025"
        elif "fila" in nome_lower:
            cor = "#F9AB00"
        else:
            cor = "#1A73E8"

        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=y_vals,
                mode="lines+markers+text",
                text=[fmt_num(v) for v in y_vals],
                textposition="top center",
                textfont=dict(size=10, color=cor),
                name=col,
                line=dict(width=2.4, color=cor),
                marker=dict(size=5, color=cor),
                customdata=x_labels,
                hovertemplate="%{customdata}<br>" + col + ": %{y}<extra></extra>",
                cliponaxis=False,
            )
        )

    passo = max(1, len(x_vals) // 20)
    tickvals = x_vals[::passo]
    ticktext = x_labels[::passo]

    janela = 25
    inicio = max(0, len(x_vals) - janela)
    fim = max(1, len(x_vals) - 1)

    fig.update_xaxes(
        tickmode="array",
        tickvals=tickvals,
        ticktext=ticktext,
        tickangle=-45,
        range=[inicio, fim],
        rangeslider=dict(visible=True),
    )

    # Reserva espaço no eixo Y para os rótulos acima dos pontos,
    # evitando corte dos valores mais altos.
    y_max = 0
    for col in cols:
        if col in df.columns:
            serie = pd.to_numeric(df[col], errors="coerce").fillna(0)
            if not serie.empty:
                y_max = max(y_max, float(serie.max()))

    if y_max > 0:
        fig.update_yaxes(range=[0, y_max * 1.28])
    else:
        fig.update_yaxes(range=[0, 1])

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=16, color="#202124"),
            y=0.98,
            x=0.01,
            xanchor="left",
        )
    )

    return fig_layout(fig, height=520)

def extrair_mensagem_inconsistencia(valor):
    texto = str(valor or "").strip()

    if not texto:
        return ""

    try:
        obj = json.loads(texto)
        if isinstance(obj, dict):
            return str(
                obj.get("mensagem")
                or obj.get("ERRORMESSAGE")
                or obj.get("WEBERRORMESSAGE")
                or texto
            ).strip()
    except Exception:
        pass

    # Alguns retornos chegam como texto quase JSON, mas sem dois-pontos.
    texto_upper = texto.upper()
    for chave in ["ERRORMESSAGE", "WEBERRORMESSAGE", "MENSAGEM"]:
        pos = texto_upper.find(chave)
        if pos >= 0:
            recorte = texto[pos + len(chave):]
            recorte = recorte.replace('":', " ").replace('";', " ")
            recorte = recorte.replace('"', " ").replace("}", " ")
            recorte = " ".join(recorte.split())
            if recorte:
                return recorte

    return texto


def quebrar_texto_longo(valor, limite=120):
    texto = extrair_mensagem_inconsistencia(valor)

    if len(texto) <= limite:
        return texto

    partes = []
    atual = ""

    for pedaco in texto.split(" "):
        if len(atual) + len(pedaco) + 1 > limite:
            partes.append(atual.strip())
            atual = pedaco
        else:
            atual += " " + pedaco

    if atual.strip():
        partes.append(atual.strip())

    return "<br>".join(partes)


def extrair_mensagem_inconsistencia(valor):
    texto = str(valor or "").strip()

    if not texto:
        return ""

    try:
        obj = json.loads(texto)
        if isinstance(obj, dict):
            return str(
                obj.get("mensagem")
                or obj.get("ERRORMESSAGE")
                or obj.get("WEBERRORMESSAGE")
                or texto
            ).strip()
    except Exception:
        pass

    texto_upper = texto.upper()
    for chave in ["ERRORMESSAGE", "WEBERRORMESSAGE", "MENSAGEM"]:
        pos = texto_upper.find(chave)
        if pos >= 0:
            recorte = texto[pos + len(chave):]
            recorte = recorte.replace('":', " ").replace('";', " ")
            recorte = recorte.replace('"', " ").replace("}", " ")
            recorte = " ".join(recorte.split())
            if recorte:
                return recorte

    return texto


def quebrar_texto_longo(valor, limite=120):
    texto = extrair_mensagem_inconsistencia(valor)

    if len(texto) <= limite:
        return texto

    partes = []
    atual = ""

    for pedaco in texto.split(" "):
        if len(atual) + len(pedaco) + 1 > limite:
            partes.append(atual.strip())
            atual = pedaco
        else:
            atual += " " + pedaco

    if atual.strip():
        partes.append(atual.strip())

    return "<br>".join(partes)


def formatar_valor_tabela(valor):
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    texto = str(valor).strip()

    if texto == "":
        return ""

    # Remove .0 de números inteiros em formato texto: 3393.0 -> 3393
    if re.fullmatch(r"-?\d+\.0+", texto):
        return texto.split(".")[0]

    # Remove .0 mesmo quando vier com espaços
    if texto.endswith(".0"):
        base = texto[:-2]
        if re.fullmatch(r"-?\d+", base):
            return base

    # Converte números: 46.0 -> 46 | 0.13 -> 0,13
    try:
        numero = float(texto.replace(",", "."))
        if numero.is_integer():
            return str(int(numero))
        return str(round(numero, 2)).replace(".", ",")
    except Exception:
        return texto

def render_tabela_escura(df_tabela: pd.DataFrame, scroll: bool = False):

    df_tabela = df_tabela.copy()

    for col in df_tabela.columns:

        nome_col = str(col).lower()

        if (
            "descr" in nome_col
            or "mensagem" in nome_col
            or "erro" in nome_col
            or "retorno" in nome_col
            or "inconsist" in nome_col
        ):
            df_tabela[col] = df_tabela[col].apply(lambda x: formatar_valor_tabela(quebrar_texto_longo(x, limite=120)))
            continue

        # Formata coluna numérica ou valores numéricos isolados, removendo .0.
        df_tabela[col] = df_tabela[col].apply(formatar_valor_tabela)

    colunas_lower = [str(c).strip().lower() for c in df_tabela.columns]

    if "data/hora" in colunas_lower and (
        "inconsistência" in colunas_lower or "inconsistencia" in colunas_lower
    ):
        classe_extra = " tabela-criticas"
    elif (
        ("descrição" in colunas_lower or "descricao" in colunas_lower)
        and "total" in colunas_lower
        and "%" in colunas_lower
    ):
        classe_extra = " tabela-descricao"
    else:
        classe_extra = ""

    html = df_tabela.to_html(index=False, escape=False)

    html = html.replace(
        '<table border="1" class="dataframe">',
        f'<table class="dark-table{classe_extra}">'
    )

    if scroll:
        html = f'<div class="table-scroll">{html}</div>'

    st.markdown(html, unsafe_allow_html=True)

def render_mensagem_tabela(texto: str):
    st.markdown(
        f"""
        <div style="
            background: #FFFFFF;
            border: 1px solid #DADCE0;
            border-radius: 18px;
            padding: 12px 14px;
            color: #5F6368;
            font-size: 13px;
            font-weight: 600;
            box-shadow: 0 1px 2px rgba(60,64,67,.10), 0 2px 6px rgba(60,64,67,.08);
        ">{texto}</div>
        """,
        unsafe_allow_html=True,
    )


def normalizar_coluna_data_para_date(df_base: pd.DataFrame) -> pd.DataFrame:
    """Garante que a coluna Data esteja sempre em formato date, evitando falha no filtro da tela Histórico."""
    if df_base is None or df_base.empty:
        return pd.DataFrame() if df_base is None else df_base

    df_base = df_base.copy()

    if "Data/Hora" in df_base.columns:
        df_base["Data/Hora"] = pd.to_datetime(
            df_base["Data/Hora"],
            dayfirst=True,
            errors="coerce",
            format="mixed",
        )
        df_base["Data"] = df_base["Data/Hora"].dt.date

    elif "Data" in df_base.columns:
        df_base["Data"] = pd.to_datetime(
            df_base["Data"],
            dayfirst=True,
            errors="coerce",
        ).dt.date

    return df_base

def _preparar_historico_full_ultimo_ciclo(df_hist: pd.DataFrame) -> pd.DataFrame:
    dfh = _df_historico_full_ultimo_ciclo(df_hist)
    if dfh.empty:
        return pd.DataFrame()

    col_servico = _coluna_existente(dfh, ["Serviço", "Servico", "Tipo Serviço", "Tipo Servico"])
    col_inc = _coluna_existente(dfh, ["Inconsistência", "Inconsistencia", "Descrição Inconsistência", "Descricao Inconsistencia"])
    col_ant = _coluna_existente(dfh, ["Total ciclo anterior", "Total Ciclo Anterior"])
    col_atual = _coluna_existente(dfh, ["Total ciclo atual", "Total Ciclo Atual"])
    col_dif = _coluna_existente(dfh, ["Diferença ciclo", "Diferenca ciclo", "Diferença", "Diferenca"])
    col_nova = _coluna_existente(dfh, [
        "Inconsistência Não Mapeada?",
        "Inconsistencia Nao Mapeada?",
        "Nova na Base de Regras?",
        "Inconsistência nova no ciclo?",
        "Inconsistencia nova no ciclo?",
    ])

    if not col_servico or not col_inc:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["Serviço"] = dfh[col_servico].astype(str)
    out["Descrição"] = dfh[col_inc].astype(str)
    out["Descrição agrupada"] = out["Descrição"].apply(normalizar_inconsistencia)
    out["Total anterior"] = pd.to_numeric(dfh[col_ant], errors="coerce").fillna(0).astype(int) if col_ant else 0
    out["Total ciclo atual"] = pd.to_numeric(dfh[col_atual], errors="coerce").fillna(0).astype(int) if col_atual else 0
    out["Diferença"] = pd.to_numeric(dfh[col_dif], errors="coerce").fillna(0).astype(int) if col_dif else 0
    out["Inconsistência desconhecida?"] = dfh[col_nova].astype(str).str.upper().str.strip().replace({"": "NÃO"}) if col_nova else "NÃO"
    out["_chave_servico"] = out["Serviço"].apply(_servico_para_chave)
    return out


def tabela_top3_ciclo_atual(df_hist: pd.DataFrame) -> pd.DataFrame:
    base = _preparar_historico_full_ultimo_ciclo(df_hist)
    if base.empty:
        return pd.DataFrame(columns=["Serviço", "Descrição", "Quantidade no ciclo"])

    base = base[base["Diferença"] > 0].copy()
    if base.empty:
        return pd.DataFrame(columns=["Serviço", "Descrição", "Quantidade no ciclo"])

    base = base.sort_values(["Diferença", "Total ciclo atual"], ascending=[False, False]).head(5)
    return base[["Serviço", "Descrição", "Diferença"]].rename(columns={"Diferença": "Quantidade no ciclo"})


def tabela_inconsistencias_desconhecidas(df_hist: pd.DataFrame) -> pd.DataFrame:
    base = _preparar_historico_full_ultimo_ciclo(df_hist)
    if base.empty:
        return pd.DataFrame(columns=["Serviço", "Descrição", "Total ciclo atual", "Diferença"])

    novas = base[base["Inconsistência desconhecida?"].astype(str).str.upper().str.strip().eq("SIM")].copy()
    if novas.empty:
        return pd.DataFrame(columns=["Serviço", "Descrição", "Total ciclo atual", "Diferença"])

    novas = novas.sort_values(["Total ciclo atual", "Diferença"], ascending=[False, False])
    return novas[["Serviço", "Descrição", "Total ciclo atual", "Diferença"]]


def _preparar_historico_full_total(df_hist: pd.DataFrame) -> pd.DataFrame:
    """
    Consolida o histórico completo das inconsistências por serviço.

    Diferente do Top 3 e da tabela de desconhecidas, esta função NÃO filtra
    apenas o último ciclo. Ela usa todo o Historico_Criticas.xlsx, pega o
    registro mais recente de cada par Serviço + Descrição e mantém os totais
    daquele último registro.
    """
    dfh = _df_historico_full(df_hist)
    if dfh.empty:
        return pd.DataFrame()

    col_servico = _coluna_existente(dfh, ["Serviço", "Servico", "Tipo Serviço", "Tipo Servico"])
    col_inc = _coluna_existente(dfh, ["Inconsistência", "Inconsistencia", "Descrição Inconsistência", "Descricao Inconsistencia"])
    col_ant = _coluna_existente(dfh, ["Total ciclo anterior", "Total Ciclo Anterior"])
    col_atual = _coluna_existente(dfh, ["Total ciclo atual", "Total Ciclo Atual"])
    col_dif = _coluna_existente(dfh, ["Diferença ciclo", "Diferenca ciclo", "Diferença", "Diferenca"])
    col_data = _coluna_existente(dfh, ["Data/Hora", "Data Hora", "Data"])

    if not col_servico or not col_inc:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["Serviço"] = dfh[col_servico].astype(str)
    out["Descrição"] = dfh[col_inc].astype(str)
    out["Descrição agrupada"] = out["Descrição"].apply(normalizar_inconsistencia)
    out["Total anterior"] = pd.to_numeric(dfh[col_ant], errors="coerce").fillna(0).astype(int) if col_ant else 0
    out["Total ciclo atual"] = pd.to_numeric(dfh[col_atual], errors="coerce").fillna(0).astype(int) if col_atual else 0
    out["Diferença"] = pd.to_numeric(dfh[col_dif], errors="coerce").fillna(0).astype(int) if col_dif else 0
    out["_chave_servico"] = out["Serviço"].apply(_servico_para_chave)
    out["_ordem"] = range(len(out))

    if col_data:
        out["_data_sort"] = pd.to_datetime(dfh[col_data], dayfirst=True, errors="coerce")
    else:
        out["_data_sort"] = pd.NaT

    out = out[out["_chave_servico"].notna()].copy()
    out = out[out["Descrição"].astype(str).str.strip() != ""].copy()
    if out.empty:
        return pd.DataFrame()

    # Mantém o registro mais recente de cada inconsistência por serviço.
    out = out.sort_values(["_data_sort", "_ordem"], ascending=[True, True])
    out = out.drop_duplicates(subset=["_chave_servico", "Descrição"], keep="last")
    
    out = (
        out.groupby(["_chave_servico", "Descrição agrupada"], as_index=False)
        .agg({
            "Serviço": "first",
            "Total anterior": "sum",
            "Total ciclo atual": "sum",
            "Diferença": "sum",
            "_data_sort": "max"
        })
    )
    
    out["Descrição"] = out["Descrição agrupada"]
    
    return out

def tabela_historico_servico(df_hist: pd.DataFrame, chave_servico: str) -> pd.DataFrame:
    base = _preparar_historico_full_total(df_hist)
    if base.empty:
        return pd.DataFrame(columns=["Descrição", "Total", "%"])

    if chave_servico == "Transferências":
        base = base[base["_chave_servico"].isin(["Transferência 2", "Transferência 3"])].copy()
    else:
        base = base[base["_chave_servico"] == chave_servico].copy()

    if base.empty:
        return pd.DataFrame(columns=["Descrição", "Total", "%"])

    base = (
        base.groupby("Descrição", as_index=False)
        .agg({
            "Total ciclo atual": "sum",
            "Diferença": "sum"
        })
    )
    
    # Remove inconsistências com total zerado
    base["Total ciclo atual"] = pd.to_numeric(
        base["Total ciclo atual"],
        errors="coerce"
    ).fillna(0).astype(int)
    
    base = base[base["Total ciclo atual"] > 0].copy()
    
    if base.empty:
        return pd.DataFrame(columns=["Descrição", "Total", "%"])
    
    base = base.sort_values(
        ["Total ciclo atual", "Diferença"],
        ascending=[False, False]
    )
    
    total_geral = base["Total ciclo atual"].sum()

    if total_geral > 0:
        base["%"] = (
            pd.to_numeric(base["Total ciclo atual"], errors="coerce").fillna(0)
            / total_geral
            * 100
        )

        base["%"] = (
            pd.to_numeric(base["Total ciclo atual"], errors="coerce").fillna(0)
            / total_geral
            * 100
        ).map(lambda x: f"{x:.2f}%".replace(".", ","))
    else:
        base["%"] = "0,00%"

    return base[["Descrição", "Total ciclo atual", "%"]].rename(columns={
        "Total ciclo atual": "Total"
    })


def tabela_historico_criticas_minuto(df_hist: pd.DataFrame):

    if df_hist is None or df_hist.empty:
        return pd.DataFrame()

    dfh = df_hist.copy()

    col_tipo = _coluna_existente(dfh, ["Tipo Histórico", "Tipo Historico"])
    col_servico = _coluna_existente(dfh, ["Serviço", "Servico"])

    def normalizar_txt(v):
        txt = str(v or "").upper().strip()
        txt = unicodedata.normalize("NFKD", txt)
        txt = "".join(c for c in txt if not unicodedata.combining(c))
        return txt

    filtro = pd.Series(False, index=dfh.index)

    if col_tipo:
        tipo = dfh[col_tipo].apply(normalizar_txt)
        filtro |= (
            tipo.str.contains("CRITICA", na=False)
            & tipo.str.contains("MINUTO", na=False)
        )

    if col_servico:
        servico = dfh[col_servico].apply(normalizar_txt)
        filtro |= (
            servico.str.contains("CRITICA", na=False)
            & servico.str.contains("MINUTO", na=False)
        )

    dfh = dfh[filtro].copy()

    if dfh.empty:
        return pd.DataFrame()

    col_data = _coluna_existente(dfh, ["Data/Hora", "Data Hora", "Data"])

    if col_data:
        dfh["_ordem"] = pd.to_datetime(
            dfh[col_data],
            dayfirst=True,
            errors="coerce"
        )
        dfh = dfh.sort_values("_ordem", ascending=False)
        dfh = dfh.drop(columns=["_ordem"], errors="ignore")

    return dfh.head(50)

# =========================
# LEITURA DOS ARQUIVOS
# =========================

try:
    caminho_monitor, meta_monitor = baixar_github_se_houver_alteracao(
        GITHUB_ARQ_MONITORAMENTO,
        ARQ_LOCAL_MONITORAMENTO,
        obrigatorio=True,
    )

    df = pd.read_excel(caminho_monitor, engine="openpyxl")
    df = normalizar_colunas(df)

    if "Horário" in df.columns:
        df["Horário"] = df["Horário"].astype(str).str.slice(0, 5)

    for col in df.columns:
        if col not in ("Horário", "Data/Hora", "Data"):
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .fillna(0)
                .astype(int)
            )

except Exception as e:
    st.error(f"Erro ao carregar o arquivo Monitoramento.xlsx do GitHub: {e}")
    st.stop()

if df.empty:
    st.error("A planilha principal está vazia.")
    st.stop()

# Arquivos complementares. Não bloqueiam o painel se ainda não existirem no GitHub.
df_media = carregar_media_padrao()

df_criticas = None
meta_criticas = {}
try:
    caminho_criticas, meta_criticas = baixar_github_se_houver_alteracao(
        GITHUB_ARQ_CRITICAS,
        ARQ_LOCAL_CRITICAS,
        obrigatorio=False,
    )
    if caminho_criticas and Path(caminho_criticas).exists():
        df_criticas = carregar_excel_generico(str(caminho_criticas), Path(caminho_criticas).stat().st_mtime)
except Exception as e:
    st.warning(f"Criticas.xlsx ainda não disponível ou não carregado: {e}")

df_hist = None
meta_historico = {}
try:
    caminho_hist, meta_historico = baixar_github_se_houver_alteracao(
        GITHUB_ARQ_HISTORICO,
        ARQ_LOCAL_HISTORICO,
        obrigatorio=False,
    )
    if caminho_hist and Path(caminho_hist).exists():
        df_hist = carregar_excel_generico(str(caminho_hist), Path(caminho_hist).stat().st_mtime)
except Exception as e:
    st.warning(f"Historico_Criticas.xlsx ainda não disponível ou não carregado: {e}")

# Evita horários duplicados após normalização do RPA.
df = agrupar_monitoramento_por_horario(df)


# =========================
# DADOS PRINCIPAIS
# =========================

ultima = df.iloc[-1]
hora_coleta = str(ultima["Horário"]) if "Horário" in df.columns else "-"
ultima_modificacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# Considera o robô de monitoramento OFF quando o arquivo principal
# fica mais de 15 minutos sem nova atualização no GitHub/cache.
ultima_sync_github = meta_monitor.get("downloaded_at", "") if isinstance(meta_monitor, dict) else ""
robo_monitoramento_online = True
try:
    if ultima_sync_github:
        dt_sync = datetime.strptime(ultima_sync_github, "%d/%m/%Y %H:%M:%S")
        minutos_sem_atualizacao = (datetime.now() - dt_sync).total_seconds() / 60
        if minutos_sem_atualizacao > 15:
            robo_monitoramento_online = False
except Exception:
    robo_monitoramento_online = False

fila_trf = int(ultima.get("Fila 2 e 3", 0))
sucesso_trf = int(ultima.get("Sucesso 2 e 3", 0))
incons_trf = int(ultima.get("Inconsistência 2 e 3", 0))
automatizado = int(ultima.get("Automatizado", 0))

fila_0km = int(ultima.get("Fila 0km", 0))
sucesso_0km = int(ultima.get("Sucesso 0km", 0))
incons_0km = int(ultima.get("Inconsistência 0km", 0))

total_sucesso = sucesso_trf + sucesso_0km
total_criticas_minuto = obter_total_criticas_minuto(df_criticas)

robos = status_robos(df, df_criticas, df_hist)


if pagina == "Monitoramento atual":
    # =========================
    # TÍTULO
    # =========================

    st.markdown(
        f"""
        <div class="hive-title">Monitoramento e-CRV</div>
        <div class="hive-subtitle">
            Última coleta: <b>{hora_coleta}</b>
        </div>
        """,
        unsafe_allow_html=True,
    )


    # =========================
    # CARDS
    # =========================

    cols = st.columns(5)

    with cols[0]:
        render_card(
            "Fila Transferências",
            fila_trf,
            cor_saude(fila_trf, media_coluna(df_media, "Fila 2 e 3", hora_coleta), "negativo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[1]:
        render_card(
            "Sucesso Transferências",
            sucesso_trf,
            cor_saude(sucesso_trf, media_coluna(df_media, "Sucesso 2 e 3", hora_coleta), "positivo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[2]:
        render_card(
            "Inconsistências Transferências",
            incons_trf,
            cor_saude(incons_trf, media_coluna(df_media, "Inconsistência 2 e 3", hora_coleta), "negativo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[3]:
        render_card(
            "Automatizado",
            automatizado,
            cor_saude(automatizado, media_coluna(df_media, "Automatizado", hora_coleta), "positivo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[4]:
        render_card(
            "Inconsistências Críticas",
            total_criticas_minuto,
            cor_criticas_minuto(total_criticas_minuto),
            "Total de críticas do minuto atual",
        )

    cols = st.columns(5)

    with cols[0]:
        render_card(
            "Fila 0KM",
            fila_0km,
            cor_saude(fila_0km, media_coluna(df_media, "Fila 0km", hora_coleta), "negativo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[1]:
        render_card(
            "Sucesso 0KM",
            sucesso_0km,
            cor_saude(sucesso_0km, media_coluna(df_media, "Sucesso 0km", hora_coleta), "positivo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[2]:
        render_card(
            "Inconsistências 0KM",
            incons_0km,
            cor_saude(incons_0km, media_coluna(df_media, "Inconsistência 0km", hora_coleta), "negativo"),
            f"Último registro: {hora_coleta}",
        )

    with cols[3]:
        render_card(
            "Total",
            total_sucesso,
            "#1A73E8",
            "Sucesso Transferências + Sucesso 0KM",
        )

    with cols[4]:
        render_robos_card(robos, robo_monitoramento_online)


    # =========================
    # GRÁFICOS
    # =========================

    c1, c2 = st.columns(2)

    with c1:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)

        st.plotly_chart(
            line_chart(
                df,
                ["Sucesso 2 e 3"],
                "Transferências - Sucesso",
            ),
            use_container_width=True,
        )

        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)

        st.plotly_chart(
            line_chart(
                df,
                ["Fila 2 e 3", "Inconsistência 2 e 3"],
                "Transferências - Fila e Inconsistências",
            ),
            use_container_width=True,
        )

        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)

        st.plotly_chart(
            line_chart(
                df,
                ["Sucesso 0km"],
                "0KM - Sucesso",
            ),
            use_container_width=True,
        )

        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)

        st.plotly_chart(
            line_chart(
                df,
                ["Fila 0km", "Inconsistência 0km"],
                "0KM - Fila e Inconsistências",
            ),
            use_container_width=True,
        )

        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)


    # =========================
    # TABELAS DE HISTÓRICO E INCONSISTÊNCIAS
    # =========================

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Histórico de Inconsistências Críticas</div>', unsafe_allow_html=True)

    if df_hist is None or df_hist.empty:
        st.info("Histórico de críticas ainda não disponível.")
    else:
        df_crit_minuto = tabela_historico_criticas_minuto(df_hist)

        if df_crit_minuto.empty:
            render_mensagem_tabela("Não há histórico de críticas do minuto para exibir.")
        else:
            # Remove valores NaN da tabela
            df_crit_minuto = df_crit_minuto.fillna("")

            # Remove colunas que não serão exibidas
            colunas_remover = [
                "Tipo Histórico",
                "Tipo Historico",
                "Serviço",
                "Servico",
                "Total ciclo anterior",
                "Total Ciclo Anterior",
                "Diferença ciclo",
                "Diferenca ciclo",
                "Inconsistência nova no ciclo?",
                "Inconsistencia nova no ciclo?",
            ]

            df_crit_minuto = df_crit_minuto.drop(
                columns=[c for c in colunas_remover if c in df_crit_minuto.columns],
                errors="ignore"
            )

            # Limpa a descrição da inconsistência, se existir
            #col_inc = _coluna_existente(
            #    df_crit_minuto,
            #    [
            #        "Inconsistência",
            #        "Inconsistencia",
            #        "Descrição Inconsistência",
            #        "Descricao Inconsistencia",
            #    ]
            #)

            #if col_inc:
             #   df_crit_minuto[col_inc] = df_crit_minuto[col_inc].apply(normalizar_inconsistencia)

    render_tabela_escura(df_crit_minuto)

    st.markdown("</div>", unsafe_allow_html=True)


    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Inconsistências - Top 5 - Ciclo atual (10 min)</div>', unsafe_allow_html=True)

    df_top3 = tabela_top3_ciclo_atual(df_hist)
    if df_top3.empty:
        render_mensagem_tabela("Não foram identificadas inconsistências com aumento no ciclo atual.")
    else:
        render_tabela_escura(df_top3)

    st.markdown("</div>", unsafe_allow_html=True)


    #st.markdown('<div class="panel">', unsafe_allow_html=True)
    #st.markdown('<div class="panel-title">Tabela de Inconsistências Novas</div>', unsafe_allow_html=True)

    #df_novas = tabela_inconsistencias_desconhecidas(df_hist)
    #if df_novas.empty:
    #    render_mensagem_tabela("Não foram encontradas inconsistências desconhecidas.")
    #else:
    #    render_tabela_escura(df_novas)
    #
    #st.markdown("</div>", unsafe_allow_html=True)


    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Histórico Total de Inconsistências por Serviço</div>', unsafe_allow_html=True)

    servicos_tabelas = [
        ("Primeiro Registro / 0KM", "0KM"),
        ("Transferências", "Transferências"),
    ]

    for titulo_servico, chave_servico in servicos_tabelas:
        st.markdown(
            f'<div style="font-size:13px; font-weight:700; color:#202124; margin:14px 0 8px 0;">{titulo_servico}</div>',
            unsafe_allow_html=True,
        )
        df_serv = tabela_historico_servico(df_hist, chave_servico)
        if df_serv.empty:
            render_mensagem_tabela("Sem histórico de inconsistências registrado para este serviço.")
        else:
            render_tabela_escura(df_serv)

    st.markdown("</div>", unsafe_allow_html=True)





elif pagina == "Histórico monitoramento":
    st.markdown(
        """
        <div class="hive-title">Histórico monitoramento</div>
        <div class="hive-subtitle">
            Selecione uma data para consultar a tabela, os gráficos e o histórico de inconsistências.
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        caminho_mon_hist, _ = baixar_github_se_houver_alteracao(
            GITHUB_ARQ_MONITORAMENTO_HIST,
            ARQ_LOCAL_MONITORAMENTO_HIST,
            obrigatorio=True,
        )

        df_mon_hist = carregar_csv_historico(
            str(caminho_mon_hist),
            Path(caminho_mon_hist).stat().st_mtime,
        )

        df_mon_hist = normalizar_coluna_data_para_date(df_mon_hist)

    except Exception as e:
        st.error(f"Erro ao carregar histórico mensal de monitoramento: {e}")
        st.stop()

    try:
        caminho_inc_hist, meta_inc_hist = baixar_github_se_houver_alteracao(
            GITHUB_ARQ_INCONSISTENCIAS_HIST,
            ARQ_LOCAL_INCONSISTENCIAS_HIST,
            obrigatorio=False,
        )

        if caminho_inc_hist and Path(caminho_inc_hist).exists():
            df_inc_hist = carregar_csv_historico(
                str(caminho_inc_hist),
                Path(caminho_inc_hist).stat().st_mtime,
            )
            df_inc_hist = normalizar_coluna_data_para_date(df_inc_hist)
        else:
            df_inc_hist = pd.DataFrame()

    except Exception as e:
        df_inc_hist = pd.DataFrame()
        st.warning(f"Histórico de inconsistências não carregado: {e}")

    if df_mon_hist.empty or "Data" not in df_mon_hist.columns:
        st.error("O histórico de monitoramento não possui registros com Data/Hora válida.")
        st.stop()

    datas_monitoramento = set(df_mon_hist["Data"].dropna().unique())

    if df_inc_hist is not None and not df_inc_hist.empty and "Data" in df_inc_hist.columns:
        datas_inconsistencias = set(df_inc_hist["Data"].dropna().unique())
    else:
        datas_inconsistencias = set()

    datas_disponiveis = sorted(
        datas_monitoramento.union(datas_inconsistencias),
        reverse=True,
    )

    if not datas_disponiveis:
        st.error("Nenhuma data disponível no histórico.")
        st.stop()

    data_selecionada = st.date_input(
        "Selecionar data",
        value=datas_disponiveis[0],
        min_value=min(datas_disponiveis),
        max_value=max(datas_disponiveis),
        format="DD/MM/YYYY",
    )

    data_ref = pd.to_datetime(data_selecionada, errors="coerce").date()

    df_dia = df_mon_hist[df_mon_hist["Data"] == data_ref].copy()
    df_dia = agrupar_monitoramento_por_horario(df_dia)

    if df_inc_hist is not None and not df_inc_hist.empty and "Data" in df_inc_hist.columns:
        df_hist_dia = df_inc_hist[df_inc_hist["Data"] == data_ref].copy()
    else:
        df_hist_dia = pd.DataFrame()

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Tabela de monitoramento da data selecionada</div>', unsafe_allow_html=True)

    if df_dia.empty:
        render_mensagem_tabela("Não há registros de monitoramento para a data selecionada.")
    else:
        colunas_exibir = [
            c for c in df_dia.columns
            if (
                c not in ("Data", "Data/Hora")
                and "Horário;Automatizado;Fila 2 e 3" not in str(c)
                and "Anexo" not in str(c)
                and "anexo" not in str(c)
            )
        ]

        df_tabela = df_dia[colunas_exibir].copy().fillna("")

        if not df_tabela.empty:
            render_tabela_escura(df_tabela, scroll=True)
        else:
            render_mensagem_tabela("Nenhuma coluna disponível para exibição.")

    st.markdown("</div>", unsafe_allow_html=True)

    if not df_dia.empty:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)
        st.plotly_chart(
            line_chart(
                df_dia,
                ["Sucesso 2 e 3", "Fila 2 e 3", "Inconsistência 2 e 3"],
                "Transferências - Histórico do dia",
            ),
            use_container_width=True,
        )
        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)
        st.plotly_chart(
            line_chart(
                df_dia,
                ["Sucesso 0km", "Fila 0km", "Inconsistência 0km"],
                "0KM - Histórico do dia",
            ),
            use_container_width=True,
        )
        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Histórico de críticas da data selecionada</div>', unsafe_allow_html=True)

    if df_hist_dia.empty:
        render_mensagem_tabela("Não há histórico de críticas para a data selecionada.")
    else:
        df_crit_minuto = tabela_historico_criticas_minuto(df_hist_dia)
        if df_crit_minuto.empty:
            render_mensagem_tabela("Não há críticas do minuto para a data selecionada.")
        else:
            colunas_ocultar_crit = [
                c for c in [
                    "Data",
                    "Tipo Histórico",
                    "Tipo Historico",
                    "Serviço",
                    "Servico",
                    "Total ciclo anterior",
                    "Diferença ciclo",
                    "Inconsistência nova no ciclo?",
                ]
                if c in df_crit_minuto.columns
            ]

            df_crit_minuto_exibir = (
                df_crit_minuto
                .drop(columns=colunas_ocultar_crit, errors="ignore")
                .fillna("")
            )

            render_tabela_escura(df_crit_minuto_exibir)

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Histórico total de inconsistências por serviço</div>', unsafe_allow_html=True)

    for titulo_servico, chave_servico in [
        ("Primeiro Registro / 0KM", "0KM"),
        ("Transferências", "Transferências"),
    ]:
        st.markdown(
            f'<div style="font-size:13px; font-weight:700; color:#202124; margin:14px 0 8px 0;">{titulo_servico}</div>',
            unsafe_allow_html=True,
        )

        if df_hist_dia.empty:
            render_mensagem_tabela("Sem histórico de inconsistências para esta data.")
        else:
            df_serv = tabela_historico_servico(df_hist_dia, chave_servico)
            if df_serv.empty:
                render_mensagem_tabela("Sem histórico de inconsistências para este serviço na data selecionada.")
            else:
                df_serv = df_serv.drop(columns=["Ciclo atual"], errors="ignore")
                render_tabela_escura(df_serv.fillna(""))

    st.markdown("</div>", unsafe_allow_html=True)

# =========================
# RODAPÉ / AUTO REFRESH
# =========================

st.markdown(
    f"""
    <div class="footer-note">
        Verificação automática a cada {INTERVALO_VERIFICACAO_SEGUNDOS}s.
        O dashboard recarrega os dados quando os arquivos Excel são modificados.
    </div>
    """,
    unsafe_allow_html=True,
)

time.sleep(INTERVALO_VERIFICACAO_SEGUNDOS)
st.rerun()
