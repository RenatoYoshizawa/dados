import os
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components


# =========================
# CONFIGURAÇÕES
# =========================

# Teste para Streamlit Cloud lendo somente o arquivo obrigatório direto do GitHub.
# Este link deve ser o RAW do arquivo no GitHub.
URL_MONITORAMENTO = "https://raw.githubusercontent.com/RenatoYoshizawa/dados/main/Monitoramento_05_05_26.xlsx"

INTERVALO_VERIFICACAO_SEGUNDOS = 10


# =========================
# STREAMLIT
# =========================

st.set_page_config(
    page_title="Monitoramento e-CRV",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =========================
# CSS
# =========================

CSS = """
<style>
:root {
    --bg: #061529;
    --panel: #0A2344;
    --panel2: #0E315F;
    --blue: #0B5ED7;
    --cyan: #33C7FF;
    --text: #EAF4FF;
    --muted: #9DB7D2;
    --green: #21D07A;
    --yellow: #F6C343;
    --red: #FF4D5E;
    --white: #FFFFFF;
    --border: rgba(91, 166, 255, 0.22);
}

.stApp {
    background: radial-gradient(circle at top left, #0B3D91 0%, #061529 28%, #030A14 100%);
    color: var(--text);
}

[data-testid="stHeader"] { background: rgba(0,0,0,0); }
[data-testid="stToolbar"] { display: none; }

.block-container {
    padding-top: 1.3rem;
    padding-bottom: 1rem;
    max-width: 100%;
}

.hive-title {
    font-size: 32px;
    font-weight: 900;
    letter-spacing: .5px;
    color: #EAF4FF;
    margin-bottom: 0px;
}

.hive-subtitle {
    color: #9DB7D2;
    font-size: 13px;
    margin-top: 2px;
    margin-bottom: 18px;
}

.kpi-card {
    background: linear-gradient(160deg, rgba(14,49,95,.96), rgba(7,25,49,.96));
    border: 1px solid var(--border);
    border-radius: 18px;
    padding: 16px 16px 14px 16px;
    min-height: 142px;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .28);
    margin-bottom: 20px;
    overflow: hidden;
}

.kpi-label {
    color: #9DB7D2;
    font-size: 12px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: .6px;
    min-height: 34px;
}

.kpi-value {
    font-size: 32px;
    font-weight: 900;
    line-height: 1.08;
    margin-top: 8px;
    white-space: nowrap;
}

.kpi-note {
    color: #9DB7D2;
    font-size: 11px;
    margin-top: 9px;
}

.panel {
    background: rgba(7, 25, 49, .92);
    border: 1px solid var(--border);
    border-radius: 18px;
    padding: 16px 18px 12px 18px;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .24);
    margin-bottom: 18px;
}

.panel-title {
    font-size: 15px;
    font-weight: 800;
    color: #EAF4FF;
    margin-bottom: 10px;
    letter-spacing: .3px;
}

.robot-line {
    display: flex;
    align-items: center;
    gap: 7px;
    margin-top: 8px;
    font-size: 13px;
    font-weight: 800;
    color: #EAF4FF;
}

.robot-dot {
    display: inline-block;
    width: 12px;
    height: 12px;
    min-width: 12px;
    border-radius: 50%;
}

.robot-dot.on {
    background: #21D07A;
    box-shadow: 0 0 10px rgba(33, 208, 122, .9);
}

.robot-dot.off {
    background: #FF4D5E;
    box-shadow: 0 0 10px rgba(255, 77, 94, .9);
}

.robot-status {
    font-size: 12px;
    font-weight: 900;
    min-width: 28px;
}

.robot-status.on {
    color: #21D07A;
}

.robot-status.off {
    color: #FF4D5E;
}

.footer-note {
    color: #9DB7D2;
    font-size: 12px;
    text-align: right;
}

div[data-testid="stDataFrame"] {
    border: 1px solid var(--border);
    border-radius: 14px;
    background: rgba(7, 25, 49, .96);
}

/* fundo geral */
div[data-testid="stDataFrame"] > div {
    background: rgba(7, 25, 49, .96) !important;
}

/* tabela */
div[data-testid="stDataFrame"] [role="table"] {
    background: rgba(7, 25, 49, .96) !important;
    color: #EAF4FF !important;
}

/* cabeçalho */
div[data-testid="stDataFrame"] [role="columnheader"] {
    background: #0E315F !important;
    color: #EAF4FF !important;
    font-weight: 800 !important;
    border-bottom: 1px solid rgba(91, 166, 255, 0.25) !important;
}

/* células */
div[data-testid="stDataFrame"] [role="gridcell"] {
    background: rgba(7, 25, 49, .96) !important;
    color: #EAF4FF !important;
    border-color: rgba(91, 166, 255, 0.10) !important;
}

/* linhas alternadas */
div[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] {
    background: rgba(10, 35, 68, .96) !important;
}

/* fundo interno real do dataframe */
div[data-testid="stDataFrame"] .glideDataEditor {
    background: rgba(7, 25, 49, .96) !important;
}

/* viewport */
div[data-testid="stDataFrame"] .dvn-scroller {
    background: rgba(7, 25, 49, .96) !important;
}

/* células */
div[data-testid="stDataFrame"] .gdg-cell {
    background: rgba(7, 25, 49, .96) !important;
    color: #EAF4FF !important;
}

/* header */
div[data-testid="stDataFrame"] .gdg-header {
    background: #0E315F !important;
    color: #EAF4FF !important;
}

/* linhas alternadas */
div[data-testid="stDataFrame"] .gdg-row:nth-child(even) .gdg-cell {
    background: rgba(10, 35, 68, .96) !important;
}

/* aumenta o respiro entre colunas de cards */
[data-testid="column"] {
    padding-left: 0.35rem;
    padding-right: 0.35rem;
}

.chart-scroll {
    overflow-x: auto;
    overflow-y: hidden;
    width: 100%;
    padding-bottom: 8px;
}

.chart-inner {
    min-width: 1600px;
}

</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


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

def carregar_media_padrao() -> pd.DataFrame:
    dados = [
        ["06:50", 12.5, 149.5, 5.5, 4.5, 8, 0, 1],
        ["07:00", 26.86, 187.29, 6.86, 7.29, 13.14, 0.43, 1.43],
        ["07:10", 23.36, 223.73, 6.36, 12.55, 17.36, 0.27, 1.27],
        ["07:20", 29.18, 294.18, 7.09, 13.09, 23, 0.55, 1.27],
        ["07:30", 38.27, 372, 9.36, 17, 29.64, 1.64, 1.27],
        ["07:40", 49.62, 460.62, 11.54, 23.15, 38.08, 0.85, 1.38],
        ["07:50", 60.15, 556.69, 12.92, 18.69, 47.38, 3.38, 1.92],
        ["08:00", 72, 646.21, 22.21, 31.07, 59.21, 4.79, 2.43],
        ["08:10", 88.64, 756.43, 28, 41.64, 72.36, 5.93, 2.93],
        ["08:20", 109.2, 867.27, 34.6, 61, 85.73, 9.6, 3.33],
        ["08:30", 134.62, 1021.94, 60.12, 83.31, 102.56, 37.12, 4.75],
        ["08:40", 165.94, 1193.69, 87, 101.38, 121.56, 20.44, 6.38],
        ["08:50", 199.44, 1394.25, 104.62, 98.69, 139.38, 19.62, 6.69],
        ["09:00", 238.69, 1569.94, 141.06, 134.25, 163.56, 25.75, 8.25],
        ["09:10", 297.75, 1795, 168.56, 125.75, 190.81, 10.06, 8.62],
        ["09:20", 363.93, 2027.6, 182.8, 143.07, 215.13, 15.33, 9.07],
        ["09:30", 434.93, 2231.79, 224, 169.13, 226.2, 38.47, 10.4],
        ["09:40", 488.73, 2423.93, 251.2, 197.6, 270, 20.53, 13],
        ["09:50", 550.2, 2623.13, 267.33, 246.67, 294.13, 27.73, 13.07],
        ["10:00", 609.47, 2784.67, 314.53, 316.4, 320.47, 45.13, 13.47],
        ["10:10", 695, 3008.93, 324.36, 351.07, 340.07, 23.07, 15.07],
        ["10:20", 778.27, 3247.2, 366.07, 361.47, 377.4, 36.6, 14.67],
        ["10:30", 880.4, 3567.71, 420.27, 378.27, 399.33, 44.93, 17.27],
        ["10:40", 974.79, 3773.43, 465.21, 366.14, 441.21, 24.93, 18.64],
        ["10:50", 1052.29, 4010.79, 470.77, 370.36, 474.21, 23.79, 19.21],
        ["11:00", 1112.71, 4181.86, 546.93, 428.21, 500.07, 48.43, 20.36],
        ["11:10", 1175, 4427.38, 589.46, 431.69, 529.85, 15, 22.38],
        ["11:20", 1254.46, 4663, 604.62, 424.69, 562.15, 30.85, 22.31],
        ["11:30", 1302.92, 4846.46, 643.31, 440.15, 604.38, 56.31, 23.62],
        ["11:40", 1379.29, 5095.21, 694.43, 436.86, 642.86, 27.79, 23.86],
        ["11:50", 1484.79, 5339.86, 710.21, 471.43, 681.64, 33.64, 21.07],
        ["12:00", 1586.07, 5518.6, 777.67, 459.73, 704.67, 58.33, 23.6],
        ["12:10", 1686.07, 5762.6, 828.2, 408.53, 737.47, 16.07, 24.07],
        ["12:20", 1776, 5990.8, 843.8, 373.13, 776.13, 23.33, 24.87],
        ["12:30", 1847.81, 6153.88, 890.06, 384.56, 800.94, 54.06, 25.88],
        ["12:40", 1956.56, 6351.25, 934.81, 374.81, 834.19, 17.06, 26.81],
        ["12:50", 2044, 6531.44, 945.38, 365.56, 863.69, 27, 26.81],
        ["13:00", 2107.44, 6611.27, 969.6, 429.94, 826.38, 51.12, 26.31],
        ["13:10", 2156, 6796.33, 1012.07, 443.6, 907.4, 21.47, 29.13],
        ["13:20", 2249.94, 7034.31, 1046.06, 398.5, 942.5, 23.19, 29.12],
        ["13:30", 2329.38, 7225.44, 1093.44, 370.5, 979.12, 40.06, 29.94],
        ["13:40", 2396.75, 7417.25, 1128, 360.12, 1009.31, 17.19, 30.81],
        ["13:50", 2485.94, 7620.62, 1142.44, 350.12, 1033.62, 29.31, 31.69],
        ["14:00", 2540, 7639.38, 1165.79, 430.93, 1042.57, 53.5, 32.86],
        ["14:10", 2790.93, 8092.43, 1207.79, 273.64, 1079.64, 30.93, 33.5],
        ["14:20", 2864.4, 8302.67, 1251.73, 285.67, 1121.33, 37, 34],
        ["14:30", 2958.8, 8506.8, 1299.6, 290.2, 1150.4, 65.13, 35.13],
        ["14:40", 3047.93, 8726.2, 1347.07, 318.67, 1180.2, 29.67, 35.8],
        ["14:50", 3062, 8824.81, 1355.62, 399.69, 1199.62, 35.81, 37.12],
        ["15:00", 3142.56, 9028.62, 1414.31, 436.25, 1228, 52.38, 38.69],
        ["15:10", 3224.43, 9205.29, 1444.79, 442.64, 1240.57, 17.71, 41.29],
        ["15:20", 3296.07, 9337.77, 1487.93, 426.29, 1282.36, 27.07, 42.86],
        ["15:30", 3442.81, 9807.94, 1550.88, 391.94, 1325.19, 52.12, 43.75],
        ["15:40", 3558.88, 10072.33, 1603.12, 394.88, 1355.69, 16.81, 45.62],
        ["15:50", 3676.62, 10334, 1629.19, 384.12, 1380.88, 34.5, 47.12],
        ["16:00", 3808, 10578, 1698.67, 420.93, 1415.13, 65.8, 49.2],
        ["16:10", 3926, 10783, 1778.47, 407.69, 1354.5, 26.12, 47.94],
        ["16:20", 4055.2, 11148.87, 1870.87, 428.07, 1474.2, 212.53, 52.8],
        ["16:30", 4236.8, 11451.93, 2003.13, 379.27, 1504.6, 558.4, 54.27],
        ["16:40", 4381.4, 11725.53, 2144.8, 346, 1535, 448.2, 56.87],
        ["16:50", 4556.25, 11989.94, 2273.94, 295.19, 1551.75, 395.75, 60.25],
        ["17:00", 4684.06, 12175.14, 2384.88, 286.44, 1573.75, 444.62, 63.44],
        ["17:10", 4844.5, 12417.44, 2509.81, 242.81, 1596.5, 344.81, 66.19],
        ["17:20", 4980.38, 12640.2, 2638, 211.44, 1617.56, 285.19, 69.38],
        ["17:30", 5108.12, 12803.19, 2765.81, 183.62, 1637.88, 316.12, 73.56],
        ["17:40", 5226.06, 12968.81, 2882.19, 171, 1654.25, 218.88, 76.75],
        ["17:50", 5344.06, 13112.81, 2975.38, 131.06, 1665.75, 187.62, 79.12],
        ["18:00", 5423.5, 13226.56, 3070.69, 115.69, 1675.81, 201.94, 82.69],
        ["18:10", 5487.12, 13340.81, 3169, 84.62, 1687.12, 115.56, 84.69],
        ["18:20", 5545.86, 13541.07, 3202.93, 60.79, 1686.29, 83.43, 85.07],
        ["18:30", 5591.87, 13579.07, 3300.2, 32, 1710.6, 72.13, 86.13],
        ["18:40", 5583.93, 13556.53, 3333.4, 16.8, 1699.8, 34.87, 88.27],
        ["18:50", 5641.06, 13645.93, 3354.88, 10.25, 1712.25, 24.75, 87.75],
        ["19:00", 5661.6, 13727.27, 3389.33, 11.27, 1730.4, 40.87, 90.6],
        ["19:10", 5677.47, 13755.33, 3411.07, 6.2, 1732.6, 25.6, 91.07],
        ["19:20", 5690.2, 13785.13, 3426.67, 6.73, 1734.8, 17.33, 90.87],
        ["19:30", 5730.36, 13865.64, 3452.36, 5.57, 1736.43, 12.07, 88.64],
        ["19:40", 5767.08, 13946.25, 3427.33, 8.25, 1747, 5.67, 87.17],
        ["19:50", 5760.38, 13918.69, 3409.15, 6.23, 1731.92, 7.08, 88.54],
        ["20:00", 5771.38, 13938.92, 3420.31, 10.08, 1734.15, 8.15, 88.69],
        ["20:10", 5784.46, 13967.77, 3424.38, 6.31, 1735.77, 1.92, 88.69],
        ["20:20", 5799.23, 13995.85, 3428.15, 4.23, 1737.23, 10.69, 88.77],
        ["20:30", 5809, 14045, 3426.75, 5.33, 1740.5, 1.42, 91.33],
        ["20:40", 5818.83, 14066.67, 3428.08, 4.92, 1742.25, 1.17, 91.33],
        ["20:50", 5833.25, 14085, 3431.33, 3.5, 1743.08, 4.42, 91.42],
        ["21:00", 5842, 14098.25, 3440.25, 5.25, 1744, 1.5, 91.42],
        ["21:10", 5884.45, 14184.64, 3488.09, 2.91, 1763.82, 0.55, 91.73],
        ["21:20", 5866.17, 14129.17, 3443.08, 1.25, 1746.33, 1.5, 91.5],
        ["21:30", 5878.75, 14234.82, 3435.09, 3.75, 1577.75, 0.33, 86],
        ["21:40", 5890.36, 14248.36, 3435.91, 3.18, 1722.73, 0.27, 93.82],
        ["21:50", 5886.88, 14326.25, 3451.62, 3.75, 1735.75, 0.12, 91.88],
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
        if col != "Horário":
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


@st.cache_data(show_spinner=False)
def carregar_excel_generico(path_str: str, mtime: float):
    df = pd.read_excel(path_str, engine="openpyxl")
    df = normalizar_colunas(df)
    return df


def media_coluna(df_media: pd.DataFrame, coluna: str):
    if df_media is None or df_media.empty or coluna not in df_media.columns:
        return None

    s = pd.to_numeric(df_media[coluna], errors="coerce").dropna()
    if s.empty:
        return None

    return float(s.mean())


def cor_saude(valor: int, media, tipo: str):
    """
    tipo='positivo': sucesso e automatizado
        verde: >= média
        amarelo: até 20% abaixo
        vermelho: abaixo de 20%

    tipo='negativo': fila e inconsistência
        verde: <= média
        amarelo: até 20% acima
        vermelho: acima de 20%
    """
    if media is None or media <= 0:
        return "#FFFFFF"

    valor = float(valor)

    if tipo == "positivo":
        if valor >= media:
            return "#21D07A"
        elif valor >= media * 0.80:
            return "#F6C343"
        else:
            return "#FF4D5E"

    if tipo == "negativo":
        if valor <= media:
            return "#21D07A"
        elif valor <= media * 1.20:
            return "#F6C343"
        else:
            return "#FF4D5E"

    return "#FFFFFF"

def cor_criticas_minuto(valor):
    valor = int(valor or 0)

    if valor == 0:
        return "#21D07A"  # verde
    elif 1 <= valor <= 5:
        return "#F6C343"  # amarelo
    else:
        return "#FF4D5E"  # vermelho

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


def stop_ativado(df_criticas: pd.DataFrame) -> bool:
    if df_criticas is None or df_criticas.empty:
        return False

    if "STOP PROCESSO ativado?" not in df_criticas.columns:
        return False

    valores = df_criticas["STOP PROCESSO ativado?"].astype(str).str.upper().str.strip()
    return valores.eq("SIM").any()


def servicos_desligados(df_criticas: pd.DataFrame) -> str:
    if df_criticas is None or df_criticas.empty:
        return ""

    if "Serviços desligados" not in df_criticas.columns:
        return ""

    s = df_criticas["Serviços desligados"].dropna().astype(str)
    s = s[s.str.strip() != ""]
    if s.empty:
        return ""

    return s.iloc[-1]


def houve_aumento(df: pd.DataFrame, coluna: str) -> bool:
    if df is None or len(df) < 2 or coluna not in df.columns:
        return False

    atual = int(pd.to_numeric(df.iloc[-1][coluna], errors="coerce") or 0)
    anterior = int(pd.to_numeric(df.iloc[-2][coluna], errors="coerce") or 0)

    return atual > anterior


def status_robos(df_monitor: pd.DataFrame, df_criticas: pd.DataFrame):
    """
    Regra:
    - Se Criticas.xlsx indicar STOP PROCESSO = SIM, os serviços desligados ficam OFF.
    - Para voltar ON, verifica se houve aumento de sucesso ou inconsistência no serviço.
    - Como o monitoramento consolida Transferência 2 e 3, as duas usam a mesma variação.
    """
    status = {
        "0KM": "ON",
        "Transferência 2": "ON",
        "Transferência 3": "ON",
    }

    if stop_ativado(df_criticas):
        desligados_txt = servicos_desligados(df_criticas).lower()

        if "1" in desligados_txt or "primeiro" in desligados_txt or "0km" in desligados_txt:
            status["0KM"] = "OFF"

        if "2" in desligados_txt or "propriet" in desligados_txt:
            status["Transferência 2"] = "OFF"

        if "3" in desligados_txt or "munic" in desligados_txt or "estado" in desligados_txt:
            status["Transferência 3"] = "OFF"

    aumento_0km = (
        houve_aumento(df_monitor, "Sucesso 0km")
        or houve_aumento(df_monitor, "Inconsistência 0km")
    )

    aumento_trf = (
        houve_aumento(df_monitor, "Sucesso 2 e 3")
        or houve_aumento(df_monitor, "Inconsistência 2 e 3")
    )

    if aumento_0km:
        status["0KM"] = "ON"

    if aumento_trf:
        status["Transferência 2"] = "ON"
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


def render_robos_card(status_dict):
    status_0km = status_dict.get("0KM", "ON")
    status_t2 = status_dict.get("Transferência 2", "ON")
    status_t3 = status_dict.get("Transferência 3", "ON")

    def bolinha(status):
        return "🟢" if status == "ON" else "🔴"

    html = f"""
    <div class="kpi-card">
        <div class="kpi-label">Robôs</div>
        <div style="margin-top:10px; font-size:14px; font-weight:800;">
            {bolinha(status_0km)} <span style="color:{'#21D07A' if status_0km == 'ON' else '#FF4D5E'};">{status_0km}</span> 0KM<br><br>
            {bolinha(status_t2)} <span style="color:{'#21D07A' if status_t2 == 'ON' else '#FF4D5E'};">{status_t2}</span> Transferência 2<br><br>
            {bolinha(status_t3)} <span style="color:{'#21D07A' if status_t3 == 'ON' else '#FF4D5E'};">{status_t3}</span> Transferência 3
        </div>
        <div class="kpi-note">Status estimado pelo STOP e movimentação do fluxo</div>
    </div>
    """

    st.markdown(html, unsafe_allow_html=True)


def fig_layout(fig, height=360):
    fig.update_layout(
        height=height,
        width=1600,
        margin=dict(l=22, r=22, t=88, b=35),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#EAF4FF", family="Arial"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.06,
            xanchor="left",
            x=0,
            font=dict(size=11, color="#9DB7D2"),
            bgcolor="rgba(0,0,0,0)",
        ),

        xaxis=dict(
            gridcolor="rgba(157,183,210,.12)",
            zerolinecolor="rgba(157,183,210,.12)",
            tickfont=dict(color="#9DB7D2"),
            tickangle=-45,

            # 🔥 ESSENCIAL → ativa slider
            rangeslider=dict(visible=True),

            # 🔥 ESSENCIAL → define janela inicial (últimos pontos)
            #range=[-30, 0],  # mostra só últimos 30 pontos
        ),

        yaxis=dict(
            gridcolor="rgba(157,183,210,.12)",
            zerolinecolor="rgba(157,183,210,.12)",
            tickfont=dict(color="#9DB7D2"),
        ),
    )
    return fig


def line_chart(df, cols, title):
    fig = go.Figure()
    colors = ["#F6C343", "#21D07A", "#FF4D5E", "#33C7FF", "#8AB4FF"]

    # 🔥 CONVERTE eixo X para índice numérico
    x_labels = df["Horário"].astype(str).tolist() if "Horário" in df.columns else [str(i) for i in df.index]
    x_vals = list(range(len(x_labels)))

    for i, col in enumerate(cols):
        if col in df.columns:
            y_vals = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=y_vals,
                    mode="lines+markers",
                    name=col,
                    line=dict(width=3, color=colors[i % len(colors)]),
                    marker=dict(size=7),

                    customdata=x_labels,
                    hovertemplate="%{customdata}<br>" + col + ": %{y}<extra></extra>",
                )
            )

    # 🔥 REDUZ LABELS (evita poluição)
    passo = max(1, len(x_vals) // 20)
    tickvals = x_vals[::passo]
    ticktext = x_labels[::passo]

    # 🔥 DEFINE JANELA INICIAL (últimos pontos)
    janela = 25
    inicio = max(0, len(x_vals) - janela)
    fim = max(1, len(x_vals) - 1)

    fig.update_xaxes(
        tickmode="array",
        tickvals=tickvals,
        ticktext=ticktext,
        tickangle=-45,
        range=[inicio, fim],  # ✅ AGORA FUNCIONA
        rangeslider=dict(visible=True),
    )

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=16, color="#EAF4FF"),
            y=0.98,
            x=0.01,
            xanchor="left",
        )
    )

    return fig_layout(fig)

def render_tabela_escura(df_tabela: pd.DataFrame):
    html = df_tabela.to_html(index=False, escape=False)

    html = html.replace(
        '<table border="1" class="dataframe">',
        '<table class="dark-table">'
    )

    st.markdown(
        """
        <style>
        .dark-table {
            width: 100%;
            border-collapse: collapse;
            background: rgba(7, 25, 49, .96);
            color: #EAF4FF;
            font-size: 12px;
        }

        .dark-table thead th {
            background: #0E315F;
            color: #EAF4FF;
            text-align: left;
            padding: 9px;
            font-weight: 800;
            border-bottom: 1px solid rgba(91, 166, 255, 0.25);
        }

        .dark-table tbody td {
            padding: 8px;
            color: #EAF4FF;
            border-bottom: 1px solid rgba(91, 166, 255, 0.10);
            max-width: 520px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .dark-table tbody tr:nth-child(even) {
            background: rgba(10, 35, 68, .96);
        }

        .dark-table tbody tr:nth-child(odd) {
            background: rgba(7, 25, 49, .96);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(html, unsafe_allow_html=True)

# =========================
# LEITURA DOS ARQUIVOS
# =========================

try:
    df = pd.read_excel(URL_MONITORAMENTO, engine="openpyxl")
    df = normalizar_colunas(df)

    if "Horário" in df.columns:
        df["Horário"] = df["Horário"].astype(str).str.slice(0, 5)

    for col in df.columns:
        if col != "Horário":
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .fillna(0)
                .astype(int)
            )

except Exception as e:
    st.error(f"Erro ao carregar o arquivo do GitHub: {e}")
    st.stop()

if df.empty:
    st.error("A planilha principal está vazia.")
    st.stop()

# Neste teste, somente o Monitoramento é obrigatório.
# A média usa a tabela padrão interna. Críticas e histórico ficam indisponíveis.
df_media = carregar_media_padrao()
df_criticas = None
df_hist = None


# =========================
# DADOS PRINCIPAIS
# =========================

ultima = df.iloc[-1]
hora_coleta = str(ultima["Horário"]) if "Horário" in df.columns else "-"
ultima_modificacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

fila_trf = int(ultima.get("Fila 2 e 3", 0))
sucesso_trf = int(ultima.get("Sucesso 2 e 3", 0))
incons_trf = int(ultima.get("Inconsistência 2 e 3", 0))
automatizado = int(ultima.get("Automatizado", 0))

fila_0km = int(ultima.get("Fila 0km", 0))
sucesso_0km = int(ultima.get("Sucesso 0km", 0))
incons_0km = int(ultima.get("Inconsistência 0km", 0))

total_sucesso = sucesso_trf + sucesso_0km
total_criticas_minuto = obter_total_criticas_minuto(df_criticas)

robos = status_robos(df, df_criticas)


# =========================
# TÍTULO
# =========================

st.markdown(
    f"""
    <div class="hive-title">Monitoramento e-CRV</div>
    <div class="hive-subtitle">
        Última coleta: <b>{hora_coleta}</b> • Arquivo modificado em: <b>{ultima_modificacao}</b>
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
        cor_saude(fila_trf, media_coluna(df_media, "Fila 2 e 3"), "negativo"),
        f"Último registro: {hora_coleta}",
    )

with cols[1]:
    render_card(
        "Sucesso Transferências",
        sucesso_trf,
        cor_saude(sucesso_trf, media_coluna(df_media, "Sucesso 2 e 3"), "positivo"),
        f"Último registro: {hora_coleta}",
    )

with cols[2]:
    render_card(
        "Inconsistências Transferências",
        incons_trf,
        cor_saude(incons_trf, media_coluna(df_media, "Inconsistência 2 e 3"), "negativo"),
        f"Último registro: {hora_coleta}",
    )

with cols[3]:
    render_card(
        "Automatizado",
        automatizado,
        cor_saude(automatizado, media_coluna(df_media, "Automatizado"), "positivo"),
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
        cor_saude(fila_0km, media_coluna(df_media, "Fila 0km"), "negativo"),
        f"Último registro: {hora_coleta}",
    )

with cols[1]:
    render_card(
        "Sucesso 0KM",
        sucesso_0km,
        cor_saude(sucesso_0km, media_coluna(df_media, "Sucesso 0km"), "positivo"),
        f"Último registro: {hora_coleta}",
    )

with cols[2]:
    render_card(
        "Inconsistências 0KM",
        incons_0km,
        cor_saude(incons_0km, media_coluna(df_media, "Inconsistência 0km"), "negativo"),
        f"Último registro: {hora_coleta}",
    )

with cols[3]:
    render_card(
        "Total",
        total_sucesso,
        "#8AB4FF",
        "Sucesso Transferências + Sucesso 0KM",
    )

with cols[4]:
    render_robos_card(robos)


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
            ["Fila 2 e 3", "Sucesso 2 e 3", "Inconsistência 2 e 3"],
            "Comparativo de Transferências",
        ),
        use_container_width=False,
    )

    st.markdown('</div></div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

with c2:
    st.markdown('<div class="panel">', unsafe_allow_html=True)

    st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)

    st.plotly_chart(
        line_chart(
            df,
            ["Fila 0km", "Sucesso 0km", "Inconsistência 0km"],
            "Comparativo de 0KM",
        ),
        use_container_width=False,
    )

    st.markdown('</div></div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# =========================
# TABELA HISTÓRICO DE CRÍTICAS
# =========================

st.markdown('<div class="panel">', unsafe_allow_html=True)
st.markdown('<div class="panel-title">Histórico de Inconsistências Críticas</div>', unsafe_allow_html=True)

if df_hist is None or df_hist.empty:
    st.info("Histórico de críticas ainda não disponível.")
else:
    render_tabela_escura(
        df_hist.tail(20).sort_index(ascending=False)
    )

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