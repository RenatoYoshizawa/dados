import os
import time
import base64
import json
import re
import unicodedata
import html
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components


# =========================
# DATA/HORA PADRÃO DO PAINEL
# =========================

FUSO_HORARIO_PADRAO = "America/Sao_Paulo"

def agora_sao_paulo() -> pd.Timestamp:
    """
    Retorna a data/hora atual no fuso de São Paulo, sem timezone.

    O RPA grava as planilhas em horário local de São Paulo.
    No Streamlit Cloud/servidores Linux, datetime.now() e pd.Timestamp.now()
    podem usar UTC, causando filtros do dia e janelas de tempo incorretas
    principalmente à noite.
    """
    return pd.Timestamp.now(tz=FUSO_HORARIO_PADRAO).tz_localize(None)

def datetime_sao_paulo() -> datetime:
    """Retorna datetime nativo no fuso de São Paulo, sem timezone."""
    return agora_sao_paulo().to_pydatetime()


# =========================#
# CONFIGURAÇÕES
# =========================

# Repositório GitHub utilizado pelo RPA de monitoramento.
# Para repositório privado, cadastre TOKEN_GITHUB em: Streamlit Cloud > Settings > Secrets.
GITHUB_OWNER = "RenatoYoshizawa"
GITHUB_REPO = "Monitoramento"
GITHUB_BRANCH = "main"

WEBHOOK_TEAMS = st.secrets.get(
    "WEBHOOK_TEAMS",
    ""
)

GITHUB_ARQ_MONITORAMENTO = "monitoramento/Monitoramento.xlsx"
GITHUB_ARQ_CRITICAS = "monitoramento/Criticas.xlsx"
GITHUB_ARQ_HISTORICO = "monitoramento/Historico_Criticas.xlsx"
GITHUB_ARQ_LOG_DIA = f"logs/Log_{agora_sao_paulo().strftime('%Y_%m_%d')}.csv"

# Arquivos usados pelo controle manual de robôs.
# controle_ecrv.json: comando PENDENTE gravado pelo dashboard para o RPA processar.
# status_ecrv.json: retorno/estado confirmado pelo RPA para o dashboard reconhecer.
GITHUB_ARQ_CONTROLE_ECRV = "comandos/controle_ecrv.json"
GITHUB_ARQ_STATUS_ECRV = "status/status_ecrv.json"

# Quantidade de meses que serão exibidos na página de Histórico.
# Exemplo: 6 = mês atual + 5 meses anteriores.
QUANTIDADE_MESES_HISTORICO = 6

def caminhos_historico_ultimos_meses(prefixo: str, quantidade_meses: int = QUANTIDADE_MESES_HISTORICO):
    """
    Monta os caminhos dos arquivos mensais de histórico dos últimos meses.

    Exemplo, em junho/2026 e quantidade_meses=6:
    historico/Monitoramento_2026_01.csv até historico/Monitoramento_2026_06.csv.
    """

    hoje = agora_sao_paulo().replace(day=1)
    caminhos = []

    for i in range(quantidade_meses - 1, -1, -1):
        mes_ref = hoje - pd.DateOffset(months=i)
        ano_mes = mes_ref.strftime("%Y_%m")
        caminhos.append(f"historico/{prefixo}_{ano_mes}.csv")

    return caminhos

CACHE_DIR = Path("/tmp/monitoramento_ecrv_cache")
CACHE_DIR.mkdir(parents=True, exist_ok=True)

ARQ_LOCAL_MONITORAMENTO = CACHE_DIR / "Monitoramento.xlsx"
ARQ_LOCAL_CRITICAS = CACHE_DIR / "Criticas.xlsx"
ARQ_LOCAL_HISTORICO = CACHE_DIR / "Historico_Criticas.xlsx"
ARQ_LOCAL_MONITORAMENTO_HIST = CACHE_DIR / "Monitoramento_Historico.csv"
ARQ_LOCAL_INCONSISTENCIAS_HIST = CACHE_DIR / "Inconsistencias_Historico.csv"
ARQ_LOCAL_LOG_DIA = CACHE_DIR / "Log_Dia.csv"
META_LOCAL = CACHE_DIR / "github_meta.json"
ARQ_ALERTA_ECRV = CACHE_DIR / "alerta_ecrv_off.json"
ARQ_LOCAL_CONTROLE_ECRV = CACHE_DIR / "controle_ecrv.json"
ARQ_LOCAL_STATUS_ECRV = CACHE_DIR / "status_ecrv.json"


def caminho_github_log_data(data_ref=None) -> str:
    """Monta o caminho do CSV de log diário no GitHub para a data informada."""
    ref = pd.Timestamp(data_ref if data_ref is not None else agora_sao_paulo())
    return f"logs/Log_{ref.strftime('%Y_%m_%d')}.csv"


def caminho_local_log_data(data_ref=None) -> Path:
    """Monta o caminho local de cache do CSV de log diário para a data informada."""
    ref = pd.Timestamp(data_ref if data_ref is not None else agora_sao_paulo())
    return CACHE_DIR / f"Log_{ref.strftime('%Y_%m_%d')}.csv"

INTERVALO_VERIFICACAO_SEGUNDOS = 30
TEMPO_MINIMO_OFF_MINUTOS = 15
JANELA_STOP_MINUTOS = 420  # considera STOPs dos últimos 420 minutos
JANELA_STATUS_RPA_MINUTOS = 60  # usa confirmação do RPA somente se for recente


# =========================
# PARÂMETROS DOS CARDS DINÂMICOS
# =========================

COR_VERDE = "#188038"
COR_AMARELO = "#F9AB00"
COR_VERMELHO = "#D93025"
COR_AZUL = "#1A73E8"

# Taxas máximas de inconsistência nos últimos 60 minutos.
LIMITES_INCONSISTENCIA = {
    "Transferências": {
        "verde": 12.0,
        "amarelo": 14.0,
        "amostra_minima": 200,
    },
    "0KM": {
        "verde": 3.0,
        "amarelo": 5.0,
        "amostra_minima": 50,
    },
    "TDV": {
        "verde": 8.0,
        "amarelo": 12.0,
        "amostra_minima": 30,
    },
}

# Fila: tempo necessário para absorção, considerando a produção dos últimos 60 minutos.
LIMITE_FILA_VERDE_MIN = 15.0
LIMITE_FILA_AMARELO_MIN = 30.0

# Sucesso: só avalia queda de desempenho quando a fila atual representa
# pelo menos 15 minutos de processamento no ritmo da última hora.
LIMITE_DEMANDA_SUCESSO_MIN = 15.0

# Tolerâncias para não alternar setas por pequenas oscilações.
TOLERANCIA_TENDENCIA_SUCESSO_PCT = 5.0
TOLERANCIA_TENDENCIA_FILA_MIN = 5.0
TOLERANCIA_TENDENCIA_INCONS_PP = 0.5


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
    text-align: center;
    width: 100%;
}

.hive-subtitle {
    color: var(--md-muted);
    font-size: 13px;
    font-weight: 400;
    margin-top: 2px;
    margin-bottom: 18px;
    text-align: center;
    width: 100%;
}

.kpi-card {
    background: var(--md-surface);
    border: none;
    border-radius: 24px;
    padding: 20px;
    min-height: 148px;
    box-shadow: var(--md-shadow);
    margin-bottom: 18px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    text-align: center;
}

.kpi-card.kpi-tall {
    min-height: 231px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
}

.kpi-card.kpi-tall .kpi-label {
    font-size: 20px;
}

.kpi-card.kpi-tall .kpi-value {
    font-size: 40px;
}

.kpi-card.kpi-tall .kpi-note {
    font-size: 13px;
}

.kpi-card.kpi-robos-tall {
    min-height: 231px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    text-align: center;
}

.kpi-card.kpi-robos-tall .kpi-label {
    font-size: 20px;
}

.kpi-card.kpi-robos-tall .kpi-value {
    font-size: 40px;
}

.kpi-card.kpi-robos-tall .kpi-note {
    font-size: 13px;
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



/* =========================
   MENU / LOGS DO ROBÔ
========================= */

.menu-divider {
    opacity: 0;
    height: 1px;
    margin: 10px 18px;
    background: #E0E3EB;
    transition: opacity 0.2s ease;
}

.hover-menu:hover .menu-divider {
    opacity: 1;
}

.menu-item-theme {
    margin-top: 4px;
    color: #5F6368 !important;
}

.log-toolbar {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
    align-items: stretch;
    margin-bottom: 16px;
}

.log-meta-card {
    background: var(--md-surface-container);
    border: 1px solid var(--md-border);
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: var(--md-shadow);
    min-width: 0;
}

.log-meta-label {
    color: var(--md-muted);
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-bottom: 5px;
}

.log-meta-value {
    color: var(--md-text);
    font-size: 15px;
    line-height: 1.35;
    font-weight: 800;
    overflow-wrap: anywhere;
}

.log-filter-box {
    background: var(--md-surface-container);
    border: 1px solid var(--md-border);
    border-radius: 18px;
    padding: 14px 16px 4px 16px;
    margin: 8px 0 14px 0;
    box-shadow: var(--md-shadow);
}

.log-list {
    display: flex;
    flex-direction: column;
    gap: 10px;
    max-height: 68vh;
    overflow-y: auto;
    padding: 2px 4px 2px 0;
}

.log-row {
    display: grid;
    grid-template-columns: 158px 96px minmax(0, 1fr);
    gap: 12px;
    align-items: start;
    background: var(--md-surface);
    border: 1px solid var(--md-border);
    border-radius: 18px;
    padding: 12px 14px;
    box-shadow: 0 1px 2px rgba(60,64,67,.08);
}

.log-time {
    color: var(--md-muted);
    font-size: 12px;
    font-weight: 700;
    white-space: nowrap;
    padding-top: 3px;
}

.log-level {
    display: inline-flex;
    justify-content: center;
    align-items: center;
    min-width: 76px;
    border-radius: 999px;
    padding: 5px 9px;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: .3px;
    text-transform: uppercase;
}

.log-msg {
    color: var(--md-text);
    font-size: 13px;
    line-height: 1.45;
    overflow-wrap: anywhere;
}

.log-level-info,
.log-level-etapa,
.log-level-step,
.log-level-sched,
.log-level-intervalo {
    background: #E8F0FE;
    color: #174EA6;
}

.log-level-ok {
    background: #E6F4EA;
    color: #137333;
}

.log-level-warn {
    background: #FEF7E0;
    color: #B06000;
}

.log-level-fail,
.log-level-erro,
.log-level-error {
    background: #FCE8E6;
    color: #B3261E;
}

.log-empty {
    background: var(--md-surface-container);
    border: 1px dashed var(--md-border);
    color: var(--md-muted);
    border-radius: 18px;
    padding: 22px;
    text-align: center;
    font-size: 13px;
    font-weight: 600;
}

@media (max-width: 1100px) {
    .log-toolbar {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }
}

@media (max-width: 900px) {
    .log-toolbar {
        grid-template-columns: 1fr;
    }

    .log-row {
        grid-template-columns: 1fr;
        gap: 6px;
    }

    .log-time {
        white-space: normal;
    }

    .log-level {
        width: fit-content;
    }
}



/* =========================
   CONTROLE MANUAL DE ROBÔS
========================= */

.controle-page-wrap {
    max-width: 1120px;
    margin: 0 auto;
}

.controle-hero {
    background: var(--md-surface);
    border: 1px solid var(--md-border);
    border-radius: 28px;
    padding: 24px 28px;
    margin: 8px auto 18px auto;
    box-shadow: var(--md-shadow);
    text-align: center;
}

.controle-hero-icon {
    width: 52px;
    height: 52px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 18px;
    background: var(--md-blue-soft);
    color: var(--md-primary);
    font-size: 25px;
    margin-bottom: 10px;
}

.controle-hero-title {
    font-size: 30px;
    font-weight: 800;
    color: var(--md-text);
    margin-bottom: 4px;
}

.controle-hero-subtitle {
    color: var(--md-muted);
    font-size: 13px;
    line-height: 1.45;
}

.controle-login-card {
    max-width: 520px;
    margin: 24px auto;
    background: var(--md-surface);
    border: 1px solid var(--md-border);
    border-radius: 28px;
    padding: 24px 26px;
    box-shadow: var(--md-shadow);
}

.controle-login-title {
    font-size: 20px;
    font-weight: 800;
    color: var(--md-text);
    margin-bottom: 4px;
    text-align: center;
}

.controle-login-subtitle {
    color: var(--md-muted);
    font-size: 13px;
    text-align: center;
    margin-bottom: 12px;
}

.controle-alerta {
    background: var(--md-surface-container);
    border: 1px solid var(--md-border);
    border-radius: 18px;
    padding: 14px 16px;
    margin: 10px 0 14px 0;
    color: var(--md-text);
    box-shadow: var(--md-shadow);
}

.controle-grid-status {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 16px;
    margin-top: 10px;
    margin-bottom: 16px;
}

.controle-status-card {
    background: var(--md-surface-container);
    border: 1px solid var(--md-border);
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: var(--md-shadow);
}

.controle-status-card-title {
    color: var(--md-muted);
    font-size: 11px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-bottom: 6px;
}

.controle-status-card-value {
    color: var(--md-text);
    font-size: 13px;
    font-weight: 700;
    line-height: 1.45;
}

.controle-mini-note {
    color: var(--md-muted);
    font-size: 12px;
    line-height: 1.45;
    margin-top: 8px;
}

.controle-form-card {
    background: var(--md-surface);
    border: 1px solid var(--md-border);
    border-radius: 24px;
    padding: 20px 22px;
    box-shadow: var(--md-shadow);
    margin-bottom: 18px;
}

.controle-section-title {
    font-size: 18px;
    font-weight: 800;
    color: var(--md-text);
    margin-bottom: 2px;
}

.controle-section-desc {
    color: var(--md-muted);
    font-size: 12px;
    line-height: 1.45;
    margin-bottom: 8px;
}

@media (max-width: 900px) {
    .controle-grid-status {
        grid-template-columns: 1fr;
    }
}

/* =========================
   BOTÕES DESLIZANTES - CONTROLE DE ROBÔS
========================= */

/* Não pinta a linha nem o texto */
div[data-testid="stToggle"] label[data-baseweb="checkbox"],
div[data-testid="stToggle"] label[data-baseweb="checkbox"] > div {
    background: transparent !important;
}

/* OFF - botão deslizante vermelho */
div[data-testid="stToggle"] label[data-baseweb="checkbox"] > span:first-child > div {
    background-color: var(--md-red) !important;
    border-color: var(--md-red) !important;
}

/* ON - botão deslizante verde */
div[data-testid="stToggle"] label[data-baseweb="checkbox"]:has(input:checked) > span:first-child > div,
div[data-testid="stToggle"] label[data-baseweb="checkbox"]:has([aria-checked="true"]) > span:first-child > div,
div[data-testid="stToggle"]:has([aria-checked="true"]) label[data-baseweb="checkbox"] > span:first-child > div {
    background-color: var(--md-green) !important;
    border-color: var(--md-green) !important;
}

/* Bolinha interna */
div[data-testid="stToggle"] label[data-baseweb="checkbox"] > span:first-child > div > div {
    background-color: #FFFFFF !important;
}
/* =========================
   ALERTA VISUAL - CRÍTICAS / ROBÔS OFF
========================= */

@keyframes piscar-borda-alerta {
    0% {
        opacity: 1;
        box-shadow: 0 0 0 rgba(217, 48, 37, 0);
    }
    50% {
        opacity: 0.30;
        box-shadow: 0 0 18px rgba(217, 48, 37, 0.85);
    }
    100% {
        opacity: 1;
        box-shadow: 0 0 0 rgba(217, 48, 37, 0);
    }
}

.kpi-card.kpi-alerta-critico {
    position: relative !important;
}

.kpi-card.kpi-alerta-critico::before {
    content: "";
    position: absolute;
    left: 0;
    top: 0;
    bottom: 0;
    width: 9px;
    background: var(--md-red);
    border-radius: 24px 0 0 24px;
    animation: piscar-borda-alerta 0.9s infinite;
    z-index: 5;
}

</style>
"""

# Tema visual do painel: claro por padrão, escuro quando selecionado no menu lateral.
query_params = st.query_params
tema = str(query_params.get("tema", "claro")).lower().strip()
if tema not in ("claro", "escuro"):
    tema = "claro"

st.markdown(CSS, unsafe_allow_html=True)

CSS_ESCURO = """
<style>
:root {
    --md-bg: #030A14;
    --md-surface: #071931;
    --md-surface-container: #0A2344;
    --md-surface-variant: #0E315F;
    --md-primary: #33C7FF;
    --md-primary-dark: #0B5ED7;
    --md-text: #EAF4FF;
    --md-muted: #9DB7D2;
    --md-border: rgba(91, 166, 255, 0.22);
    --md-green: #21D07A;
    --md-yellow: #F6C343;
    --md-red: #FF4D5E;
    --md-blue-soft: rgba(51, 199, 255, 0.12);
    --md-shadow: 0 12px 28px rgba(0, 0, 0, .32);
}

.stApp {
    background: radial-gradient(circle at top left, #0B3D91 0%, #061529 28%, #030A14 100%) !important;
    color: var(--md-text) !important;
}

.hover-menu {
    background: #071931 !important;
    border-right: 1px solid rgba(91, 166, 255, 0.22) !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .36) !important;
}

.menu-icon,
.menu-title,
.menu-item {
    color: #EAF4FF !important;
}

.menu-item:hover {
    background: rgba(51, 199, 255, 0.14) !important;
    color: #33C7FF !important;
}

.hive-title {
    font-size: 32px !important;
    font-weight: 900 !important;
    letter-spacing: .5px !important;
    color: #EAF4FF !important;
    margin-bottom: 0px !important;
}

.hive-subtitle {
    color: #9DB7D2 !important;
    font-size: 13px !important;
    margin-top: 2px !important;
    margin-bottom: 18px !important;
}

.kpi-label {
    color: #9DB7D2 !important;
    font-size: 12px !important;
    font-weight: 800 !important;
    text-transform: uppercase !important;
    letter-spacing: .6px !important;
}

.kpi-value {
    font-size: 32px !important;
    font-weight: 900 !important;
    line-height: 1.08 !important;
    margin-top: 8px !important;
    white-space: nowrap !important;
}

.kpi-note {
    color: #9DB7D2 !important;
    font-size: 11px !important;
    margin-top: 9px !important;
}

.panel-title {
    font-size: 15px !important;
    font-weight: 800 !important;
    color: #EAF4FF !important;
    margin-bottom: 10px !important;
    letter-spacing: .3px !important;
}

.kpi-card {
    background: linear-gradient(160deg, rgba(14,49,95,.96), rgba(7,25,49,.96)) !important;
    border: 1px solid rgba(91, 166, 255, 0.22) !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .32) !important;
}

.kpi-label,
.kpi-note { color: #9DB7D2 !important; }

.panel {
    background: rgba(7, 25, 49, .94) !important;
    border: 1px solid rgba(91, 166, 255, 0.22) !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .28) !important;
}

.dark-table {
    background: #071931 !important;
    color: #EAF4FF !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .24) !important;
}

.dark-table thead th,
.table-scroll .dark-table thead th {
    background: #0E315F !important;
    color: #EAF4FF !important;
    border-bottom: 1px solid rgba(91, 166, 255, 0.25) !important;
}

.dark-table tbody td {
    color: #EAF4FF !important;
    border-bottom: 1px solid rgba(91, 166, 255, 0.12) !important;
}

.dark-table tbody tr:nth-child(even) { background: rgba(10, 35, 68, .96) !important; }
.dark-table tbody tr:nth-child(odd) { background: rgba(7, 25, 49, .96) !important; }

/* Ajusta textos inline usados nos subtítulos internos dos painéis. */
.panel div[style*="color:#202124"],
.panel div[style*="color: #202124"] {
    color: #EAF4FF !important;
}

/* Componentes nativos do Streamlit no tema escuro */
.stDateInput label,
.stSelectbox label,
.stMultiSelect label,
.stCheckbox label,
.stTextInput label {
    color: #EAF4FF !important;
}

[data-baseweb="input"],
[data-baseweb="select"] {
    background-color: #071931 !important;
    color: #EAF4FF !important;
}

/* Dataframes nativos, quando usados. */
div[data-testid="stDataFrame"] {
    border: 1px solid rgba(91, 166, 255, 0.22) !important;
    border-radius: 14px !important;
    background: rgba(7, 25, 49, .96) !important;
}

div[data-testid="stDataFrame"] [role="table"],
div[data-testid="stDataFrame"] [role="gridcell"],
div[data-testid="stDataFrame"] .glideDataEditor,
div[data-testid="stDataFrame"] .dvn-scroller,
div[data-testid="stDataFrame"] .gdg-cell {
    background: rgba(7, 25, 49, .96) !important;
    color: #EAF4FF !important;
}

div[data-testid="stDataFrame"] [role="columnheader"],
div[data-testid="stDataFrame"] .gdg-header {
    background: #0E315F !important;
    color: #EAF4FF !important;
}


/* =========================
   LOGS DO ROBÔ - TEMA ESCURO
========================= */

.menu-divider {
    background: rgba(91, 166, 255, 0.22) !important;
}

.menu-item-theme {
    color: #9DB7D2 !important;
}

.log-meta-card,
.log-filter-box {
    background: rgba(10, 35, 68, .96) !important;
    border: 1px solid rgba(91, 166, 255, 0.22) !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .28) !important;
}

.log-meta-label {
    color: #9DB7D2 !important;
}

.log-meta-value {
    color: #EAF4FF !important;
}

.log-row {
    background: rgba(7, 25, 49, .96) !important;
    border: 1px solid rgba(91, 166, 255, 0.18) !important;
    box-shadow: 0 8px 20px rgba(0, 0, 0, .24) !important;
}

.log-time {
    color: #9DB7D2 !important;
}

.log-msg {
    color: #EAF4FF !important;
}

.log-level-info,
.log-level-etapa,
.log-level-step,
.log-level-sched,
.log-level-intervalo {
    background: rgba(51, 199, 255, 0.14) !important;
    color: #33C7FF !important;
}

.log-level-ok {
    background: rgba(33, 208, 122, 0.14) !important;
    color: #21D07A !important;
}

.log-level-warn {
    background: rgba(246, 195, 67, 0.16) !important;
    color: #F6C343 !important;
}

.log-level-fail,
.log-level-erro,
.log-level-error {
    background: rgba(255, 77, 94, 0.16) !important;
    color: #FF4D5E !important;
}

.log-empty {
    background: rgba(10, 35, 68, .80) !important;
    border: 1px dashed rgba(91, 166, 255, 0.25) !important;
    color: #9DB7D2 !important;
}


.controle-hero,
.controle-login-card,
.controle-form-card,
.controle-alerta,
.controle-status-card {
    background: rgba(10, 35, 68, .96) !important;
    border: 1px solid rgba(91, 166, 255, 0.22) !important;
    box-shadow: 0 12px 28px rgba(0, 0, 0, .28) !important;
}

.controle-hero-title,
.controle-login-title,
.controle-section-title,
.controle-status-card-value,
.controle-alerta {
    color: #EAF4FF !important;
}

.controle-hero-subtitle,
.controle-login-subtitle,
.controle-status-card-title,
.controle-mini-note,
.controle-section-desc {
    color: #9DB7D2 !important;
}
</style>
"""

if tema == "escuro":
    st.markdown(CSS_ESCURO, unsafe_allow_html=True)


# =========================
# MENU
# =========================

# query_params e tema já foram definidos acima, antes da aplicação do CSS.

pagina = query_params.get(
    "pagina",
    "Monitoramento atual"
)

proximo_tema = "claro" if tema == "escuro" else "escuro"
rotulo_tema = "☀️ Modo claro" if tema == "escuro" else "🌙 Modo escuro"

menu_html = f"""
<div class="hover-menu">

    <div class="menu-icon">
        ☰
    </div>

    <div class="menu-title">
        Monitoramento
    </div>

    <a class="menu-item"
       href="?pagina=Monitoramento atual&tema={tema}">
       📊 Monitoramento
    </a>

    <a class="menu-item"
       href="?pagina=Histórico monitoramento&tema={tema}">
       📁 Histórico
    </a>

    <a class="menu-item"
       href="?pagina=Logs&tema={tema}">
       🧾 Logs
    </a>

    <div class="menu-divider"></div>

    <a class="menu-item"
       href="?pagina=Ligar/Desligar&tema={tema}">
       🔐 Ligar/Desligar
    </a>

    <a class="menu-item menu-item-theme"
       href="?pagina={pagina}&tema={proximo_tema}">
       {rotulo_tema}
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
@st.cache_data(ttl=60)
def listar_datas_logs_disponiveis():
    """
    Lista no GitHub os arquivos existentes em /logs
    e retorna somente as datas que possuem Log_YYYY_MM_DD.csv.
    """
    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/logs"
    )

    try:
        resp = requests.get(
            api_url,
            headers=github_headers(),
            params={"ref": GITHUB_BRANCH},
            timeout=30,
        )

        if resp.status_code == 404:
            return []

        resp.raise_for_status()

        dados = resp.json()

        if not isinstance(dados, list):
            return []

        datas = []

        for item in dados:
            nome = str(item.get("name", "") or "")

            m = re.fullmatch(r"Log_(\d{4})_(\d{2})_(\d{2})\.csv", nome)

            if not m:
                continue

            ano, mes, dia = map(int, m.groups())

            try:
                datas.append(datetime(ano, mes, dia).date())
            except Exception:
                continue

        return sorted(set(datas), reverse=True)

    except Exception:
        return []

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

def carregar_alerta_ecrv():
    try:
        if ARQ_ALERTA_ECRV.exists():
            return json.loads(
                ARQ_ALERTA_ECRV.read_text(encoding="utf-8")
            )
    except Exception:
        pass

    return {"alerta_enviado": False}


def salvar_alerta_ecrv(dados):
    try:
        ARQ_ALERTA_ECRV.write_text(
            json.dumps(dados, ensure_ascii=False),
            encoding="utf-8"
        )
    except Exception:
        pass

def normalizar_status_on_off(valor, padrao="ON") -> str:
    """Normaliza qualquer valor recebido para ON/OFF."""
    txt = str(valor if valor is not None else "").strip().upper()

    if txt in ("ON", "LIGADO", "TRUE", "1", "SIM"):
        return "ON"

    if txt in ("OFF", "DESLIGADO", "FALSE", "0", "NAO", "NÃO"):
        return "OFF"

    return padrao


def parse_data_hora_painel(valor):
    """Converte datas usadas pelo dashboard/RPA em Timestamp, quando possível."""
    if valor is None or str(valor).strip() == "":
        return None

    try:
        dt = pd.to_datetime(
            valor,
            dayfirst=True,
            errors="coerce",
            format="mixed",
        )
    except Exception:
        return None

    if pd.isna(dt):
        return None

    try:
        return pd.Timestamp(dt).tz_localize(None)
    except Exception:
        return pd.Timestamp(dt)


def rerun_streamlit():
    """Compatibilidade entre versões do Streamlit."""
    try:
        st.rerun()
    except Exception:
        try:
            st.experimental_rerun()
        except Exception:
            pass


def carregar_json_github(caminho_repo: str, destino_local: Path, obrigatorio: bool = False) -> tuple[dict, dict]:
    """Carrega um JSON do GitHub usando o mesmo mecanismo de cache do painel."""
    try:
        caminho_local, meta = baixar_github_se_houver_alteracao(
            caminho_repo,
            destino_local,
            obrigatorio=obrigatorio,
        )

        if not caminho_local or not Path(caminho_local).exists():
            return {}, meta or {"status": "ausente", "path": caminho_repo}

        conteudo = Path(caminho_local).read_text(encoding="utf-8").strip()

        if not conteudo:
            return {}, meta or {"path": caminho_repo}

        return json.loads(conteudo), meta or {"path": caminho_repo}

    except FileNotFoundError:
        if obrigatorio:
            raise
        return {}, {"status": "ausente", "path": caminho_repo}

    except Exception as e:
        return {}, {"status": "erro", "path": caminho_repo, "erro": str(e)}


def salvar_json_github(caminho_repo: str, payload: dict, mensagem_commit: str) -> tuple[bool, str]:
    """
    Cria/atualiza um JSON no GitHub.

    Necessita TOKEN_GITHUB com permissão de escrita no repositório.
    """
    token = obter_token_github()

    if not token:
        return False, "TOKEN_GITHUB não configurado. Não foi possível gravar o comando no GitHub."

    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/{caminho_repo}"
    )

    headers = github_headers()

    try:
        resp_get = requests.get(
            api_url,
            headers=headers,
            params={"ref": GITHUB_BRANCH},
            timeout=30,
        )

        sha_atual = None

        if resp_get.status_code == 200:
            sha_atual = resp_get.json().get("sha")
        elif resp_get.status_code != 404:
            return False, f"Erro ao consultar arquivo no GitHub: HTTP {resp_get.status_code} - {resp_get.text[:300]}"

        conteudo_json = json.dumps(
            payload,
            ensure_ascii=False,
            indent=2,
        )

        body = {
            "message": mensagem_commit,
            "content": base64.b64encode(conteudo_json.encode("utf-8")).decode("ascii"),
            "branch": GITHUB_BRANCH,
        }

        if sha_atual:
            body["sha"] = sha_atual

        resp_put = requests.put(
            api_url,
            headers=headers,
            json=body,
            timeout=30,
        )

        if resp_put.status_code not in (200, 201):
            return False, f"Erro ao gravar arquivo no GitHub: HTTP {resp_put.status_code} - {resp_put.text[:300]}"

        dados_resp = resp_put.json()
        sha_novo = (
            dados_resp.get("content", {}) or {}
        ).get("sha", "")

        # Atualiza cache local/metadados para evitar leitura defasada.
        destino = (
            ARQ_LOCAL_CONTROLE_ECRV
            if caminho_repo == GITHUB_ARQ_CONTROLE_ECRV
            else CACHE_DIR / caminho_repo.replace("/", "_")
        )

        try:
            destino.write_text(conteudo_json, encoding="utf-8")

            meta = carregar_meta_cache()
            meta[caminho_repo] = {
                "sha": sha_novo,
                "downloaded_at": agora_sao_paulo().strftime("%d/%m/%Y %H:%M:%S"),
                "size": len(conteudo_json.encode("utf-8")),
            }
            salvar_meta_cache(meta)
        except Exception:
            pass

        return True, "Comando gravado no GitHub com sucesso."

    except Exception as e:
        return False, f"Erro ao gravar comando no GitHub: {e}"


def status_card_robos(status_dict, robo_monitoramento_online=True) -> dict:
    """Retorna o status no mesmo formato exibido no card Robôs."""
    return {
        "Transferência 2": normalizar_status_on_off(status_dict.get("Transferência 2", "ON")),
        "Transferência 3": normalizar_status_on_off(status_dict.get("Transferência 3", "ON")),
        "0KM": normalizar_status_on_off(status_dict.get("0KM", "ON")),
        "Monitoramento e-CRV": "ON" if robo_monitoramento_online else "OFF",
    }


def servicos_controle_robos() -> list[str]:
    return [
        "Transferência 2",
        "Transferência 3",
        "0KM",
        "Monitoramento e-CRV",
    ]


def aplicar_status_confirmado_rpa(
    status_dict,
    robo_monitoramento_online,
    status_ecrv: dict | None,
    df_criticas=None,
    df_hist=None,
):
    """
    Aplica no painel o status confirmado pelo RPA, quando existir.

    Regra de segurança:
    - se houver STOP posterior à confirmação do RPA, o STOP prevalece;
    - o status confirmado não impede novo desligamento automático.
    """
    if not isinstance(status_ecrv, dict) or not status_ecrv:
        return status_dict, robo_monitoramento_online

    servicos = status_ecrv.get("servicos")

    if not isinstance(servicos, dict):
        return status_dict, robo_monitoramento_online

    atualizado_em = (
        status_ecrv.get("atualizado_em")
        or status_ecrv.get("processado_em")
        or status_ecrv.get("data_hora")
    )

    dt_confirmacao = parse_data_hora_painel(atualizado_em)

    if dt_confirmacao is None:
        return status_dict, robo_monitoramento_online

    try:
        if (agora_sao_paulo() - dt_confirmacao).total_seconds() / 60 > JANELA_STATUS_RPA_MINUTOS:
            return status_dict, robo_monitoramento_online
    except Exception:
        return status_dict, robo_monitoramento_online

    status_final = dict(status_dict or {})

    for servico in ["Transferência 2", "Transferência 3", "0KM"]:
        if servico not in servicos:
            continue

        desejado = normalizar_status_on_off(servicos.get(servico), padrao=status_final.get(servico, "ON"))

        if desejado == "OFF":
            status_final[servico] = "OFF"
            continue

        # Se houver STOP posterior à confirmação, mantém a regra automática.
        ultimo_stop = _ultimo_stop_servico(df_criticas, df_hist, servico)

        if ultimo_stop is not None and dt_confirmacao is not None and pd.Timestamp(ultimo_stop) > dt_confirmacao:
            continue

        status_final[servico] = "ON"

    if "Monitoramento e-CRV" in servicos:
        status_monitor = normalizar_status_on_off(
            servicos.get("Monitoramento e-CRV"),
            padrao=("ON" if robo_monitoramento_online else "OFF"),
        )
        robo_monitoramento_online = status_monitor == "ON"

    return status_final, robo_monitoramento_online


def aplicar_controle_dashboard_manual(
    status_dict,
    robo_monitoramento_online,
    df_criticas=None,
    df_hist=None,
):
    """
    Aplica controle visual temporário do Dashboard.

    Regras preservadas:
    - OFF manual força OFF temporariamente no painel;
    - ON manual não vence STOP novo posterior à intervenção;
    - após TEMPO_MINIMO_OFF_MINUTOS, volta à regra automática.
    """
    controle = st.session_state.get("controle_dash_status")
    ts_manual = st.session_state.get("controle_dash_ts")
    ate_manual = st.session_state.get("controle_dash_ate")

    if not isinstance(controle, dict) or not ts_manual or not ate_manual:
        return status_dict, robo_monitoramento_online

    agora = agora_sao_paulo()
    ts_manual = parse_data_hora_painel(ts_manual)
    ate_manual = parse_data_hora_painel(ate_manual)

    if ts_manual is None or ate_manual is None or agora > ate_manual:
        return status_dict, robo_monitoramento_online

    status_final = dict(status_dict or {})

    for servico in ["Transferência 2", "Transferência 3", "0KM"]:
        if servico not in controle:
            continue

        desejado = normalizar_status_on_off(controle.get(servico), padrao=status_final.get(servico, "ON"))

        if desejado == "OFF":
            status_final[servico] = "OFF"
            continue

        ultimo_stop = _ultimo_stop_servico(df_criticas, df_hist, servico)

        if ultimo_stop is None or pd.Timestamp(ultimo_stop) <= ts_manual:
            status_final[servico] = "ON"

    if "Monitoramento e-CRV" in controle:
        status_monitor = normalizar_status_on_off(
            controle.get("Monitoramento e-CRV"),
            padrao=("ON" if robo_monitoramento_online else "OFF"),
        )
        robo_monitoramento_online = status_monitor == "ON"

    return status_final, robo_monitoramento_online


def status_base_ecrv_para_controle(status_painel: dict, status_ecrv: dict | None = None) -> dict:
    """
    Status inicial dos botões e-CRV.

    Usa primeiro a confirmação do RPA em status/status_ecrv.json.
    Se não houver retorno confirmado, usa o status atual do card Robôs.
    """
    base = dict(status_painel or {})

    if isinstance(status_ecrv, dict):
        atualizado_em = (
            status_ecrv.get("atualizado_em")
            or status_ecrv.get("processado_em")
            or status_ecrv.get("data_hora")
        )
        dt_confirmacao = parse_data_hora_painel(atualizado_em)

        status_rpa_recente = False
        try:
            status_rpa_recente = (
                dt_confirmacao is not None
                and (agora_sao_paulo() - dt_confirmacao).total_seconds() / 60 <= JANELA_STATUS_RPA_MINUTOS
            )
        except Exception:
            status_rpa_recente = False

        servicos = status_ecrv.get("servicos")
        if status_rpa_recente and isinstance(servicos, dict):
            for servico in servicos_controle_robos():
                if servico in servicos:
                    base[servico] = normalizar_status_on_off(
                        servicos.get(servico),
                        padrao=base.get(servico, "ON"),
                    )

    for servico in servicos_controle_robos():
        base[servico] = normalizar_status_on_off(base.get(servico), padrao="ON")

    return base


def fechar_controle_robos():
    """Reseta o estado interno do controle de robôs."""
    st.session_state["controle_robos_aberto"] = False
    st.session_state["controle_robos_inicializado"] = False
    st.session_state["controle_robos_reinicializar"] = True


def inicializar_chaves_controle_robos(status_painel: dict, status_ecrv_base: dict):
    """Inicializa os botões ON/OFF com o status atual do painel/e-CRV."""
    reinicializar = bool(st.session_state.pop("controle_robos_reinicializar", False))

    if st.session_state.get("controle_robos_inicializado") and not reinicializar:
        return

    for servico in servicos_controle_robos():
        st.session_state[f"dash_{servico}"] = (
            normalizar_status_on_off(status_painel.get(servico), padrao="ON") == "ON"
        )
        st.session_state[f"ecrv_{servico}"] = (
            normalizar_status_on_off(status_ecrv_base.get(servico), padrao="ON") == "ON"
        )

    st.session_state["controle_robos_inicializado"] = True


def render_controle_robos(
    status_painel: dict,
    status_ecrv: dict | None = None,
    comando_ecrv: dict | None = None,
):
    """Renderiza a página protegida de Ligar/Desligar robôs."""
    st.markdown(
        """
<div class="controle-page-wrap">
    <div class="controle-hero">
        <div class="controle-hero-icon">🔐</div>
        <div class="controle-hero-title">Controle de Robôs</div>
        <div class="controle-hero-subtitle">
            Controle visual do Dashboard e envio de comandos para execução pelo RPA de monitoramento.
        </div>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    senha_configurada = str(
        st.secrets.get(
            "SENHA_CONTROLE_ROBOS",
            os.getenv("SENHA_CONTROLE_ROBOS", ""),
        )
    ).strip()

    if not senha_configurada:
        st.error("Configure SENHA_CONTROLE_ROBOS nos Secrets do Streamlit para habilitar esta página.")
        return

    if not st.session_state.get("controle_robos_autorizado"):
        col_login_1, col_login_2, col_login_3 = st.columns([1, 1.15, 1])

        with col_login_2:
            st.markdown(
                """
<div class="controle-login-card">
    <div class="controle-login-title">Acesso restrito</div>
    <div class="controle-login-subtitle">
        Informe a senha para acessar as opções de ligar/desligar.
    </div>
""",
                unsafe_allow_html=True,
            )

            with st.form("form_senha_controle_robos"):
                senha_digitada = st.text_input(
                    "Senha",
                    type="password",
                    key="senha_controle_robos",
                )
                entrar = st.form_submit_button("Entrar", type="primary", use_container_width=True)

            st.markdown("</div>", unsafe_allow_html=True)

            if entrar:
                if senha_digitada == senha_configurada:
                    st.session_state["controle_robos_autorizado"] = True
                    st.session_state["controle_robos_reinicializar"] = True
                    rerun_streamlit()
                else:
                    st.error("Senha incorreta.")

        return

    status_painel = {
        servico: normalizar_status_on_off(status_painel.get(servico), padrao="ON")
        for servico in servicos_controle_robos()
    }

    status_ecrv_base = status_base_ecrv_para_controle(status_painel, status_ecrv)
    inicializar_chaves_controle_robos(status_painel, status_ecrv_base)

    comando_ecrv = comando_ecrv or {}
    ultimo_status_comando = str(comando_ecrv.get("status", "") or "").strip().upper()
    ultimo_comando_id = str(comando_ecrv.get("id", "") or "").strip()

    status_rpa_txt = "Sem confirmação do RPA"
    if isinstance(status_ecrv, dict) and status_ecrv:
        status_rpa_txt = (
            f"Atualizado em {status_ecrv.get('atualizado_em', '-')}"
            f" | último comando: {status_ecrv.get('ultimo_comando_status', '-')}"
        )

    st.markdown('<div class="controle-page-wrap">', unsafe_allow_html=True)

    st.markdown(
        f"""
<div class="controle-grid-status">
    <div class="controle-status-card">
        <div class="controle-status-card-title">Último comando e-CRV/RPA</div>
        <div class="controle-status-card-value">
            Status: {html.escape(ultimo_status_comando or "Sem comando")}<br>
            ID: {html.escape(ultimo_comando_id or "-")}
        </div>
    </div>
    <div class="controle-status-card">
        <div class="controle-status-card-title">Confirmação recebida do RPA</div>
        <div class="controle-status-card-value">{html.escape(status_rpa_txt)}</div>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    if ultimo_status_comando == "PENDENTE":
        st.warning("Há comando pendente para o RPA processar. Evite enviar novo comando antes do retorno, salvo necessidade operacional.")

    st.markdown('<div class="controle-form-card">', unsafe_allow_html=True)

    toggle_fn = getattr(st, "toggle", st.checkbox)

    with st.form("form_controle_robos"):
        col_dash, col_ecrv = st.columns(2)

        with col_dash:
            st.markdown('<div class="controle-section-title">Dashboard</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="controle-section-desc">Altera somente o status visual do painel. Não executa ação no e-CRV.</div>',
                unsafe_allow_html=True,
            )

            dash_valores = {}
            for servico in servicos_controle_robos():
                dash_valores[servico] = toggle_fn(
                    servico,
                    key=f"dash_{servico}",
                )

        with col_ecrv:
            st.markdown('<div class="controle-section-title">e-CRV</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="controle-section-desc">Grava comando no GitHub para o RPA de monitoramento executar/confirmar depois.</div>',
                unsafe_allow_html=True,
            )

            ecrv_valores = {}
            for servico in servicos_controle_robos():
                ecrv_valores[servico] = toggle_fn(
                    servico,
                    key=f"ecrv_{servico}",
                )

        st.markdown("---")
        col_aplicar, col_sair = st.columns([1, 1])
        with col_aplicar:
            aplicar = st.form_submit_button("Aplicar", type="primary", use_container_width=True)
        with col_sair:
            sair = st.form_submit_button("Bloquear acesso", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    if sair:
        st.session_state["controle_robos_autorizado"] = False
        st.session_state["senha_controle_robos"] = ""
        rerun_streamlit()
        return

    if not aplicar:
        st.markdown(
            f"""
<div class="controle-alerta">
    <b>Observação:</b> os botões abrem com base no status atual exibido no card Robôs.<br>
    O status real do e-CRV somente será considerado confirmado após o RPA atualizar
    <b>{html.escape(GITHUB_ARQ_STATUS_ECRV)}</b>.
</div>
""",
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    agora = agora_sao_paulo()
    agora_txt = agora.strftime("%d/%m/%Y %H:%M:%S")

    # 1) Controle visual do Dashboard.
    st.session_state["controle_dash_status"] = {
        servico: ("ON" if ativo else "OFF")
        for servico, ativo in dash_valores.items()
    }
    st.session_state["controle_dash_ts"] = agora
    st.session_state["controle_dash_ate"] = agora + pd.Timedelta(
        minutes=TEMPO_MINIMO_OFF_MINUTOS
    )
    st.session_state["controle_robos_reinicializar"] = True

    # 2) Comando real/sinal para o RPA via GitHub.
    servicos_desejados = {
        servico: ("ON" if ativo else "OFF")
        for servico, ativo in ecrv_valores.items()
    }

    acoes = []
    for servico, status_desejado in servicos_desejados.items():
        status_anterior = status_ecrv_base.get(servico, status_painel.get(servico, "ON"))
        status_anterior = normalizar_status_on_off(status_anterior, padrao="ON")
    
        if status_anterior != status_desejado:
            acoes.append({
                "servico": servico,
                "acao": "LIGAR" if status_desejado == "ON" else "DESLIGAR",
                "status_anterior": status_anterior,
                "status_desejado": status_desejado,
            })
    
    if not acoes:
        st.success(
            "Controle do Dashboard aplicado. Nenhum comando foi enviado ao RPA, "
            "pois não houve alteração no bloco e-CRV."
        )
    
        st.markdown(
            """
    <div class="controle-alerta">
        <b>Resultado:</b> alteração aplicada somente no Dashboard.<br>
        <span class="controle-mini-note">
            O bloco e-CRV não foi alterado, portanto nenhum comando PENDENTE foi gravado no GitHub.
        </span>
    </div>
    """,
            unsafe_allow_html=True,
        )
    
        st.markdown("</div>", unsafe_allow_html=True)
    
        st.query_params["pagina"] = "Monitoramento atual"
        st.query_params["tema"] = tema
        rerun_streamlit()
        return
    
    payload = {
        "id": agora.strftime("%Y%m%d_%H%M%S"),
        "origem": "dashboard",
        "tipo": "CONTROLE_ROBOS",
        "status": "PENDENTE",
        "solicitado_em": agora_txt,
        "expira_em": (agora + pd.Timedelta(minutes=30)).strftime("%d/%m/%Y %H:%M:%S"),
        "servicos_desejados": servicos_desejados,
        "acoes": acoes,
        "observacao": (
            "Comando gerado pelo dashboard. O RPA de monitoramento deve ler este arquivo, "
            "executar as ações aplicáveis no STOP Processo/e-CRV ou no próprio monitoramento, "
            "e gravar o retorno em status/status_ecrv.json."
        ),
    }

    ok, msg = salvar_json_github(
        GITHUB_ARQ_CONTROLE_ECRV,
        payload,
        f"Controle robôs dashboard {payload['id']}",
    )

    if ok:
        st.success("Controle do Dashboard aplicado e comando gravado no GitHub para o RPA processar.")
    else:
        st.warning("Controle do Dashboard aplicado, mas o comando para o RPA não foi gravado no GitHub.")
        st.error(msg)

    st.markdown(
        f"""
<div class="controle-alerta">
    <b>Resultado:</b> {html.escape(msg)}<br>
    <span class="controle-mini-note">
        O status real do e-CRV será considerado confirmado somente após o RPA atualizar
        <b>{html.escape(GITHUB_ARQ_STATUS_ECRV)}</b>.
    </span>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown("</div>", unsafe_allow_html=True)


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

    downloaded_at = agora_sao_paulo().strftime("%d/%m/%Y %H:%M:%S")
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

def carregar_historico_multimes(prefixo: str, obrigatorio: bool = False) -> pd.DataFrame:
    """
    Carrega o histórico dos últimos meses, conforme QUANTIDADE_MESES_HISTORICO,
    juntando os CSVs disponíveis em um único DataFrame.
    """

    dfs = []

    for caminho_repo in caminhos_historico_ultimos_meses(prefixo):
        nome_local = caminho_repo.replace("/", "_").replace(".csv", "")
        destino_local = CACHE_DIR / f"{nome_local}.csv"

        try:
            caminho_local, _ = baixar_github_se_houver_alteracao(
                caminho_repo,
                destino_local,
                obrigatorio=False,
            )

            if caminho_local and Path(caminho_local).exists():
                df_tmp = carregar_csv_historico(
                    str(caminho_local),
                    Path(caminho_local).stat().st_mtime,
                )

                if df_tmp is not None and not df_tmp.empty:
                    dfs.append(df_tmp)

        except FileNotFoundError:
            continue

        except Exception:
            continue

    if not dfs:
        if obrigatorio:
            raise FileNotFoundError(
                f"Nenhum arquivo de histórico encontrado para {prefixo}."
            )
        return pd.DataFrame()

    df_final = pd.concat(dfs, ignore_index=True)
    df_final = normalizar_coluna_data_para_date(df_final)

    return df_final


def classe_nivel_log(nivel: str) -> str:
    nivel_norm = str(nivel or "INFO").strip().lower()

    nivel_norm = unicodedata.normalize("NFKD", nivel_norm)
    nivel_norm = "".join(c for c in nivel_norm if not unicodedata.combining(c))
    nivel_norm = re.sub(r"[^a-z0-9]+", "-", nivel_norm).strip("-")

    if not nivel_norm:
        nivel_norm = "info"

    return f"log-level-{nivel_norm}"


def carregar_log_diario_dashboard(data_log=None) -> tuple[pd.DataFrame, dict]:
    """
    Carrega o CSV diário de logs gerado pelo robô.

    Arquivo esperado no GitHub:
      logs/Log_YYYY_MM_DD.csv
    """
    meta_log = {}
    caminho_repo = caminho_github_log_data(data_log)
    arquivo_local = caminho_local_log_data(data_log)

    try:
        caminho_log, meta_log = baixar_github_se_houver_alteracao(
            caminho_repo,
            arquivo_local,
            obrigatorio=False,
        )

        meta_log = meta_log or {}
        meta_log["path"] = caminho_repo

        if not caminho_log or not Path(caminho_log).exists():
            return pd.DataFrame(), meta_log

        df_log = pd.read_csv(
            caminho_log,
            sep=";",
            encoding="utf-8-sig",
        )

        df_log = normalizar_colunas(df_log)

        if df_log.empty:
            return pd.DataFrame(), meta_log

        for col in ["Data/Hora", "Horário", "Nível", "Mensagem"]:
            if col not in df_log.columns:
                df_log[col] = ""

        df_log["Data/Hora"] = df_log["Data/Hora"].astype(str).str.strip()
        df_log["Horário"] = df_log["Horário"].astype(str).str.strip()
        df_log["Nível"] = df_log["Nível"].astype(str).str.strip().replace("", "INFO")
        df_log["Mensagem"] = df_log["Mensagem"].astype(str).str.strip()

        df_log["_ordem"] = pd.to_datetime(
            df_log["Data/Hora"],
            dayfirst=True,
            errors="coerce",
            format="mixed",
        )

        df_log = (
            df_log
            .sort_values("_ordem", ascending=False)
            .drop(columns=["_ordem"], errors="ignore")
            .reset_index(drop=True)
        )

        return df_log, meta_log

    except Exception:
        return pd.DataFrame(), meta_log or {"path": caminho_repo}


def render_logs_dashboard(df_log: pd.DataFrame, meta_log: dict | None = None):
    meta_log = meta_log or {}

    total_logs = 0 if df_log is None or df_log.empty else len(df_log)
    ultima_atualizacao = meta_log.get("downloaded_at", "") or "Não sincronizado"
    arquivo_log = meta_log.get("path") or GITHUB_ARQ_LOG_DIA

    if df_log is None or df_log.empty:
        ultimo_log = "-"
    else:
        ultimo_log = str(df_log.iloc[0].get("Data/Hora", "-"))

    html_meta = (
        '<div class="log-toolbar">'
        '<div class="log-meta-card">'
        '<div class="log-meta-label">Arquivo</div>'
        f'<div class="log-meta-value">{html.escape(str(arquivo_log))}</div>'
        '</div>'
        '<div class="log-meta-card">'
        '<div class="log-meta-label">Linhas carregadas</div>'
        f'<div class="log-meta-value">{html.escape(fmt_num(total_logs))}</div>'
        '</div>'
        '<div class="log-meta-card">'
        '<div class="log-meta-label">Último log</div>'
        f'<div class="log-meta-value">{html.escape(str(ultimo_log))}</div>'
        '</div>'
        '<div class="log-meta-card">'
        '<div class="log-meta-label">Sincronizado em</div>'
        f'<div class="log-meta-value">{html.escape(str(ultima_atualizacao))}</div>'
        '</div>'
        '</div>'
    )
    st.markdown(html_meta, unsafe_allow_html=True)

    if df_log is None or df_log.empty:
        st.markdown(
            '<div class="log-empty">Nenhum log encontrado para a data selecionada.</div>',
            unsafe_allow_html=True,
        )
        return

    niveis_disponiveis = sorted(
        [n for n in df_log["Nível"].dropna().astype(str).unique() if n.strip()]
    )

    st.markdown('<div class="log-filter-box">', unsafe_allow_html=True)
    col_filtro_1, col_filtro_2 = st.columns([2, 1])

    with col_filtro_1:
        niveis_selecionados = st.multiselect(
            "Filtrar por nível",
            options=niveis_disponiveis,
            default=[],
            key="logs_filtro_nivel",
        )

    with col_filtro_2:
        limite_linhas = st.selectbox(
            "Quantidade",
            options=[50, 100, 200, 500],
            index=1,
            key="logs_limite_linhas",
        )

    st.markdown('</div>', unsafe_allow_html=True)

    df_view = df_log.copy()

    if niveis_selecionados:
        df_view = df_view[df_view["Nível"].isin(niveis_selecionados)].copy()

    df_view = df_view.head(int(limite_linhas)).copy()

    if df_view.empty:
        st.markdown(
            '<div class="log-empty">Nenhum log encontrado com os filtros selecionados.</div>',
            unsafe_allow_html=True,
        )
        return

    linhas_html = []

    for _, row in df_view.iterrows():
        data_hora = str(row.get("Data/Hora", "") or "").strip()
        nivel = str(row.get("Nível", "INFO") or "INFO").strip()
        msg = str(row.get("Mensagem", "") or "").strip()

        classe = classe_nivel_log(nivel)

        linhas_html.append(
            '<div class="log-row">'
            f'<div class="log-time">{html.escape(data_hora)}</div>'
            f'<div><span class="log-level {html.escape(classe)}">{html.escape(nivel)}</span></div>'
            f'<div class="log-msg">{html.escape(msg)}</div>'
            '</div>'
        )

    st.markdown(
        '<div class="log-list">' + ''.join(linhas_html) + '</div>',
        unsafe_allow_html=True,
    )

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



def preparar_dados_cards(df_base: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara somente o dia mais recente para os cálculos dos cards.

    Também remove coletas repetidas na coluna Data/Hora. Isso evita que uma
    única coleta replicada em vários intervalos distorça as janelas móveis.
    """
    if df_base is None or df_base.empty:
        return pd.DataFrame()

    df_cards = df_base.copy()

    if "Data/Hora" in df_cards.columns:
        df_cards["_momento"] = pd.to_datetime(
            df_cards["Data/Hora"],
            dayfirst=True,
            errors="coerce",
            format="mixed",
        )
    elif "Horário" in df_cards.columns:
        hoje = agora_sao_paulo().normalize()
        df_cards["_momento"] = pd.to_datetime(
            hoje.strftime("%Y-%m-%d")
            + " "
            + df_cards["Horário"].astype(str).str.slice(0, 5),
            errors="coerce",
        )
    else:
        return pd.DataFrame()

    df_cards = df_cards.dropna(subset=["_momento"])

    if df_cards.empty:
        return df_cards

    data_mais_recente = df_cards["_momento"].max().normalize()
    df_cards = df_cards[
        df_cards["_momento"].dt.normalize() == data_mais_recente
    ].copy()

    return (
        df_cards
        .sort_values("_momento")
        .drop_duplicates(subset=["_momento"], keep="last")
        .reset_index(drop=True)
    )


def linha_proxima_janela(
    df_cards: pd.DataFrame,
    momento_alvo: pd.Timestamp,
    tolerancia_minutos: int = 20,
):
    """Retorna a coleta mais próxima do momento desejado, dentro da tolerância."""
    if df_cards is None or df_cards.empty:
        return None

    distancias = (df_cards["_momento"] - momento_alvo).abs()
    indice = distancias.idxmin()

    if distancias.loc[indice] > pd.Timedelta(minutes=tolerancia_minutos):
        return None

    return df_cards.loc[indice]


def valor_numerico_linha(linha, coluna: str) -> float:
    if linha is None or coluna not in linha:
        return 0.0

    valor = pd.to_numeric(linha.get(coluna, 0), errors="coerce")
    return 0.0 if pd.isna(valor) else float(valor)


def minutos_escoamento(fila: float, processados_60min: float):
    """Calcula quantos minutos a fila representa no ritmo da última hora."""
    fila = float(fila or 0)
    processados_60min = float(processados_60min or 0)

    if fila <= 0:
        return 0.0

    if processados_60min <= 0:
        return float("inf")

    return (fila / processados_60min) * 60.0


def percentual_br(valor, casas=1) -> str:
    if valor is None or pd.isna(valor):
        return "-"
    return f"{float(valor):.{casas}f}".replace(".", ",") + "%"


def decimal_br(valor, casas=1) -> str:
    if valor is None or pd.isna(valor):
        return "-"
    return f"{float(valor):.{casas}f}".replace(".", ",")


def tendencia_html(simbolo: str, cor: str, texto: str = "") -> str:
    complemento = f" {texto}" if texto else ""
    return (
        f'<span style="color:{cor}; font-weight:800; white-space:nowrap;">'
        f'{simbolo}{complemento}</span>'
    )


def tendencia_sucesso_html(variacao_pct) -> str:
    if variacao_pct is None or pd.isna(variacao_pct):
        return tendencia_html("●", COR_AZUL)

    if variacao_pct > TOLERANCIA_TENDENCIA_SUCESSO_PCT:
        return tendencia_html("▲", COR_VERDE, percentual_br(abs(variacao_pct)))

    if variacao_pct < -TOLERANCIA_TENDENCIA_SUCESSO_PCT:
        return tendencia_html("▼", COR_VERMELHO, percentual_br(abs(variacao_pct)))

    return tendencia_html("●", COR_AZUL)


def tendencia_inconsistencia_html(taxa_60, taxa_dia) -> str:
    if taxa_60 is None or taxa_dia is None:
        return tendencia_html("●", COR_AZUL)

    diferenca_pp = float(taxa_60) - float(taxa_dia)

    if diferenca_pp > TOLERANCIA_TENDENCIA_INCONS_PP:
        return tendencia_html("▲", COR_VERMELHO)

    if diferenca_pp < -TOLERANCIA_TENDENCIA_INCONS_PP:
        return tendencia_html("▼", COR_VERDE)

    return tendencia_html("●", COR_AZUL)


def tendencia_fila_html(tempo_atual, tempo_anterior) -> str:
    if tempo_atual is None or tempo_anterior is None:
        return tendencia_html("●", COR_AZUL)

    infinito_atual = tempo_atual == float("inf")
    infinito_anterior = tempo_anterior == float("inf")

    if infinito_atual and infinito_anterior:
        return tendencia_html("●", COR_AZUL)

    if infinito_atual and not infinito_anterior:
        return tendencia_html("▲", COR_VERMELHO)

    if not infinito_atual and infinito_anterior:
        return tendencia_html("▼", COR_VERDE)

    diferenca = float(tempo_atual) - float(tempo_anterior)

    if diferenca > TOLERANCIA_TENDENCIA_FILA_MIN:
        return tendencia_html("▲", COR_VERMELHO)

    if diferenca < -TOLERANCIA_TENDENCIA_FILA_MIN:
        return tendencia_html("▼", COR_VERDE)

    return tendencia_html("●", COR_AZUL)


def calcular_indicadores_servico(
    df_cards: pd.DataFrame,
    nome_servico: str,
    coluna_sucesso: str,
    coluna_fila: str,
    coluna_inconsistencia: str,
) -> dict:
    """
    Calcula acumulados, janela móvel atual e janela móvel anterior.

    Janela atual: últimos 60 minutos.
    Janela anterior: de 120 a 60 minutos atrás.
    """
    resultado = {
        "servico": nome_servico,
        "sucesso_total": 0,
        "fila_atual": 0,
        "inconsistencia_total": 0,
        "sucesso_60": None,
        "inconsistencia_60": None,
        "total_60": 0,
        "sucesso_60_anterior": None,
        "variacao_sucesso_pct": None,
        "taxa_inconsistencia_60": None,
        "taxa_inconsistencia_dia": None,
        "escoamento_min": None,
        "escoamento_anterior_min": None,
        "cobertura_fila_min": None,
        "demanda_suficiente": False,
    }

    if df_cards is None or df_cards.empty:
        return resultado

    atual = df_cards.iloc[-1]
    momento_atual = pd.Timestamp(atual["_momento"])

    linha_60 = linha_proxima_janela(
        df_cards,
        momento_atual - pd.Timedelta(minutes=60),
    )
    linha_120 = linha_proxima_janela(
        df_cards,
        momento_atual - pd.Timedelta(minutes=120),
    )

    sucesso_atual = valor_numerico_linha(atual, coluna_sucesso)
    fila_atual = valor_numerico_linha(atual, coluna_fila)
    incons_atual = valor_numerico_linha(atual, coluna_inconsistencia)

    resultado["sucesso_total"] = int(sucesso_atual)
    resultado["fila_atual"] = int(fila_atual)
    resultado["inconsistencia_total"] = int(incons_atual)

    total_dia = sucesso_atual + incons_atual
    if total_dia > 0:
        resultado["taxa_inconsistencia_dia"] = (
            incons_atual / total_dia
        ) * 100.0

    if linha_60 is not None:
        sucesso_base_60 = valor_numerico_linha(linha_60, coluna_sucesso)
        incons_base_60 = valor_numerico_linha(linha_60, coluna_inconsistencia)

        # Evita percentual incorreto após eventual reinício dos contadores.
        if sucesso_atual >= sucesso_base_60 and incons_atual >= incons_base_60:
            sucesso_60 = sucesso_atual - sucesso_base_60
            incons_60 = incons_atual - incons_base_60
            total_60 = sucesso_60 + incons_60

            resultado["sucesso_60"] = int(sucesso_60)
            resultado["inconsistencia_60"] = int(incons_60)
            resultado["total_60"] = int(total_60)

            if total_60 > 0:
                resultado["taxa_inconsistencia_60"] = (
                    incons_60 / total_60
                ) * 100.0

            resultado["escoamento_min"] = minutos_escoamento(
                fila_atual,
                sucesso_60,
            )
            resultado["cobertura_fila_min"] = resultado["escoamento_min"]

    if linha_60 is not None and linha_120 is not None:
        sucesso_60_base = valor_numerico_linha(linha_60, coluna_sucesso)
        sucesso_120_base = valor_numerico_linha(linha_120, coluna_sucesso)

        if sucesso_60_base >= sucesso_120_base:
            sucesso_anterior = sucesso_60_base - sucesso_120_base
            resultado["sucesso_60_anterior"] = int(sucesso_anterior)

            if sucesso_anterior > 0 and resultado["sucesso_60"] is not None:
                resultado["variacao_sucesso_pct"] = (
                    (resultado["sucesso_60"] - sucesso_anterior)
                    / sucesso_anterior
                ) * 100.0

            fila_60 = valor_numerico_linha(linha_60, coluna_fila)
            resultado["escoamento_anterior_min"] = minutos_escoamento(
                fila_60,
                sucesso_anterior,
            )

    cobertura = resultado["cobertura_fila_min"]

    if cobertura == float("inf") and fila_atual > 0:
        resultado["demanda_suficiente"] = True
    elif cobertura is not None:
        resultado["demanda_suficiente"] = (
            cobertura >= LIMITE_DEMANDA_SUCESSO_MIN
        )

    return resultado


def cor_card_inconsistencia(metricas: dict) -> str:
    configuracao = LIMITES_INCONSISTENCIA.get(
        metricas.get("servico"),
        LIMITES_INCONSISTENCIA["Transferências"],
    )

    taxa = metricas.get("taxa_inconsistencia_60")
    amostra = int(metricas.get("total_60") or 0)

    if taxa is None or amostra < configuracao["amostra_minima"]:
        return COR_AZUL

    if taxa <= configuracao["verde"]:
        return COR_VERDE

    if taxa <= configuracao["amarelo"]:
        return COR_AMARELO

    return COR_VERMELHO


def cor_card_fila(metricas: dict) -> str:
    tempo = metricas.get("escoamento_min")

    if tempo is None:
        return COR_AZUL

    if tempo == float("inf"):
        return COR_VERMELHO

    if tempo <= LIMITE_FILA_VERDE_MIN:
        return COR_VERDE

    if tempo <= LIMITE_FILA_AMARELO_MIN:
        return COR_AMARELO

    return COR_VERMELHO


def cor_card_sucesso(metricas: dict) -> str:
    """
    A queda de sucesso somente gera alerta quando há fila suficiente.

    Com demanda:
      verde: aumento, estabilidade ou queda de até 10%;
      amarelo: queda acima de 10% até 25%;
      vermelho: queda superior a 25%.
    """
    if not metricas.get("demanda_suficiente"):
        return COR_AZUL

    variacao = metricas.get("variacao_sucesso_pct")

    if variacao is None:
        return COR_AZUL

    if variacao < -25.0:
        return COR_VERMELHO

    if variacao < -10.0:
        return COR_AMARELO

    return COR_VERDE


def nota_card_inconsistencia(metricas: dict) -> str:
    taxa_60 = metricas.get("taxa_inconsistencia_60")
    taxa_dia = metricas.get("taxa_inconsistencia_dia")
    amostra = int(metricas.get("total_60") or 0)
    configuracao = LIMITES_INCONSISTENCIA.get(
        metricas.get("servico"),
        LIMITES_INCONSISTENCIA["Transferências"],
    )

    if taxa_60 is None:
        linha_1 = (
            f'<b>Últimos 60 min:</b> sem dados '
            f'{tendencia_html("●", COR_AZUL)}'
        )
    else:
        linha_1 = (
            f'<b>Últimos 60 min:</b> {percentual_br(taxa_60, 2)} '
            f'{tendencia_inconsistencia_html(taxa_60, taxa_dia)}'
        )

    linha_2 = f'Acumulado do dia: {percentual_br(taxa_dia, 2)}'

    if amostra < configuracao["amostra_minima"]:
        linha_2 += " • amostra insuficiente"

    return f"{linha_1}<br>{linha_2}"


def nota_card_fila(metricas: dict) -> str:
    tempo = metricas.get("escoamento_min")
    tempo_anterior = metricas.get("escoamento_anterior_min")
    seta = tendencia_fila_html(tempo, tempo_anterior)

    if tempo is None:
        linha_1 = f'<b>Escoamento:</b> sem dados {seta}'
    elif tempo == float("inf"):
        linha_1 = f'<b>Escoamento:</b> sem processamento {seta}'
    elif tempo == 0:
        linha_1 = f'<b>Escoamento:</b> sem fila {seta}'
    else:
        linha_1 = (
            f'<b>Escoamento estimado:</b> '
            f'{decimal_br(tempo)} min {seta}'
        )

    if tempo_anterior is None:
        sucesso_60 = metricas.get("sucesso_60")
        linha_2 = (
            "Processados nos últimos 60 min: "
            + (fmt_num(sucesso_60) if sucesso_60 is not None else "-")
        )
    elif tempo_anterior == float("inf"):
        linha_2 = "Há 1 hora: sem processamento"
    elif tempo_anterior == 0:
        linha_2 = "Há 1 hora: sem fila"
    else:
        linha_2 = f"Há 1 hora: {decimal_br(tempo_anterior)} min"

    return f"{linha_1}<br>{linha_2}"


def nota_card_sucesso(metricas: dict) -> str:
    sucesso_60 = metricas.get("sucesso_60")
    sucesso_anterior = metricas.get("sucesso_60_anterior")
    variacao = metricas.get("variacao_sucesso_pct")

    if sucesso_60 is None:
        linha_1 = (
            f'<b>Últimos 60 min:</b> sem dados '
            f'{tendencia_html("●", COR_AZUL)}'
        )
    else:
        linha_1 = (
            f'<b>Últimos 60 min:</b> {fmt_num(sucesso_60)} '
            f'{tendencia_sucesso_html(variacao)}'
        )

    if not metricas.get("demanda_suficiente"):
        linha_2 = (
            f'Demanda insuficiente para avaliação '
            f'{tendencia_html("●", COR_AZUL)}'
        )
    elif sucesso_anterior is None:
        linha_2 = "Hora anterior: sem dados suficientes"
    else:
        linha_2 = f"Hora anterior: {fmt_num(sucesso_anterior)}"

    return f"{linha_1}<br>{linha_2}"

def cor_criticas_minuto(valor):
    valor = int(valor or 0)

    if valor == 0:
        return "#188038"  # verde
    elif 1 <= valor <= 5:
        return "#F9AB00"  # amarelo
    else:
        return "#D93025"  # vermelho

def obter_total_criticas_minuto(df_criticas: pd.DataFrame, tolerancia_minutos: int = 3) -> int:
    if df_criticas is None or df_criticas.empty:
        return 0

    dfc = df_criticas.copy()

    col_data = _coluna_existente(dfc, ["Data/Hora", "Data Hora", "Data"])
    col_total = _coluna_existente(dfc, ["Total críticas no minuto", "Total ciclo atual"])

    if not col_total:
        return 0

    if "Tipo Linha" in dfc.columns:
        dfc = dfc[dfc["Tipo Linha"].astype(str).str.upper().str.strip() == "RESUMO"].copy()

    if "Tipo Histórico" in dfc.columns:
        dfc = dfc[dfc["Tipo Histórico"].astype(str).str.upper().str.contains("CRÍTICA_MINUTO|CRITICA_MINUTO", na=False)].copy()

    if dfc.empty:
        return 0

    if col_data:
        dfc["_data"] = pd.to_datetime(dfc[col_data], dayfirst=True, errors="coerce")
        dfc = dfc.dropna(subset=["_data"]).sort_values("_data")

        if dfc.empty:
            return 0

        ultima_data = dfc.iloc[-1]["_data"]
        minutos = (agora_sao_paulo() - ultima_data).total_seconds() / 60

        if minutos > tolerancia_minutos:
            return 0

    valor = pd.to_numeric(dfc.iloc[-1][col_total], errors="coerce")
    return int(valor) if pd.notna(valor) else 0


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

def _descricao_historico_valida(valor) -> bool:
    """
    Remove cabeçalhos/valores inválidos que entraram no histórico como inconsistência.
    Ex.: DESCRICAO, DESCRIÇÃO, INCONSISTENCIA, etc.
    """
    raw = str(valor or "").strip()

    if not raw:
        return False

    raw_norm = unicodedata.normalize("NFKD", raw)
    raw_norm = "".join(c for c in raw_norm if not unicodedata.combining(c))
    raw_norm = re.sub(r"\s+", " ", raw_norm).upper().strip()

    invalidos = {
        "",
        "NAN",
        "NONE",
        "DESCRICAO",
        "DESCRIÇÃO",
        "DESCRICAO INCONSISTENCIA",
        "DESCRIÇÃO INCONSISTÊNCIA",
        "INCONSISTENCIA",
        "INCONSISTÊNCIA",
    }

    if raw_norm in invalidos:
        return False

    desc_norm = normalizar_inconsistencia(raw)

    if desc_norm in {
        "DESCRICAO",
        "DESCRIÇÃO",
        "DESCRICAO INCONSISTENCIA",
        "DESCRIÇÃO INCONSISTÊNCIA",
        "INCONSISTENCIA",
        "INCONSISTÊNCIA",
    }:
        return False

    return True
    
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

    # Remove linhas de cabeçalho capturadas indevidamente como inconsistência.
    dfh = dfh[dfh[col_inc].apply(_descricao_historico_valida)].copy()

    if dfh.empty:
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
    txt = str(valor or "").lower().strip()

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

    # 0KM
    if any(x in txt for x in [
        "primeiro registro",
        "primeiro registro do veículo",
        "primeiro registro do veiculo",
        "0km",
        "0 km"
    ]):
        return "0KM"

    # PR somente se for exatamente PR
    if txt == "pr":
        return "0KM"

    return None
    
def _servicos_para_chaves(valor: str) -> list[str]:
    """
    Identifica todos os serviços citados no texto.
    Necessário porque a coluna 'Serviços desligados' pode trazer
    0KM, Transferência 2 e Transferência 3 na mesma célula.
    """

    txt = str(valor or "").lower().strip()

    # Remove acentos para padronizar:
    # transferência -> transferencia
    # município -> municipio
    txt_norm = unicodedata.normalize("NFKD", txt)
    txt_norm = "".join(
        c for c in txt_norm
        if not unicodedata.combining(c)
    )

    servicos = []

    # 0KM / Primeiro Registro
    if (
        "primeiro registro" in txt_norm
        or "primeiro" in txt_norm
        or "1-primeiro" in txt_norm
        or "1 - primeiro" in txt_norm
        or "0km" in txt_norm
        or "0 km" in txt_norm
        or txt_norm == "pr"
    ):
        servicos.append("0KM")

    # Transferência de Proprietário
    if (
        "transferencia de proprietario" in txt_norm
        or "transferencia proprietario" in txt_norm
        or "proprietario" in txt_norm
        or "2-transferencia" in txt_norm
        or "2 - transferencia" in txt_norm
        or txt_norm == "tp"
    ):
        servicos.append("Transferência 2")

    # Transferência de Município/Estado
    if (
        "transferencia de municipio" in txt_norm
        or "transferencia municipio" in txt_norm
        or "transferencia de estado" in txt_norm
        or "transferencia estado" in txt_norm
        or "municipio" in txt_norm
        or "estado" in txt_norm
        or "3-transferencia" in txt_norm
        or "3 - transferencia" in txt_norm
        or txt_norm == "tm"
        or txt_norm == "te"
    ):
        servicos.append("Transferência 3")

    # Remove duplicidades preservando a ordem.
    return list(dict.fromkeys(servicos))
    

def _servicos_stop_sim(df_criticas: pd.DataFrame, df_hist: pd.DataFrame) -> set[str]:
    """
    Identifica serviços com STOP PROCESSO ativado = SIM.
    Mantém o STOP dentro da janela configurada em JANELA_STOP_MINUTOS.
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

        # Apenas linhas com STOP = SIM.
        tmp = tmp[
            tmp[col_stop].astype(str).str.upper().str.strip() == "SIM"
        ].copy()

        if tmp.empty:
            continue

        # Data do registro.
        if col_data:
            tmp["_data"] = pd.to_datetime(
                tmp[col_data],
                dayfirst=True,
                errors="coerce"
            )
        else:
            tmp["_data"] = agora_sao_paulo()

        agora = agora_sao_paulo()

        # Considera somente STOPs dentro da janela configurada.
        tmp = tmp[
            tmp["_data"].notna()
            & ((agora - tmp["_data"]).dt.total_seconds() / 60 <= JANELA_STOP_MINUTOS)
        ].copy()

        if tmp.empty:
            continue

        for _, row in tmp.iterrows():

            textos = []

            if col_servico:
                textos.append(str(row.get(col_servico, "")))

            if col_desligados:
                textos.append(str(row.get(col_desligados, "")))

            combinado = " | ".join(textos)

            for chave in _servicos_para_chaves(combinado):
                registros.append({
                    "servico": chave,
                    "data": row["_data"]
                })

    if not registros:
        return set()

    df_reg = pd.DataFrame(registros)

    # Mantém somente o STOP mais recente de cada serviço.
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

        tmp["_data"] = pd.to_datetime(
            tmp[col_data],
            dayfirst=True,
            errors="coerce"
        )

        agora = agora_sao_paulo()

        tmp = tmp[
            tmp["_data"].notna()
            & ((agora - tmp["_data"]).dt.total_seconds() / 60 <= JANELA_STOP_MINUTOS)
        ].copy()

        if tmp.empty:
            continue

        for _, row in tmp.iterrows():
            textos = []

            if col_servico:
                textos.append(str(row.get(col_servico, "")))

            if col_desligados:
                textos.append(str(row.get(col_desligados, "")))

            combinado = " | ".join(textos)

            if chave_servico in _servicos_para_chaves(combinado):
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
    - OFF imediatamente quando houver STOP PROCESSO ativado = SIM para o serviço.
    - Permanece OFF por pelo menos 15 minutos.
    - Após 15 minutos, volta ON somente se houver aumento em Sucesso
      em ciclo posterior ao STOP.
    """

    status = {
        "0KM": "ON",
        "Transferência 2": "ON",
        "Transferência 3": "ON",
    }

    servicos_com_stop = _servicos_stop_sim(df_criticas, df_hist)


    # Primeiro marca OFF pelos STOPs encontrados.
    for servico in servicos_com_stop:
        status[servico] = "OFF"

    if not servicos_com_stop:
        return status

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

    dfm = (
        dfm
        .dropna(subset=["_data_hora_monitor"])
        .sort_values("_data_hora_monitor")
        .copy()
    )

    if dfm.empty:
        return status

    def houve_sucesso_apos_15_min(chave_servico: str, coluna_sucesso: str) -> bool:
        ultimo_stop = _ultimo_stop_servico(df_criticas, df_hist, chave_servico)

        if ultimo_stop is None:
            return False

        limite_retorno = ultimo_stop + pd.Timedelta(minutes=TEMPO_MINIMO_OFF_MINUTOS)

        # Antes dos 15 minutos, mantém OFF obrigatoriamente.
        if agora_sao_paulo() < limite_retorno:
            return False

        if coluna_sucesso not in dfm.columns:
            return False

        # Usa somente ciclos posteriores ao STOP, mas calcula a diferença
        # antes de filtrar os registros após 15 minutos.
        # Assim, o primeiro ciclo pós-15 compara contra o ciclo anterior,
        # inclusive o ciclo executado durante o STOP.
        tmp = dfm[
            dfm["_data_hora_monitor"] > ultimo_stop
        ].copy()

        if len(tmp) < 2:
            return False

        tmp[coluna_sucesso] = pd.to_numeric(
            tmp[coluna_sucesso],
            errors="coerce"
        ).fillna(0)

        tmp["_dif_sucesso"] = tmp[coluna_sucesso].diff().fillna(0)

        tmp_pos_15 = tmp[
            tmp["_data_hora_monitor"] >= limite_retorno
        ].copy()

        if tmp_pos_15.empty:
            return False

        return tmp_pos_15["_dif_sucesso"].gt(0).any()

    # 0KM volta ON se, após 15 minutos, houver aumento em Sucesso 0km.
    if status["0KM"] == "OFF":
        if houve_sucesso_apos_15_min("0KM", "Sucesso 0km"):
            status["0KM"] = "ON"

    # Transferência 2 e 3 usam a coluna agregada Sucesso 2 e 3.
    if status["Transferência 2"] == "OFF":
        if houve_sucesso_apos_15_min("Transferência 2", "Sucesso 2 e 3"):
            status["Transferência 2"] = "ON"

    if status["Transferência 3"] == "OFF":
        if houve_sucesso_apos_15_min("Transferência 3", "Sucesso 2 e 3"):
            status["Transferência 3"] = "ON"

    return status


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

def render_card(label, value, color, note="Último registro", extra_class=""):
    """
    Renderiza card KPI com faixa lateral esquerda, sem faixa inferior.

    A faixa é aplicada como sombra interna horizontal do próprio card.
    Assim não cria fundo colorido embaixo e preserva a curvatura do card.
    """
    cor_borda = color or "#1A73E8"

    st.markdown(
        f"""
        <div class="kpi-card {extra_class}" style="
            box-shadow: inset 7px 0 0 0 {cor_borda}, var(--md-shadow) !important;
            overflow:hidden !important;
        ">
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

    def normalizar_status(status):
        return str(status or "").strip().upper()

    def bolinha(status):
        return "🟢" if normalizar_status(status) == "ON" else "🔴"

    def cor(status):
        return "#188038" if normalizar_status(status) == "ON" else "#D93025"

    todos_status = [status_t2, status_t3, status_0km, status_monitoramento]

    tem_robo_off = any(
        normalizar_status(s) == "OFF"
        for s in todos_status
    )

    cor_borda_robos = "#D93025" if tem_robo_off else "#188038"
    classe_alerta_robos = " kpi-alerta-critico" if tem_robo_off else ""

    html = f"""
    <div class="kpi-card kpi-robos-tall{classe_alerta_robos}" style="
        box-shadow: inset 7px 0 0 0 {cor_borda_robos}, var(--md-shadow) !important;
        overflow:hidden !important;
    ">
        <div class="kpi-label">Robôs</div>
        <div style="display:grid; grid-template-columns:1fr; gap:12px; margin-top:14px; align-items:center; justify-items:center; width:100%;">
            <div style="font-size:16px; font-weight:600; line-height:1.45; white-space:nowrap;">
                {bolinha(status_t2)} <span style="color:{cor(status_t2)}; font-weight:700;">{status_t2}</span> Transferência 2
            </div>
            <div style="font-size:16px; font-weight:600; line-height:1.45; white-space:nowrap;">
                {bolinha(status_t3)} <span style="color:{cor(status_t3)}; font-weight:700;">{status_t3}</span> Transferência 3
            </div>
            <div style="font-size:16px; font-weight:600; line-height:1.45; white-space:nowrap;">
                {bolinha(status_0km)} <span style="color:{cor(status_0km)}; font-weight:700;">{status_0km}</span> 0KM
            </div>
            <div style="font-size:16px; font-weight:600; line-height:1.45; white-space:nowrap;">
                {bolinha(status_monitoramento)} <span style="color:{cor(status_monitoramento)}; font-weight:700;">{status_monitoramento}</span> Monitoramento e-CRV
            </div>
        </div>
    </div>
    """

    st.markdown(html, unsafe_allow_html=True)


def enviar_alerta_robo_ecrv_off(
    robo_monitoramento_online,
    minutos_sem_atualizacao=None
):
    estado = carregar_alerta_ecrv()

    # Se voltou ON, libera novo alerta futuro
    if robo_monitoramento_online:
        if estado.get("alerta_enviado"):
            estado["alerta_enviado"] = False
            salvar_alerta_ecrv(estado)
        return

    # Se já enviou enquanto está OFF, não repete
    if estado.get("alerta_enviado"):
        return

    try:
        msg = (
            "🚨 **Alerta – Robô e-CRV OFF**\n\n"
            "O robô de monitoramento e-CRV está **OFF** no painel.\n"
            "Não foi identificada atualização recente do arquivo principal de monitoramento."
        )

        if minutos_sem_atualizacao is not None:
            msg += f"\n\nTempo sem atualização: **{int(minutos_sem_atualizacao)} minutos**."

        payload = {"text": msg}

        r = requests.post(
            WEBHOOK_TEAMS,
            json=payload,
            timeout=20
        )

        if r.status_code in (200, 202):
            estado["alerta_enviado"] = True
            salvar_alerta_ecrv(estado)

    except Exception:
        pass

def fig_layout(fig, height=520):
    modo_escuro = str(globals().get("tema", "claro")).lower() == "escuro"

    plot_bg = "#071931" if modo_escuro else "#FFFFFF"
    fonte_cor = "#EAF4FF" if modo_escuro else "#202124"
    fonte_muted = "#9DB7D2" if modo_escuro else "#5F6368"
    fonte_family = "Arial" if modo_escuro else "Google Sans, Roboto, Arial"
    grid_cor = "rgba(91, 166, 255, 0.16)" if modo_escuro else "#E0E0E0"

    fig.update_layout(
        height=height,
        autosize=True,
        margin=dict(l=28, r=28, t=125, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=plot_bg,
        font=dict(color=fonte_cor, family=fonte_family),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.06,
            xanchor="left",
            x=0,
            font=dict(size=11, color=fonte_muted),
            bgcolor="rgba(255,255,255,0)",
        ),

        xaxis=dict(
            gridcolor=grid_cor,
            zerolinecolor=grid_cor,
            tickfont=dict(color=fonte_muted),
            tickangle=-45,
            rangeslider=dict(visible=True),
        ),

        yaxis=dict(
            gridcolor=grid_cor,
            zerolinecolor=grid_cor,
            tickfont=dict(color=fonte_muted),
        ),
    )
    return fig



def adicionar_quantidade_processos(df_base: pd.DataFrame) -> pd.DataFrame:
    if df_base is None or df_base.empty:
        return df_base

    df_base = df_base.copy()

    if "Sucesso 2 e 3" in df_base.columns:
        df_base["Quantidade de processos - Transferências"] = (
            pd.to_numeric(df_base["Sucesso 2 e 3"], errors="coerce")
            .fillna(0)
            .diff()
            .fillna(0)
            .clip(lower=0)
            .astype(int)
        )

    if "Sucesso 0km" in df_base.columns:
        df_base["Quantidade de processos - 0KM"] = (
            pd.to_numeric(df_base["Sucesso 0km"], errors="coerce")
            .fillna(0)
            .diff()
            .fillna(0)
            .clip(lower=0)
            .astype(int)
        )

    if "Sucesso TDV" in df_base.columns:
        df_base["Quantidade de processos - TDV"] = (
            pd.to_numeric(df_base["Sucesso TDV"], errors="coerce")
            .fillna(0)
            .diff()
            .fillna(0)
            .clip(lower=0)
            .astype(int)
        )

    return df_base

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
            font=dict(size=16, color="#EAF4FF" if str(globals().get("tema", "claro")).lower() == "escuro" else "#202124"),
            y=0.98,
            x=0.01,
            xanchor="left",
        )
    )

    return fig_layout(fig, height=520)

def line_chart_comparativo_horario(
    df_base,
    df_comp,
    cols,
    title,
    label_base="Data selecionada",
    label_comp="Data comparativa",
):
    """
    Gráfico comparativo por Horário.
    Quando houver data comparativa, mantém somente os horários existentes nos dois dias.
    """
    fig = go.Figure()

    if df_base is None or df_base.empty:
        return fig_layout(fig, height=520)

    if "Horário" not in df_base.columns:
        return line_chart(df_base, cols, title)

    df_base = df_base.copy()
    df_base["Horário"] = df_base["Horário"].astype(str).str.slice(0, 5)

    tem_comparativo = df_comp is not None and not df_comp.empty and "Horário" in df_comp.columns

    if tem_comparativo:
        df_comp = df_comp.copy()
        df_comp["Horário"] = df_comp["Horário"].astype(str).str.slice(0, 5)

        colunas_comuns = [
            c for c in cols
            if c in df_base.columns and c in df_comp.columns
        ]

        if not colunas_comuns:
            tem_comparativo = False
            df_plot = df_base.copy()
        else:
            df_plot = pd.merge(
                df_base[["Horário"] + colunas_comuns],
                df_comp[["Horário"] + colunas_comuns],
                on="Horário",
                how="inner",
                suffixes=("_base", "_comp"),
            )

            if df_plot.empty:
                fig.update_layout(
                    title=dict(
                        text=f"{title} - sem horários coincidentes para comparação",
                        font=dict(size=16, color="#EAF4FF" if str(globals().get("tema", "claro")).lower() == "escuro" else "#202124"),
                        y=0.98,
                        x=0.01,
                        xanchor="left",
                    )
                )
                return fig_layout(fig, height=520)
    else:
        df_plot = df_base.copy()

    x_labels = df_plot["Horário"].astype(str).tolist()
    x_vals = list(range(len(x_labels)))

    def cor_coluna(col):
        nome_lower = str(col).lower()
        if "sucesso" in nome_lower:
            return "#188038"
        if "incons" in nome_lower:
            return "#D93025"
        if "fila" in nome_lower:
            return "#F9AB00"
        return "#1A73E8"

    y_max = 0

    for col in cols:
        cor = cor_coluna(col)
        col_base = f"{col}_base" if tem_comparativo else col

        if col_base not in df_plot.columns:
            continue

        y_base = pd.to_numeric(df_plot[col_base], errors="coerce").fillna(0).astype(int)
        if not y_base.empty:
            y_max = max(y_max, float(y_base.max()))

        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=y_base,
                mode="lines+markers+text",
                text=[fmt_num(v) for v in y_base],
                textposition="top center",
                textfont=dict(size=10, color=cor),
                name=f"{col} - {label_base}",
                line=dict(width=2.6, color=cor),
                marker=dict(size=5, color=cor),
                customdata=x_labels,
                hovertemplate="%{customdata}<br>" + f"{col} - {label_base}: " + "%{y}<extra></extra>",
                cliponaxis=False,
            )
        )

        col_comp = f"{col}_comp"
        if tem_comparativo and col_comp in df_plot.columns:
            y_comp = pd.to_numeric(df_plot[col_comp], errors="coerce").fillna(0).astype(int)
            if not y_comp.empty:
                y_max = max(y_max, float(y_comp.max()))

            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=y_comp,
                    mode="lines+markers+text",
                    text=[fmt_num(v) for v in y_comp],
                    textposition="bottom center",
                    textfont=dict(size=10, color=cor),
                    name=f"{col} - {label_comp}",
                    line=dict(width=2.4, color=cor, dash="dash"),
                    marker=dict(size=5, color=cor, symbol="circle-open"),
                    customdata=x_labels,
                    hovertemplate="%{customdata}<br>" + f"{col} - {label_comp}: " + "%{y}<extra></extra>",
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

    if y_max > 0:
        fig.update_yaxes(range=[0, y_max * 1.28])
    else:
        fig.update_yaxes(range=[0, 1])

    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=16, color="#EAF4FF" if str(globals().get("tema", "claro")).lower() == "escuro" else "#202124"),
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

def filtrar_historico_por_dia(df):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    col_data = _coluna_existente(
        df,
        ["Data/Hora", "Data Hora", "Data"]
    )

    if not col_data:
        return df

    df["_data"] = pd.to_datetime(
        df[col_data],
        dayfirst=True,
        errors="coerce",
        format="mixed",
    )

    df = df[df["_data"].notna()].copy()

    if df.empty:
        return pd.DataFrame()

    hoje = agora_sao_paulo().date()

    return (
        df[df["_data"].dt.date == hoje]
        .drop(columns=["_data"], errors="ignore")
        .copy()
    )

def _preparar_historico_full_ultimo_ciclo(df_hist: pd.DataFrame) -> pd.DataFrame:
    dfh = _df_historico_full(df_hist)
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

def houve_aumento_inconsistencias_no_ultimo_ciclo(df_monitor: pd.DataFrame) -> bool:
    """
    Verifica se houve aumento real de inconsistências no último ciclo do Monitoramento.xlsx.

    Evita exibir no Top 5 inconsistências antigas gravadas no histórico,
    quando o total atual do painel não aumentou no último ciclo.
    """
    if df_monitor is None or df_monitor.empty or len(df_monitor) < 2:
        return False

    atual = df_monitor.iloc[-1]
    anterior = df_monitor.iloc[-2]

    colunas_incons = [
        "Inconsistência 2 e 3",
        "Inconsistência 0km",
        "Inconsistencia TDV",
    ]

    for col in colunas_incons:
        if col not in df_monitor.columns:
            continue

        valor_atual = pd.to_numeric(atual.get(col, 0), errors="coerce")
        valor_anterior = pd.to_numeric(anterior.get(col, 0), errors="coerce")

        valor_atual = 0 if pd.isna(valor_atual) else int(valor_atual)
        valor_anterior = 0 if pd.isna(valor_anterior) else int(valor_anterior)

        if valor_atual > valor_anterior:
            return True

    return False
    
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

    # Mantém somente o último valor da inconsistência por serviço
    out = out.sort_values(
        ["_data_sort", "_ordem"],
        ascending=[True, True]
    )
    
    out = out.drop_duplicates(
        subset=["_chave_servico", "Descrição"],
        keep="last"
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

    # Total do serviço (0KM ou Transferências)
    total_servico = int(base["Total ciclo atual"].sum())
    
    return (
        base[
            ["Descrição", "Total ciclo atual", "%"]
        ]
        .rename(
            columns={
                "Total ciclo atual": f"Total ({fmt_num(total_servico)})",
                "%": "% (100%)"
            }
        )
    )


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


# Logs são carregados sob demanda na página Logs, conforme a data selecionada.

# Evita horários duplicados após normalização do RPA.
df = agrupar_monitoramento_por_horario(df)
df = adicionar_quantidade_processos(df)


# =========================
# DADOS PRINCIPAIS
# =========================

ultima = df.iloc[-1]
hora_coleta = str(ultima["Horário"]) if "Horário" in df.columns else "-"
ultima_modificacao = agora_sao_paulo().strftime("%d/%m/%Y %H:%M:%S")

# Considera o robô de monitoramento OFF quando o arquivo principal
# fica mais de 15 minutos sem nova atualização no GitHub/cache.
ultima_sync_github = meta_monitor.get("downloaded_at", "") if isinstance(meta_monitor, dict) else ""
robo_monitoramento_online = True
minutos_sem_atualizacao = None

try:
    if ultima_sync_github:
        dt_sync = datetime.strptime(ultima_sync_github, "%d/%m/%Y %H:%M:%S")
        minutos_sem_atualizacao = (datetime_sao_paulo() - dt_sync).total_seconds() / 60

        if minutos_sem_atualizacao > 15:
            robo_monitoramento_online = False
    else:
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

fila_tdv = int(ultima.get("Fila TDV", 0))
sucesso_tdv = int(ultima.get("Sucesso TDV", 0))
incons_tdv = int(ultima.get("Inconsistencia TDV", 0))


# Base temporal utilizada pelos cards dinâmicos.
df_cards = preparar_dados_cards(df)

metricas_trf = calcular_indicadores_servico(
    df_cards,
    "Transferências",
    "Sucesso 2 e 3",
    "Fila 2 e 3",
    "Inconsistência 2 e 3",
)

metricas_0km = calcular_indicadores_servico(
    df_cards,
    "0KM",
    "Sucesso 0km",
    "Fila 0km",
    "Inconsistência 0km",
)

metricas_tdv = calcular_indicadores_servico(
    df_cards,
    "TDV",
    "Sucesso TDV",
    "Fila TDV",
    "Inconsistencia TDV",
)

tdv_hora = 0
try:
    if "Horário" in df.columns and "Quantidade de processos - TDV" in df.columns:
        hora_ref = str(hora_coleta)[:2]
        tdv_hora = int(
            df[df["Horário"].astype(str).str.slice(0, 2) == hora_ref]["Quantidade de processos - TDV"]
            .sum()
        )
except Exception:
    tdv_hora = 0

total_sucesso = sucesso_trf + sucesso_0km + sucesso_tdv
total_criticas_minuto = obter_total_criticas_minuto(df_criticas)

robos = status_robos(df, df_criticas, df_hist)

status_ecrv_rpa, meta_status_ecrv = carregar_json_github(
    GITHUB_ARQ_STATUS_ECRV,
    ARQ_LOCAL_STATUS_ECRV,
    obrigatorio=False,
)

comando_ecrv_atual, meta_comando_ecrv = carregar_json_github(
    GITHUB_ARQ_CONTROLE_ECRV,
    ARQ_LOCAL_CONTROLE_ECRV,
    obrigatorio=False,
)

# Reconhece status confirmado pelo RPA e, depois, aplica eventual controle
# visual temporário feito no próprio Dashboard.
robos, robo_monitoramento_online = aplicar_status_confirmado_rpa(
    robos,
    robo_monitoramento_online,
    status_ecrv_rpa,
    df_criticas,
    df_hist,
)

robos, robo_monitoramento_online = aplicar_controle_dashboard_manual(
    robos,
    robo_monitoramento_online,
    df_criticas,
    df_hist,
)

status_painel_robos = status_card_robos(
    robos,
    robo_monitoramento_online,
)

enviar_alerta_robo_ecrv_off(
    robo_monitoramento_online,
    minutos_sem_atualizacao
)


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

    cols = st.columns(
        5,
        vertical_alignment="top"
    )

    with cols[0]:
        render_card(
            "Fila Transferências",
            fila_trf,
            cor_card_fila(metricas_trf),
            nota_card_fila(metricas_trf),
        )
        render_card(
            "Fila 0KM",
            fila_0km,
            cor_card_fila(metricas_0km),
            nota_card_fila(metricas_0km),
        )
        render_card(
            "Fila TDV",
            fila_tdv,
            cor_card_fila(metricas_tdv),
            nota_card_fila(metricas_tdv),
        )

    with cols[1]:
        render_card(
            "Sucesso Transferências",
            sucesso_trf,
            cor_card_sucesso(metricas_trf),
            nota_card_sucesso(metricas_trf),
        )
        render_card(
            "Sucesso 0KM",
            sucesso_0km,
            cor_card_sucesso(metricas_0km),
            nota_card_sucesso(metricas_0km),
        )
        render_card(
            "Sucesso TDV",
            sucesso_tdv,
            cor_card_sucesso(metricas_tdv),
            nota_card_sucesso(metricas_tdv),
        )

    with cols[2]:
        render_card(
            "Inconsistências Transferências",
            incons_trf,
            cor_card_inconsistencia(metricas_trf),
            nota_card_inconsistencia(metricas_trf),
        )
        render_card(
            "Inconsistências 0KM",
            incons_0km,
            cor_card_inconsistencia(metricas_0km),
            nota_card_inconsistencia(metricas_0km),
        )
        render_card(
            "Inconsistências TDV",
            incons_tdv,
            cor_card_inconsistencia(metricas_tdv),
            nota_card_inconsistencia(metricas_tdv),
        )

    with cols[3]:
        render_card(
            "Automatizado e-CRV",
            automatizado,
            cor_saude(automatizado, media_coluna(df_media, "Automatizado", hora_coleta), "positivo"),
            f"Último registro: {hora_coleta}",
        )
        render_card(
            "Total",
            total_sucesso,
            "#1A73E8",
            "Soma das Transferências, 0KM e TDV"
        )
        render_card(
            "TDV (Hora)",
            tdv_hora,
            "#1A73E8",
            "Emissões TDV na hora atual",
        )

    with cols[4]:
        render_card(
            "Inconsistências Críticas",
            total_criticas_minuto,
            cor_criticas_minuto(total_criticas_minuto),
            "Total de críticas do minuto atual",
            extra_class=(
                "kpi-tall kpi-alerta-critico"
                if int(total_criticas_minuto or 0) > 0
                else "kpi-tall"
            ),
        )
        render_robos_card(robos, robo_monitoramento_online)


    # =========================
    # GRÁFICOS
    # =========================

    c1, c2, c3 = st.columns(3)

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
                ["Quantidade de processos - Transferências"],
                "Emissões Transferências",
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
                ["Quantidade de processos - 0KM"],
                "Emissões 0KM",
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

    with c3:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)

        st.plotly_chart(
            line_chart(
                df,
                ["Sucesso TDV"],
                "TDV - Sucesso",
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
                ["Quantidade de processos - TDV"],
                "Emissões TDV",
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
                ["Fila TDV", "Inconsistencia TDV"],
                "TDV - Fila e Inconsistências",
            ),
            use_container_width=True,
        )

        st.markdown('</div></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)


    # =========================
    # TABELAS DE HISTÓRICO E INCONSISTÊNCIAS
    # =========================

    df_hist_hoje = filtrar_historico_por_dia(df_hist)
    
    
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Histórico de Inconsistências Críticas</div>', unsafe_allow_html=True)

    if df_hist is None or df_hist.empty:
        st.info("Histórico de críticas ainda não disponível.")
    else:
        df_crit_minuto = tabela_historico_criticas_minuto(df_hist_hoje)

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

    if houve_aumento_inconsistencias_no_ultimo_ciclo(df):
        df_top3 = tabela_top3_ciclo_atual(df_hist_hoje)
    else:
        df_top3 = pd.DataFrame(columns=["Serviço", "Descrição", "Quantidade no ciclo"])
    
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
        df_serv = tabela_historico_servico(df_hist_hoje, chave_servico)
        if df_serv.empty:
            render_mensagem_tabela("Sem histórico de inconsistências registrado para este serviço.")
        else:
            render_tabela_escura(df_serv)

    st.markdown("</div>", unsafe_allow_html=True)






elif pagina == "Ligar/Desligar":
    render_controle_robos(
        status_painel_robos,
        status_ecrv_rpa,
        comando_ecrv_atual,
    )


elif pagina == "Logs":
    st.markdown(
        '''
<div class="hive-title">Logs do Monitoramento e-CRV</div>
<div class="hive-subtitle">
    Acompanhe os registros operacionais do monitoramento e-CRV na data selecionada.
</div>
''',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        '<div class="panel-title">Selecionar data do log</div>',
        unsafe_allow_html=True,
    )

    datas_logs_disponiveis = listar_datas_logs_disponiveis()
    
    if not datas_logs_disponiveis:
        st.markdown(
            '<div class="log-empty">Nenhum arquivo de log encontrado no GitHub.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()
    
    hoje = agora_sao_paulo().date()
    
    try:
        indice_padrao_log = datas_logs_disponiveis.index(hoje)
    except ValueError:
        indice_padrao_log = 0
    
    data_log_selecionada = st.selectbox(
        "Data do log",
        options=datas_logs_disponiveis,
        index=indice_padrao_log,
        format_func=lambda d: d.strftime("%d/%m/%Y"),
        key="logs_data_selecionada",
    )

    st.markdown("</div>", unsafe_allow_html=True)

    df_log_dia = pd.DataFrame()
    meta_log_dia = {}

    try:
        df_log_dia, meta_log_dia = carregar_log_diario_dashboard(data_log_selecionada)
    except Exception:
        df_log_dia = pd.DataFrame()
        meta_log_dia = {"path": caminho_github_log_data(data_log_selecionada)}

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        '<div class="panel-title">Log diário do robô</div>',
        unsafe_allow_html=True,
    )

    render_logs_dashboard(df_log_dia, meta_log_dia)

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
        df_mon_hist = carregar_historico_multimes(
            "Monitoramento",
            obrigatorio=True,
        )
    
    except Exception as e:
        st.error(f"Erro ao carregar histórico de monitoramento: {e}")
        st.stop()

    try:
        df_inc_hist = carregar_historico_multimes(
            "Inconsistencias",
            obrigatorio=False,
        )
    
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

    col_data_1, col_data_2 = st.columns(2)

    with col_data_1:
        data_selecionada = st.date_input(
            "Selecionar data",
            value=datas_disponiveis[0],
            min_value=min(datas_disponiveis),
            max_value=max(datas_disponiveis),
            format="DD/MM/YYYY",
            key="data_historico_principal",
        )

    data_ref = pd.to_datetime(data_selecionada, errors="coerce").date()

    datas_comparacao = [
        d for d in datas_disponiveis
        if d != data_ref
    ]

    data_comparativa_padrao = datas_comparacao[0] if datas_comparacao else data_ref

    with col_data_2:
        data_comparativa = st.date_input(
            "Comparar com a data",
            value=data_comparativa_padrao,
            min_value=min(datas_disponiveis),
            max_value=max(datas_disponiveis),
            format="DD/MM/YYYY",
            key="data_historico_comparativa",
        )

    data_ref_comp = pd.to_datetime(data_comparativa, errors="coerce").date()
    comparar_datas = data_ref_comp != data_ref

    df_dia = df_mon_hist[df_mon_hist["Data"] == data_ref].copy()
    df_dia = agrupar_monitoramento_por_horario(df_dia)

    if comparar_datas:
        df_dia_comp = df_mon_hist[df_mon_hist["Data"] == data_ref_comp].copy()
        df_dia_comp = agrupar_monitoramento_por_horario(df_dia_comp)

        if df_dia_comp.empty:
            st.warning(
                f"Não há registros de monitoramento para comparação em {data_ref_comp.strftime('%d/%m/%Y')}."
            )
    else:
        df_dia_comp = pd.DataFrame()

    if df_inc_hist is not None and not df_inc_hist.empty and "Data" in df_inc_hist.columns:
        df_hist_dia = df_inc_hist[df_inc_hist["Data"] == data_ref].copy()
    else:
        df_hist_dia = pd.DataFrame()

    df_dia = adicionar_quantidade_processos(df_dia)

    if comparar_datas and df_dia_comp is not None and not df_dia_comp.empty:
        df_dia_comp = adicionar_quantidade_processos(df_dia_comp)

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Tabela de monitoramento da data selecionada</div>', unsafe_allow_html=True)

    if df_dia.empty:
        render_mensagem_tabela("Não há registros de monitoramento para a data selecionada.")
    else:
        df_dia = adicionar_quantidade_processos(df_dia)
        if comparar_datas and df_dia_comp is not None and not df_dia_comp.empty:
            df_dia_comp = adicionar_quantidade_processos(df_dia_comp)

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
        graficos_historico = [
            ("Transferências - Sucesso", ["Sucesso 2 e 3"]),
            ("Transferências - Emissões Transferências", ["Quantidade de processos - Transferências"]),
            ("Transferências - Fila", ["Fila 2 e 3"]),
            ("Transferências - Inconsistências", ["Inconsistência 2 e 3"]),
            ("0KM - Sucesso", ["Sucesso 0km"]),
            ("0KM - Emissões 0KM", ["Quantidade de processos - 0KM"]),
            ("0KM - Fila", ["Fila 0km"]),
            ("0KM - Inconsistências", ["Inconsistência 0km"]),
            ("TDV - Sucesso", ["Sucesso TDV"]),
            ("TDV - Emissões TDV", ["Quantidade de processos - TDV"]),
            ("TDV - Fila", ["Fila TDV"]),
            ("TDV - Inconsistências", ["Inconsistencia TDV"]),
        ]

        for titulo_grafico, colunas_grafico in graficos_historico:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown('<div class="chart-scroll"><div class="chart-inner">', unsafe_allow_html=True)
            st.plotly_chart(
                line_chart_comparativo_horario(
                    df_dia,
                    df_dia_comp if comparar_datas else None,
                    colunas_grafico,
                    titulo_grafico,
                    data_ref.strftime("%d/%m/%Y"),
                    data_ref_comp.strftime("%d/%m/%Y") if comparar_datas else "",
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

if pagina != "Ligar/Desligar":
    time.sleep(INTERVALO_VERIFICACAO_SEGUNDOS)
    st.rerun()
