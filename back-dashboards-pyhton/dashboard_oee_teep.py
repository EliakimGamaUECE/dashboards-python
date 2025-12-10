from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output


# =========================
# HELPERS DE DADOS / FIGURAS
# =========================

def get_linhas_disponiveis(base_dir: Path):
    """
    Lê oee.xlsx e devolve a lista de linhas existentes (como strings).
    Usado para popular o dropdown no Dash.
    """
    try:
        file = base_dir / "oee.xlsx"
        if not file.exists():
            return []

        df = pd.read_excel(file, engine="openpyxl")
        df.columns = df.columns.str.strip().str.lower()
        df = df.rename(columns={"linha": "Linha"})
        df["Linha"] = df["Linha"].astype(str)
        return sorted(df["Linha"].unique().tolist())
    except Exception:
        return []


def create_fig_oee_mensal(base_dir: Path, linha=None):
    """
    Cria a figura do gráfico 'IM OEE Mensal da Planta'.
    Se linha for None -> considera todas as linhas (planta).
    Se linha for algo como 'Linha 1' -> filtra antes de agregar.
    """
    file = base_dir / "oee.xlsx"
    fig = go.Figure()

    if not file.exists():
        fig.add_annotation(
            text="Arquivo oee.xlsx não encontrado",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
        )
        return fig

    df = pd.read_excel(file, engine="openpyxl")
    df.columns = df.columns.str.strip().str.lower()

    esperadas = ["data", "linha", "disponibilidade", "performance", "qualidade", "oee"]
    faltando = [c for c in esperadas if c not in df.columns]
    if faltando:
        fig.add_annotation(
            text=f"Colunas faltando no Excel OEE: {faltando}",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
        )
        return fig

    df = df.rename(
        columns={
            "data": "Data",
            "linha": "Linha",
            "disponibilidade": "Disponibilidade",
            "performance": "Performance",
            "qualidade": "Qualidade",
            "oee": "OEE",
        }
    )

    df["Data"] = pd.to_datetime(df["Data"])
    df["Linha"] = df["Linha"].astype(str)

    # filtro opcional por linha
    if linha is not None:
        df = df[df["Linha"] == linha]

    df_mes = (
        df.assign(ano_mes=df["Data"].dt.to_period("M"))
        .groupby("ano_mes")["OEE"]
        .mean()
        .reset_index()
    )

    if df_mes.empty:
        fig.add_annotation(
            text="Sem dados de OEE para plotar",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
        )
        return fig

    df_mes["mes_label"] = df_mes["ano_mes"].dt.strftime("%b/%Y")
    df_mes["oee_pct"] = df_mes["OEE"] * 100

    meta_oee = 0.75
    meta_pct = meta_oee * 100

    fig = go.Figure()
    fig.add_bar(x=df_mes["mes_label"], y=df_mes["oee_pct"], name="OEE")
    fig.add_scatter(
        x=df_mes["mes_label"],
        y=df_mes["oee_pct"],
        mode="lines+markers",
        name="Tendência",
    )
    fig.add_scatter(
        x=df_mes["mes_label"],
        y=[meta_pct] * len(df_mes),
        mode="lines",
        name=f"Target {meta_pct:.0f}%",
        line=dict(dash="dash"),
    )

    titulo = "IM OEE Mensal da Planta" if linha is None else f"OEE Mensal - {linha}"
    fig.update_layout(
        title=titulo,
        xaxis_title="Mês",
        yaxis_title="OEE (%)",
        template="plotly_white",
        margin=dict(l=40, r=10, t=50, b=40),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
        height=360,
    )
    fig.update_yaxes(range=[0, 100])
    return fig


def create_fig_teep_mensal(base_dir: Path, linha=None):
    """
    Cria a figura 'TEEP Mensal IM - Utilização x TEEP'.
    Se linha for None -> planta.
    Se linha for 'Linha 1' etc -> filtra antes.
    """
    file = base_dir / "teep.xlsx"
    fig = go.Figure()

    if not file.exists():
        fig.add_annotation(
            text="Arquivo teep.xlsx não encontrado",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
        )
        return fig

    df = pd.read_excel(file, engine="openpyxl")
    df.columns = df.columns.str.strip().str.lower()

    esperadas = ["data", "linha", "utilizacao", "teep"]
    faltando = [c for c in esperadas if c not in df.columns]
    if faltando:
        fig.add_annotation(
            text=f"Colunas faltando no Excel TEEP: {faltando}",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
        )
        return fig

    df = df.rename(
        columns={
            "data": "Data",
            "linha": "Linha",
            "utilizacao": "Utilizacao",
            "teep": "TEEP",
        }
    )

    df["Data"] = pd.to_datetime(df["Data"])
    df["Linha"] = df["Linha"].astype(str)

    # filtro opcional por linha
    if linha is not None:
        df = df[df["Linha"] == linha]

    df_mes = (
        df.assign(ano_mes=df["Data"].dt.to_period("M"))
        .groupby("ano_mes")[["Utilizacao", "TEEP"]]
        .mean()
        .reset_index()
    )

    if df_mes.empty:
        fig.add_annotation(
            text="Sem dados de TEEP/Utilização para plotar",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
        )
        return fig

    df_mes["mes_label"] = df_mes["ano_mes"].dt.strftime("%b/%Y")
    df_mes["util_pct"] = df_mes["Utilizacao"] * 100
    df_mes["teep_pct"] = df_mes["TEEP"] * 100

    meta_teep = 0.75
    meta_pct = meta_teep * 100

    fig = go.Figure()
    fig.add_bar(x=df_mes["mes_label"], y=df_mes["util_pct"], name="Utilização")
    fig.add_bar(x=df_mes["mes_label"], y=df_mes["teep_pct"], name="TEEP")
    fig.add_scatter(
        x=df_mes["mes_label"],
        y=df_mes["teep_pct"],
        mode="lines+markers",
        name="Tendência TEEP",
    )
    fig.add_scatter(
        x=df_mes["mes_label"],
        y=[meta_pct] * len(df_mes),
        mode="lines",
        name=f"Target {meta_pct:.0f}%",
        line=dict(dash="dash"),
    )

    titulo = "TEEP Mensal IM - Utilização x TEEP" if linha is None else f"TEEP Mensal - {linha}"
    fig.update_layout(
        title=titulo,
        xaxis_title="Mês",
        yaxis_title="Percentual (%)",
        barmode="group",
        template="plotly_white",
        margin=dict(l=40, r=10, t=50, b=40),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
        height=360,
    )
    fig.update_yaxes(range=[0, 100])
    return fig


def compute_kpis(base_dir: Path):
    """
    Lê oee.xlsx e teep.xlsx e monta 4 KPIs:
    - OEE SMT
    - OEE IM
    - TEEP SMT
    - TEEP IM
    """
    meta = 0.77  # 77%

    def formata_kpi(title, valor_0a1):
        meta_0a1 = meta
        valor_pct = valor_0a1 * 100
        meta_pct = meta_0a1 * 100
        delta_pct = valor_pct - meta_pct

        is_positive = valor_0a1 >= meta_0a1
        delta_label = "Acima da meta" if is_positive else "Abaixo da meta"

        value_str = f"{valor_pct:.1f}".replace(".", ",") + "%"
        meta_str = f"Meta: ≥ {meta_pct:.0f}%"
        delta_str = f"{delta_pct:+.1f}".replace(".", ",") + "%"

        return {
            "title": title,
            "value": value_str,
            "meta": meta_str,
            "delta_label": delta_label,
            "delta_value": delta_str,
            "is_positive": is_positive,
        }

    # valores default caso dê erro
    oee_smt_val = oee_im_val = teep_smt_val = teep_im_val = 0.0

    try:
        # ---------- OEE ----------
        df_oee = pd.read_excel(base_dir / "oee.xlsx", engine="openpyxl")
        df_oee.columns = df_oee.columns.str.strip().str.lower()
        df_oee = df_oee.rename(columns={"data": "Data", "linha": "Linha", "oee": "OEE"})
        df_oee["Data"] = pd.to_datetime(df_oee["Data"])
        df_oee["Linha"] = df_oee["Linha"].astype(str)

        ultimos_oee = df_oee.sort_values("Data").groupby("Linha").tail(1)
        linhas_oee = sorted(ultimos_oee["Linha"].unique().tolist())

        if linhas_oee:
            linha_smt = linhas_oee[0]
            oee_smt_val = float(
                ultimos_oee[ultimos_oee["Linha"] == linha_smt]["OEE"].iloc[0]
            )
            if len(linhas_oee) > 1:
                linha_im = linhas_oee[1]
                oee_im_val = float(
                    ultimos_oee[ultimos_oee["Linha"] == linha_im]["OEE"].iloc[0]
                )
            else:
                oee_im_val = oee_smt_val

        # ---------- TEEP ----------
        df_teep = pd.read_excel(base_dir / "teep.xlsx", engine="openpyxl")
        df_teep.columns = df_teep.columns.str.strip().str.lower()
        df_teep = df_teep.rename(columns={"data": "Data", "linha": "Linha", "teep": "TEEP"})
        df_teep["Data"] = pd.to_datetime(df_teep["Data"])
        df_teep["Linha"] = df_teep["Linha"].astype(str)

        ultimos_teep = df_teep.sort_values("Data").groupby("Linha").tail(1)
        linhas_teep = sorted(ultimos_teep["Linha"].unique().tolist())

        if linhas_teep:
            linha_smt_t = linhas_teep[0]
            teep_smt_val = float(
                ultimos_teep[ultimos_teep["Linha"] == linha_smt_t]["TEEP"].iloc[0]
            )
            if len(linhas_teep) > 1:
                linha_im_t = linhas_teep[1]
                teep_im_val = float(
                    ultimos_teep[ultimos_teep["Linha"] == linha_im_t]["TEEP"].iloc[0]
                )
            else:
                teep_im_val = teep_smt_val

    except Exception:
        # se der ruim, ficam todos 0 mesmo
        pass

    return [
        formata_kpi("OEE SMT", oee_smt_val),
        formata_kpi("OEE IM", oee_im_val),
        formata_kpi("TEEP SMT", teep_smt_val),
        formata_kpi("TEEP IM", teep_im_val),
    ]


def make_kpi_card(kpi):
    card_class = (
        "kpi-card kpi-card-positive"
        if kpi["is_positive"]
        else "kpi-card kpi-card-negative"
    )

    return html.Div(
        className=card_class,
        children=[
            html.Div(
                children=html.H3(
                    kpi["title"],
                    style={"margin": 0, "fontSize": "0.9rem", "fontWeight": 600},
                )
            ),
            html.Div(
                style={"marginTop": "4px", "marginBottom": "6px"},
                children=html.Span(
                    kpi["value"],
                    style={"fontSize": "1.3rem", "fontWeight": 700},
                ),
            ),
            html.Div(
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "baseline",
                    "fontSize": "0.75rem",
                },
                children=[
                    html.Span(kpi["meta"], className="kpi-meta"),
                    html.Span(
                        [
                            html.Span(kpi["delta_label"] + ": "),
                            html.Span(kpi["delta_value"], className="kpi-delta-value"),
                        ]
                    ),
                ],
            ),
        ],
    )



# =========================
# CRIAÇÃO DO APP DASH
# =========================

def create_dash_app(base_dir: Path) -> Dash:
    dash_app = Dash(
        __name__,
        requests_pathname_prefix="/dash/",
    )

    linhas = get_linhas_disponiveis(base_dir)
    linha_options = [{"label": "Todas as linhas", "value": "ALL"}] + [
        {"label": linha, "value": linha} for linha in linhas
    ]

    kpis_data = compute_kpis(base_dir)
    kpi_cards = [make_kpi_card(k) for k in kpis_data]

    dash_app.layout = html.Div(
        className="dashboard-page",
        children=[
            html.H1("Painel de KPIs da Fábrica", className="page-title"),
            html.P(
                "OEE e TEEP · Nível 3 · Resumo Mensal",
                className="dashboard-subtitle",
            ),

            html.Div(className="spacer-sm"),
            
            # === MENU DE ABA (fake só pra visual) ===
        html.Div(
            className="top-menu",
            children=[
                html.Div("Todos os KPIs", className="top-menu-tab top-menu-tab--active"),
                html.Div("Produtividade", className="top-menu-tab"),
                html.Div("Qualidade", className="top-menu-tab"),
                html.Div("Pessoas", className="top-menu-tab"),
                html.Div("Logística", className="top-menu-tab"),
                html.Div("Controladoria", className="top-menu-tab"),
            ],
        ),

            # 🔹 FILTROS EM UMA BARRA HORIZONTAL (CARD CHEIO)
            html.Div(
                className="card card--small-radius card-padding-sm filters-card",
                children=[
                    html.Div(
                        className="card-header",
                        children=[
                            html.Span("Filtros", className="card-title"),
                            html.Span(
                                "Fake · ajustar depois",
                                className="card-badge",
                            ),
                        ],
                    ),
                    html.Div(
                        className="filters-row",
                        children=[
                            # Filtro de linha
                            html.Div(
                                className="filters-field",
                                children=[
                                    html.Label("Linha", className="filters-label"),
                                    dcc.Dropdown(
                                        id="linha-dropdown",
                                        options=linha_options,
                                        value="ALL",
                                        clearable=False,
                                        style={"fontSize": "0.85rem"},
                                    ),
                                ],
                            ),
                            # Filtro fake de período
                            html.Div(
                                className="filters-field",
                                children=[
                                    html.Label(
                                        "Período (fake)", className="filters-label"
                                    ),
                                    dcc.Dropdown(
                                        id="periodo-dropdown",
                                        options=[
                                            {
                                                "label": "Últimos 3 meses",
                                                "value": "3M",
                                            },
                                            {
                                                "label": "Últimos 6 meses",
                                                "value": "6M",
                                            },
                                            {
                                                "label": "Últimos 12 meses",
                                                "value": "12M",
                                            },
                                        ],
                                        value="6M",
                                        clearable=False,
                                        style={"fontSize": "0.85rem"},
                                    ),
                                ],
                            ),
                            # Filtro fake de turno
                            html.Div(
                                className="filters-field",
                                children=[
                                    html.Label("Turno (fake)", className="filters-label"),
                                    dcc.Dropdown(
                                        id="turno-dropdown",
                                        options=[
                                            {"label": "Todos", "value": "ALL"},
                                            {"label": "Turno A", "value": "A"},
                                            {"label": "Turno B", "value": "B"},
                                            {"label": "Turno C", "value": "C"},
                                        ],
                                        value="ALL",
                                        clearable=False,
                                        style={"fontSize": "0.85rem"},
                                    ),
                                ],
                            ),
                        ],
                    ),
                ],
            ),

            html.Div(className="spacer-sm"),

            # 🔹 GRID PRINCIPAL: ESQUERDA KPIs / DIREITA GRÁFICOS
            html.Div(
                className="dashboard-grid",
                children=[
                    # COLUNA ESQUERDA: KPIs
                    html.Div(
                        children=[
                            html.Div(
                                className="card card-padding",
                                children=[
                                    html.Div(
                                        className="card-header",
                                        children=[
                                            html.Span(
                                                "Resumo OEE / TEEP",
                                                className="card-title",
                                            ),
                                            html.Span(
                                                "KPIs calculados em Python",
                                                className="card-badge",
                                            ),
                                        ],
                                    ),
                                    html.Div(
                                        className="kpis-grid",
                                        children=kpi_cards,
                                    ),
                                ],
                            ),
                        ]
                    ),

                    # COLUNA DIREITA: GRÁFICOS
                    html.Div(
                        className="charts-column",
                        children=[
                            html.Div(
                                className="card card-padding",
                                children=[
                                    dcc.Graph(
                                        id="oee-graph",
                                        figure=create_fig_oee_mensal(base_dir),
                                        style={"height": "360px"},
                                    ),
                                ],
                            ),
                            html.Div(
                                className="card card-padding",
                                children=[
                                    dcc.Graph(
                                        id="teep-graph",
                                        figure=create_fig_teep_mensal(base_dir),
                                        style={"height": "360px"},
                                    ),
                                ],
                            ),
                        ],
                    ),
                ],
            ),
        ],
    )


    @dash_app.callback(
        [Output("oee-graph", "figure"), Output("teep-graph", "figure")],
        Input("linha-dropdown", "value"),
    )
    def atualizar_graficos(linha_value):
        if linha_value in (None, "ALL"):
            linha = None
        else:
            linha = linha_value
        return (
            create_fig_oee_mensal(base_dir, linha),
            create_fig_teep_mensal(base_dir, linha),
        )

    return dash_app
