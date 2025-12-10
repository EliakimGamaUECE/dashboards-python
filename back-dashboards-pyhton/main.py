from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import plotly.io as pio
from pydantic import BaseModel
from typing import List

from dash import Dash, dcc, html, Input, Output
from starlette.middleware.wsgi import WSGIMiddleware


app = FastAPI()

origins = [
    "http://localhost:4200",  # futuro Angular
    "http://127.0.0.1:4200",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).parent


@app.get("/health")
def health():
    return {"status": "ok", "message": "Laboratório de gráficos Plotly ativo!"}


@app.get("/dashboards/grafico1_producao_diaria.html", response_class=HTMLResponse)
def grafico1_producao_diaria():
    file = BASE_DIR / "producao_diaria.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo producao_diaria.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")
    df["data"] = pd.to_datetime(df["data"])

    fig = px.line(
        df,
        x="data",
        y="total_produzido",
        title="Produção diária da fábrica",
        labels={"data": "Data", "total_produzido": "Produção (peças)"},
    )

    fig.update_traces(mode="lines+markers", hovertemplate="Data: %{x|%d/%m/%Y}<br>Produção: %{y} peças<extra></extra>")

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)


@app.get("/dashboards/grafico2_producao_por_linha.html", response_class=HTMLResponse)
def grafico2_producao_por_linha():
    file = BASE_DIR / "producao_por_linha.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo producao_por_linha.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")

    fig = px.bar(
        df,
        x="linha",
        y="producao",
        title="Produção total por linha",
        labels={"linha": "Linha de produção", "producao": "Produção (peças)"},
    )

    fig.update_traces(hovertemplate="Linha: %{x}<br>Produção: %{y} peças<extra></extra>")

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)


@app.get("/dashboards/grafico3_refugo_por_linha.html", response_class=HTMLResponse)
def grafico3_refugo_por_linha():
    file = BASE_DIR / "refugo_por_linha.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo refugo_por_linha.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df["linha"],
        y=df["ok"],
        name="OK",
        hovertemplate="Linha: %{x}<br>OK: %{y} peças<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        x=df["linha"],
        y=df["refugo"],
        name="Refugo",
        hovertemplate="Linha: %{x}<br>Refugo: %{y} peças<extra></extra>",
    ))

    fig.update_layout(
        barmode="stack",
        title="Produção OK x Refugo por linha",
        xaxis_title="Linha",
        yaxis_title="Quantidade (peças)",
        legend_title="Status",
    )

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)

@app.get("/dashboards/producao_por_linha_interativo.html", response_class=HTMLResponse)
def producao_por_linha_interativo():
    try:
        file = BASE_DIR / "producao_diaria_por_linha.xlsx"

        if not file.exists():
            return HTMLResponse(
                "<h1>Arquivo producao_diaria_por_linha.xlsx não encontrado na pasta do main.py</h1>",
                status_code=500
            )

        df = pd.read_excel(file, engine="openpyxl")

        # garante tipos corretos
        df["data"] = pd.to_datetime(df["data"])
        df["linha"] = df["linha"].astype(str)

        linhas = df["linha"].unique().tolist()
        fig = go.Figure()

        # um trace por linha
        for linha in linhas:
            df_linha = df[df["linha"] == linha]
            fig.add_trace(go.Scatter(
                x=df_linha["data"],
                y=df_linha["producao"],
                mode="lines+markers",
                name=linha,
                visible=True if linha == linhas[0] else False,  # só a primeira visível
            ))

        # dropdown para alternar linha
        buttons = []
        for i, linha in enumerate(linhas):
            visible = [False] * len(linhas)
            visible[i] = True
            buttons.append(dict(
                label=linha,
                method="update",
                args=[{"visible": visible},
                      {"title": f"Produção diária - {linha}"}],
            ))

        fig.update_layout(
            title=f"Produção diária - {linhas[0]}",
            xaxis_title="Data",
            yaxis_title="Produção (peças)",
            updatemenus=[dict(
                buttons=buttons,
                direction="down",
                showactive=True,
                x=0.0,
                xanchor="left",
                y=1.15,
                yanchor="top"
            )]
        )

        html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
        return HTMLResponse(content=html)

    except Exception as e:
        # se der algum erro, mostra o motivo na tela pra ficar mais fácil de debugar
        return HTMLResponse(
            f"<h1>Erro ao gerar gráfico interativo</h1><pre>{e!r}</pre>",
            status_code=500
        )




@app.get("/dashboards/grafico4_refugo_por_causa.html", response_class=HTMLResponse)
def grafico4_refugo_por_causa():
    file = BASE_DIR / "refugo_por_causa.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo refugo_por_causa.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")

    fig = px.pie(
        df,
        names="causa",
        values="quantidade",
        title="Distribuição do refugo por causa",
        hole=0.4,  # donut
    )

    fig.update_traces(hovertemplate="Causa: %{label}<br>Quantidade: %{value} peças<br>%: %{percent}<extra></extra>")

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)

#NÃO TA FUNCIONANDO
@app.get("/dashboards/grafico5_correlacao_temp_defeitos.html", response_class=HTMLResponse)
def grafico5_correlacao_temp_defeitos():
    file = BASE_DIR / "correlacao_temp_defeitos.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo correlacao_temp_defeitos.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")

    fig = px.scatter(
        df,
        x="temperatura_forno",
        y="defeitos",
        title="Correlação entre temperatura do forno e defeitos",
        labels={"temperatura_forno": "Temperatura do forno (°C)", "defeitos": "Número de defeitos"},
        trendline="ols",  # linha de tendência (se quiser)
    )

    fig.update_traces(hovertemplate="Temp: %{x} °C<br>Defeitos: %{y}<extra></extra>")

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)

@app.get("/dashboards/grafico6_eficiencia_turno.html", response_class=HTMLResponse)
def grafico6_eficiencia_turno():
    file = BASE_DIR / "eficiencia_turno.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo eficiencia_turno.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")

    # Pivot: linhas = linha, colunas = turno, valores = eficiência
    tabela = df.pivot(index="linha", columns="turno", values="eficiencia")

    fig = px.imshow(
        tabela,
        text_auto=True,
        aspect="auto",
        title="Eficiência por linha e turno",
        labels=dict(x="Turno", y="Linha", color="Eficiência"),
        zmin=0.8,  # ajuste conforme seus dados
        zmax=1.0,
    )

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)

@app.get("/dashboards/grafico7_oee_velocimetro.html", response_class=HTMLResponse)
def grafico7_oee_velocimetro():
    file = BASE_DIR / "oee_velocimetro.xlsx"
    if not file.exists():
        return HTMLResponse("<h1>Arquivo oee_velocimetro.xlsx não encontrado</h1>", status_code=500)

    df = pd.read_excel(file, engine="openpyxl")

    # pega o primeiro valor
    valor = float(df["valor"].iloc[0])  # assume 0-1
    percentual = valor * 100

    # Configura faixas de cor (verde, amarelo, vermelho, por exemplo)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=percentual,
        number={"suffix": "%"},
        title={"text": "OEE - Eficiência Global do Equipamento"},
        gauge={
            "axis": {"range": [0, 100]},
            "bar": {"color": "darkblue"},
            "steps": [
                {"range": [0, 60], "color": "#ff4d4f"},    # vermelho
                {"range": [60, 80], "color": "#faad14"},   # amarelo
                {"range": [80, 100], "color": "#52c41a"},  # verde
            ],
            "threshold": {
                "line": {"color": "black", "width": 4},
                "thickness": 0.75,
                "value": percentual,
            }
        }
    ))

    fig.update_layout(
        height=400,
        margin=dict(t=80, b=20, l=20, r=20)
    )

    html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
    return HTMLResponse(content=html)

@app.get("/dashboards/grafico8_teep_utilizacao.html", response_class=HTMLResponse)
def grafico8_teep_utilizacao():
    try:
        file = BASE_DIR / "teep.xlsx"

        if not file.exists():
            return HTMLResponse(
                "<h1>Arquivo teep.xlsx não encontrado</h1>",
                status_code=500
            )

        df = pd.read_excel(file, engine="openpyxl")

        # Normaliza os nomes das colunas
        df.columns = df.columns.str.strip().str.lower()

        # Espera essas colunas (sem acento mesmo: utilizacao)
        esperadas = ["data", "linha", "utilizacao", "teep"]
        faltando = [c for c in esperadas if c not in df.columns]
        if faltando:
            raise ValueError(
                f"Colunas faltando no Excel TEEP: {faltando}. "
                f"Colunas encontradas: {list(df.columns)}"
            )

        # Renomeia para a forma “bonita” que o gráfico usa
        df = df.rename(columns={
            "data": "Data",
            "linha": "Linha",
            "utilizacao": "Utilizacao",
            "teep": "TEEP",
        })

        df["Data"] = pd.to_datetime(df["Data"])
        df["Linha"] = df["Linha"].astype(str)

        # =========================
        #  AGRUPA POR MÊS (PLANTA)
        # =========================
        df_mes = (
            df.assign(ano_mes=df["Data"].dt.to_period("M"))
              .groupby("ano_mes")[["Utilizacao", "TEEP"]]
              .mean()
              .reset_index()
        )

        if df_mes.empty:
            return HTMLResponse("<h1>Sem dados de TEEP/Utilização para plotar</h1>", status_code=200)

        df_mes["mes_label"] = df_mes["ano_mes"].dt.strftime("%b/%Y")
        df_mes["util_pct"] = df_mes["Utilizacao"] * 100
        df_mes["teep_pct"] = df_mes["TEEP"] * 100

        meta_teep = 0.75   # 75% de meta – ajuste se quiser
        meta_pct = meta_teep * 100

        # =========================
        #  FIGURA: BARRAS + LINHAS
        # =========================
        fig = go.Figure()

        # Barras de Utilização
        fig.add_bar(
            x=df_mes["mes_label"],
            y=df_mes["util_pct"],
            name="Utilização",
        )

        # Barras de TEEP
        fig.add_bar(
            x=df_mes["mes_label"],
            y=df_mes["teep_pct"],
            name="TEEP",
        )

        # Linha de tendência do TEEP
        fig.add_scatter(
            x=df_mes["mes_label"],
            y=df_mes["teep_pct"],
            mode="lines+markers",
            name="Tendência TEEP",
        )

        # Linha horizontal de meta TEEP
        fig.add_scatter(
            x=df_mes["mes_label"],
            y=[meta_pct] * len(df_mes),
            mode="lines",
            name=f"Target {meta_pct:.0f}%",
            line=dict(dash="dash"),
        )

        fig.update_layout(
            title="TEEP Mensal IM - Utilização x TEEP",
            xaxis_title="Mês",
            yaxis_title="Percentual (%)",
            barmode="group",  # barras lado a lado
            template="plotly_white",
            margin=dict(l=40, r=10, t=50, b=40),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            height=360,
        )

        fig.update_yaxes(range=[0, 100])

        html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
        return HTMLResponse(content=html)

    except Exception as e:
        return HTMLResponse(
            f"<h1>Erro ao gerar gráfico de TEEP</h1><pre>{e!r}</pre>",
            status_code=500
        )



@app.get("/dashboards/grafico9_oee_diario.html", response_class=HTMLResponse)
def grafico9_oee_diario():
    try:
        file = BASE_DIR / "oee.xlsx"

        if not file.exists():
            return HTMLResponse(
                "<h1>Arquivo oee.xlsx não encontrado</h1>",
                status_code=500
            )

        df = pd.read_excel(file, engine="openpyxl")

        # Normaliza nomes de colunas
        df.columns = df.columns.str.strip().str.lower()

        esperadas = ["data", "linha", "disponibilidade", "performance", "qualidade", "oee"]
        faltando = [c for c in esperadas if c not in df.columns]
        if faltando:
            raise ValueError(
                f"Colunas faltando no Excel OEE: {faltando}. "
                f"Colunas encontradas: {list(df.columns)}"
            )

        df = df.rename(columns={
            "data": "Data",
            "linha": "Linha",
            "disponibilidade": "Disponibilidade",
            "performance": "Performance",
            "qualidade": "Qualidade",
            "oee": "OEE",
        })

        df["Data"] = pd.to_datetime(df["Data"])
        df["Linha"] = df["Linha"].astype(str)

        # =========================
        #  AGRUPA POR MÊS (PLANTA)
        # =========================
        df_mes = (
            df.assign(ano_mes=df["Data"].dt.to_period("M"))
              .groupby("ano_mes")["OEE"]
              .mean()
              .reset_index()
        )

        if df_mes.empty:
            return HTMLResponse("<h1>Sem dados de OEE para plotar</h1>", status_code=200)

        df_mes["mes_label"] = df_mes["ano_mes"].dt.strftime("%b/%Y")
        df_mes["oee_pct"] = df_mes["OEE"] * 100

        meta_oee = 0.75  # 75% de meta
        meta_pct = meta_oee * 100

        # =========================
        #  FIGURA: BARRAS + LINHAS
        # =========================
        fig = go.Figure()

        # Barras de OEE (%) por mês
        fig.add_bar(
            x=df_mes["mes_label"],
            y=df_mes["oee_pct"],
            name="OEE",
        )

        # Linha de tendência (mesmos valores das barras)
        fig.add_scatter(
            x=df_mes["mes_label"],
            y=df_mes["oee_pct"],
            mode="lines+markers",
            name="Tendência",
        )

        # Linha horizontal de meta (pontilhada)
        fig.add_scatter(
            x=df_mes["mes_label"],
            y=[meta_pct] * len(df_mes),
            mode="lines",
            name=f"Target {meta_pct:.0f}%",
            line=dict(dash="dash"),
        )

        fig.update_layout(
            title="IM OEE Mensal da Planta",
            xaxis_title="Mês",
            yaxis_title="OEE (%)",
            template="plotly_white",
            margin=dict(l=40, r=10, t=50, b=40),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            height=360,
        )

        fig.update_yaxes(range=[0, 100])

        html = pio.to_html(fig, full_html=True, include_plotlyjs="cdn")
        return HTMLResponse(content=html)

    except Exception as e:
        return HTMLResponse(
            f"<h1>Erro ao gerar gráfico de OEE</h1><pre>{e!r}</pre>",
            status_code=500
        )


class KpiItem(BaseModel):
    title: str
    value: str        # ex: "81,2%"
    meta: str         # ex: "Meta: ≥ 77%"
    delta_label: str  # ex: "Acima da meta" / "Abaixo da meta"
    delta_value: str  # ex: "+1,2%" / "-0,4%"
    is_positive: bool

class KpiResponse(BaseModel):
    kpis: List[KpiItem]


@app.get("/api/kpis-resumo", response_model=KpiResponse)
def kpis_resumo():
    """
    Lê oee.xlsx e teep.xlsx e monta 4 KPIs:
    - OEE SMT
    - OEE IM
    - TEEP SMT
    - TEEP IM

    Mapeamento:
    - pega as linhas existentes no Excel (Linha 1, Linha 2, ...)
    - a primeira é usada como SMT
    - a segunda (se existir) é usada como IM
    - se só tiver uma, SMT e IM recebem o mesmo valor
    """
    meta = 0.77  # 77%

    # ---------- OEE ----------
    df_oee = pd.read_excel(BASE_DIR / "oee.xlsx", engine="openpyxl")
    df_oee.columns = df_oee.columns.str.strip().str.lower()
    df_oee = df_oee.rename(columns={
        "data": "Data",
        "linha": "Linha",
        "oee": "OEE",
    })
    df_oee["Data"] = pd.to_datetime(df_oee["Data"])
    df_oee["Linha"] = df_oee["Linha"].astype(str)

    # último valor de OEE por linha
    ultimos_oee = df_oee.sort_values("Data").groupby("Linha").tail(1)
    linhas_oee = sorted(ultimos_oee["Linha"].unique().tolist())

    if not linhas_oee:
        # fallback hardcore: nada na planilha, tudo 0
        oee_smt_val = 0.0
        oee_im_val = 0.0
    else:
        # primeira linha = SMT
        linha_smt = linhas_oee[0]
        oee_smt_val = float(
            ultimos_oee[ultimos_oee["Linha"] == linha_smt]["OEE"].iloc[0]
        )
        # segunda linha = IM (se existir), senão copia SMT
        if len(linhas_oee) > 1:
            linha_im = linhas_oee[1]
            oee_im_val = float(
                ultimos_oee[ultimos_oee["Linha"] == linha_im]["OEE"].iloc[0]
            )
        else:
            oee_im_val = oee_smt_val

    # ---------- TEEP ----------
    df_teep = pd.read_excel(BASE_DIR / "teep.xlsx", engine="openpyxl")
    df_teep.columns = df_teep.columns.str.strip().str.lower()
    df_teep = df_teep.rename(columns={
        "data": "Data",
        "linha": "Linha",
        "teep": "TEEP",
    })
    df_teep["Data"] = pd.to_datetime(df_teep["Data"])
    df_teep["Linha"] = df_teep["Linha"].astype(str)

    ultimos_teep = df_teep.sort_values("Data").groupby("Linha").tail(1)
    linhas_teep = sorted(ultimos_teep["Linha"].unique().tolist())

    if not linhas_teep:
        teep_smt_val = 0.0
        teep_im_val = 0.0
    else:
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

    # ---------- formatação ----------
    def formata_kpi(title: str, valor_0a1: float) -> KpiItem:
        meta_0a1 = meta
        valor_pct = valor_0a1 * 100
        meta_pct = meta_0a1 * 100
        delta_pct = valor_pct - meta_pct

        is_positive = valor_0a1 >= meta_0a1
        delta_label = "Acima da meta" if is_positive else "Abaixo da meta"

        value_str = f"{valor_pct:.1f}".replace(".", ",") + "%"
        meta_str = f"Meta: ≥ {meta_pct:.0f}%"
        delta_str = f"{delta_pct:+.1f}".replace(".", ",") + "%"

        return KpiItem(
            title=title,
            value=value_str,
            meta=meta_str,
            delta_label=delta_label,
            delta_value=delta_str,
            is_positive=is_positive,
        )

    kpis = [
        formata_kpi("OEE SMT", oee_smt_val),
        formata_kpi("OEE IM", oee_im_val),
        formata_kpi("TEEP SMT", teep_smt_val),
        formata_kpi("TEEP IM", teep_im_val),
    ]

    return KpiResponse(kpis=kpis)

# =========================
# HELPERS PARA APP DASH
# =========================

def get_linhas_disponiveis():
    """
    Lê oee.xlsx e devolve a lista de linhas existentes (como strings).
    Usado para popular o dropdown no Dash.
    """
    try:
        file = BASE_DIR / "oee.xlsx"
        if not file.exists():
            return []

        df = pd.read_excel(file, engine="openpyxl")
        df.columns = df.columns.str.strip().str.lower()
        df = df.rename(columns={"linha": "Linha"})
        df["Linha"] = df["Linha"].astype(str)
        return sorted(df["Linha"].unique().tolist())
    except Exception:
        return []


def create_fig_oee_mensal(linha=None):
    """
    Cria a figura do gráfico 'IM OEE Mensal da Planta'.
    Se linha for None -> considera todas as linhas (planta).
    Se linha for algo como 'Linha 1' -> filtra antes de agregar.
    """
    file = BASE_DIR / "oee.xlsx"
    fig = go.Figure()

    if not file.exists():
        fig.add_annotation(
            text="Arquivo oee.xlsx não encontrado",
            showarrow=False,
            x=0.5, y=0.5, xref="paper", yref="paper",
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
            x=0.5, y=0.5, xref="paper", yref="paper",
        )
        return fig

    df = df.rename(columns={
        "data": "Data",
        "linha": "Linha",
        "disponibilidade": "Disponibilidade",
        "performance": "Performance",
        "qualidade": "Qualidade",
        "oee": "OEE",
    })

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
            x=0.5, y=0.5, xref="paper", yref="paper",
        )
        return fig

    df_mes["mes_label"] = df_mes["ano_mes"].dt.strftime("%b/%Y")
    df_mes["oee_pct"] = df_mes["OEE"] * 100

    meta_oee = 0.75
    meta_pct = meta_oee * 100

    fig = go.Figure()
    fig.add_bar(x=df_mes["mes_label"], y=df_mes["oee_pct"], name="OEE")
    fig.add_scatter(
        x=df_mes["mes_label"], y=df_mes["oee_pct"],
        mode="lines+markers", name="Tendência",
    )
    fig.add_scatter(
        x=df_mes["mes_label"], y=[meta_pct] * len(df_mes),
        mode="lines", name=f"Target {meta_pct:.0f}%",
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


def create_fig_teep_mensal(linha=None):
    """
    Cria a figura 'TEEP Mensal IM - Utilização x TEEP'.
    Se linha for None -> planta.
    Se linha for 'Linha 1' etc -> filtra antes.
    """
    file = BASE_DIR / "teep.xlsx"
    fig = go.Figure()

    if not file.exists():
        fig.add_annotation(
            text="Arquivo teep.xlsx não encontrado",
            showarrow=False,
            x=0.5, y=0.5, xref="paper", yref="paper",
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
            x=0.5, y=0.5, xref="paper", yref="paper",
        )
        return fig

    df = df.rename(columns={
        "data": "Data",
        "linha": "Linha",
        "utilizacao": "Utilizacao",
        "teep": "TEEP",
    })

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
            x=0.5, y=0.5, xref="paper", yref="paper",
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
        x=df_mes["mes_label"], y=df_mes["teep_pct"],
        mode="lines+markers", name="Tendência TEEP",
    )
    fig.add_scatter(
        x=df_mes["mes_label"], y=[meta_pct] * len(df_mes),
        mode="lines", name=f"Target {meta_pct:.0f}%",
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


# =========================
# APP DASH INTEGRADO AO FASTAPI
# =========================

dash_app = Dash(
    __name__,
    requests_pathname_prefix="/dash/",  # importante pro FastAPI + mount
)

# opções do dropdown (lidas da planilha)
_linhas = get_linhas_disponiveis()
linha_options = [{"label": "Todas as linhas", "value": "ALL"}] + [
    {"label": linha, "value": linha} for linha in _linhas
]

dash_app.layout = html.Div(
    style={
        "fontFamily": "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        "backgroundColor": "#f3f4f6",
        "minHeight": "100vh",
        "padding": "16px 24px",
    },
    children=[
        html.H1("Painel de KPIs da Fábrica", style={"marginBottom": "4px"}),
        html.P("OEE e TEEP · Nível 3 · Resumo Mensal", style={"color": "#6b7280"}),

        html.Div(style={"height": "8px"}),

        # Filtro de linha
        html.Div(
            style={
                "backgroundColor": "white",
                "borderRadius": "12px",
                "padding": "8px 12px",
                "marginBottom": "12px",
                "boxShadow": "0 1px 3px rgba(15,23,42,0.08)",
                "display": "flex",
                "alignItems": "center",
                "gap": "12px",
            },
            children=[
                html.Span("Linha:", style={"fontSize": "0.85rem", "fontWeight": 600}),
                dcc.Dropdown(
                    id="linha-dropdown",
                    options=linha_options,
                    value="ALL",
                    clearable=False,
                    style={"width": "220px", "fontSize": "0.85rem"},
                ),
            ],
        ),

        # Gráficos
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr", "gap": "16px"},
            children=[
                html.Div(
                    style={
                        "backgroundColor": "white",
                        "borderRadius": "16px",
                        "padding": "12px 14px 16px",
                        "boxShadow": "0 1px 4px rgba(15, 23, 42, 0.08)",
                    },
                    children=[
                        dcc.Graph(
                            id="oee-graph",
                            figure=create_fig_oee_mensal(),
                            style={"height": "360px"},
                        ),
                    ],
                ),
                html.Div(
                    style={
                        "backgroundColor": "white",
                        "borderRadius": "16px",
                        "padding": "12px 14px 16px",
                        "boxShadow": "0 1px 4px rgba(15, 23, 42, 0.08)",
                    },
                    children=[
                        dcc.Graph(
                            id="teep-graph",
                            figure=create_fig_teep_mensal(),
                            style={"height": "360px"},
                        ),
                    ],
                ),
            ],
        ),
    ],
)

# callback: atualiza os dois gráficos quando muda a linha
@dash_app.callback(
    [Output("oee-graph", "figure"), Output("teep-graph", "figure")],
    Input("linha-dropdown", "value"),
)
def atualizar_graficos(linha_value):
    if linha_value in (None, "ALL"):
        linha = None
    else:
        linha = linha_value
    return create_fig_oee_mensal(linha), create_fig_teep_mensal(linha)


# monta o app Dash dentro do FastAPI em /dash
app.mount("/dash", WSGIMiddleware(dash_app.server))
