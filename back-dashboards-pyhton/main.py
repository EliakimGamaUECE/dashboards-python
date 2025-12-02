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

        # Normaliza os nomes das colunas: tira espaços e deixa minúsculo
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

        df_melt = df.melt(
            id_vars=["Data", "Linha"],
            value_vars=["Utilizacao", "TEEP"],
            var_name="Indicador",
            value_name="valor"
        )

        fig = px.line(
            df_melt,
            x="Data",
            y="valor",
            color="Indicador",
            line_group="Linha",
            title="Utilização x TEEP por dia",
            labels={
                "Data": "Data",
                "valor": "Valor",
                "Indicador": "Indicador"
            }
        )

        fig.update_yaxes(title_text="Percentual", tickformat=".0%")

        fig.update_traces(
            mode="lines+markers",
            hovertemplate="Data: %{x|%d/%m/%Y}<br>%{legendgroup}: %{y:.1%}<extra></extra>"
        )

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

        df_melt = df.melt(
            id_vars=["Data", "Linha"],
            value_vars=["Disponibilidade", "Performance", "Qualidade", "OEE"],
            var_name="Indicador",
            value_name="valor",
        )

        fig = px.line(
            df_melt,
            x="Data",
            y="valor",
            color="Indicador",
            line_group="Linha",
            title="Disponibilidade, Performance, Qualidade e OEE - Diário",
            labels={
                "Data": "Data",
                "valor": "Percentual",
                "Indicador": "Indicador",
            },
        )

        fig.update_yaxes(title_text="Percentual", tickformat=".0%")

        fig.update_traces(
            mode="lines+markers",
            hovertemplate="Data: %{x|%d/%m/%Y}<br>%{legendgroup}: %{y:.1%}<extra></extra>"
        )

        # Linha de meta como trace (compatível com qualquer versão do Plotly)
        meta_oee = 0.75  # 75% de meta – ajusta se quiser
        fig.add_trace(go.Scatter(
            x=[df["Data"].min(), df["Data"].max()],
            y=[meta_oee, meta_oee],
            mode="lines",
            name=f"Meta OEE {meta_oee:.0%}",
            line=dict(dash="dash")
        ))

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
    Lê oee.xlsx e teep.xlsx e monta os 4 KPIs:
    - OEE SMT
    - OEE IM
    - TEEP SMT
    - TEEP IM
    """
    meta = 0.77  # 77%

    # ---- OEE ----
    df_oee = pd.read_excel(BASE_DIR / "oee.xlsx", engine="openpyxl")
    df_oee.columns = df_oee.columns.str.strip().str.lower()
    df_oee = df_oee.rename(columns={
        "data": "Data",
        "linha": "Linha",
        "oee": "OEE",
    })
    df_oee["Data"] = pd.to_datetime(df_oee["Data"])

    # pega o último valor por linha (por data)
    ultimos_oee = df_oee.sort_values("Data").groupby("Linha").tail(1)

    def pega_oee(linha_nome: str) -> float:
        row = ultimos_oee[ultimos_oee["Linha"] == linha_nome]
        if row.empty:
            return 0.0
        return float(row["OEE"].iloc[0])  # assume 0–1

    oee_smt = pega_oee("SMT")
    oee_im = pega_oee("IM")

    # ---- TEEP ----
    df_teep = pd.read_excel(BASE_DIR / "teep.xlsx", engine="openpyxl")
    df_teep.columns = df_teep.columns.str.strip().str.lower()
    df_teep = df_teep.rename(columns={
        "data": "Data",
        "linha": "Linha",
        "teep": "TEEP",
    })
    df_teep["Data"] = pd.to_datetime(df_teep["Data"])
    ultimos_teep = df_teep.sort_values("Data").groupby("Linha").tail(1)

    def pega_teep(linha_nome: str) -> float:
        row = ultimos_teep[ultimos_teep["Linha"] == linha_nome]
        if row.empty:
            return 0.0
        return float(row["TEEP"].iloc[0])  # assume 0–1

    teep_smt = pega_teep("SMT")
    teep_im = pega_teep("IM")

    def formata_kpi(title: str, valor_0a1: float) -> KpiItem:
        meta_0a1 = meta
        valor_pct = valor_0a1 * 100          # ex: 0.812 → 81.2
        meta_pct = meta_0a1 * 100           # 77
        delta_pct = valor_pct - meta_pct    # diferença em pontos percentuais

        is_positive = valor_0a1 >= meta_0a1
        delta_label = "Acima da meta" if is_positive else "Abaixo da meta"

        # strings já prontas para o front (com vírgula)
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
        formata_kpi("OEE SMT", oee_smt),
        formata_kpi("OEE IM", oee_im),
        formata_kpi("TEEP SMT", teep_smt),
        formata_kpi("TEEP IM", teep_im),
    ]

    return KpiResponse(kpis=kpis)