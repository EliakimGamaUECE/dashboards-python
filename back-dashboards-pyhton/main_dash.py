from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
from starlette.middleware.wsgi import WSGIMiddleware
from dashboard_oee_teep import create_dash_app

app = FastAPI()

origins = [
    "http://localhost:4200",
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
    return {"status": "ok", "message": "Laboratório de gráficos Plotly + Dash ativo!"}


# cria a instância do Dash e monta em /dash
dash_app = create_dash_app(BASE_DIR)
app.mount("/dash", WSGIMiddleware(dash_app.server))
