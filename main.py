from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from pathlib import Path

from app.routers import cleaner

app = FastAPI(
    title="Fresh Excel API",
    description="Limpia y normaliza archivos Excel con formatos inconsistentes",
    version="1.0.0",
)

app.include_router(cleaner.router, prefix="/api")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = Path("static/index.html")
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))
