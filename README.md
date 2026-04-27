# 🌿 Fresh Excel

App para limpiar archivos Excel con FastAPI + pandas + openpyxl.

## ¿Qué limpia?

| Problema | Solución |
|---|---|
| Fechas en múltiples formatos (dd/mm/yyyy, yyyy-mm-dd, etc.) | Normaliza todas a `YYYY-MM-DD` |
| Celdas combinadas (merged cells) | Las separa y rellena con el valor original |
| Caracteres rotos / de control | Elimina caracteres no imprimibles y normaliza unicode |

## Requisitos

- Python 3.11+
- [uv](https://docs.astral.sh/uv/) instalado

## Instalación y uso

```bash
# 1. Clonar / entrar al proyecto
cd excel-cleaner

# 2. Instalar dependencias con uv
uv sync

# 3. Iniciar el servidor
uv run uvicorn main:app --reload --port 8000
```

Luego abrir **http://localhost:8000** en el navegador.

## Endpoints API

| Método | Ruta | Descripción |
|---|---|---|
| `POST` | `/api/analyze` | Analiza el archivo y reporta problemas (sin modificarlo) |
| `POST` | `/api/clean` | Limpia y devuelve el archivo procesado |

### Parámetros de `/api/clean`

| Campo | Tipo | Default | Descripción |
|---|---|---|---|
| `file` | File | — | Archivo `.xlsx` o `.xls` |
| `fix_dates` | bool | `true` | Normalizar fechas inconsistentes |
| `fix_broken_chars` | bool | `true` | Eliminar caracteres rotos |
| `unmerge_cells` | bool | `true` | Separar celdas combinadas |
| `output_mode` | string | `"new"` | `"new"` = archivo nuevo, `"overwrite"` = mismo nombre |

## Estructura del proyecto

```
excel-cleaner/
├── main.py                    # Entry point FastAPI
├── pyproject.toml             # Dependencias (uv)
├── static/
│   └── index.html             # Frontend
├── app/
│   ├── routers/
│   │   └── cleaner.py         # Endpoints /api/analyze y /api/clean
│   └── services/
│       └── cleaner_service.py # Lógica de limpieza con pandas + openpyxl
└── temp/                      # Archivos temporales (auto-generado)
```
