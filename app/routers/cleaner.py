import shutil
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

from app.services.cleaner_service import analyze_file, clean_file

router = APIRouter(tags=["cleaner"])

TEMP_DIR = Path("temp")
TEMP_DIR.mkdir(exist_ok=True)


def _save_upload(upload: UploadFile) -> Path:
    suffix = Path(upload.filename).suffix
    dest = TEMP_DIR / f"{uuid.uuid4()}{suffix}"
    with dest.open("wb") as f:
        shutil.copyfileobj(upload.file, f)
    return dest


@router.post("/analyze")
async def analyze(file: UploadFile = File(...)):
    """Preview what issues will be fixed before cleaning."""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos .xlsx o .xls")

    path = _save_upload(file)
    try:
        result = analyze_file(path)
    finally:
        path.unlink(missing_ok=True)

    return result


@router.post("/clean")
async def clean(
    file: UploadFile = File(...),
    fix_dates: bool = Form(True),
    fix_broken_chars: bool = Form(True),
    unmerge_cells: bool = Form(True),
    output_mode: str = Form("new"),
    custom_filename: str = Form(""),
):
    """
    Clean the Excel file and return it with the correct filename.
    - custom_filename provided  -> use that name
    - output_mode='overwrite'   -> keep original filename
    - output_mode='new'         -> original name + '_limpio' suffix
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos .xlsx o .xls")

    original_name = Path(file.filename)
    input_path = _save_upload(file)

    if custom_filename.strip():
        stem = Path(custom_filename.strip()).stem
        output_name = stem + original_name.suffix
    elif output_mode == "overwrite":
        output_name = original_name.name
    else:
        output_name = original_name.stem + "_limpio" + original_name.suffix

    output_path = TEMP_DIR / f"{uuid.uuid4()}.xlsx"

    try:
        report = clean_file(
            file_path=input_path,
            output_path=output_path,
            fix_dates=fix_dates,
            fix_broken_chars=fix_broken_chars,
            unmerge_cells=unmerge_cells,
        )
    except Exception as e:
        raise HTTPException(500, f"Error procesando el archivo: {e}")
    finally:
        input_path.unlink(missing_ok=True)

    if not output_path.exists():
        raise HTTPException(500, "No se pudo generar el archivo limpio")

    content = output_path.read_bytes()
    output_path.unlink(missing_ok=True)

    safe_ascii = output_name.encode("ascii", "ignore").decode()
    content_disposition = (
        f'attachment; filename="{safe_ascii}"; '
        f"filename*=UTF-8''{output_name}"
    )

    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": content_disposition,
            "X-Clean-Report": str(report.get("summary", {})),
            "X-Output-Filename": output_name,
            "Access-Control-Expose-Headers": "X-Clean-Report, X-Output-Filename",
        },
    )
