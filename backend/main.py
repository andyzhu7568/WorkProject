import io
import os
import shutil
import subprocess
import tempfile
from datetime import datetime

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from .converter import pptx_to_test_excel


def _ppt_to_pptx_bytes(ppt_bytes: bytes) -> bytes:
    """Convert .ppt (binary) to .pptx using LibreOffice. Returns .pptx file bytes."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise HTTPException(
            status_code=400,
            detail=(
                "Converting .ppt files requires LibreOffice. "
                "Install LibreOffice and ensure 'soffice' or 'libreoffice' is on your PATH, "
                "or save the file as .pptx in PowerPoint and upload again."
            ),
        )
    with tempfile.TemporaryDirectory() as tmpdir:
        ppt_path = os.path.join(tmpdir, "input.ppt")
        with open(ppt_path, "wb") as f:
            f.write(ppt_bytes)
        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir, exist_ok=True)
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pptx", "--outdir", out_dir, ppt_path],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            raise HTTPException(
                status_code=400,
                detail=f"LibreOffice conversion failed: {result.stderr or result.stdout or 'unknown error'}",
            )
        pptx_path = os.path.join(out_dir, "input.pptx")
        if not os.path.isfile(pptx_path):
            raise HTTPException(
                status_code=400,
                detail="LibreOffice did not produce a .pptx file. Try saving as .pptx in PowerPoint.",
            )
        with open(pptx_path, "rb") as f:
            return f.read()


app = FastAPI(title="PPT to Test Excel Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/convert")
async def convert_ppt_to_excel(file: UploadFile = File(...)):
    filename = file.filename or ""
    lower_name = filename.lower()

    if not (lower_name.endswith(".pptx") or lower_name.endswith(".ppt")):
        raise HTTPException(
            status_code=400,
            detail="Only .pptx or .ppt files are supported. Please upload a valid file.",
        )

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        if lower_name.endswith(".ppt"):
            content = _ppt_to_pptx_bytes(content)

        excel_bytes = pptx_to_test_excel(content)
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(
            status_code=500,
            detail=f"Conversion error: {exc}",
        ) from exc

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = filename.rsplit(".", 1)[0] or "converted"
    out_name = f"{base_name}_test_sheet_{timestamp}.xlsx"

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


# Static frontend (run from project root: uvicorn backend.main:app --reload)
app.mount(
    "/",
    StaticFiles(directory="frontend", html=True),
    name="static",
)

