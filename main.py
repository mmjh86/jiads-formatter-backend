# main.py
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
from formatter import format_jiads_manuscript

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Use /tmp on Render (ephemeral filesystem)
UPLOAD_DIR = os.environ.get("MANUSCRIPT_UPLOAD_DIR", "/tmp/uploads")
OUTPUT_DIR = os.environ.get("MANUSCRIPT_OUTPUT_DIR", "/tmp/outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.get("/")
def root():
    return {"message": "JIADS Formatter API", "status": "running"}


@app.post("/format")
async def format_manuscript(file: UploadFile = File(...)):
    if not file.filename.endswith(".docx"):
        raise HTTPException(400, "Only .docx files are accepted")

    input_id    = str(uuid.uuid4())
    input_path  = os.path.join(UPLOAD_DIR, f"{input_id}.docx")
    output_path = os.path.join(OUTPUT_DIR, f"{input_id}_formatted.docx")

    with open(input_path, "wb") as f:
        f.write(await file.read())

    api_key = os.environ.get("ANTHROPIC_API_KEY")

    try:
        format_jiads_manuscript(input_path, output_path, api_key)
        return FileResponse(
            output_path,
            filename="JIADS_formatted.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        raise HTTPException(500, f"Formatting failed: {str(e)}")
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)
