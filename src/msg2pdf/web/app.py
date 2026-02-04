"""FastAPI web application for MSG to PDF conversion."""

import html
import logging
import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime, timedelta
from pathlib import Path, PurePath
from typing import Annotated

from fastapi import FastAPI, File, HTTPException, Request, UploadFile, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse

from msg2pdf.core.converter import MSGToPDFConverter
from msg2pdf.core.exceptions import MSG2PDFError

# Configuration constants
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB
MAX_FILES_PER_SESSION = 50
CLEANUP_AGE_HOURS = 1
CHUNK_SIZE = 8192  # 8 KB chunks for file reading

logger = logging.getLogger(__name__)

app = FastAPI(
    title="MSG to PDF Converter",
    description="Convert Outlook MSG files to PDF",
    version="0.1.0",
)

# Temporary storage for converted files
TEMP_DIR = Path(tempfile.gettempdir()) / "msg2pdf_web"
TEMP_DIR.mkdir(exist_ok=True)

# Store job info (in production, use Redis or database)
jobs: dict[str, dict] = {}


def sanitize_filename(filename: str) -> str:
    """Sanitize filename to prevent path traversal attacks."""
    # Extract just the filename, removing any path components
    safe_name = PurePath(filename).name
    # Remove any remaining dangerous characters
    safe_name = safe_name.replace("\x00", "").strip()
    if not safe_name:
        safe_name = "unnamed.msg"
    return safe_name


def cleanup_old_files():
    """Remove files older than configured age."""
    cutoff = datetime.now() - timedelta(hours=CLEANUP_AGE_HOURS)
    for job_dir in TEMP_DIR.iterdir():
        if job_dir.is_dir():
            try:
                mtime = datetime.fromtimestamp(job_dir.stat().st_mtime)
                if mtime < cutoff:
                    shutil.rmtree(job_dir)
                    logger.info(f"Cleaned up old job directory: {job_dir.name}")
            except Exception as e:
                logger.warning(f"Failed to clean up {job_dir}: {e}")


async def save_upload_file(file: UploadFile, destination: Path) -> int:
    """Save uploaded file with chunked reading and size limit.

    Returns the file size in bytes.
    Raises HTTPException if file is too large.
    """
    total_size = 0
    with open(destination, "wb") as f:
        while chunk := await file.read(CHUNK_SIZE):
            total_size += len(chunk)
            if total_size > MAX_FILE_SIZE:
                # Clean up partial file
                f.close()
                destination.unlink(missing_ok=True)
                raise HTTPException(
                    status_code=413,
                    detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)} MB.",
                )
            f.write(chunk)
    return total_size


@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main upload page."""
    return """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Convertir vos emails en PDF</title>
    <style>
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background: linear-gradient(135deg, #4a90a4 0%, #2c5f72 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }

        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 50px;
            max-width: 700px;
            width: 100%;
        }

        h1 {
            color: #2c5f72;
            margin-bottom: 12px;
            font-size: 32px;
            text-align: center;
        }

        .subtitle {
            color: #666;
            margin-bottom: 35px;
            font-size: 18px;
            text-align: center;
            line-height: 1.5;
        }

        .instructions {
            background: #f0f7fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 25px;
            font-size: 16px;
            color: #444;
            line-height: 1.6;
        }

        .instructions strong {
            color: #2c5f72;
        }

        input[type="file"] {
            display: none;
        }

        .upload-area {
            display: block;
            border: 3px dashed #4a90a4;
            border-radius: 16px;
            padding: 50px 30px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #f8fbfc;
        }

        .upload-area:hover, .upload-area.dragover {
            border-color: #2c5f72;
            background: #e8f4f8;
        }

        .upload-area svg {
            width: 64px;
            height: 64px;
            color: #4a90a4;
            margin-bottom: 20px;
        }

        .upload-area p {
            color: #333;
            margin-bottom: 10px;
            font-size: 20px;
            font-weight: 500;
        }

        .upload-area .hint {
            font-size: 16px;
            color: #666;
        }

        .file-list {
            margin-top: 25px;
            max-height: 300px;
            overflow-y: auto;
        }

        .file-item {
            display: flex;
            align-items: center;
            padding: 15px 18px;
            background: #f5f5f5;
            border-radius: 10px;
            margin-bottom: 10px;
            font-size: 17px;
        }

        .file-item.converting {
            background: #fff3cd;
            border: 2px solid #ffc107;
        }

        .file-item.done {
            background: #d4edda;
            border: 2px solid #28a745;
        }

        .file-item.error {
            background: #f8d7da;
            border: 2px solid #dc3545;
        }

        .file-item .name {
            flex: 1;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-weight: 500;
        }

        .file-item .status {
            margin-left: 15px;
            font-size: 15px;
            color: #555;
            font-weight: 500;
        }

        .file-item .size {
            color: #888;
            margin-left: 15px;
            font-size: 15px;
        }

        .file-item .remove {
            margin-left: 15px;
            color: #999;
            cursor: pointer;
            font-size: 24px;
            font-weight: bold;
            padding: 0 8px;
        }

        .file-item .remove:hover {
            color: #e74c3c;
        }

        .btn {
            display: inline-block;
            padding: 20px 40px;
            background: linear-gradient(135deg, #4a90a4 0%, #2c5f72 100%);
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 20px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            margin-top: 25px;
            width: 100%;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(74, 144, 164, 0.4);
        }

        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }

        .progress {
            margin-top: 25px;
            display: none;
        }

        .progress-bar {
            height: 12px;
            background: #eee;
            border-radius: 6px;
            overflow: hidden;
        }

        .progress-bar-fill {
            height: 100%;
            background: linear-gradient(135deg, #4a90a4 0%, #2c5f72 100%);
            width: 0%;
            transition: width 0.3s;
        }

        .progress-text {
            text-align: center;
            margin-top: 12px;
            font-size: 18px;
            color: #555;
            font-weight: 500;
        }

        .result {
            margin-top: 25px;
            padding: 25px;
            border-radius: 12px;
            display: none;
            font-size: 18px;
            text-align: center;
        }

        .result.success {
            background: #d4edda;
            color: #155724;
            border: 2px solid #28a745;
        }

        .result.error {
            background: #f8d7da;
            color: #721c24;
            border: 2px solid #dc3545;
        }

        .download-btn {
            display: inline-block;
            margin-top: 18px;
            padding: 16px 35px;
            background: #28a745;
            color: white;
            text-decoration: none;
            border-radius: 10px;
            font-weight: 600;
            font-size: 19px;
        }

        .download-btn:hover {
            background: #218838;
        }

        .spinner {
            display: inline-block;
            width: 18px;
            height: 18px;
            border: 3px solid #666;
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-right: 10px;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .footer {
            margin-top: 30px;
            text-align: center;
            color: #999;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Convertir vos emails en PDF</h1>
        <p class="subtitle">Transformez vos fichiers Outlook (.msg) en documents PDF</p>

        <div class="instructions">
            <strong>Comment faire :</strong><br>
            1. Cliquez sur la zone bleue ci-dessous pour choisir vos fichiers<br>
            2. Cliquez sur le bouton "Convertir en PDF"<br>
            3. Attendez la fin de la conversion<br>
            4. Cliquez sur "Telecharger" pour recuperer vos PDF
        </div>

        <input type="file" id="fileInput" multiple accept=".msg">
        <div class="upload-area" id="uploadArea">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <p>Cliquez ici pour choisir vos fichiers</p>
            <p class="hint">ou glissez-deposez vos fichiers .msg ici</p>
        </div>

        <div class="file-list" id="fileList"></div>

        <button class="btn" id="convertBtn" disabled>Convertir en PDF</button>

        <div class="progress" id="progress">
            <div class="progress-bar">
                <div class="progress-bar-fill" id="progressFill"></div>
            </div>
            <p class="progress-text" id="progressText">Conversion en cours...</p>
        </div>

        <div class="result" id="result"></div>

        <div class="footer">
            Les fichiers sont automatiquement supprimes apres 1 heure
        </div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const convertBtn = document.getElementById('convertBtn');
        const progress = document.getElementById('progress');
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        const result = document.getElementById('result');

        let selectedFiles = [];
        let sessionId = null;

        // Escape HTML to prevent XSS
        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        // Click on upload area - explicitly trigger file input
        uploadArea.addEventListener('click', function(e) {
            console.log('Upload area clicked, target:', e.target);
            console.log('fileInput element:', fileInput);
            fileInput.click();
            console.log('fileInput.click() called');
        });

        // Drag and drop - need both dragenter and dragover prevented
        uploadArea.addEventListener('dragenter', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        // File input change
        fileInput.addEventListener('change', (e) => {
            console.log('Change event fired, files:', e.target.files);
            if (e.target.files && e.target.files.length > 0) {
                handleFiles(e.target.files);
            }
            // Reset input to allow selecting the same file again
            e.target.value = '';
        });

        function handleFiles(files) {
            console.log('handleFiles called with', files.length, 'files');
            let addedCount = 0;
            let skippedCount = 0;
            for (const file of files) {
                console.log('Processing file:', file.name);
                if (file.name.toLowerCase().endsWith('.msg')) {
                    if (!selectedFiles.find(f => f.file.name === file.name)) {
                        selectedFiles.push({ file, status: 'pending' });
                        console.log('Added file:', file.name);
                        addedCount++;
                    } else {
                        console.log('File already in list:', file.name);
                    }
                } else {
                    console.log('Skipped non-MSG file:', file.name);
                    skippedCount++;
                }
            }
            console.log('selectedFiles now has', selectedFiles.length, 'files');

            // Show message if files were skipped
            if (skippedCount > 0 && addedCount === 0) {
                alert('Les fichiers selectionnes ne sont pas des fichiers .msg\\n\\nVeuillez selectionner des fichiers Outlook (.msg)');
            }

            updateFileList();
        }

        function updateFileList() {
            console.log('updateFileList called, selectedFiles:', selectedFiles.length);
            fileList.innerHTML = '';
            selectedFiles.forEach((item, index) => {
                const div = document.createElement('div');
                div.className = 'file-item ' + item.status;
                div.id = 'file-' + index;

                // Create elements safely to prevent XSS
                const nameSpan = document.createElement('span');
                nameSpan.className = 'name';
                nameSpan.textContent = item.file.name;  // Safe: textContent escapes HTML

                const sizeSpan = document.createElement('span');
                sizeSpan.className = 'size';
                sizeSpan.textContent = formatSize(item.file.size);

                div.appendChild(nameSpan);
                div.appendChild(sizeSpan);

                if (item.status === 'pending') {
                    const removeBtn = document.createElement('span');
                    removeBtn.className = 'remove';
                    removeBtn.textContent = '×';
                    removeBtn.onclick = () => removeFile(index);
                    div.appendChild(removeBtn);
                } else if (item.status === 'converting') {
                    const spinner = document.createElement('span');
                    spinner.className = 'spinner';
                    div.appendChild(spinner);
                    const statusSpan = document.createElement('span');
                    statusSpan.className = 'status';
                    statusSpan.textContent = 'En cours...';
                    div.appendChild(statusSpan);
                } else if (item.status === 'done') {
                    const statusSpan = document.createElement('span');
                    statusSpan.className = 'status';
                    statusSpan.textContent = 'Termine !';
                    div.appendChild(statusSpan);
                } else if (item.status === 'error') {
                    const statusSpan = document.createElement('span');
                    statusSpan.className = 'status';
                    statusSpan.textContent = 'Erreur';
                    div.appendChild(statusSpan);
                }

                fileList.appendChild(div);
            });
            convertBtn.disabled = selectedFiles.length === 0 || selectedFiles.some(f => f.status === 'converting');
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            updateFileList();
        }

        function formatSize(bytes) {
            if (bytes < 1024) return bytes + ' o';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' Ko';
            return (bytes / (1024 * 1024)).toFixed(1) + ' Mo';
        }

        // Convert files one by one
        convertBtn.addEventListener('click', async () => {
            if (selectedFiles.length === 0) return;

            convertBtn.disabled = true;
            progress.style.display = 'block';
            result.style.display = 'none';

            // Create a new session
            sessionId = null;
            let successCount = 0;
            let errorCount = 0;

            for (let i = 0; i < selectedFiles.length; i++) {
                const item = selectedFiles[i];
                item.status = 'converting';
                updateFileList();

                const progressPct = Math.round((i / selectedFiles.length) * 100);
                progressFill.style.width = progressPct + '%';
                progressText.textContent = 'Conversion du fichier ' + (i + 1) + ' sur ' + selectedFiles.length + '...';

                const formData = new FormData();
                formData.append('file', item.file);
                if (sessionId) {
                    formData.append('session_id', sessionId);
                }

                try {
                    const response = await fetch('/convert-single', {
                        method: 'POST',
                        body: formData
                    });

                    const data = await response.json();

                    if (response.ok) {
                        item.status = 'done';
                        sessionId = data.session_id;
                        successCount++;
                    } else {
                        item.status = 'error';
                        item.error = data.detail;
                        errorCount++;
                    }
                } catch (error) {
                    item.status = 'error';
                    item.error = error.message;
                    errorCount++;
                }

                updateFileList();
            }

            progressFill.style.width = '100%';
            progressText.textContent = 'Termine !';

            result.style.display = 'block';

            if (successCount > 0 && sessionId) {
                result.className = 'result success';
                if (successCount === 1) {
                    result.innerHTML =
                        '<strong>Conversion reussie !</strong>' +
                        (errorCount > 0 ? '<br>' + errorCount + ' fichier(s) en erreur.' : '') +
                        '<br><a href="/download/' + escapeHtml(sessionId) + '" class="download-btn">Telecharger le PDF</a>';
                } else {
                    result.innerHTML =
                        '<strong>' + successCount + ' fichiers convertis avec succes !</strong>' +
                        (errorCount > 0 ? '<br>' + errorCount + ' fichier(s) en erreur.' : '') +
                        '<br><a href="/download/' + escapeHtml(sessionId) + '" class="download-btn">Telecharger le fichier ZIP</a>';
                }
            } else {
                result.className = 'result error';
                result.innerHTML = '<strong>La conversion a echoue.</strong><br>Veuillez reessayer avec d\'autres fichiers.';
            }

            // Reset for next batch
            selectedFiles = selectedFiles.filter(f => f.status !== 'done');
            if (selectedFiles.length === 0) {
                setTimeout(() => {
                    updateFileList();
                }, 2000);
            }
            convertBtn.disabled = false;
        });
    </script>
</body>
</html>
"""


@app.post("/convert-single")
async def convert_single_file(
    background_tasks: BackgroundTasks,
    file: Annotated[UploadFile, File(description="MSG file to convert")],
    session_id: str | None = None,
):
    """Convert a single MSG file to PDF."""
    # Sanitize filename to prevent path traversal
    safe_filename = sanitize_filename(file.filename or "unnamed.msg")

    if not safe_filename.lower().endswith(".msg"):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Only .msg files are accepted.",
        )

    # Create or reuse session directory
    if session_id and session_id in jobs:
        job_id = session_id
        job_dir = TEMP_DIR / job_id
        output_dir = job_dir / "output"

        # Check file count limit
        if jobs[job_id]["file_count"] >= MAX_FILES_PER_SESSION:
            raise HTTPException(
                status_code=400,
                detail=f"Maximum {MAX_FILES_PER_SESSION} files per session.",
            )
    else:
        job_id = str(uuid.uuid4())
        job_dir = TEMP_DIR / job_id
        input_dir = job_dir / "input"
        output_dir = job_dir / "output"
        input_dir.mkdir(parents=True)
        output_dir.mkdir(parents=True)
        jobs[job_id] = {
            "created": datetime.now(),
            "file_count": 0,
            "output_dir": output_dir,
            "files": [],
        }

    # Save uploaded file with chunked reading
    input_dir = job_dir / "input"
    input_dir.mkdir(exist_ok=True)
    file_path = input_dir / safe_filename

    try:
        await save_upload_file(file, file_path)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Failed to save uploaded file: {e}")
        raise HTTPException(status_code=500, detail="Failed to save file.")

    # Convert file
    converter = MSGToPDFConverter()
    try:
        pdf_path = converter.convert(
            file_path,
            output_dir,
            merge_attachments=True,
            show_source=False,
        )

        # Update job info
        jobs[job_id]["file_count"] += 1
        jobs[job_id]["files"].append(pdf_path.name)

        # Schedule cleanup
        background_tasks.add_task(cleanup_old_files)

        return {
            "session_id": job_id,
            "filename": pdf_path.name,
            "success": True,
        }

    except Exception as e:
        logger.error(f"Conversion failed for {safe_filename}: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Conversion failed: {str(e)}",
        )


@app.post("/convert")
async def convert_files(
    background_tasks: BackgroundTasks,
    files: Annotated[list[UploadFile], File(description="MSG files to convert")],
):
    """Convert uploaded MSG files to PDF (legacy endpoint for single/few files)."""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    if len(files) > MAX_FILES_PER_SESSION:
        raise HTTPException(
            status_code=400,
            detail=f"Maximum {MAX_FILES_PER_SESSION} files per request.",
        )

    # Validate and sanitize filenames
    for file in files:
        safe_name = sanitize_filename(file.filename or "unnamed.msg")
        if not safe_name.lower().endswith(".msg"):
            raise HTTPException(
                status_code=400,
                detail="Invalid file type. Only .msg files are accepted.",
            )

    # Create job directory
    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    input_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)

    # Save uploaded files with chunked reading
    saved_files = []
    for file in files:
        safe_filename = sanitize_filename(file.filename or "unnamed.msg")
        file_path = input_dir / safe_filename
        try:
            await save_upload_file(file, file_path)
            saved_files.append(file_path)
        except HTTPException:
            # Clean up and re-raise
            shutil.rmtree(job_dir, ignore_errors=True)
            raise

    # Convert files
    converter = MSGToPDFConverter()
    converted = []
    errors = []

    for msg_path in saved_files:
        try:
            pdf_path = converter.convert(
                msg_path,
                output_dir,
                merge_attachments=True,
                show_source=False,
            )
            converted.append(pdf_path)
        except Exception as e:
            errors.append(f"{msg_path.name}: {str(e)}")

    if not converted:
        shutil.rmtree(job_dir)
        raise HTTPException(
            status_code=500,
            detail=f"All conversions failed: {'; '.join(errors)}",
        )

    # Store job info
    jobs[job_id] = {
        "created": datetime.now(),
        "file_count": len(converted),
        "output_dir": output_dir,
        "files": [p.name for p in converted],
    }

    # Schedule cleanup
    background_tasks.add_task(cleanup_old_files)

    return {
        "job_id": job_id,
        "file_count": len(converted),
        "files": [p.name for p in converted],
        "errors": errors if errors else None,
    }


@app.get("/download/{job_id}")
async def download_result(job_id: str):
    """Download converted PDF(s)."""
    # Validate job_id format (UUID)
    try:
        uuid.UUID(job_id)
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid job ID format")

    if job_id not in jobs:
        # Check if directory exists (for restarts)
        job_dir = TEMP_DIR / job_id / "output"
        if not job_dir.exists():
            raise HTTPException(status_code=404, detail="Job not found or expired")

        # Rebuild job info
        pdf_files = list(job_dir.glob("*.pdf"))
        if not pdf_files:
            raise HTTPException(status_code=404, detail="No files found")

        jobs[job_id] = {
            "created": datetime.now(),
            "file_count": len(pdf_files),
            "output_dir": job_dir,
            "files": [p.name for p in pdf_files],
        }

    job = jobs[job_id]
    output_dir = job["output_dir"]

    if job["file_count"] == 1:
        # Single file - return PDF directly
        pdf_path = output_dir / job["files"][0]
        return FileResponse(
            pdf_path,
            media_type="application/pdf",
            filename=job["files"][0],
        )
    else:
        # Multiple files - create ZIP
        zip_path = output_dir.parent / "converted.zip"
        if not zip_path.exists():
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename in job["files"]:
                    pdf_path = output_dir / filename
                    zf.write(pdf_path, filename)

        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename="converted_emails.zip",
        )


@app.get("/health")
async def health():
    """Health check endpoint."""
    return {"status": "ok"}
