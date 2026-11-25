"""Browser interface for the PPTX to HTML converter."""

import shutil
import tempfile
import threading
import zipfile
import re
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from starlette.background import BackgroundTask

from converter import convert

app = FastAPI(title="PPTX to HTML")
app.add_middleware(CORSMiddleware, allow_origins=["https://ecemozturk.github.io"], allow_methods=["POST", "PUT", "DELETE"])
conversion_lock = threading.Lock()
uploads_lock = threading.Lock()
uploads = {}
MAX_UPLOAD = 200 * 1024 * 1024
CHUNK_SIZE = 8 * 1024 * 1024
MAX_UNCOMPRESSED = 1024 * 1024 * 1024
UPLOAD_TTL = 60 * 60

PAGE = r"""<!doctype html>
<html lang="en">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PowerPoint to HTML</title>
<style>
*{box-sizing:border-box}
:root{font:16px/1.55 system-ui,sans-serif;color:#182c2a;background:#eef4f1}
body{margin:0;padding:clamp(24px,7vw,85px) 16px}
main{max-width:680px;margin:auto;background:white;padding:clamp(20px,5vw,44px);border:1px solid #d6e4df;border-radius:24px;box-shadow:0 18px 55px #1d413014}
h1{font-size:clamp(29px,6vw,44px);line-height:1.08;margin:0;letter-spacing:-.045em}
p{margin:.75rem 0 1.5rem;color:#455c54}
.drop{display:block;padding:24px 16px;text-align:center;border:2px dashed #88ac9d;border-radius:14px;background:#f5faf7}
.drop.drag{background:#dcefe5;border-color:#167052}
.drop strong{display:block}
.drop small{display:block;color:#455c54;margin:.4rem 0 .9rem}
input[type=file]{max-width:100%;font:inherit;font-size:.875rem}
input::file-selector-button{border:1px solid #8daea0;border-radius:7px;padding:.45rem .65rem;background:white;color:#174f3b;cursor:pointer}
.settings{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:24px 0 12px}
.settings label{font-weight:650}
select{width:100%;margin-top:6px;padding:10px;background:#fff;border:1px solid #91ada2;border-radius:8px;font:inherit;color:inherit}
.check{display:flex;gap:8px;align-items:center;margin:20px 0}
input[type=checkbox]{accent-color:#136549;width:17px;height:17px}
button,.download{display:inline-block;padding:12px 18px;border:0;border-radius:9px;background:#126449;color:white;font:inherit;font-weight:650;cursor:pointer}
button:disabled{opacity:.65;cursor:wait}
progress{width:100%;height:12px;margin-top:12px;accent-color:#126449}
.download{color:#12573f;background:#e7f4ed}
[hidden]{display:none!important}
#status{margin:12px 0 0}
#status[data-error=true]{color:#982c26}
:focus-visible{outline:3px solid #196eaa;outline-offset:3px}
@media(max-width:470px){.settings{grid-template-columns:1fr}}
</style>
<main>
<h1>PowerPoint to HTML</h1>
<p>Convert Bryan Ventures PowerPoint course materials into web pages and matching PNG images.</p>
<form id="form" action="/convert" method="post" enctype="multipart/form-data">
<label class="drop" id="drop">
  <strong>Drop a PowerPoint file here</strong>
  <small>Or select a .pptx file up to 200 MB.</small>
  <input id="file" required type="file" name="file" accept=".pptx" aria-label="Choose a PowerPoint file">
</label>
<div class="settings">
<label>Output format
  <select name="format"><option value="legacy">Learning Center</option><option value="standard">Standard</option></select>
</label>
<label>PNG resolution
  <select name="dpi"><option value="96">96 DPI</option></select>
</label>
<label>Image width
  <select name="image_width"><option value="607">607 px (original size)</option><option value="480">480 px</option><option value="720">720 px</option></select>
</label>
<label>Slide text font
  <select name="slide_font"><option value="georgia">Georgia</option><option value="arial">Arial</option><option value="times">Times New Roman</option><option value="verdana">Verdana</option></select>
</label>
<label>Body font size
  <select name="body_font_size"><option value="14">14 px</option><option value="17" selected>17 px</option><option value="20">20 px</option></select>
</label>
</div>
<label class="check"><input type="checkbox" name="new_tab" value="true" checked> Open links in a new tab</label>
<label class="check"><input type="checkbox" name="create_index" value="true"> Create an index page with links to every slide</label>
<button type="submit">Convert PowerPoint</button>
<progress id="progress" hidden max="100"></progress>
<p id="status" role="status"></p>
<a id="download" class="download" hidden>Download ZIP</a>
</form>
</main>
<script>
const form = document.querySelector('#form');
const file = document.querySelector('#file');
const drop = document.querySelector('#drop');
const button = form.querySelector('button');
const statusNode = document.querySelector('#status');
const progressNode = document.querySelector('#progress');
const download = document.querySelector('#download');
let downloadUrl;

function showStatus(message, error = false) {
  statusNode.textContent = message;
  statusNode.dataset.error = error;
}

function showProgress(message, percent) {
  progressNode.hidden = false;
  if (percent == null) progressNode.removeAttribute('value');
  else progressNode.value = percent;
  showStatus(message);
}

function clearDownload() {
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = undefined;
  download.hidden = true;
  download.removeAttribute('href');
}

function validateFile() {
  const selected = file.files[0];
  let error = '';
  if (selected && !/\.pptx$/i.test(selected.name)) error = 'Choose a .pptx file.';
  if (selected && selected.size > 200 * 1024 * 1024) error = 'Choose a file under 200 MB.';
  file.setCustomValidity(error);
  showStatus(error || (selected ? 'Selected: ' + selected.name : ''), Boolean(error));
  return Boolean(selected && !error);
}

file.addEventListener('change', () => {
  clearDownload();
  validateFile();
});
drop.addEventListener('dragover', event => {
  event.preventDefault();
  if (!button.disabled) drop.classList.add('drag');
});
drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
drop.addEventListener('drop', event => {
  event.preventDefault();
  drop.classList.remove('drag');
  if (button.disabled || !event.dataTransfer.files.length) return;
  if (event.dataTransfer.files.length !== 1) {
    showStatus('Choose one PowerPoint file at a time.', true);
    return;
  }
  file.files = event.dataTransfer.files;
  clearDownload();
  validateFile();
});

async function uploadInChunks(selected, apiBase) {
  const start = await fetch(new URL('/uploads', apiBase), {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name: selected.name, size: selected.size})
  });
  if (!start.ok) throw new Error((await start.json()).detail || 'Could not start the upload.');
  const {upload} = await start.json();
  const chunkSize = 8 * 1024 * 1024;
  try {
    for (let offset = 0; offset < selected.size; offset += chunkSize) {
      const part = selected.slice(offset, offset + chunkSize);
      const response = await fetch(new URL('/uploads/' + upload + '?offset=' + offset, apiBase), {
        method: 'PUT', body: part
      });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.detail || 'Upload failed.');
      }
      showProgress('Uploading: ' + Math.round(100 * Math.min(offset + part.size, selected.size) / selected.size) + '%',
                   100 * Math.min(offset + part.size, selected.size) / selected.size);
    }
    return upload;
  } catch (error) {
    await fetch(new URL('/uploads/' + upload, apiBase), {method: 'DELETE'}).catch(() => {});
    throw error;
  }
}

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (button.disabled || !validateFile()) return;
  clearDownload();
  const selected = file.files[0];
  const body = new FormData(form);
  body.delete('file');
  const name = file.files[0].name.replace(/\.pptx$/i, '');
  button.disabled = true;
  file.disabled = true;
  showProgress('Uploading: 0%', 0);
  try {
    const upload = await uploadInChunks(selected, window.location.origin);
    body.set('upload', upload);
    showProgress('Converting slides and notes…', null);
    const response = await fetch(new URL('/convert', window.location.origin), {method: 'POST', body});
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(typeof data.detail === 'string' ? data.detail : 'Conversion failed. Try again.');
    }
    downloadUrl = URL.createObjectURL(await response.blob());
    download.href = downloadUrl;
    download.download = name + '-html.zip';
    download.hidden = false;
    progressNode.hidden = true;
    showStatus('Your ZIP is ready.');
  } catch (error) {
    progressNode.hidden = true;
    showStatus(error.message, true);
  } finally {
    button.disabled = false;
    file.disabled = false;
  }
});
</script>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def home():
    return PAGE


def discard_upload(upload):
    entry = uploads.pop(upload, None)
    if entry:
        shutil.rmtree(entry["directory"], ignore_errors=True)


@app.post("/uploads")
async def start_upload(request: Request):
    data = await request.json()
    name, size = data.get("name"), data.get("size")
    if not isinstance(name, str) or Path(name).suffix.lower() != ".pptx":
        raise HTTPException(400, "Please upload a .pptx file")
    if type(size) is not int or not 0 < size <= MAX_UPLOAD:
        raise HTTPException(413, "PPTX must be between 1 byte and 200 MB")
    with uploads_lock:
        for key, entry in list(uploads.items()):
            if time.monotonic() - entry["updated"] > UPLOAD_TTL:
                discard_upload(key)
        upload = uuid.uuid4().hex
        directory = Path(tempfile.mkdtemp(prefix="pptx-upload-"))
        uploads[upload] = {"directory": directory, "name": name, "size": size,
                           "received": 0, "updated": time.monotonic()}
    return {"upload": upload}


@app.put("/uploads/{upload}")
async def append_upload(upload: str, offset: int, request: Request):
    if int(request.headers.get("content-length", "0")) > CHUNK_SIZE:
        raise HTTPException(413, "Upload chunk exceeds its limit")
    content = await request.body()
    with uploads_lock:
        entry = uploads.get(upload)
        if not entry:
            raise HTTPException(404, "Upload not found or expired")
        if offset != entry["received"]:
            raise HTTPException(409, "Upload offset does not match")
        if not content or len(content) > CHUNK_SIZE or offset + len(content) > entry["size"]:
            raise HTTPException(413, "Upload chunk exceeds its limit")
        with (entry["directory"] / "source.pptx").open("ab") as target:
            target.write(content)
        entry["received"] += len(content)
        entry["updated"] = time.monotonic()
        return {"received": entry["received"]}


@app.delete("/uploads/{upload}")
def cancel_upload(upload: str):
    with uploads_lock:
        discard_upload(upload)
    return {"cancelled": True}


@app.post("/convert")
def upload_and_convert(file: UploadFile | None = File(None), upload: str | None = Form(None), dpi: int = Form(96),
                       new_tab: bool = Form(False), format: str = Form("legacy"),
                       image_width: int = Form(607), slide_font: str = Form("georgia"),
                       body_font_size: int = Form(17), create_index: bool = Form(False)):
    if not upload and (not file or not file.filename or Path(file.filename).suffix.lower() != ".pptx"):
        raise HTTPException(400, "Please upload a .pptx file")
    if format not in ("legacy", "standard"):
        raise HTTPException(400, "Invalid output format")
    if dpi != 96:
        raise HTTPException(400, "Invalid resolution")
    if image_width not in (480, 607, 720) or slide_font not in ("arial", "georgia", "times", "verdana") or body_font_size not in (14, 17, 20):
        raise HTTPException(400, "Invalid display setting")
    if not conversion_lock.acquire(blocking=False):
        raise HTTPException(503, "Another conversion is running. Please retry shortly")
    temp = Path(tempfile.mkdtemp(prefix="pptx-to-html-"))
    try:
        if upload:
            with uploads_lock:
                entry = uploads.get(upload)
                if not entry:
                    raise HTTPException(404, "Upload not found or expired")
                if entry["received"] != entry["size"]:
                    raise HTTPException(400, "Upload is incomplete")
                uploads.pop(upload)
            filename = entry["name"]
            uploaded_source = entry["directory"] / "source.pptx"
        else:
            filename = file.filename
        safe_stem = re.sub(r"[^A-Za-z0-9_-]", "_", Path(filename).stem)[:80] or "presentation"
        source = temp / f"{safe_stem}.pptx"
        if upload:
            shutil.move(uploaded_source, source)
            shutil.rmtree(entry["directory"], ignore_errors=True)
        else:
            with source.open("wb") as output:
                total = 0
                while chunk := file.file.read(1024 * 1024):
                    total += len(chunk)
                    if total > MAX_UPLOAD:
                        raise HTTPException(413, "PPTX exceeds the 200 MB limit")
                    output.write(chunk)
        if not zipfile.is_zipfile(source):
            raise HTTPException(400, "The file is not a valid PPTX")
        with zipfile.ZipFile(source) as archive:
            if (len(archive.infolist()) > 3000 or
                    sum(item.file_size for item in archive.infolist()) > MAX_UNCOMPRESSED):
                raise HTTPException(413, "PPTX contents exceed the safety limits")
            if "[Content_Types].xml" not in archive.namelist() or "ppt/presentation.xml" not in archive.namelist():
                raise HTTPException(400, "The file is not a valid PPTX")
        output_dir = temp / "output"
        convert(source, output_dir, dpi=dpi, new_tab=new_tab, legacy=format == "legacy",
                image_width=image_width, slide_font=slide_font, body_font_size=body_font_size,
                create_index=create_index)
        bundle = temp / "converted.zip"
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
            for item in sorted(output_dir.iterdir()):
                archive.write(item, item.name)
        return FileResponse(bundle, media_type="application/zip",
                            filename="converted-slides.zip",
                            background=BackgroundTask(shutil.rmtree, temp, ignore_errors=True))
    except HTTPException:
        shutil.rmtree(temp, ignore_errors=True)
        raise
    except Exception as error:
        shutil.rmtree(temp, ignore_errors=True)
        raise HTTPException(500, "Conversion failed. Check that the presentation opens normally in PowerPoint") from error
    finally:
        conversion_lock.release()
