# PPTX to HTML web app

Upload a PPTX in a browser and download a ZIP containing `pg_1.html` through `pg_N.html`, and one PNG per slide. Choose **Create an index page** to add `index.html` with links to every page. Each page displays the slide, its speaker notes, and a visible page number. External hyperlinks in speaker notes retain their stored destinations, and plain URLs become clickable. Choose whether links open in a new tab.

## Run locally on macOS

Install Python 3.10+ and LibreOffice (place `LibreOffice.app` in `/Applications`). In Terminal, from this directory:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
uvicorn web:app --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000` in your browser. Stop with Ctrl+C.

## Run with Docker

```sh
docker build -t pptx-to-html .
docker run --rm -p 8000:8000 pptx-to-html
```

Open `http://localhost:8000`. Put these project files in a GitHub repository, then deploy the Docker app to a service that runs containers. GitHub Pages only hosts static files and cannot run this server-side conversion. Add access controls before a public deployment if conversions should be limited to authorized users.

## CLI option

```sh
python converter.py presentation.pptx output-directory
```

The web interface limits uploads to 50 MB and allows one conversion at a time. It uses LibreOffice to render slides, so output may vary with installed fonts. Compare the pages with the existing Click to Convert HTML before treating it as a replacement. The converter produces static images. Links attached to individual text runs on the slide itself are outside the current scope.

## Learning Center export

Select **Learning Center** in the browser (the default). The ZIP contains numbered HTML pages and their matching `backgroundN.png` assets. Extract the ZIP and open `pg_1.html` to review the output. Select **Standard** to include a contents page.

Learning Center output moves selectable horizontal slide text from the PDF into styled HTML. Artwork and text already embedded in artwork stay in the PNG. Image width controls the displayed slide size; PNG resolution controls its pixel density. The slide font setting affects extracted text only, while the body font controls speaker notes and page text. The browser uses fallback fonts and the exporter composites artwork into each background. Compare the layout in your Learning Center before replacing course content.
