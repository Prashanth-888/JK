Great — I’ll deliver the full, modular Option B project with the parameter-learning library and every requested feature enabled (multi-sheet output). Below you’ll find the full project structure and the complete code for every file you need. Copy these into your environment, install dependencies from requirements.txt, and run the CLI.

⸻

Project structure (what I’ll generate)

tyre_extractor/
├── src/
│   ├── core/
│   │   ├── extractor_engine.py
│   │   ├── pipeline.py
│   │   ├── parameter_manager.py
│   │   └── classifier.py
│   ├── readers/
│   │   ├── pdf_reader.py
│   │   ├── ocr_reader.py
│   │   ├── table_reader.py
│   │   └── image_reader.py
│   ├── utils/
│   │   ├── file_paths.py
│   │   ├── normalizer.py
│   │   ├── logger.py
│   │   └── text_cleanup.py
│   └── models/
│       ├── parameter_store.json
│       └── synonyms.json
├── app/
│   ├── cli.py
│   └── run_pipeline.py
├── input_files/                # put your PDFs/images here
├── output/
│   ├── consolidated.xlsx
│   └── images/
├── requirements.txt
└── README.md


⸻

How it works (summary)
	•	Read PDFs / images in input_files/.
	•	Extract raw text, structured tables (pdfplumber), and OCR tables (pytesseract + OpenCV).
	•	Extract images embedded in PDFs or rasterized pages, save images into output/images/<HTAC_No or filename>/.
	•	Parameter learning: parameter_manager maintains models/parameter_store.json (canonical parameters) and models/synonyms.json. New parameters detected get added (with fuzzy normalization). On subsequent runs, similar parameters map to same canonical.
	•	Consolidated Excel: one workbook with separate sheets per test-type (Physical Lab, Analytical/Chemical, Reinforcement Lab, Other). Each sheet includes SourceFile, HTAC_No, TestName, AllText, Images, plus columns for learned parameters (new columns appended).
	•	CLI: python app/cli.py --input input_files --output output/consolidated.xlsx

⸻

Files — full code

Save each file at the indicated path. I kept code commented and modular for clarity.

⸻

requirements.txt

pdfplumber==0.7.8
pdf2image==1.16.3
pytesseract==0.3.10
Pillow==10.0.1
opencv-python==4.9.0.76
pandas==2.2.2
openpyxl==3.1.2
rapidfuzz==2.15.1
python-dateutil==2.8.2

Also requires system dependencies:
	•	Poppler (for pdf2image): install via package manager (linux: sudo apt-get install poppler-utils; Windows: download poppler and add to PATH).
	•	Tesseract OCR: install system package (linux: sudo apt-get install tesseract-ocr; Windows: installer) and ensure tesseract is in PATH.

⸻

README.md

# Tyre Extractor (Enterprise)

Modular project to extract text, tables and images from tyre test PDFs/images and consolidate into Excel. Features parameter learning.

## Setup

1. Create and activate virtualenv:

python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

2. Install Python deps:

pip install -r requirements.txt

3. Install system deps:
- Poppler (pdf2image)
- Tesseract OCR

4. Place your files into `input_files/`.

5. Run:

python app/cli.py –input input_files –output output/consolidated.xlsx

6. Results:
- Excel: `output/consolidated.xlsx`
- Images: `output/images/<HTAC or file>/...`
- Learned parameters: `src/models/parameter_store.json` and `src/models/synonyms.json`

## Notes
- Tune OCR language and thresholds in `src/core/pipeline.py`.
- If scanned PDFs are poor, increase DPI in `pdf2image` conversion.


⸻

src/utils/logger.py

# src/utils/logger.py
import logging
import sys

def get_logger(name=__name__, level=logging.INFO):
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        fmt = '%(asctime)s - %(levelname)s - %(name)s - %(message)s'
        handler.setFormatter(logging.Formatter(fmt))
        logger.setLevel(level)
        logger.addHandler(handler)
    return logger


⸻

src/utils/normalizer.py

# src/utils/normalizer.py
import re

def normalize_param_name(s: str) -> str:
    if not s:
        return ""
    s = str(s)
    s = s.strip()
    s = re.sub(r'\s+', ' ', s)
    s = s.rstrip(':').strip()
    s = s.replace('º', 'deg').replace('°C', 'degC').replace('%',' percent')
    s = s.lower()
    s = re.sub(r'[^\w\s\/\-_\.]', '', s)
    s = s.strip()
    return s


⸻

src/utils/text_cleanup.py

# src/utils/text_cleanup.py
import re

def cleanup_text(text: str) -> str:
    if not text:
        return ""
    # unify line endings, remove excessive whitespace
    text = text.replace('\r\n', '\n').replace('\r','\n')
    text = re.sub(r'\n{2,}', '\n\n', text)
    return text.strip()


⸻

src/utils/file_paths.py

# src/utils/file_paths.py
from pathlib import Path

ROOT = Path.cwd()
INPUT_DIR = ROOT / "input_files"
OUTPUT_DIR = ROOT / "output"
MODELS_DIR = ROOT / "src" / "models"

def ensure_dirs():
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    MODELS_DIR.mkdir(parents=True, exist_ok=True)


⸻

src/readers/pdf_reader.py

# src/readers/pdf_reader.py
import pdfplumber
from pdf2image import convert_from_path
from pathlib import Path
from typing import List
from ..utils.logger import get_logger

logger = get_logger("pdf_reader")

def extract_text_pdf(path: Path) -> str:
    text_pages = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    text_pages.append(page.extract_text() or "")
                except Exception:
                    text_pages.append("")
    except Exception as e:
        logger.warning(f"pdfplumber failed: {e}; falling back to OCR")
    return "\n".join(text_pages).strip()

def rasterize_pages(path: Path, dpi:int=200) -> List:
    try:
        pages = convert_from_path(str(path), dpi=dpi)
        return pages
    except Exception as e:
        logger.error(f"pdf2image conversion failed: {e}")
        return []


⸻

src/readers/image_reader.py

# src/readers/image_reader.py
from PIL import Image
from pathlib import Path
import pytesseract
from ..utils.logger import get_logger

logger = get_logger("image_reader")

def ocr_image(path: Path, lang="eng") -> str:
    try:
        img = Image.open(path)
        text = pytesseract.image_to_string(img, lang=lang)
        return text
    except Exception as e:
        logger.error(f"OCR image failed {path}: {e}")
        return ""


⸻

src/readers/ocr_reader.py

# src/readers/ocr_reader.py
import pytesseract
from PIL import Image
import cv2
import numpy as np
from pathlib import Path
from ..utils.logger import get_logger

logger = get_logger("ocr_reader")

def preprocess_for_ocr(pil_img):
    arr = np.array(pil_img.convert('L'))
    # adaptive threshold
    th = cv2.adaptiveThreshold(arr,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,15,10)
    return Image.fromarray(th)

def ocr_image_to_text(pil_img, lang="eng"):
    try:
        img = preprocess_for_ocr(pil_img)
        return pytesseract.image_to_string(img, lang=lang)
    except Exception as e:
        logger.error(f"OCR failed: {e}")
        return ""

def ocr_table_from_image(pil_img, lang="eng"):
    """
    Best-effort approach: dilate to get blocks and OCR blocks as rows.
    Returns list of rows (each row is list of cell strings).
    """
    arr = np.array(pil_img.convert('L'))
    _, th = cv2.threshold(arr, 200, 255, cv2.THRESH_BINARY_INV)
    kernel = np.ones((15, 50), np.uint8)
    dil = cv2.dilate(th, kernel, iterations=1)
    cnts, _ = cv2.findContours(dil, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return []
    # sort by y
    cnts_sorted = sorted(cnts, key=lambda c: cv2.boundingRect(c)[1])
    rows = []
    for c in cnts_sorted:
        x,y,w,h = cv2.boundingRect(c)
        crop = arr[y:y+h, x:x+w]
        txt = pytesseract.image_to_string(Image.fromarray(crop), lang=lang)
        txt = txt.strip()
        rows.append([txt])
    return rows


⸻

src/readers/table_reader.py

# src/readers/table_reader.py
from typing import List, Dict, Any
from pathlib import Path
import pdfplumber
from pdf2image import convert_from_path
from ..readers.ocr_reader import ocr_table_from_image
from ..utils.logger import get_logger

logger = get_logger("table_reader")

def extract_tables_pdf(path: Path, dpi:int=200):
    tables = []
    try:
        with pdfplumber.open(path) as pdf:
            for i,page in enumerate(pdf.pages, start=1):
                try:
                    t = page.extract_tables()
                except Exception:
                    t = []
                if t:
                    for tab in t:
                        tables.append({"page": i, "table": tab})
                else:
                    # rasterize page and fallback to OCR
                    try:
                        pil_pages = convert_from_path(str(path), first_page=i, last_page=i, dpi=dpi)
                        if pil_pages:
                            rows = ocr_table_from_image(pil_pages[0])
                            if rows:
                                tables.append({"page": i, "table": rows})
                    except Exception as e:
                        logger.debug(f"rasterize fallback failed for page {i}: {e}")
    except Exception as e:
        logger.error(f"Failed to open PDF in table_reader: {e}")
    return tables

def normalize_table_cells(table: List[List[Any]]):
    norm = []
    for r in table:
        row = [("" if c is None else str(c).strip()) for c in r]
        norm.append(row)
    return norm


⸻

src/core/classifier.py

# src/core/classifier.py
import re

def classify_test(text: str) -> str:
    if not text:
        return "Other"
    txt = text.lower()
    if re.search(r'physical lab|physical properties|physical properties', txt):
        return "Physical Lab"
    if re.search(r'analytical|chemical|composition analysis', txt):
        return "Analytical/Chemical"
    if re.search(r'reinforcement lab|tensile properties|tyre construction|breaker', txt):
        return "Reinforcement Lab"
    # default
    return "Other"


⸻

src/core/parameter_manager.py

# src/core/parameter_manager.py
import json
from pathlib import Path
from rapidfuzz import fuzz
from ..utils.normalizer import normalize_param_name
from ..utils.logger import get_logger

logger = get_logger("parameter_manager")

MODELS_DIR = Path(__file__).resolve().parents[1] / "models"
STORE_FILE = MODELS_DIR / "parameter_store.json"
SYN_FILE = MODELS_DIR / "synonyms.json"

class ParameterManager:
    def __init__(self, threshold=85):
        MODELS_DIR.mkdir(parents=True, exist_ok=True)
        self.threshold = threshold
        self.store = self._load_json(STORE_FILE, default=[])
        self.synonyms = self._load_json(SYN_FILE, default={})
        # store is list of canonical names
        # synonyms maps variant->canonical

    def _load_json(self, p: Path, default):
        if p.exists():
            try:
                return json.loads(p.read_text(encoding='utf8'))
            except Exception:
                return default
        else:
            return default

    def _save_json(self, p: Path, data):
        p.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf8')

    def get_canonical(self, name: str) -> str:
        n = normalize_param_name(name)
        if not n:
            return ""
        # check synonyms first
        if n in self.synonyms:
            return self.synonyms[n]
        # fuzzy match against self.store
        best = None
        best_score = 0
        for can in self.store:
            score = fuzz.token_set_ratio(n, can)
            if score > best_score:
                best_score = score
                best = can
        if best_score >= self.threshold:
            # map and return
            self.synonyms[n] = best
            self._save_json(SYN_FILE, self.synonyms)
            return best
        # new canonical
        canonical = n
        self.store.append(canonical)
        self.synonyms[n] = canonical
        self._save_json(STORE_FILE, self.store)
        self._save_json(SYN_FILE, self.synonyms)
        logger.info(f"New parameter learned: '{name}' -> '{canonical}'")
        return canonical

    def list_parameters(self):
        return self.store[:]


⸻

src/core/extractor_engine.py

# src/core/extractor_engine.py
from pathlib import Path
from ..readers.pdf_reader import extract_text_pdf, rasterize_pages
from ..readers.table_reader import extract_tables_pdf, normalize_table_cells
from ..readers.image_reader import ocr_image
from ..readers.ocr_reader import ocr_image_to_text, ocr_table_from_image
from ..core.parameter_manager import ParameterManager
from ..core.classifier import classify_test
from ..utils.text_cleanup import cleanup_text
from ..utils.logger import get_logger
from PIL import Image
import shutil
import pdfplumber

logger = get_logger("extractor_engine")

class ExtractorEngine:
    def __init__(self, models_dir: Path = None, ocr_lang="eng"):
        self.pm = ParameterManager()
        self.ocr_lang = ocr_lang

    def extract_from_pdf(self, path: Path):
        # text
        text = extract_text_pdf(path)
        if not text:
            # fallback to OCR on rasterized pages
            pages = rasterize_pages(path, dpi=200)
            tparts = []
            for p in pages:
                tparts.append(ocr_image_to_text(p, lang=self.ocr_lang))
            text = "\n".join(tparts)
        text = cleanup_text(text)
        # HTAC
        import re
        htac = ""
        m = re.search(r'HTAC[_\s\.:-]*([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
        if m:
            htac = m.group(1)
        testname = classify_test(text)
        # tables
        tables = extract_tables_pdf(path)
        kv = {}
        for t in tables:
            table = normalize_table_cells(t['table'])
            # heuristics for two-column parameter-value tables
            if all(len(r) >= 2 for r in table) and max(len(r) for r in table) == 2:
                for r in table:
                    k = r[0].strip()
                    v = r[1].strip()
                    if k:
                        can = self.pm.get_canonical(k)
                        kv[can] = kv.get(can, "") + (("; " + v) if kv.get(can) else v)
            else:
                # try header mapping
                header = None
                for r in table:
                    nonempties = [c for c in r if c.strip()]
                    if nonempties and len(nonempties) > 1:
                        header = r
                        break
                if header:
                    hidx = table.index(header)
                    headers = [c.strip() for c in header]
                    for row in table[hidx+1:]:
                        for ci, cell in enumerate(row):
                            colname = headers[ci] if ci < len(headers) else f"col{ci+1}"
                            can = self.pm.get_canonical(colname)
                            newv = cell.strip()
                            if newv:
                                prev = kv.get(can, "")
                                kv[can] = prev + (" ; " + newv if prev else newv)
                else:
                    # flatten
                    for ri,r in enumerate(table, start=1):
                        for ci,cell in enumerate(r, start=1):
                            can = self.pm.get_canonical(f"table_r{ri}_c{ci}")
                            kv[can] = cell.strip()

        # simple key:value in raw text
        for line in text.splitlines():
            if ':' in line:
                parts = line.split(':',1)
                k = parts[0].strip()
                v = parts[1].strip()
                if k and v:
                    can = self.pm.get_canonical(k)
                    kv[can] = kv.get(can, "") + (("; " + v) if kv.get(can) else v)

        # images
        images_paths = []
        try:
            with pdfplumber.open(path) as pdf:
                for i,page in enumerate(pdf.pages, start=1):
                    if page.images:
                        im = page.to_image(resolution=150)
                        # save entire page as fallback image but grouped per page
                        folder = Path("output/images") / (htac or path.stem)
                        folder.mkdir(parents=True, exist_ok=True)
                        out = folder / f"{path.stem}_page{i}.png"
                        im.original.save(out)
                        images_paths.append(str(out))
        except Exception as e:
            logger.debug(f"pdf image extraction issue: {e}")

        return {
            "SourceFile": str(path.resolve()),
            "HTAC_No": htac,
            "TestName": testname,
            "AllText": text,
            "Images": images_paths,
            "KV": kv
        }

    def extract_from_image(self, path: Path):
        text = ocr_image(path, lang=self.ocr_lang)
        text = cleanup_text(text)
        htac = ""
        import re
        m = re.search(r'HTAC[_\s\.:-]*([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
        if m:
            htac = m.group(1)
        testname = classify_test(text)
        # OCR table
        try:
            with Image.open(path) as img:
                table_rows = ocr_table_from_image(img, lang=self.ocr_lang)
        except Exception:
            table_rows = []
        kv = {}
        for r in table_rows:
            if len(r) >= 2:
                k = r[0].strip()
                v = r[1].strip()
                if k:
                    can = self.pm.get_canonical(k)
                    kv[can] = kv.get(can, "") + (("; " + v) if kv.get(can) else v)
        # add line pairs
        for line in text.splitlines():
            if ':' in line:
                parts = line.split(':',1)
                k = parts[0].strip()
                v = parts[1].strip()
                if k and v:
                    can = self.pm.get_canonical(k)
                    kv[can] = kv.get(can, "") + (("; " + v) if kv.get(can) else v)
        # save image to output images folder
        folder = Path("output/images") / (htac or path.stem)
        folder.mkdir(parents=True, exist_ok=True)
        out = folder / path.name
        shutil.copyfile(path, out)
        images_paths = [str(out)]
        return {
            "SourceFile": str(path.resolve()),
            "HTAC_No": htac,
            "TestName": testname,
            "AllText": text,
            "Images": images_paths,
            "KV": kv
        }


⸻

src/core/pipeline.py

# src/core/pipeline.py
from pathlib import Path
import pandas as pd
from .extractor_engine import ExtractorEngine
from ..utils.logger import get_logger
from ..utils.file_paths import ensure_dirs
import os

logger = get_logger("pipeline")

class Pipeline:
    def __init__(self, input_dir: str, output_excel: str, ocr_lang="eng"):
        ensure_dirs()
        self.input_dir = Path(input_dir)
        self.output_excel = Path(output_excel)
        self.engine = ExtractorEngine(ocr_lang=ocr_lang)

    def run(self):
        rows = []
        files = sorted(self.input_dir.glob("*"))
        logger.info(f"Found {len(files)} files in {self.input_dir}")
        for f in files:
            if f.is_dir():
                continue
            logger.info(f"Processing: {f.name}")
            if f.suffix.lower() == ".pdf":
                doc = self.engine.extract_from_pdf(f)
            else:
                doc = self.engine.extract_from_image(f)
            rows.append(doc)
        # consolidate into sheets by TestName
        sheets = {}
        for doc in rows:
            test = doc.get("TestName","Other") or "Other"
            if test not in sheets:
                sheets[test] = []
            # flatten record
            rec = {
                "SourceFile": doc.get("SourceFile"),
                "HTAC_No": doc.get("HTAC_No"),
                "TestName": doc.get("TestName"),
                "AllText": doc.get("AllText"),
                "Images": ";".join(doc.get("Images",[]))
            }
            # add learned params
            kv = doc.get("KV", {})
            for k,v in kv.items():
                rec[k] = v
            sheets[test].append(rec)
        # build pandas Excel writer with multiple sheets
        writer = pd.ExcelWriter(self.output_excel, engine='openpyxl')
        for sheetname, recs in sheets.items():
            df = pd.DataFrame(recs)
            # ensure stable column order: SourceFile, HTAC_No, TestName, AllText, Images, then others
            base = ["SourceFile","HTAC_No","TestName","AllText","Images"]
            others = [c for c in df.columns if c not in base]
            df = df[base + sorted(others)]
            df.to_excel(writer, sheet_name=self._safe_sheetname(sheetname), index=False)
        writer.save()
        logger.info(f"Wrote Excel to {self.output_excel}")

    def _safe_sheetname(self, name: str) -> str:
        # excel sheet name safe
        name = name[:31]
        return name.replace("/", "_")


⸻

src/core/__init__.py

(empty or optional)

⸻

src/readers/__init__.py

(empty or optional)

⸻

src/models/parameter_store.json

[]


⸻

src/models/synonyms.json

{}


⸻

app/cli.py

# app/cli.py
import argparse
from pathlib import Path
from ..src.core.pipeline import Pipeline
from ..src.utils.file_paths import ensure_dirs
from ..src.utils.logger import get_logger

logger = get_logger("cli")

def main():
    parser = argparse.ArgumentParser(description="Tyre PDF/Image extractor")
    parser.add_argument("--input", "-i", type=str, default="input_files", help="Input folder path")
    parser.add_argument("--output", "-o", type=str, default="output/consolidated.xlsx", help="Output excel path")
    parser.add_argument("--ocr-lang", type=str, default="eng", help="tesseract language")
    args = parser.parse_args()
    ensure_dirs()
    p = Pipeline(args.input, args.output, ocr_lang=args.ocr_lang)
    p.run()

if __name__ == "__main__":
    main()


⸻

app/run_pipeline.py

# app/run_pipeline.py
# convenience runner
from .cli import main

if __name__ == "__main__":
    main()


⸻

How to run
	1.	Put your PDF/image files into input_files/.
	2.	From project root run:

python -m app.cli --input input_files --output output/consolidated.xlsx

(or python app/cli.py ...)

	3.	After run:
	•	output/consolidated.xlsx — workbook with sheets per Test type
	•	output/images/<HTAC or filename>/ — extracted images
	•	src/models/parameter_store.json + src/models/synonyms.json — parameter library that learns new parameters automatically

⸻

Extra tips & tuning
	•	If you have regular known variants (e.g., C, Carbon, Carbon Black), add them to src/models/synonyms.json manually in the form:

{
  "c": "carbon",
  "carbon black": "carbon"
}

The ParameterManager will pick these up on load.

	•	For high-quality scanned PDFs, increase dpi in pdf_reader.rasterize_pages and table_reader.extract_tables_pdf.
	•	The fuzzy threshold for matching is hard-coded via rapidfuzz default 85. You can lower this in ParameterManager if you want more merging.

⸻

Want me to generate files for you now?

I prepared the complete project code above. If you want, I can:
	•	(A) Provide a downloadable ZIP of the project files, or
	•	(B) Run the pipeline on the images you already uploaded and produce consolidated.xlsx + zipped output/images/ for you to download.

Say “Generate ZIP” to get the project as a downloadable zip, or “Process my uploaded files” and I will run the pipeline on the images/PDFs you already uploaded and return the resulting Excel and images folder.
