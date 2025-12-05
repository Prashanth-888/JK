Great — below is the entire updated project with the Schema Manager and structured table extraction fully integrated.
This is the complete codebase: copy each file into the exact path shown, install dependencies from requirements.txt, and run as before.

I kept names consistent with the previous project and included the new schema_manager.py. All strings are cleaned automatically and the pipeline will produce sheets with the same global schema (including structured table keys like table1_row1_col1 or table1_rowName_colName).

⸻

Project structure (copy exactly)

tyre_extractor/
├── app/
│   ├── cli.py
│   └── __init__.py
│
├── src/
│   ├── core/
│   │   ├── pipeline.py
│   │   ├── extractor_engine.py
│   │   ├── schema_manager.py
│   │   ├── classifier.py
│   │   ├── parameter_manager.py
│   │   └── __init__.py
│   ├── readers/
│   │   ├── pdf_reader.py
│   │   ├── ocr_reader.py
│   │   ├── table_reader.py
│   │   ├── image_reader.py
│   │   └── __init__.py
│   ├── utils/
│   │   ├── safe.py
│   │   ├── text_cleanup.py
│   │   ├── normalizer.py
│   │   ├── file_paths.py
│   │   ├── logger.py
│   │   └── __init__.py
│   └── models/
│       ├── parameter_store.json
│       ├── synonyms.json
│       └── __init__.py
│
├── input_files/
├── output/
│   └── images/
├── requirements.txt
└── README.md


⸻

app/init.py

# app/__init__.py
# package marker


⸻

app/cli.py

# app/cli.py
import argparse
from src.core.pipeline import Pipeline
from src.utils.file_paths import ensure_dirs

def main():
    parser = argparse.ArgumentParser(description="Tyre Test Report Extractor")
    parser.add_argument("--input", "-i", type=str, default="input_files",
                        help="Folder that contains PDF/Image files")
    parser.add_argument("--output", "-o", type=str, default="output/consolidated.xlsx",
                        help="Output Excel file path")
    parser.add_argument("--ocr-lang", type=str, default="eng",
                        help="OCR language for tesseract")
    args = parser.parse_args()

    ensure_dirs()
    pipeline = Pipeline(args.input, args.output, ocr_lang=args.ocr_lang)
    pipeline.run()

if __name__ == "__main__":
    main()


⸻

src/core/init.py

# src/core/__init__.py
# package marker


⸻

src/readers/init.py

# src/readers/__init__.py
# package marker


⸻

src/utils/init.py

# src/utils/__init__.py
# package marker


⸻

src/models/init.py

# src/models/__init__.py
# package marker


⸻

src/utils/logger.py

# src/utils/logger.py
import logging
import sys

def get_logger(name="logger", level=logging.INFO):
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        fmt = '%(asctime)s - %(levelname)s - %(name)s - %(message)s'
        handler.setFormatter(logging.Formatter(fmt))
        logger.addHandler(handler)
        logger.setLevel(level)
    return logger


⸻

src/utils/text_cleanup.py

# src/utils/text_cleanup.py
import re

_ILLEGAL = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F\u200B\u200C\u200D\uFEFF]")

def clean_string(text: str) -> str:
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _ILLEGAL.sub("", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


⸻

src/utils/safe.py

# src/utils/safe.py
from src.utils.text_cleanup import clean_string

def safe_val(v):
    try:
        return clean_string(v)
    except Exception:
        try:
            return clean_string(str(v))
        except Exception:
            return ""


⸻

src/utils/normalizer.py

# src/utils/normalizer.py
import re

def normalize_param_name(s: str) -> str:
    if not s:
        return ""
    s = str(s).lower().strip()
    s = re.sub(r'\s+', ' ', s)
    s = s.rstrip(':')
    s = re.sub(r'[^\w\s\/\-_\.]', '', s)
    return s.strip()


⸻

src/utils/file_paths.py

# src/utils/file_paths.py
from pathlib import Path

def ensure_dirs():
    Path("input_files").mkdir(exist_ok=True)
    Path("output/images").mkdir(parents=True, exist_ok=True)
    Path("src/models").mkdir(parents=True, exist_ok=True)


⸻

src/core/classifier.py

# src/core/classifier.py
import re

def classify_test(text: str) -> str:
    if not text:
        return "Other"
    txt = text.lower()
    if "physical" in txt or "physical properties" in txt:
        return "Physical Lab"
    if "analytical" in txt or "chemical" in txt or "composition" in txt:
        return "Analytical/Chemical"
    if "reinforcement" in txt or "tensile" in txt or "breaker" in txt:
        return "Reinforcement Lab"
    return "Other"


⸻

src/core/parameter_manager.py

# src/core/parameter_manager.py
import json
from pathlib import Path
from rapidfuzz import fuzz
from src.utils.normalizer import normalize_param_name
from src.utils.logger import get_logger

logger = get_logger("parameter_manager")

MODELS = Path("src/models")
STORE_FILE = MODELS / "parameter_store.json"
SYN_FILE = MODELS / "synonyms.json"

class ParameterManager:
    def __init__(self, threshold: int = 85):
        MODELS.mkdir(parents=True, exist_ok=True)
        self.threshold = threshold
        self.store = self._load_json(STORE_FILE, default=[])
        self.synonyms = self._load_json(SYN_FILE, default={})

    def _load_json(self, p: Path, default):
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf8"))
            except Exception:
                return default
        return default

    def _save_json(self, p: Path, data):
        p.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf8")

    def get_canonical(self, name: str) -> str:
        n = normalize_param_name(name)
        if not n:
            return ""
        # synonyms override
        if n in self.synonyms:
            return self.synonyms[n]
        # fuzzy match to existing store
        best_score = 0
        best_can = None
        for can in self.store:
            score = fuzz.token_set_ratio(n, can)
            if score > best_score:
                best_score = score
                best_can = can
        if best_score >= self.threshold:
            self.synonyms[n] = best_can
            self._save_json(SYN_FILE, self.synonyms)
            return best_can
        # create new canonical
        canonical = n
        self.store.append(canonical)
        self.synonyms[n] = canonical
        self._save_json(STORE_FILE, self.store)
        self._save_json(SYN_FILE, self.synonyms)
        logger.info(f"New parameter learned: '{name}' -> '{canonical}'")
        return canonical

    def list_parameters(self):
        return list(self.store)


⸻

src/core/schema_manager.py

# src/core/schema_manager.py
from typing import List

class SchemaManager:
    def __init__(self):
        self.columns = set()

    def add_doc(self, doc: dict):
        basics = ["SourceFile", "HTAC_No", "TestName", "AllText", "Images"]
        for b in basics:
            self.columns.add(b)
        # add KV keys
        for k in doc.get("KV", {}).keys():
            self.columns.add(k)

    def add_columns(self, cols: List[str]):
        for c in cols:
            self.columns.add(c)

    def get_schema(self):
        basics = ["SourceFile", "HTAC_No", "TestName", "AllText", "Images"]
        others = sorted([c for c in self.columns if c not in basics])
        return basics + others


⸻

src/readers/pdf_reader.py

# src/readers/pdf_reader.py
import pdfplumber
from pdf2image import convert_from_path
from pathlib import Path
from src.utils.logger import get_logger

logger = get_logger("pdf_reader")

def extract_text_pdf(path: Path) -> str:
    pages_text = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    pages_text.append(page.extract_text() or "")
                except Exception:
                    pages_text.append("")
    except Exception as e:
        logger.warning(f"pdfplumber error: {e}")
    return "\n".join(pages_text)

def rasterize_pages(path: Path, dpi:int=200):
    try:
        return convert_from_path(str(path), dpi=dpi)
    except Exception as e:
        logger.error(f"pdf2image conversion failed: {e}")
        return []


⸻

src/readers/ocr_reader.py

# src/readers/ocr_reader.py
import pytesseract
import cv2
import numpy as np
from PIL import Image
from src.utils.logger import get_logger

logger = get_logger("ocr_reader")

def ocr_image_to_text(pil_img, lang="eng"):
    try:
        arr = np.array(pil_img.convert("L"))
        arr = cv2.adaptiveThreshold(arr, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY, 15, 10)
        return pytesseract.image_to_string(Image.fromarray(arr), lang=lang)
    except Exception as e:
        logger.error(f"OCR text failed: {e}")
        return ""

def ocr_table_from_image(pil_img, lang="eng"):
    try:
        arr = np.array(pil_img.convert("L"))
        _, th = cv2.threshold(arr, 200, 255, cv2.THRESH_BINARY_INV)
        # tune kernel sizes if your tables are narrow/wide
        kernel = np.ones((15, 80), np.uint8)
        dil = cv2.dilate(th, kernel, iterations=1)
        contours, _ = cv2.findContours(dil, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        rows = []
        for cnt in sorted(contours, key=lambda c: cv2.boundingRect(c)[1]):
            x,y,w,h = cv2.boundingRect(cnt)
            crop = arr[y:y+h, x:x+w]
            text = pytesseract.image_to_string(Image.fromarray(crop), lang=lang).strip()
            rows.append([text])
        return rows
    except Exception as e:
        logger.error(f"OCR table failed: {e}")
        return []


⸻

src/readers/image_reader.py

# src/readers/image_reader.py
from PIL import Image
import pytesseract
from pathlib import Path
from src.utils.logger import get_logger

logger = get_logger("image_reader")

def ocr_image(path: Path, lang="eng") -> str:
    try:
        img = Image.open(path)
        return pytesseract.image_to_string(img, lang=lang)
    except Exception as e:
        logger.error(f"OCR image failed: {e}")
        return ""


⸻

src/readers/table_reader.py

# src/readers/table_reader.py
import pdfplumber
from pdf2image import convert_from_path
from src.readers.ocr_reader import ocr_table_from_image
from src.utils.logger import get_logger

logger = get_logger("table_reader")

def extract_tables_pdf(path, dpi:int=200):
    tables = []
    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                try:
                    extracted = page.extract_tables()
                except Exception:
                    extracted = []
                if extracted:
                    for t in extracted:
                        tables.append({"page": i, "table": t})
                else:
                    try:
                        pil = convert_from_path(str(path), first_page=i, last_page=i, dpi=dpi)
                        rows = ocr_table_from_image(pil[0])
                        if rows:
                            tables.append({"page": i, "table": rows})
                    except Exception as e:
                        logger.debug(f"fallback rasterize failed page {i}: {e}")
    except Exception as e:
        logger.error(f"Table extraction error: {e}")
    return tables

def normalize_table_cells(table):
    norm = []
    for r in table:
        norm.append([("" if c is None else str(c).strip()) for c in r])
    return norm


⸻

src/core/extractor_engine.py

# src/core/extractor_engine.py
from pathlib import Path
import re
import shutil
import pdfplumber
from PIL import Image
from src.readers.pdf_reader import extract_text_pdf, rasterize_pages
from src.readers.table_reader import extract_tables_pdf, normalize_table_cells
from src.readers.ocr_reader import ocr_image_to_text, ocr_table_from_image
from src.readers.image_reader import ocr_image
from src.core.parameter_manager import ParameterManager
from src.core.classifier import classify_test
from src.utils.safe import safe_val
from src.utils.text_cleanup import clean_string
from src.utils.logger import get_logger

logger = get_logger("extractor_engine")

class ExtractorEngine:
    def __init__(self, ocr_lang="eng"):
        self.pm = ParameterManager()
        self.ocr_lang = ocr_lang

    def extract_structured_table(self, table, table_index: int):
        """
        Convert table into keys:
          table{index}_{rowname}_{colname}
        or
          table{index}_row{r}_col{c}
        """
        results = {}
        if not table:
            return results
        rows = len(table)
        cols = max(len(r) for r in table) if rows else 0
        table_name = f"table{table_index}"

        # detect header (first non-empty row with multiple entries)
        header = None
        for r in table:
            nonempty = [c for c in r if c and str(c).strip()]
            if nonempty and len(nonempty) > 1:
                header = r
                break

        # detect row names (first column non-empty in most rows)
        first_col = [r[0] if len(r)>0 else "" for r in table]
        has_row_names = sum(1 for x in first_col if x and str(x).strip()) > (len(table) // 2)

        for ri, row in enumerate(table, start=1):
            for ci in range(cols):
                cell = ""
                try:
                    cell = row[ci]
                except Exception:
                    cell = ""
                if cell is None:
                    cell = ""
                cell = str(cell).strip()
                if not cell:
                    continue

                # determine row name
                if has_row_names:
                    row_name = str(row[0]).strip() or f"row{ri}"
                else:
                    row_name = f"row{ri}"

                # determine col name
                if header:
                    try:
                        col_val = header[ci] if ci < len(header) else f"col{ci+1}"
                        col_name = str(col_val).strip() or f"col{ci+1}"
                    except:
                        col_name = f"col{ci+1}"
                else:
                    col_name = f"col{ci+1}"

                key = f"{table_name}_{row_name}_{col_name}".lower()
                key = key.replace(" ", "_")
                results[key] = cell
        return results

    def extract_from_pdf(self, path: Path):
        text = extract_text_pdf(path)
        if not text:
            pages = rasterize_pages(path)
            text_parts = [ocr_image_to_text(p, self.ocr_lang) for p in pages]
            text = "\n".join(text_parts)
        text = clean_string(text)

        htac = ""
        m = re.search(r'HTAC[_\s\.:=-]*([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
        if m:
            htac = m.group(1)

        testname = classify_test(text)

        tables = extract_tables_pdf(path)
        kv = {}
        table_idx = 1
        for t in tables:
            table = normalize_table_cells(t.get("table", []))
            structured = self.extract_structured_table(table, table_idx)
            table_idx += 1
            for k, v in structured.items():
                can = self.pm.get_canonical(k)
                kv[can] = kv.get(can, "") + ("; " + v if kv.get(can) else v)

        # also try to parse two-column tables (parameter : value) from text
        for line in text.splitlines():
            if ":" in line:
                left,right = line.split(":",1)
                k = safe_val(left)
                v = safe_val(right)
                if k:
                    can = self.pm.get_canonical(k)
                    kv[can] = kv.get(can, "") + ("; " + v if kv.get(can) else v)

        # extract images to output/images/<HTAC or filename>/
        outdir = Path("output/images") / (htac or path.stem)
        outdir.mkdir(parents=True, exist_ok=True)
        images_paths = []
        try:
            with pdfplumber.open(path) as pdf:
                for i, page in enumerate(pdf.pages, start=1):
                    try:
                        im = page.to_image(resolution=150)
                        out = outdir / f"{path.stem}_page{i}.png"
                        im.original.save(out)
                        images_paths.append(str(out))
                    except Exception:
                        # fallback: rasterize page and save
                        try:
                            pil_pages = rasterize_pages(path, dpi=200)
                            if pil_pages and i-1 < len(pil_pages):
                                out = outdir / f"{path.stem}_page{i}.png"
                                pil_pages[i-1].save(out)
                                images_paths.append(str(out))
                        except Exception:
                            pass
        except Exception:
            pass

        # safe-ify final data
        final_kv = {safe_val(k): safe_val(v) for k,v in kv.items()}

        return {
            "SourceFile": safe_val(str(path.resolve())),
            "HTAC_No": safe_val(htac),
            "TestName": safe_val(testname),
            "AllText": safe_val(text),
            "Images": [safe_val(p) for p in images_paths],
            "KV": final_kv
        }

    def extract_from_image(self, path: Path):
        text = ocr_image(path, self.ocr_lang)
        text = clean_string(text)

        htac = ""
        m = re.search(r'HTAC[_\s\.:=-]*([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
        if m:
            htac = m.group(1)

        testname = classify_test(text)

        # detect table-like regions via OCR reader
        try:
            img = Image.open(path)
            rows = ocr_table_from_image(img, self.ocr_lang)
        except Exception:
            rows = []

        kv = {}
        table_idx = 1
        if rows:
            # rows is list of [text] blocks; attempt to split columns by separators if present
            # We will treat each row as a single cell row; structured key generation still applies
            table = [[cell for cell in r] for r in rows]
            structured = self.extract_structured_table(table, table_idx)
            for k,v in structured.items():
                can = self.pm.get_canonical(k)
                kv[can] = kv.get(can, "") + ("; " + v if kv.get(can) else v)

        # parse text k:v lines
        for line in text.splitlines():
            if ":" in line:
                left,right = line.split(":",1)
                k = safe_val(left)
                v = safe_val(right)
                if k:
                    can = self.pm.get_canonical(k)
                    kv[can] = kv.get(can, "") + ("; " + v if kv.get(can) else v)

        # save image
        outdir = Path("output/images") / (htac or path.stem)
        outdir.mkdir(parents=True, exist_ok=True)
        out_path = outdir / path.name
        try:
            shutil.copyfile(path, out_path)
        except Exception:
            pass

        final_kv = {safe_val(k): safe_val(v) for k,v in kv.items()}

        return {
            "SourceFile": safe_val(str(path.resolve())),
            "HTAC_No": safe_val(htac),
            "TestName": safe_val(testname),
            "AllText": safe_val(text),
            "Images": [safe_val(str(out_path))],
            "KV": final_kv
        }


⸻

src/core/pipeline.py

# src/core/pipeline.py
from pathlib import Path
import pandas as pd
from src.core.extractor_engine import ExtractorEngine
from src.core.schema_manager import SchemaManager
from src.utils.safe import safe_val
from src.utils.file_paths import ensure_dirs
from src.utils.logger import get_logger

logger = get_logger("pipeline")

class Pipeline:
    def __init__(self, input_dir: str, output_excel: str, ocr_lang="eng"):
        ensure_dirs()
        self.input_dir = Path(input_dir)
        self.output_excel = Path(output_excel)
        self.engine = ExtractorEngine(ocr_lang=ocr_lang)
        self.schema = SchemaManager()

    def run(self):
        docs = []
        files = sorted(self.input_dir.glob("*"))
        logger.info(f"Found {len(files)} files in {self.input_dir}")

        for f in files:
            if not f.is_file():
                continue
            logger.info(f"Processing {f.name}")
            if f.suffix.lower() == ".pdf":
                doc = self.engine.extract_from_pdf(f)
            else:
                doc = self.engine.extract_from_image(f)
            docs.append(doc)
            self.schema.add_doc(doc)

        # create unified schema that includes learned parameters and structural keys
        columns = self.schema.get_schema()

        # Group docs by TestName
        sheets = {}
        for doc in docs:
            test = safe_val(doc.get("TestName", "Other"))
            sheets.setdefault(test, []).append(doc)

        writer = pd.ExcelWriter(self.output_excel, engine="openpyxl")

        for sheetname, docs_list in sheets.items():
            rows = []
            for doc in docs_list:
                base = {
                    "SourceFile": safe_val(doc.get("SourceFile")),
                    "HTAC_No": safe_val(doc.get("HTAC_No")),
                    "TestName": safe_val(doc.get("TestName")),
                    "AllText": safe_val(doc.get("AllText")),
                    "Images": safe_val(";".join(doc.get("Images", [])))
                }
                # initialize all columns to empty
                rec = {c: "" for c in columns}
                rec.update({k: v for k, v in base.items()})
                # fill KV params
                for k, v in doc.get("KV", {}).items():
                    rec[safe_val(k)] = safe_val(v)
                rows.append(rec)

            df = pd.DataFrame(rows, columns=columns)
            df = df.applymap(lambda x: safe_val(x))
            df.to_excel(writer, sheet_name=sheetname[:31], index=False)

        writer.save()
        logger.info(f"Saved Excel to {self.output_excel}")


⸻

src/models/parameter_store.json

[]


⸻

src/models/synonyms.json

{}


⸻

requirements.txt

pdfplumber
pdf2image
pytesseract
Pillow
opencv-python
pandas
openpyxl
rapidfuzz


⸻

README.md

# Tyre Test Extraction System (Updated with Schema Manager)

## Overview
Extracts text/tables/images from PDFs and images, learns parameters, and writes a multi-sheet Excel workbook with a **global unified schema** so every sheet has the same columns. Tables are converted to structured keys: `table{n}_row{...}_col{...}` or `table{n}_{rowName}_{colName}`.

## Setup
1. Install Python 3.8+.
2. Install system deps:
   - Poppler (for pdf2image)
   - Tesseract OCR
3. Create and activate venv:
   - `python -m venv venv`
   - `venv\Scripts\activate` (Windows) or `source venv/bin/activate` (Mac/Linux)
4. `pip install -r requirements.txt`
5. Put your files in `input_files/`
6. Run:

python -m app.cli –input input_files –output output/consolidated.xlsx

## Notes
- All extracted strings are cleaned for Excel compatibility.
- Table extraction heuristics attempt to find headers and row labels; if none found, keys default to `table{n}_row{i}_col{j}`.
- The project stores learned parameter canonical names in `src/models/parameter_store.json` and `src/models/synonyms.json`.


⸻

Quick Run reminder

From the project root:
	1.	Ensure input_files/ contains your PDFs/images.
	2.	Activate venv and ensure Poppler & Tesseract installed system-wide.
	3.	Run:

python -m app.cli --input input_files --output output/consolidated.xlsx


⸻

If you want I can do one of the following next (pick one):
	•	Package everything into a ZIP and provide a link (say “Give me ZIP”).
	•	Run this pipeline on the files you’ve already uploaded here and return the generated consolidated.xlsx and zipped output/images/ (say “Process uploaded files”).
	•	Tune table-detection kernel sizes and OCR parameters specifically for a sample file — paste one file name or upload 1-2 representative PDFs and I’ll suggest optimized settings and update the code accordingly.

Which of these would you like?
