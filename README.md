# JK
Great — below I’m pasting the complete, production-ready project (all files) straight in chat, ready for you to copy into a folder and run. I used the behavior we agreed on:
	•	canonical/fuzzy merging across files
	•	table extraction that uses row name (first column) + column header + test category as the final column name
	•	final column format: testcategory_ColumnHeader_RowName (normalized & canonicalized)
	•	append mode: new runs append new files into the same Excel (no duplicate rows by source_file)
	•	images saved into output/extracted_images/<HTAC>/... and image paths stored in the sheet
	•	OCR confidence and extract status saved (OK / LOW_CONFIDENCE / FAILED)
	•	fallback methods for reading PDFs & images
	•	simple logging

I chose auto-normalized test categories (recommended): physical, chemical, analytical, dimension, valve, general — you can edit mappings in config.py.

⸻

How to use (short)
	1.	Create folder test_report_extractor/ and subfolders extractor/, input_reports/, output/.
	2.	Copy each file below to its path.
	3.	Install OS system packages (Tesseract, Ghostscript) then Python deps (requirements.txt).
	4.	Put PDF/image files into input_reports/.
	5.	Run python main.py.
	6.	Outputs: output/extracted_report_final.xlsx, output/extracted_report_final.csv, output/extraction_issues.csv, output/extracted_images/.

⸻

Project file list (copy each file into the given path)

Create folder structure:

test_report_extractor/
├── main.py
├── config.py
├── requirements.txt
├── README.md
├── extractor/
│   ├── __init__.py
│   ├── logger.py
│   ├── file_utils.py
│   ├── text_parser.py
│   ├── pdf_reader.py
│   ├── image_reader.py
│   ├── table_extractor.py
│   ├── canonicalizer.py
│   └── merger.py
└── input_reports/   # put your files here


⸻

requirements.txt

pandas
openpyxl
pdfplumber
pytesseract
Pillow
camelot-py[cv]
pdf2image
rapidfuzz

Note: camelot requires ghostscript and tk installed at the system level.

⸻

config.py

# config.py
# -----------------------------
# Tweak these settings for your environment
# -----------------------------

# OCR thresholds
MIN_TEXT_LENGTH = 30           # minimum text length to be considered readable
OCR_OK_CONFIDENCE = 40         # avg OCR confidence >= this -> OK
OCR_LOW_CONFIDENCE = 25        # avg >= this and <OK -> LOW_CONFIDENCE

# Fuzzy threshold for merging column names (0-100)
FUZZY_THRESHOLD = 85

# Canonical mapping. Keys should be normalized (lower, no punctuation).
# Add variants as needed. Left side: normalized key; right side: desired canonical name.
CANONICAL_MAP = {
    "carbonblack": "Carbon Black (%)",
    "carbon": "Carbon Black (%)",
    "c": "Carbon Black (%)",
    "polymer": "Polymer (%)",
    "polymer%": "Polymer (%)",
    "ash": "Ash (%)",
    "ash%": "Ash (%)",
    "wt": "Weight (g)",
    "weight": "Weight (g)",
    "tensilestrength": "Tensile Strength (MPa)",
    "stressat100elongation": "Stress @100% Elongation (kg/cm2)",
    "stressat300elongation": "Stress @300% Elongation (kg/cm2)",
    # Add more mappings from examples you frequently see
}

# Test name normalization map:
TEST_NAME_MAP = {
    "physical": "physical",
    "physicalproperties": "physical",
    "physical properties": "physical",
    "chemical": "chemical",
    "composition": "chemical",
    "analytical": "analytical",
    "analysis": "analytical",
    "dimension": "dimension",
    "valve": "valve",
    "general": "general",
}


⸻

extractor/__init__.py

# extractor package


⸻

extractor/logger.py

# extractor/logger.py
import logging

def get_logger(name="extractor"):
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s | %(levelname)s | %(message)s")
    return logging.getLogger(name)


⸻

extractor/file_utils.py

# extractor/file_utils.py
import os
from pathlib import Path

def ensure_dir(p):
    Path(p).mkdir(parents=True, exist_ok=True)

def list_supported_files(input_dir):
    exts = {'.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp'}
    files = []
    for root, _, filenames in os.walk(input_dir):
        for fn in filenames:
            if Path(fn).suffix.lower() in exts:
                files.append(os.path.join(root, fn))
    return files


⸻

extractor/text_parser.py

# extractor/text_parser.py
import re
from PIL import ImageOps
import pytesseract

HTAC_REGEXES = [
    r'HTAC\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\-_]+)',
    r'HTAC_No\.?\s*[:\-]?\s*([A-Za-z0-9\-_]+)',
    r'HTAC No[:\s]*([A-Za-z0-9\-_]+)',
    r'TC[-_\s]?\d{3,,}'
]

def extract_kv_pairs(text):
    kv = {}
    if not text:
        return kv
    for line in text.splitlines():
        line = line.strip()
        if not line or len(line) > 300:
            continue
        # look for "Key : Value" or "Key - Value"
        m = re.match(r'([\w\-\s\(\)%/\.]+?)\s*[:\-]\s*(.+)', line)
        if m:
            k = ' '.join(m.group(1).split())
            v = m.group(2).strip()
            kv[k] = v
    return kv

def find_htac(text):
    if not text:
        return None
    for pattern in HTAC_REGEXES:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None

def find_test_name(text):
    if not text:
        return 'general'
    t = text.lower()
    if 'physical' in t:
        return 'physical'
    if 'chemical' in t or 'composition' in t:
        return 'chemical'
    if 'analytical' in t or 'analysis' in t:
        return 'analytical'
    if 'dimension' in t or 'dimension analysis' in t:
        return 'dimension'
    if 'valve' in t:
        return 'valve'
    return 'general'

def ocr_image_get_text_and_conf(pil_img):
    """
    Returns (text_string, average_confidence) for a PIL image using pytesseract.
    """
    gray = ImageOps.grayscale(pil_img)
    data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT, lang='eng')
    texts = []
    confs = []
    n = len(data['text'])
    for i in range(n):
        txt = (data['text'][i] or '').strip()
        try:
            conf = float(data['conf'][i])
        except Exception:
            conf = -1.0
        if txt:
            texts.append(txt)
            if conf >= 0:
                confs.append(conf)
    avg_conf = sum(confs)/len(confs) if confs else -1.0
    return ' '.join(texts), avg_conf


⸻

extractor/pdf_reader.py

# extractor/pdf_reader.py
import pdfplumber
from extractor.text_parser import ocr_image_get_text_and_conf
from extractor.file_utils import ensure_dir
import os

def extract_text_and_conf_from_pdf(path, ocr_resolution=200):
    """
    Try pdfplumber text extraction first; if empty use OCR per page.
    Returns: (text_combined, avg_confidence)
    """
    text_all = []
    confs = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ''
                if txt and len(txt.strip()) > 10:
                    text_all.append(txt)
                    confs.append(95.0)  # assume high for digital text
                else:
                    img = page.to_image(resolution=ocr_resolution).original
                    t, c = ocr_image_get_text_and_conf(img)
                    text_all.append(t)
                    if c >= 0:
                        confs.append(c)
    except Exception:
        # fallback: try OCR of all pages if pdfplumber had trouble
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    img = page.to_image(resolution=ocr_resolution).original
                    t, c = ocr_image_get_text_and_conf(img)
                    text_all.append(t)
                    if c >= 0:
                        confs.append(c)
        except Exception:
            return '', -1.0
    combined = '\n'.join(text_all).strip()
    avg_conf = (sum([c for c in confs if c>0]) / len([c for c in confs if c>0])) if any(c>0 for c in confs) else -1.0
    return combined, avg_conf

def extract_images_from_pdf(path, out_dir, htac):
    saved = []
    ensure_dir(out_dir)
    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                try:
                    img = page.to_image(resolution=150).original
                    fname = f"{htac}_page{i}.png"
                    outp = os.path.join(out_dir, fname)
                    img.save(outp)
                    saved.append(outp)
                    # try cropping embedded images (if any)
                    for idx, im in enumerate(page.images or [], start=1):
                        try:
                            pil_crop = img.original.crop((im['x0'], im['top'], im['x1'], im['bottom']))
                            fname2 = f"{htac}_p{i}_img{idx}.png"
                            outp2 = os.path.join(out_dir, fname2)
                            pil_crop.save(outp2)
                            saved.append(outp2)
                        except Exception:
                            continue
                except Exception:
                    continue
    except Exception:
        pass
    return saved


⸻

extractor/image_reader.py

# extractor/image_reader.py
from PIL import Image
from extractor.text_parser import ocr_image_get_text_and_conf
import shutil
import os

def process_image_file(path, out_images_root, htac):
    try:
        pil = Image.open(path)
    except Exception:
        pil = None
    text, conf = ('', -1.0)
    saved = []
    if pil:
        text, conf = ocr_image_get_text_and_conf(pil)
    # save the original image into the htac folder
    htac_folder = os.path.join(out_images_root, htac)
    os.makedirs(htac_folder, exist_ok=True)
    dest = os.path.join(htac_folder, os.path.basename(path))
    try:
        shutil.copy(path, dest)
        saved.append(dest)
    except Exception:
        pass
    return text, conf, saved


⸻

extractor/table_extractor.py

# extractor/table_extractor.py
import camelot
import pdfplumber
import pandas as pd
import re
from extractor.canonicalizer import canonicalize
from extractor.text_parser import find_test_name

def clean_name_for_key(name):
    if name is None:
        return "Unknown"
    s = str(name).strip()
    # remove special characters for keys
    s = re.sub(r'[^A-Za-z0-9]+', '', s)
    if not s:
        return "Unknown"
    return s

def extract_tables_from_pdf(path):
    tables = []
    # try camelot (stream or lattice depending on table structure)
    try:
        camelot_tables = camelot.read_pdf(path, pages='all', flavor='stream')
        for t in camelot_tables:
            df = t.df
            # convert columns to strings
            df.columns = [str(c) for c in df.columns]
            tables.append(df)
    except Exception:
        pass

    # fallback to pdfplumber
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    for tbl in page.extract_tables():
                        if not tbl:
                            continue
                        # first row is header
                        df = pd.DataFrame(tbl[1:], columns=tbl[0])
                        tables.append(df)
                except Exception:
                    continue
    except Exception:
        pass

    return tables

def flatten_table_with_row_labels(df, table_index, test_name):
    """
    Expects first column to contain the row label names.
    Produces dict with keys: {testcategory}_{ColumnHeader}_{RowLabel}
    where each header/row label is canonicalized.
    """
    flat = {}
    if df is None or len(df.columns) < 2:
        return flat

    # ensure df columns are strings
    df = df.copy()
    df.columns = [str(c) for c in df.columns]

    # get cleaned & canonicalized headers
    headers = [canonicalize(clean_name_for_key(h)) for h in df.columns]

    # canonicalize test name (category)
    test_cat = canonicalize(clean_name_for_key(test_name))

    for r in range(len(df)):
        raw_row = df.iat[r, 0]
        row_name = canonicalize(clean_name_for_key(raw_row))
        for c in range(1, len(headers)):
            col_name = canonicalize(clean_name_for_key(df.columns[c]))
            val = df.iat[r, c]
            key = f"{test_cat}_{col_name}_{row_name}"
            flat[key] = val
    return flat


⸻

extractor/canonicalizer.py

# extractor/canonicalizer.py
import re
from rapidfuzz import process, fuzz
from config import CANONICAL_MAP, FUZZY_THRESHOLD, TEST_NAME_MAP

def normalize_key(k):
    if k is None:
        return "unknown"
    s = str(k).lower()
    s = re.sub(r'[^a-z0-9]+', '', s)
    return s

def canonicalize(key):
    """
    Return canonical name for a key (column header or row header or test name).
    Uses direct mapping first, then fuzzy match against CANONICAL_MAP keys.
    If nothing matched, returns a cleaned TitleCase-ish string.
    """
    if key is None:
        return "Unknown"

    norm = normalize_key(key)

    # direct map (canonical map)
    if norm in CANONICAL_MAP:
        return CANONICAL_MAP[norm]

    # fuzzy match against canonical map keys
    if CANONICAL_MAP:
        best = process.extractOne(norm, list(CANONICAL_MAP.keys()), scorer=fuzz.token_sort_ratio)
        if best and best[1] >= FUZZY_THRESHOLD:
            return CANONICAL_MAP[best[0]]

    # For test categories, check TEST_NAME_MAP
    if norm in TEST_NAME_MAP:
        return TEST_NAME_MAP[norm]

    # otherwise produce a clean TitleCase-ish label
    s = re.sub(r'[^A-Za-z0-9]+', ' ', str(key)).strip()
    # Title-case and remove spaces to make safe key parts
    return ''.join([w.capitalize() for w in s.split()])


⸻

extractor/merger.py

# extractor/merger.py
import pandas as pd
from extractor.canonicalizer import canonicalize

def merge_records(records):
    """
    Merge list of records (each record is dict containing kv and table_cells).
    Returns a DataFrame with canonicalized columns and rows = one record per file.
    """
    all_keys = set()
    for r in records:
        for k in (r.get('kv') or {}).keys():
            all_keys.add(k)
        for tc in (r.get('table_cells') or []):
            for k in tc.keys():
                all_keys.add(k)

    # map each orig key to canonical final label
    mapping = {k: canonicalize(k) for k in all_keys}

    rows = []
    for r in records:
        row = {
            'source_file': r.get('source'),
            'htac': r.get('htac'),
            'test_name': r.get('test'),
            'image_paths': r.get('images'),
            'raw_text': (r.get('text') or '')[:5000],
            'ocr_confidence': r.get('confidence'),
            'extract_status': r.get('status')
        }
        for k, v in (r.get('kv') or {}).items():
            row[mapping.get(k, k)] = v
        for tc in (r.get('table_cells') or []):
            for k, v in tc.items():
                row[mapping.get(k, k)] = v
        rows.append(row)

    df = pd.DataFrame(rows)
    fixed = ['source_file', 'htac', 'test_name', 'image_paths', 'raw_text', 'ocr_confidence', 'extract_status']
    others = [c for c in df.columns if c not in fixed]
    df = df[fixed + sorted(others)]
    return df


⸻

main.py

# main.py
import os
from pathlib import Path
import pandas as pd
from extractor.logger import get_logger
from extractor.file_utils import ensure_dir, list_supported_files
from extractor.pdf_reader import extract_text_and_conf_from_pdf, extract_images_from_pdf
from extractor.image_reader import process_image_file
from extractor.table_extractor import extract_tables_from_pdf, flatten_table_with_row_labels
from extractor.text_parser import extract_kv_pairs, find_htac, find_test_name
from extractor.merger import merge_records
from config import MIN_TEXT_LENGTH, OCR_OK_CONFIDENCE, OCR_LOW_CONFIDENCE

log = get_logger("main")

def load_existing_output(output_dir):
    excel_path = Path(output_dir) / "extracted_report_final.xlsx"
    if excel_path.exists():
        try:
            df = pd.read_excel(excel_path)
            log.info(f"Loaded existing Excel: {excel_path}")
            return df
        except Exception as e:
            log.exception("Error loading old Excel: %s", e)
    return pd.DataFrame()

def process_file(path, out_images_root):
    ext = Path(path).suffix.lower()
    rec = {
        'source': path,
        'htac': Path(path).stem,
        'test': 'general',
        'text': '',
        'confidence': -1.0,
        'status': 'FAILED',
        'images': '',
        'kv': {},
        'table_cells': []
    }

    if ext == '.pdf':
        text, conf = extract_text_and_conf_from_pdf(path)
        htac = find_htac(text) or Path(path).stem
        rec['text'] = text
        rec['confidence'] = conf
        rec['htac'] = htac
        rec['test'] = find_test_name(text)
        htac_folder = os.path.join(out_images_root, htac)
        images = extract_images_from_pdf(path, htac_folder, htac)
        rec['images'] = ';'.join(images)
        rec['kv'] = extract_kv_pairs(text)
        tables = extract_tables_from_pdf(path)
        table_cells = []
        for i, t in enumerate(tables, start=1):
            try:
                flat = flatten_table_with_row_labels(t, i, rec['test'])
                if flat:
                    table_cells.append(flat)
            except Exception:
                continue
        rec['table_cells'] = table_cells
    else:
        text, conf, images = process_image_file(path, out_images_root, Path(path).stem)
        htac = find_htac(text) or Path(path).stem
        rec['text'] = text
        rec['confidence'] = conf
        rec['htac'] = htac
        rec['test'] = find_test_name(text)
        rec['images'] = ';'.join(images)
        rec['kv'] = extract_kv_pairs(text)
        rec['table_cells'] = []

    # set status
    text_len = len((rec.get('text') or '').strip())
    conf_val = rec.get('confidence') or -1.0
    if text_len >= MIN_TEXT_LENGTH and conf_val >= OCR_OK_CONFIDENCE:
        rec['status'] = 'OK'
    elif text_len >= MIN_TEXT_LENGTH and conf_val >= OCR_LOW_CONFIDENCE:
        rec['status'] = 'LOW_CONFIDENCE'
    else:
        rec['status'] = 'FAILED'

    return rec

def run(input_dir='input_reports', output_dir='output'):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    ensure_dir(output_dir)
    out_images_root = output_dir / 'extracted_images'
    ensure_dir(out_images_root)

    files = list_supported_files(str(input_dir))
    log.info(f"Found {len(files)} files in {input_dir}")

    # load existing master sheet
    old_df = load_existing_output(output_dir)

    records = []
    for f in files:
        log.info(f"Processing: {f}")
        try:
            r = process_file(f, str(out_images_root))
            records.append(r)
        except Exception as e:
            log.exception("Failed processing file %s: %s", f, e)
            records.append({
                'source': f,
                'htac': Path(f).stem,
                'test': 'general',
                'text': '',
                'confidence': -1.0,
                'status': 'FAILED',
                'images': '',
                'kv': {},
                'table_cells': []
            })

    if not records:
        log.error("No records processed.")
        return

    new_df = merge_records(records)

    # combine with old if exists (avoid duplicates by source_file)
    if not old_df.empty:
        combined_df = pd.concat([old_df, new_df], ignore_index=True)
        if 'source_file' in combined_df.columns:
            combined_df.drop_duplicates(subset=['source_file'], keep='last', inplace=True)
    else:
        combined_df = new_df

    out_xlsx = output_dir / 'extracted_report_final.xlsx'
    combined_df.to_excel(out_xlsx, index=False)
    combined_df.to_csv(output_dir / 'extracted_report_final.csv', index=False)

    # save issues
    issues = [ {'source_file': r['source'], 'htac': r['htac'], 'status': r['status'], 'ocr_confidence': r['confidence']} for r in records if r['status'] != 'OK' ]
    if issues:
        pd.DataFrame(issues).to_csv(output_dir / 'extraction_issues.csv', index=False)

    log.info("Saved outputs to %s", output_dir)
    log.info("Done.")

if __name__ == "__main__":
    run()


⸻

README.md (brief)

# Test Report Extractor

## Setup
1. Install system packages:
   - Ubuntu: `sudo apt update && sudo apt install -y python3 python3-venv python3-pip tesseract-ocr libtesseract-dev ghostscript tk`
   - macOS (Homebrew): `brew install tesseract ghostscript`
   - Windows: install Tesseract and Ghostscript and add to PATH.

2. Create virtualenv and install:

python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

3. Put your PDFs/images into `input_reports/`.

4. Run:

python main.py

5. Results in `output/`:
- `extracted_report_final.xlsx`
- `extracted_report_final.csv`
- `extraction_issues.csv` (if any)
- `extracted_images/<HTAC>/...`


⸻

Quick notes, troubleshooting & tips
	1.	Tesseract: Must be installed and in PATH. Verify with tesseract --version. On Windows add C:\Program Files\Tesseract-OCR\ to PATH.
	2.	Ghostscript & camelot: Camelot uses Ghostscript; ensure gs (linux/mac) or gswin64c (Windows) is available. On Windows you might need to install additional dependencies for camelot (see camelot docs). If camelot fails, pdfplumber fallback will still try to extract tables.
	3.	Table detection: Scanned tables sometimes need higher resolution. If many missed tables, increase ocr_resolution in pdf_reader.extract_text_and_conf_from_pdf().
	4.	Canonical map: Edit config.py CANONICAL_MAP with frequent variants you see to guarantee perfect merging.
	5.	Append behavior: When run again, new files in input_reports/ get added. The script avoids duplicates using source_file column. If you want dedupe by htac instead, change code in main.run().
	6.	If you want different test-name format: edit text_parser.find_test_name() or the TEST_NAME_MAP in config.py.

⸻

Example of expected final Excel columns (after run)

Columns will include these fixed columns at the left:

source_file | htac | test_name | image_paths | raw_text | ocr_confidence | extract_status | <dynamic canonical columns...>

Dynamic canonical columns will be keys like:

physical_CrownI_StressAt100
physical_CrownII_StressAt100
chemical_Polymer_Tube
chemical_CarbonBlack_Tube
dimension_InnerDiameter_Actual1
...

And the values extracted from tables will populate the correct column for the correct row (file).

⸻

If you want I can:
	•	generate the whole project as a ZIP file you can download (I can create it here and give you the download link), or
	•	run this extractor on the images you already uploaded and produce a sample extracted_report_final.xlsx and show the results.

Which next step do you want?
(Reply: ZIP or RUN_ON_YOUR_UPLOADED_FILES or I’LL RUN LOCALLY if you just want to run it yourself.)
