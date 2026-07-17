#!/usr/bin/env python3
"""
Local web UI for renumbering .docx citations by order of first appearance
(or a manual map), and re-sorting the reference list to match.

Runs 100% locally. Nothing leaves your computer.

Setup:  python -m pip install flask python-docx
Run:    python app.py
Open:   http://127.0.0.1:5000
"""

import io
import re
import sys

try:
    from flask import Flask, request, jsonify, send_file, render_template_string
    from docx import Document
    from docx.document import Document as _DocumentClass
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("Missing dependencies. Run:  python -m pip install flask python-docx")

# ===========================================================================
# Core logic (same engine as the command-line version)
# ===========================================================================
CITATION_RE = re.compile(r"\[\s*\d+(?:\s*[,;\u2013\-]\s*\d+)*\s*\]")
NUMBER_RE = re.compile(r"(?<!\w)\d+(?!\w)")
PLACEHOLDER_RE = re.compile(r"TEMP_NUM_(\d+)_END")
REF_LABEL_RE = re.compile(r"^\s*\[(\d+)\]")
REFERENCE_HEADINGS = {"references", "reference", "bibliography", "works cited"}

def iter_block_items(parent):
    if isinstance(parent, _DocumentClass):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Unsupported parent type")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def iter_paragraphs_in_order(parent):
    for block in iter_block_items(parent):
        if isinstance(block, Paragraph):
            yield block
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    yield from iter_paragraphs_in_order(cell)

def make_transformer(number_map):
    def pass1(m):
        n = int(m.group(0))
        return f"TEMP_NUM_{n}_END" if n in number_map else m.group(0)

    def pass2(m):
        return str(number_map[int(m.group(1))])

    def transform(text):
        stage1 = CITATION_RE.sub(lambda c: NUMBER_RE.sub(pass1, c.group(0)), text)
        return PLACEHOLDER_RE.sub(pass2, stage1)

    return transform

def _build_intervals(runs):
    intervals, pos = [], 0
    for r in runs:
        intervals.append((pos, pos + len(r.text)))
        pos += len(r.text)
    return intervals

def _replace_span(intervals, runs, start, end, new_text):
    inserted = False
    for i, (rs, re_) in enumerate(intervals):
        if re_ <= start or rs >= end:
            continue
        lo, hi = max(start, rs) - rs, min(end, re_) - rs
        run = runs[i]
        if not inserted:
            run.text = run.text[:lo] + new_text + run.text[hi:]
            inserted = True
        else:
            run.text = run.text[:lo] + run.text[hi:]

def rewrite_paragraph(paragraph, transform):
    runs = paragraph.runs
    if not runs:
        return 0
    full = "".join(r.text for r in runs)
    if "[" not in full:
        return 0
    changes = []
    for cit in CITATION_RE.finditer(full):
        new_cit = transform(cit.group(0))
        if new_cit != cit.group(0):
            changes.append((cit.start(), cit.end(), new_cit))
    if not changes:
        return 0
    for start, end, new_cit in sorted(changes, key=lambda c: c[0], reverse=True):
        _replace_span(_build_intervals(runs), runs, start, end, new_cit)
    return len(changes)

def relabel_reference(paragraph, new_num):
    runs = paragraph.runs
    if not runs:
        return
    full = "".join(r.text for r in runs)
    m = REF_LABEL_RE.match(full)
    if m:
        _replace_span(_build_intervals(runs), runs, m.start(1), m.end(1), str(new_num))

def author_snippet(paragraph, max_len=60):
    text = REF_LABEL_RE.sub("", paragraph.text, count=1).strip()
    return text[:max_len] + ("..." if len(text) > max_len else "")

def find_heading_index(paragraphs):
    for i, p in enumerate(paragraphs):
        norm = re.sub(r"^[\d.\s]+", "", p.text.strip().lower()).rstrip(":").strip()
        if norm in REFERENCE_HEADINGS:
            return i
    return None

def parse_manual_map(text):
    m = {}
    for chunk in re.split(r"[\n,;]+", text or ""):
        nums = re.findall(r"\d+", chunk)
        if len(nums) >= 2:
            m[int(nums[0])] = int(nums[1])
    return m

def analyze(doc, mode, manual_map):
    """Read the doc and compute the renumbering plan without modifying it."""
    all_paras = list(iter_paragraphs_in_order(doc))
    heading_idx = find_heading_index(all_paras)
    if heading_idx is None:
        raise ValueError("Could not find a 'REFERENCES' heading in the document.")

    body_paras = all_paras[:heading_idx]
    ref_region = all_paras[heading_idx + 1:]

    ref_entries, spacers = [], []
    for p in ref_region:
        m = REF_LABEL_RE.match(p.text)
        if m:
            ref_entries.append((int(m.group(1)), p))
        elif p.text.strip() == "":
            spacers.append(p)
        else:
            break

    ref_old_numbers = [old for old, _ in ref_entries]
    ref_lookup = {old: author_snippet(p) for old, p in ref_entries}

    if mode == "manual":
        if not manual_map:
            raise ValueError("Manual mode selected but no pairs were provided.")
        mapping = dict(manual_map)
    else:
        mapping, nxt = {}, 1
        for p in body_paras:
            text = "".join(r.text for r in p.runs)
            if "[" not in text:
                continue
            for cit in CITATION_RE.finditer(text):
                for num in NUMBER_RE.findall(cit.group(0)):
                    old = int(num)
                    if old not in mapping:
                        mapping[old] = nxt
                        nxt += 1

    cited = set(mapping)
    nxt = (max(mapping.values()) + 1) if mapping else 1
    for old in sorted(o for o in ref_old_numbers if o not in mapping):
        mapping[old] = nxt
        nxt += 1

    warnings = []
    missing = sorted(o for o in cited if o not in ref_old_numbers)
    if missing:
        warnings.append(f"Cited in body but missing from reference list: {missing}")
    uncited = sorted(o for o in ref_old_numbers if o not in cited)
    if uncited:
        warnings.append(f"In reference list but never cited (placed at end): {uncited}")

    plan = [
        {"old": old, "new": new, "ref": ref_lookup.get(old, "(no reference entry)")}
        for old, new in sorted(mapping.items(), key=lambda kv: kv[1])
    ]
    return {
        "heading_idx": heading_idx,
        "all_paras": all_paras,
        "body_paras": body_paras,
        "ref_entries": ref_entries,
        "spacers": spacers,
        "mapping": mapping,
        "plan": plan,
        "warnings": warnings,
    }

def apply_changes(doc, info, reorder, remove_blanks):
    transform = make_transformer(info["mapping"])
    changed = sum(rewrite_paragraph(p, transform) for p in info["body_paras"])
    for old, p in info["ref_entries"]:
        relabel_reference(p, info["mapping"][old])

    if reorder and info["ref_entries"]:
        if remove_blanks:
            for p in info["spacers"]:
                p._p.getparent().remove(p._p)
        anchor = info["all_paras"][info["heading_idx"]]._p
        for old, p in sorted(info["ref_entries"], key=lambda e: info["mapping"][e[0]]):
            elm = p._p
            elm.getparent().remove(elm)
            anchor.addnext(elm)
            anchor = elm
    return changed

# ===========================================================================
# Flask app
# ===========================================================================
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB

PAGE = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Citation Renumberer</title>
<style>
  :root { color-scheme: light dark; }
  body { font-family: system-ui, sans-serif; max-width: 820px; margin: 2rem auto;
         padding: 0 1rem; line-height: 1.5; }
  h1 { font-size: 1.5rem; }
  .card { border: 1px solid #8883; border-radius: 10px; padding: 1.2rem; margin: 1rem 0; }
  label { display: block; margin: .6rem 0 .2rem; font-weight: 600; }
  input[type=file], textarea { width: 100%; box-sizing: border-box; }
  textarea { height: 90px; font-family: monospace; }
  .row { display: flex; gap: 1.5rem; flex-wrap: wrap; align-items: center; }
  .row label { font-weight: 400; margin: 0; }
  button { padding: .6rem 1.1rem; border-radius: 8px; border: 1px solid #8886;
           cursor: pointer; font-size: 1rem; }
  button.primary { background: #2563eb; color: #fff; border-color: #2563eb; }
  table { border-collapse: collapse; width: 100%; margin-top: .5rem; font-size: .92rem; }
  th, td { border: 1px solid #8883; padding: .35rem .5rem; text-align: left; }
  th { background: #8881; }
  .warn { background: #f59e0b22; border: 1px solid #f59e0b88; padding: .6rem .8rem;
          border-radius: 8px; margin: .5rem 0; }
  .muted { opacity: .7; font-size: .9rem; }
  #status { margin-left: .6rem; }
  .hidden { display: none; }
</style>
</head>
<body>
  <h1>Citation Renumberer</h1>
  <p class="muted">Everything runs locally on your machine. Your document is not sent to the internet.</p>

  <form id="form" action="/process" method="post" enctype="multipart/form-data">
    <div class="card">
      <label for="file">Word document (.docx)</label>
      <input id="file" name="file" type="file" accept=".docx" required>

      <label>Renumbering mode</label>
      <div class="row">
        <label><input type="radio" name="mode" value="appearance" checked> Order of first appearance (automatic)</label>
        <label><input type="radio" name="mode" value="manual"> Manual mapping</label>
      </div>

      <div id="manualBox" class="hidden">
        <label for="manual_map">Manual pairs — one per line, "old: new"</label>
        <textarea id="manual_map" name="manual_map" placeholder="52: 1&#10;1: 2&#10;2: 3"></textarea>
      </div>

      <div class="row" style="margin-top:.8rem">
        <label><input type="checkbox" name="reorder" checked> Re-sort reference list by new number</label>
        <label><input type="checkbox" name="remove_blanks" checked> Remove blank lines between entries</label>
      </div>

      <div class="row" style="margin-top:1rem">
        <button type="button" class="primary" onclick="preview()">Preview plan</button>
        <button type="submit" formaction="/process">Download renumbered .docx</button>
        <span id="status"></span>
      </div>
    </div>
  </form>

  <div id="result" class="card hidden">
    <h2 style="font-size:1.1rem">Renumbering plan</h2>
    <div id="warnings"></div>
    <div id="planTable"></div>
  </div>

<script>
  const modeRadios = document.querySelectorAll('input[name=mode]');
  const manualBox = document.getElementById('manualBox');
  modeRadios.forEach(r => r.addEventListener('change', () => {
    manualBox.classList.toggle('hidden', document.querySelector('input[name=mode]:checked').value !== 'manual');
  }));

  async function preview() {
    const form = document.getElementById('form');
    const status = document.getElementById('status');
    const result = document.getElementById('result');
    if (!form.file.files.length) { alert('Choose a .docx file first.'); return; }
    status.textContent = 'Analyzing...';
    try {
      const res = await fetch('/preview', { method: 'POST', body: new FormData(form) });
      const data = await res.json();
      if (!res.ok) { status.textContent = ''; alert(data.error || 'Error'); return; }
      const warns = document.getElementById('warnings');
      warns.innerHTML = (data.warnings || []).map(w => '<div class="warn">' + w + '</div>').join('');
      let rows = data.plan.map(p =>
        `<tr><td>[${p.old}]</td><td>&rarr; [${p.new}]</td><td>${p.ref}</td></tr>`).join('');
      document.getElementById('planTable').innerHTML =
        `<table><thead><tr><th>Old</th><th>New</th><th>Reference</th></tr></thead><tbody>${rows}</tbody></table>`;
      result.classList.remove('hidden');
      status.textContent = data.plan.length + ' entries.';
    } catch (e) { status.textContent = ''; alert('Request failed: ' + e); }
  }
</script>
</body>
</html>
"""

def _read_request():
    file = request.files.get("file")
    if not file or not file.filename:
        raise ValueError("No file uploaded.")
    mode = request.form.get("mode", "appearance")
    manual_map = parse_manual_map(request.form.get("manual_map", "")) if mode == "manual" else {}
    reorder = request.form.get("reorder") is not None
    remove_blanks = request.form.get("remove_blanks") is not None
    return file, mode, manual_map, reorder, remove_blanks

@app.route("/")
def index():
    return render_template_string(PAGE)

@app.route("/preview", methods=["POST"])
def preview():
    try:
        file, mode, manual_map, _, _ = _read_request()
        doc = Document(io.BytesIO(file.read()))
        info = analyze(doc, mode, manual_map)
        return jsonify({"plan": info["plan"], "warnings": info["warnings"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.route("/process", methods=["POST"])
def process():
    try:
        file, mode, manual_map, reorder, remove_blanks = _read_request()
        raw = file.read()
        doc = Document(io.BytesIO(raw))
        info = analyze(doc, mode, manual_map)
        apply_changes(doc, info, reorder, remove_blanks)
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        base = re.sub(r"\.docx$", "", file.filename, flags=re.I) or "document"
        return send_file(
            out,
            as_attachment=True,
            download_name=f"{base}_renumbered.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        return f"<h3>Error</h3><p>{e}</p><p><a href='/'>Back</a></p>", 400

if __name__ == "__main__":
    # 127.0.0.1 keeps it reachable only from this computer.
    app.run(host="127.0.0.1", port=5000, debug=False)
