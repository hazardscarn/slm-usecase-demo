"""Builds LLM_vs_SLM_Cost_Analysis.xlsx from sourced research data (research/*.json).

All numbers on the pricing/spec sheets come from the JSON files in ./research (each record
carries its source URL + access date). All calculations in the workbook are live Excel formulas.
Run:  python3 build_workbook.py
"""
import json, re, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule

HERE = os.path.dirname(os.path.abspath(__file__))
R = os.path.join(HERE, "research")
def load(n): return json.load(open(os.path.join(R, n + ".json")))

gem_note = open(os.path.join(R, "gemini_bigquery_notes.md")).read()
claude = [r for r in load("claude_pricing") if "_meta" not in r]
openai = [r for r in load("openai_pricing") if "_meta" not in r]
slm = load("slm_models"); bench = load("slm_benchmarks")
gpu_raw = load("gpu_pricing")
ACCESSED = "2026-11-08"

# ---------- styles ----------
FONT = "Arial"
f_norm = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_hdr = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="1F3864")
f_in = Font(name=FONT, size=10, color="0000FF")
f_link = Font(name=FONT, size=10, color="008000")
f_url = Font(name=FONT, size=9, color="0563C1", underline="single")
f_note = Font(name=FONT, size=9, italic=True, color="595959")
fill_hdr = PatternFill("solid", fgColor="1F3864")
fill_sub = PatternFill("solid", fgColor="D9E2F3")
fill_key = PatternFill("solid", fgColor="FFFF00")
fill_warn = PatternFill("solid", fgColor="FCE4D6")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
USD = '$#,##0.00;($#,##0.00);"-"'
USD4 = '$#,##0.0000;($#,##0.0000);"-"'
USD0 = '$#,##0;($#,##0);"-"'
NUM = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%'

wb = Workbook()
wb.remove(wb.active)

def sheet(name, title, subtitle=None, tab=None):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = f_title
    if subtitle:
        ws["A2"] = subtitle; ws["A2"].font = f_note
    ws.sheet_view.showGridLines = False
    if tab: ws.sheet_properties.tabColor = tab
    return ws

def hdr(ws, row, headers, col=1, height=42):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font = f_hdr; c.fill = fill_hdr; c.alignment = center; c.border = border
    ws.row_dimensions[row].height = height

def put(ws, row, col, val, fmt=None, font=None, fill=None, al=None, bd=True):
    c = ws.cell(row=row, column=col, value=val)
    c.font = font or f_norm
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    c.alignment = al or Alignment(vertical="top")
    if bd: c.border = border
    return c

def url_cell(ws, row, col, text):
    c = put(ws, row, col, text, font=f_url, al=wrap)
    m = re.search(r"https?://[^\s;,)]+", text or "")
    if m: c.hyperlink = m.group(0)
    return c

def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1): ws.column_dimensions[L(i)].width = w

# =====================================================================
# 1. ASSUMPTIONS
# =====================================================================
wsA = sheet("Assumptions", "Assumptions & Scenario Inputs",
            "Blue = editable input · Black = formula · Yellow fill = key lever. Every default below is an assumption (not a vendor fact) unless the Basis column cites a source.", "FFC000")
widths(wsA, [46, 22, 16, 100])
hdr(wsA, 4, ["Parameter", "Value", "Unit", "Basis / source / note"], height=22)
A = {}          # key -> absolute reference
row = [5]

def sec(title):
    r = row[0]
    for c in range(1, 5):
        cell = wsA.cell(row=r, column=c); cell.fill = fill_sub; cell.border = border
    wsA.cell(row=r, column=1, value=title).font = f_bold
    row[0] += 1

def add(key, label, value, unit, note, kind="input", fmt=None, key_lever=False):
    r = row[0]
    A[key] = f"Assumptions!$B${r}"
    put(wsA, r, 1, label, al=wrap)
    if isinstance(value, str) and value.startswith("="):
        value = value.format(**{k: v.replace("Assumptions!", "") for k, v in A.items()})
    font = f_in if kind == "input" else f_norm
    c = put(wsA, r, 2, value, fmt=fmt, font=font, fill=fill_key if key_lever else None)
    put(wsA, r, 3, unit)
    put(wsA, r, 4, note, al=wrap)
    row[0] += 1
    return r

sec("1. Volume scenarios")
add("vol_step", "Call volume step for scenario columns", 1000, "calls/day", "User request: model 1K to 15K calls/day in increments of 1K. Cost sheets show 15 columns = step x 1..15.", fmt=NUM)
add("ref_vol", "Reference volume for monthly/annual figures", 10000, "calls/day", "Mid-point of the user's 1K-15K range; change to see monthly/annual views on every sheet.", fmt=NUM, key_lever=True)
add("days", "Operating days per month", 30, "days", "Calendar-month convention; use 21-22 if calls are business-days only.", fmt=NUM)

sec("2. Call transcript profile (energy retail: residential + C&I)")
add("res_min", "Residential avg handle time", 8, "minutes", "ASSUMPTION - typical energy-retail residential call (billing query, move-in/out, plan change, outage). Replace with your contact-centre AHT.", fmt=NUM1, key_lever=True)
add("ci_min", "C&I avg handle time", 15, "minutes", "ASSUMPTION - commercial & industrial calls are longer (contract/tariff/metering/dispute discussions).", fmt=NUM1, key_lever=True)
add("wpm", "Spoken words per minute (both parties)", 130, "words/min", "Conversational speech 120-150 wpm (common speech-rate range); 130 used.", fmt=NUM)
add("tpw", "Tokens per English word", 1.33, "tokens/word", "Common rule of thumb ~0.75 words/token (OpenAI tokenizer guidance) -> 1.33 tokens/word.", fmt='0.00')
add("fmt_oh", "Transcript formatting overhead factor", 1.25, "x", "ASSUMPTION - speaker labels, timestamps, ASR disfluencies add ~25% tokens.", fmt='0.00')
add("mix_res", "Share of calls that are residential", 0.80, "%", "ASSUMPTION - 80/20 residential / C&I by call count. Change to your mix.", fmt=PCT, key_lever=True)
add("tok_res", "Residential transcript tokens", "={res_min}*{wpm}*{tpw}*{fmt_oh}", "tokens", "Calculated.", kind="calc", fmt=NUM)
add("tok_ci", "C&I transcript tokens", "={ci_min}*{wpm}*{tpw}*{fmt_oh}", "tokens", "Calculated.", kind="calc", fmt=NUM)
add("tok_blend", "Blended transcript tokens per call", "={mix_res}*{tok_res}+(1-{mix_res})*{tok_ci}", "tokens", "Calculated, weighted by call mix.", kind="calc", fmt=NUM)

sec("3. Closed-model (API) assumptions")
add("think", "Thinking/reasoning tokens per task call (billed as output)", 300, "tokens", "ASSUMPTION - all current Gemini, Claude 5.x and GPT-5.x+ models bill reasoning/thinking tokens as output (see Closed_Pricing notes). 300 = low reasoning effort. Set 0 if thinking is disabled; 1,000+ for high effort.", fmt=NUM, key_lever=True)
add("cache_hit", "Share of input tokens served from prompt cache", 0.0, "%", "User expectation: caching is NOT used (each transcript is unique). Left 0%. If >0, cache-read price is applied to that share; cache-write/storage fees are NOT modelled.", fmt=PCT)
add("basis", "Closed-model price basis", "Current", "list", "'Current' = list/promo prices valid on the access date; 'Post-promo' = prices after intro promos end (Gemini 3.6-3.8 Flash from 2027-01-01, Gemini 4 Argon from 2026-12-01, gpt-5.6-sol after 2026-11-30 [inferred]).", key_lever=True)
add("bq_bytes", "BigQuery bytes processed per call (transcript row)", 20000, "bytes", "ASSUMPTION - ~20 KB (transcript + metadata columns scanned per row).", fmt=NUM)
add("bq_tib", "BigQuery on-demand analysis price", 6.25, "$/TiB", "https://cloud.google.com/bigquery/pricing (accessed 2026-11-08). First 1 TiB/month free is ignored (conservative).", fmt=USD)

sec("4. SLM hosting & throughput assumptions")
add("host", "GPU hosting price option", "GCP On-demand", "list", "Choose: GCP On-demand | GCP Spot | GCP 1-yr CUD | GCP 3-yr CUD | Databricks serving (DBU). Drives every SLM cost sheet.", key_lever=True)
add("dbu", "Databricks $/DBU for GPU Model Serving", 0.07, "$/DBU", "UNVERIFIED - third-party (dbuwatch.com) for GCP US regions; official Databricks tiles are JS-rendered. Confirm with your Databricks account team.", fmt='$0.000')
add("prec", "Serving weight precision used for GPU-fit check", "FP8", "list", "BF16 | FP8 | INT4. FP8 is the production default on H100/L4/RTX PRO 6000 (A100 lacks native FP8 - treat as approximate). gpt-oss-20b uses its native MXFP4 size under 'FP8'.", key_lever=True)
add("fit", "Max share of VRAM usable for weights", 0.80, "%", "ASSUMPTION - remainder reserved for KV-cache, activations, CUDA graphs.", fmt=PCT)
add("prefill", "Prefill speed multiple of decode throughput", 5, "x", "DERIVED ASSUMPTION - prefill is compute-bound and batches far better than decode; 4-6x aggregate tok/s is typical for vLLM on H100. Validate with `vllm bench serve` on real transcripts.", fmt=NUM1)
add("util", "Production utilisation derate", 0.65, "%", "Methodology range 0.6-0.7 (throughput_refs: peak benchmark -> sustained production).", fmt=PCT)
add("startup", "Cluster start-up/idle overhead per batch run", 0.25, "GPU-hours/day", "ASSUMPTION - model load + warm-up billed on each on-demand run.", fmt='0.00')
add("slm_think", "Thinking tokens per task call for fine-tuned SLM", 0, "tokens", "Assumption: SFT teaches direct JSON/summary output with no reasoning trace.", fmt=NUM)

sec("5. One-time SFT (fine-tuning) assumptions")
add("sft_n", "SFT training samples", 50000, "samples", "User requirement: ~50K samples, one-time.", fmt=NUM)
add("sft_tok", "Avg tokens per SFT sample (prompt+transcript+target)", 3000, "tokens", "ASSUMPTION - one transcript (~2.0K blended) + instructions + target output; matches agent SFT references (50K x 3K).", fmt=NUM, key_lever=True)
add("sft_ep", "Training epochs", 3, "epochs", "Typical 2-3 for LoRA SFT; 2 lowers cost by a third.", fmt=NUM)
add("sft_iter", "Training runs incl. re-runs / hyper-parameter sweeps", 3, "x", "ASSUMPTION - budget 3 full-size runs (v1, fix, final) per model.", fmt=NUM)
add("amort", "Months to amortise SFT cost over", 12, "months", "Used for 'all-in monthly cost' on Comparison sheet.", fmt=NUM)

sec("6. Comparison baseline")
add("base", "Closed model used as baseline on Comparison sheet", "Gemini 3.1 Flash-Lite", "list", "Pick any model from the Closed_Pricing sheet (dropdown). The 'cheapest closed model' is computed automatically as well.", key_lever=True)

# --- task table
sec("7. LLM tasks run per call (every call is processed by each enabled task)")
tr = row[0]
hdr(wsA, tr, ["Task", "Enabled (1/0)", "Prompt/instr. tokens", "Output tokens - Residential", "Output tokens - C&I",
              "Input tokens/call (blended)", "Output tokens/call (blended)"], height=44)
wsA.column_dimensions["E"].width = 18; wsA.column_dimensions["F"].width = 18; wsA.column_dimensions["G"].width = 18
tasks = [("Call transcript summarization", 1, 350, 220, 380),
         ("Entity extraction (JSON: account, meter, NMI/MPRN, product, dates, amounts...)", 1, 800, 350, 600),
         ("Call disposition / intent classification", 1, 300, 40, 60)]
t0 = tr + 1
for i, (n, en, p, o1, o2) in enumerate(tasks):
    r = t0 + i
    put(wsA, r, 1, n, al=wrap)
    for j, v in enumerate([en, p, o1, o2]):
        put(wsA, r, 2 + j, v, fmt=NUM, font=f_in)
    put(wsA, r, 6, f"=C{r}+{A['tok_blend'].replace('Assumptions!','')}", fmt=NUM)
    put(wsA, r, 7, f"={A['mix_res'].replace('Assumptions!','')}*D{r}+(1-{A['mix_res'].replace('Assumptions!','')})*E{r}", fmt=NUM)
t1 = t0 + len(tasks) - 1
row[0] = t1 + 1
put(wsA, row[0], 1, "Task assumptions: ASSUMPTIONS. Extraction needs a longer prompt (schema + few-shot); output sizes extrapolated from typical summary (150-400 tok) and JSON (300-600 tok) lengths. Edit the blue cells.", font=f_note, bd=False)
row[0] += 2
sec("8. Derived per-call token totals (feed all cost sheets)")
add("n_tasks", "Enabled tasks per call", f"=SUM(B{t0}:B{t1})", "tasks", "Calculated.", kind="calc", fmt=NUM)
add("tot_in", "Input tokens per call (all enabled tasks)", f"=SUMPRODUCT(B{t0}:B{t1},F{t0}:F{t1})", "tokens", "Transcript is re-sent for each task (no cross-task caching assumed).", kind="calc", fmt=NUM)
add("tot_out", "Output tokens per call - visible (all enabled tasks)", f"=SUMPRODUCT(B{t0}:B{t1},G{t0}:G{t1})", "tokens", "Calculated.", kind="calc", fmt=NUM)
add("out_closed", "Output tokens per call billed - closed models", "={tot_out}+{think}*{n_tasks}", "tokens", "Visible output + thinking tokens.", kind="calc", fmt=NUM)
add("out_slm", "Output tokens per call generated - SLM", "={tot_out}+{slm_think}*{n_tasks}", "tokens", "Visible output + SLM thinking tokens.", kind="calc", fmt=NUM)

for rr in range(5, row[0]):
    wsA.row_dimensions[rr].height = None
dv = DataValidation(type="list", formula1='"Current,Post-promo"', allow_blank=False); wsA.add_data_validation(dv); dv.add(A["basis"].split("!")[1].replace("$", ""))
dv = DataValidation(type="list", formula1='"GCP On-demand,GCP Spot,GCP 1-yr CUD,GCP 3-yr CUD,Databricks serving (DBU)"'); wsA.add_data_validation(dv); dv.add(A["host"].split("!")[1].replace("$", ""))
dv = DataValidation(type="list", formula1='"BF16,FP8,INT4"'); wsA.add_data_validation(dv); dv.add(A["prec"].split("!")[1].replace("$", ""))
wsA.freeze_panes = "A5"

# =====================================================================
# 2. CLOSED PRICING
# =====================================================================
wsP = sheet("Closed_Pricing", "Closed-Model API Pricing (USD per 1M tokens)",
            f"Sources: official provider pricing pages, accessed {ACCESSED}. Blue = hard-coded from source; Column T/U = exact source URL per row. Thinking/reasoning tokens are billed as OUTPUT on all models listed.", "5B9BD5")
PH = ["Provider", "Model", "Model ID", "Status (as of 2026-11-08)", "Input $/1M", "Output $/1M", "Cached-input read $/1M",
      "Batch input $/1M", "Batch output $/1M", "Input $/1M (post-promo)", "Output $/1M (post-promo)", "Cache read $/1M (post-promo)",
      "Batch input $/1M (post-promo)", "Batch output $/1M (post-promo)", "Cache write (5-min) $/1M", "Batch API supported?", "Via BigQuery batch (50%)?",
      "Context window", "Notes / caveats", "Source URL(s)", "Accessed",
      "EFFECTIVE input", "EFFECTIVE output", "EFFECTIVE cache read", "EFFECTIVE batch input", "EFFECTIVE batch output"]
P0 = 5
hdr(wsP, P0, PH, height=56)
widths(wsP, [11, 24, 24, 34, 10, 10, 11, 10, 10, 11, 11, 11, 11, 11, 11, 12, 12, 11, 70, 60, 11, 11, 11, 11, 11, 11])

GEM = [  # name,id,status,in,out,cache,bin,bout,(post in,out,cache,bin,bout)|None, batch_supported, ctx, note, src
 ("Gemini 2.5 Flash-Lite", "gemini-2.5-flash-lite", "GA (older generation)", .10, .40, .01, .05, .20, None, "1M", "Cache storage $1.00/1M tok/hr. Vertex batch audio price differs from Gemini API ($0.05 vs $0.15) - text used here."),
 ("Gemini 2.5 Flash", "gemini-2.5-flash", "DEPRECATED - from 2026-12-01 requests auto-route to gemini-3.5-flash-lite and bill at its rates", .30, 2.50, .03, .15, 1.25, None, "1M", "Cache storage $1.00/1M tok/hr. Thinking billed as output."),
 ("Gemini 2.5 Pro", "gemini-2.5-pro", "DEPRECATED - from 2026-12-01 auto-routes to gemini-3.8-flash", 1.25, 10.00, .125, .625, 5.00, None, "1M", "<=200K-token prompt tier shown; >200K: $2.50/$15 (batch $1.25/$7.50). Cache storage $4.50/1M tok/hr."),
 ("Gemini 3 Flash Preview", "gemini-3-flash-preview", "Preview ('legacy Flash')", .50, 3.00, .05, .25, 1.50, None, "1M", "Preview model - no GA SLA."),
 ("Gemini 3.1 Flash-Lite", "gemini-3.1-flash-lite", "GA per pricing page (no preview label - inferred)", .25, 1.50, .025, .125, .75, None, "1M", "Status label inferred from absence of 'Preview' on pricing page."),
 ("Gemini 3.5 Flash-Lite", "gemini-3.5-flash-lite", "GA per pricing page (inferred)", .30, 2.50, .03, .15, 1.25, None, "1M", "Target model for 2.5 Flash auto-routing from 2026-12-01."),
 ("Gemini 3.1 Pro Preview", "gemini-3.1-pro-preview", "Preview", 2.00, 12.00, .20, 1.00, 6.00, None, "1M", "<=200K tier shown; >200K: $4/$18 (batch $2/$9). Cache storage $4.50/1M tok/hr."),
 ("Gemini 3.5 Flash", "gemini-3.5-flash", "Listed on Vertex pricing page only", 1.50, 9.00, .15, .75, 4.50, None, "1M", "Not on Gemini API page; Vertex price shown."),
 ("Gemini 3.6 Flash", "gemini-3.6-flash", "Listed on both pages ('previous generation Flash')", .75, 3.75, .075, .375, 1.875, (1.5, 7.5, .15, .75, 3.75), "1M", "INTRO PRICE through 2026-12-31; list price from 2027-01-01 shown in post-promo columns."),
 ("Gemini 3.7 Flash", "gemini-3.7-flash", "Vertex page only", .75, 3.75, .075, .375, 1.875, (1.5, 7.5, .15, .75, 3.75), "1M", "INTRO PRICE through 2026-12-31 (as 3.6)."),
 ("Gemini 3.8 Flash", "gemini-3.8-flash", "Newest Flash; status label not stated", .75, 3.75, .075, .375, 1.875, (1.5, 7.5, .15, .75, 3.75), "1M", "INTRO PRICE through 2026-12-31. 2.5 Pro auto-routes here from 2026-12-01."),
 ("Gemini 4 Argon", "gemini-4-argon", "Flagship; status label not stated", 2.00, 10.00, .10, 1.00, 5.00, (4.0, 20.0, .20, 2.0, 10.0), "1M", "PROMO through 2026-11-30 (on Vertex it is '50% credits back on net spend', not a list price). Implicit caching only - no storage fee."),
]
GEM_SRC = "https://ai.google.dev/gemini-api/docs/pricing ; https://cloud.google.com/vertex-ai/generative-ai/pricing"
prow = {}    # model name -> row
rows_meta = []   # (provider, model, row)
r = P0 + 1
def pricing_row(r, prov, name, mid, status, vals, post, cwrite, batch_ok, bq, ctx, note, src):
    inp, out, cache, bin_, bout = vals
    put(wsP, r, 1, prov); put(wsP, r, 2, name, font=f_bold); put(wsP, r, 3, mid); put(wsP, r, 4, status, al=wrap)
    for j, v in enumerate([inp, out, cache, bin_, bout]):
        put(wsP, r, 5 + j, v, fmt=USD4 if (v is not None and v < 0.1) else USD, font=f_in)
    # E..I =5..9 ; post-promo J..N = 10..14
    cols_cur = ["E", "F", "G", "H", "I"]
    for j in range(5):
        if post:
            put(wsP, r, 10 + j, post[j], fmt=USD, font=f_in, fill=fill_warn)
        else:
            put(wsP, r, 10 + j, f'=IF(ISNUMBER({cols_cur[j]}{r}),{cols_cur[j]}{r},"")', fmt=USD)
    put(wsP, r, 15, cwrite, fmt=USD, font=f_in)
    put(wsP, r, 16, batch_ok); put(wsP, r, 17, bq)
    put(wsP, r, 18, ctx); put(wsP, r, 19, note, al=wrap); url_cell(wsP, r, 20, src); put(wsP, r, 21, ACCESSED)
    b = A["basis"]
    # effective columns V..Z (22..26)
    put(wsP, r, 22, f'=IF({b}="Current",E{r},J{r})', fmt=USD4, fill=fill_sub)
    put(wsP, r, 23, f'=IF({b}="Current",F{r},K{r})', fmt=USD4, fill=fill_sub)
    put(wsP, r, 24, f'=IF({b}="Current",IF(ISNUMBER(G{r}),G{r},E{r}),IF(ISNUMBER(L{r}),L{r},J{r}))', fmt=USD4, fill=fill_sub)
    put(wsP, r, 25, f'=IF({b}="Current",H{r},M{r})', fmt=USD4, fill=fill_sub)
    put(wsP, r, 26, f'=IF({b}="Current",I{r},N{r})', fmt=USD4, fill=fill_sub)

for g in GEM:
    pricing_row(r, "Google (Gemini)", g[0], g[1], g[2], g[3:8], g[8], None, "Yes", "Yes", g[9], g[10], GEM_SRC)
    rows_meta.append(("Google (Gemini)", g[0], r)); prow[g[0]] = r; r += 1
for c in claude:
    name = c["model"].replace(" (prompt <=100K tokens)", "")
    if ">100K" in c["model"]: continue
    note = c["notes"]
    if "Haiku 5.5" in name: note = "Prompt <=100K-token tier shown ($0.50/$2.50 above 100K). " + note
    post = None
    pricing_row(r, "Anthropic (Claude)", name, c["model_id"], c["status"], (c["input_per_mtok"], c["output_per_mtok"], c["cache_read_per_mtok"], c["batch_input_per_mtok"], c["batch_output_per_mtok"]),
                post, c["cache_write_5m_per_mtok"], "Yes", "No", c["context_window"], note, "https://platform.claude.com/docs/en/about-claude/pricing ; " + c["source_url"].split(";")[-1].strip())
    rows_meta.append(("Anthropic (Claude)", name, r)); prow[name] = r; r += 1
for o in openai:
    name = o["model_id"]
    cache = o["cache_read_per_mtok"]
    bok = "Yes" if o["batch_endpoint_supported"] else "No"
    bin_, bout = o["batch_input_per_mtok"], o["batch_output_per_mtok"]
    note = o["notes"]
    if name in ("gpt-5.2-pro", "gpt-5.4-pro"):
        bok = "Unverified"; note = "CONFLICT: pricing page lists batch rates but model page says Batch endpoint not supported. " + note
    if name == "gpt-5.3-codex": note = "Batch endpoint not supported. " + note
    post = None
    if name == "gpt-5.6-sol":
        post = (5.0, 30.0, 0.5, 2.5, 15.0); note = "PROMOTIONAL $4/$20 until 2026-11-30; post-promo ~$5/$30 is INFERRED from the page's '20%/33% reduction' wording (not a stated price). " + note
    pricing_row(r, "OpenAI (GPT)", name, name, o["status"], (o["input_per_mtok"], o["output_per_mtok"], cache, bin_, bout),
                post, None, bok, "No", o["context_window"], note, o["source_url"].split(" ; ")[0] + " ; https://developers.openai.com/api/docs/models/" + name)
    rows_meta.append(("OpenAI (GPT)", name, r)); prow[name] = r; r += 1
P_last = r - 1
put(wsP, r + 1, 1, "Notes", font=f_bold, bd=False)
notes = [
    "BigQuery path: Gemini models called from BigQuery (ML.GENERATE_TEXT, AI.GENERATE, AI.GENERATE_TABLE etc.) are billed at Vertex AI 'batch' rates - 50% off - per https://cloud.google.com/bigquery/pricing and https://cloud.google.com/bigquery/docs/choose-text-generation-function. Claude/open models via BigQuery bill at standard Vertex prices (no discount).",
    "Batch API: Anthropic Message Batches and OpenAI Batch both give 50% off input+output (async, up to 24h). Not every OpenAI model supports Batch - see column P.",
    "Cache storage (Gemini): $1.00 per 1M tokens per hour on Flash/Flash-Lite, $4.50 on Pro. Cache-write fees: Claude 1.25x input (5 min) / 2x (1 h); OpenAI gpt-5.6+: 1.25x (30 min) / 2.5x (12 h). Not modelled because caching is assumed unused.",
    "Long-context premiums (>200K Gemini Pro, >272K OpenAI 5.4+, >100K Haiku 5.5) do not apply: calls here are ~2-8K tokens.",
    "Other modifiers NOT modelled: Vertex non-global endpoints +10% (Gemini 3.x+), Anthropic inference_geo=us +10%, OpenAI regional endpoints +10%, OpenAI/Gemini priority/fast tiers (~1.8-2x).",
    "Retirements: Claude Sonnet 4.5 retires 2026-11-30; gpt-5/5-mini/5-nano/5-pro removed 2026-12-11; Gemini 2.5 Flash/Pro auto-route from 2026-12-01 - treat as already-obsolete baselines.",
    "Models the user asked for that do not exist as named: no gpt-5.3 base model and no gpt-5.1 mini/nano exist; OpenAI's 5.6/6.x families use sol/terra/luna/astra tier names instead.",
]
for i, n in enumerate(notes):
    c = wsP.cell(row=r + 2 + i, column=1, value="• " + n); c.font = f_note
wsP.freeze_panes = wsP.cell(row=P0 + 1, column=3)
wsP.auto_filter.ref = f"A{P0}:{L(26)}{P_last}"
dv = DataValidation(type="list", formula1=f"=Closed_Pricing!$B${P0+1}:$B${P_last}"); wsA.add_data_validation(dv); dv.add(A["base"].split("!")[1].replace("$", ""))

# =====================================================================
# 3. CLOSED COST SHEETS
# =====================================================================
NV = 15
def cost_sheet(name, title, sub, mode, tab):
    ws = sheet(name, title, sub, tab)
    heads = ["Provider", "Model", "Status", "$ per call", "$ per 1K calls"] + [None] * NV + ["Monthly @ ref volume", "Annual @ ref volume", "Rank (cheapest=1)", "Flag"]
    hdr(ws, 5, heads, height=44)
    # volume header: row 5 cols F..T, numeric formulas
    for k in range(NV):
        c = ws.cell(row=5, column=6 + k)
        c.value = f"={A['vol_step']}" if k == 0 else f"={L(5 + k)}5+{A['vol_step']}"
        c.number_format = '#,##0" calls/day"'; c.font = f_hdr
    ws["F4"] = "Daily cost (USD) at calls/day  →"; ws["F4"].font = f_bold
    widths(ws, [16, 26, 30, 12, 12] + [11] * NV + [15, 15, 11, 38])
    i = 6
    out_rows = []
    for prov, mdl, pr in rows_meta:
        if mode == "bq" and prov != "Google (Gemini)": continue
        P = lambda col: f"Closed_Pricing!${col}${pr}"
        ti, to, h = A["tot_in"], A["out_closed"], A["cache_hit"]
        if mode == "std":
            f = f"=({ti}*(1-{h})*{P('V')}+{ti}*{h}*{P('X')}+{to}*{P('W')})/1000000"
            flag = f'=IF(ISNUMBER(SEARCH("DEPRECATED",{P("D")})),"Deprecated/retiring",IF(ISNUMBER(SEARCH("Preview",{P("D")})),"Preview",""))'
        else:
            base = f"({ti}*(1-{h})*{P('Y')}+{ti}*{h}*{P('X')}*0.5+{to}*{P('Z')})/1000000"
            if mode == "batch":
                f = f'=IF({P("P")}="No","n/a",{base})'
                flag = f'=IF({P("P")}="No","Batch API not supported",IF({P("P")}="Unverified","Batch support conflicting in OpenAI docs",IF(ISNUMBER(SEARCH("DEPRECATED",{P("D")})),"Deprecated/retiring","")))'
            else:
                bq = f"{A['bq_bytes']}/1099511627776*{A['bq_tib']}"
                f = f"={base}+{bq}"
                flag = f'=IF(ISNUMBER(SEARCH("DEPRECATED",{P("D")})),"Deprecated/retiring",IF(ISNUMBER(SEARCH("Preview",{P("D")})),"Preview",""))'
        put(ws, i, 1, f"={P('A')}", font=f_link); put(ws, i, 2, f"={P('B')}", font=f_link); put(ws, i, 3, f"={P('D')}", font=f_link, al=Alignment(wrap_text=False))
        put(ws, i, 4, f, fmt=USD4)
        put(ws, i, 5, f'=IF(ISNUMBER(D{i}),D{i}*1000,"n/a")', fmt=USD)
        for k in range(NV):
            put(ws, i, 6 + k, f'=IF(ISNUMBER($D{i}),$D{i}*{L(6+k)}$5,"n/a")', fmt=USD0 if True else USD)
        put(ws, i, 6 + NV, f'=IF(ISNUMBER($D{i}),$D{i}*{A["ref_vol"]}*{A["days"]},"n/a")', fmt=USD0)
        put(ws, i, 7 + NV, f'=IF(ISNUMBER({L(6+NV)}{i}),{L(6+NV)}{i}*12/{A["days"]}*30.4/30,"n/a")' if False else f'=IF(ISNUMBER({L(6+NV)}{i}),$D{i}*{A["ref_vol"]}*365,"n/a")', fmt=USD0)
        out_rows.append(i)
        put(ws, i, 9 + NV, flag, font=f_note)
        i += 1
    last = i - 1
    for rr in out_rows:
        put(ws, rr, 8 + NV, f'=IFERROR(RANK({L(6+NV)}{rr},${L(6+NV)}$6:${L(6+NV)}${last},1),"")', fmt='0')
    # totals / stats row
    ws.cell(row=last + 2, column=2, value="Cheapest model").font = f_bold
    ws.cell(row=last + 3, column=2, value="Cheapest $ per call").font = f_bold
    put(ws, last + 2, 4, f'=INDEX($B$6:$B${last},MATCH(MIN($D$6:$D${last}),$D$6:$D${last},0))', bd=False, font=f_bold)
    put(ws, last + 3, 4, f"=MIN($D$6:$D${last})", fmt=USD4, bd=False, font=f_bold)
    ws.conditional_formatting.add(f"F6:{L(5+NV)}{last}", ColorScaleRule(start_type="min", start_color="C6EFCE", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="max", end_color="F8CBAD"))
    ws.freeze_panes = "D6"
    ws.auto_filter.ref = f"A5:{L(9+NV)}{last}"
    return ws, last

wsS, S_last = cost_sheet("Closed_Std", "Closed Models - Standard (real-time) API Cost",
    "Cost/call = [input tokens x input price + output(+thinking) tokens x output price] / 1M, using token profile on Assumptions. Cache applied only if cache-hit % > 0. Prices flow from Closed_Pricing (green = cross-sheet link).", "std", "2E75B6")
wsB, B_last = cost_sheet("Closed_Batch", "Closed Models - Provider Batch API Cost (50% discount where supported)",
    "Anthropic Message Batches / OpenAI Batch / Gemini Batch API. Async (<=24h) - suits nightly processing. n/a = Batch endpoint not supported for the model.", "batch", "2E75B6")
wsQ, Q_last = cost_sheet("Closed_BigQuery", "Gemini via BigQuery (AI.GENERATE / ML.GENERATE_TEXT) - Batch-rated Cost + BigQuery compute",
    "BigQuery bills Gemini (2.0+) at Vertex BATCH rates (-50%) plus BigQuery on-demand bytes-scanned ($6.25/TiB). Token charges are billed by Vertex (label billing_service='bigquery_ml'). Claude/open models via BigQuery get NO discount, so only Gemini is shown.", "bq", "2E75B6")

# =====================================================================
# 4. SLM MODELS
# =====================================================================
wsM = sheet("SLM_Models", "SLM Candidate Models - Specs, Licence, Origin & Accuracy Proxies",
            f"Sources: Hugging Face model cards / vendor pages / benchmark leaderboards, accessed {ACCESSED}. Numeric columns (blue) are parsed from the model cards; verification status is shown per row.", "70AD47")
NUMS = [  # total_B, active_B, type, bf16, fp8, int4, tps_h100, tps_basis, lora_toks, short
 (29.6, 29.6, "Dense", 59.6, 32.8, 17, 1500, "Derived: 27-32B dense class (Qwen3-32B measured 1,131 BF16 / 2,060 tuned FP8 on H100)", 1200),
 (27, 27, "Dense", 55, 27.5, 16, 1500, "Derived: 27-32B dense class (measured proxy Qwen3-32B)", 1300),
 (30.7, 30.7, "Dense", 69.9, 34.9, 17.5, 1500, "Derived: 27-32B dense class (measured proxy Qwen3-32B)", 1200),
 (35, 3, "MoE", 70, 35, 19, 5500, "Derived: 30B-A3B MoE class (gpt-oss-20b measured 7.8k; planning range 4-8k)", 3500),
 (25.2, 3.8, "MoE", 57.7, 28.8, 14.4, 5500, "Derived: 30B-A3B MoE class (planning range 4-8k)", 3500),
 (20.9, 3.6, "MoE", 42, 13.5, 13.5, 7800, "MEASURED: 7,817 out tok/s on H100, vLLM 0.11, ShareGPT (GPUStack). FP8 col = native MXFP4 size", 4500),
 (30, 30, "Dense", 60, 30, 17.5, 1500, "Derived: 27-32B dense class (measured proxy Qwen3-32B)", 1200),
 (30, 3, "MoE", 60, 30, 18, 6000, "Derived: 30B-A3B hybrid Mamba MoE class (planning range 4-8k)", 4000),
 (12, 12, "Dense", 26.7, 13.4, 6.7, 4000, "Derived from measured Qwen3-14B (3,174) -> 12B class 3.5-4.5k", 5500),
 (8, 8, "Dense", 16, 8, 5, 6000, "Derived: 8B class 5-7k (scaled from Qwen3-14B measured)", 8000),
 (8.9, 8.9, "Dense", 18, 9, 5, 5500, "Derived: 8-9B class", 7500),
 (3, 3, "Dense", 6, 3, 2, 12000, "Derived: scaled from 8B class (~2x)", 20000),
 (8, 4.5, "Dense (PLE)", 17.9, 8.9, 4.5, 9000, "Derived: 4.5B effective params; ~1.5x the 8B class", 14000),
 (14, 14, "Dense", 28, 14, 8, 3200, "MEASURED proxy: Qwen3-14B 3,174 out tok/s H100 (GPUStack)", 4500),
 (31.6, 3.5, "MoE", 63, 32, 18, 6000, "Derived: 30B-A3B hybrid Mamba MoE class", 4000),
 (32, 32, "Dense", 64, 32, 18, 1400, "Derived: 32B dense class (Olmo 3 32B); 7B variant would be ~6,000", 1150),
 (24, 24, "Dense", 48, 24, 13, 1800, "Derived: 24B dense class (between 14B 3.2k and 32B 1.1-1.5k)", 2000),
]
BM = {0: ["Muse Glimmer"], 1: ["Qwen3.8-27B"], 2: ["Gemma 4 31B"], 3: ["Qwen3.6-35B-A3B"], 4: ["Gemma 4 26B A4B"], 5: ["gpt-oss-20b"],
      6: ["Granite 4.2 30B", "Granite 4.2 (all)"], 7: ["Nemotron 3.5 Lightning"], 8: ["Gemma 4 12B"], 9: ["Granite 4.2 8B"], 10: ["Nemotron Nano 9B v2"],
      11: ["Granite 4.2 3B"], 12: ["Gemma 4 E4B"], 13: ["Phi-4"], 14: ["Nemotron 3 Nano"], 15: [], 16: []}
def digest(i):
    out = []
    for b in bench:
        if any(k in b["model"] for k in BM[i]) and "speed" not in b["benchmark"].lower() and "5090" not in b["benchmark"]:
            sc = str(b["score"]); sc = sc if len(sc) < 60 else sc[:57] + "..."
            out.append(f'{b["benchmark"].split("(")[0].strip()}: {sc}{" [proxy]" if b["is_proxy"] else ""}')
    return " | ".join(out) if out else "No published benchmark found - evaluate in-house"
MH = ["#", "User-listed name", "Verified model", "Verification status", "Developer", "Origin / US-corp security note", "Release", "Licence", "Commercial use allowed?", "Licence restrictions",
      "Total params (B)", "Architecture", "Active params (B)", "MoE experts", "Context window",
      "Weights GB BF16", "Weights GB FP8", "Weights GB INT4", "Min GPU class (from model card / weights size)", "Fine-tuning support",
      "Output tok/s on 1xH100 (batch serving) - PLANNING", "Basis of tok/s", "Accuracy proxies (see SLM_Benchmarks)", "Hugging Face URL", "Licence URL", "Other sources", "Notes"]
M0 = 5
hdr(wsM, M0, MH, height=56)
widths(wsM, [4, 22, 30, 38, 18, 34, 14, 26, 12, 30, 10, 14, 10, 14, 18, 10, 10, 10, 40, 48, 14, 44, 70, 42, 42, 50, 60])
for i, m in enumerate(slm):
    r = M0 + 1 + i; n = NUMS[i]
    put(wsM, r, 1, i + 1)
    put(wsM, r, 2, (m["user_listed_name"] if m["user_listed_name"] and "N/A" not in m["user_listed_name"][:4] else "(suggested extra)"), al=wrap)
    put(wsM, r, 3, m["model"] if not m["model"].startswith("NO SUCH") else "Qwen3.6-35B-A3B (closest real match)", font=f_bold, al=wrap)
    put(wsM, r, 4, m["verification_status"], al=wrap, fill=fill_warn if "NOT" in m["verification_status"].upper() or "NO SUCH" in m["model"] else None)
    put(wsM, r, 5, m["developer"], al=wrap); put(wsM, r, 6, m["country_origin"], al=wrap)
    put(wsM, r, 7, m["release_date"], al=wrap); put(wsM, r, 8, m["license"], al=wrap)
    put(wsM, r, 9, "Yes" if m["commercial_use_allowed"] else "No"); put(wsM, r, 10, m["license_restrictions"], al=wrap)
    put(wsM, r, 11, n[0], fmt=NUM1, font=f_in); put(wsM, r, 12, n[2], font=f_in); put(wsM, r, 13, n[1], fmt=NUM1, font=f_in)
    put(wsM, r, 14, str(m["num_experts"]) if m["num_experts"] else "-", al=wrap)
    put(wsM, r, 15, m["context_window"], al=wrap)
    put(wsM, r, 16, n[3], fmt=NUM1, font=f_in); put(wsM, r, 17, n[4], fmt=NUM1, font=f_in); put(wsM, r, 18, n[5], fmt=NUM1, font=f_in)
    put(wsM, r, 19, json.dumps(m["min_gpu_class"]).replace('"', "").strip("{}"), al=wrap)
    put(wsM, r, 20, m["fine_tuning"], al=wrap)
    put(wsM, r, 21, n[6], fmt=NUM, font=f_in, fill=fill_key); put(wsM, r, 22, n[7], al=wrap)
    put(wsM, r, 23, digest(i), al=wrap)
    url_cell(wsM, r, 24, m["hf_url"]); url_cell(wsM, r, 25, m["license_url"])
    put(wsM, r, 26, m["source_urls"] if isinstance(m["source_urls"], str) else " ; ".join(m["source_urls"]), al=wrap, font=f_note)
    put(wsM, r, 27, m["notes"], al=wrap)
    wsM.row_dimensions[r].height = 110
M_last = M0 + len(slm)
SLM_N = len(slm)
wsM.freeze_panes = wsM.cell(row=M0 + 1, column=4)
r = M_last + 2
for t in ["Key corrections to the requested list: 'Qwen 2.5-25B-A3B' does NOT exist (Qwen 2.5 never had an A3B MoE) - replaced by Qwen3.6-35B-A3B; 'Gemma 4 26B-A4B' is 25.2B total / 3.8B active; 'Gemma 4 E4B' is ~8B incl. per-layer embeddings (4.5B effective); 'GPT-oss-20B high' is a reasoning-effort setting of gpt-oss-20b, not a different model; Granite 4.2 is dense (not hybrid).",
          "Output tok/s: NO same-machine published benchmark exists for all models. Column U is a PLANNING figure on one reference machine (1xH100 80GB, vLLM, high-concurrency batch) derived from the few measured points (GPUStack Inference Performance Lab: Qwen3-32B 1,132, Qwen3-14B 3,174, gpt-oss-20b 7,817 tok/s) scaled by model class. +/-50%. Replace with your own `vllm bench serve` results (yellow cells drive all cost sheets).",
          "Origin: Qwen models are Alibaba (China) - Apache 2.0 licensed with local weights, but many US enterprises require a vendor-risk/provenance review. Mistral Small 3.2 is French (EU). All others US-origin.",
          "Accuracy: no summarization- or entity-extraction-specific scores (CNN/DM, SAMSum, DialogSum, NER) were found for these models in vendor cards or public leaderboards. Proxies used: Vectara HHEM summarization hallucination (lower = better), IFBench/IFEval (instruction/format following -> JSON extraction), BFCL (structured output), MMLU-Pro, long-context (RULER/MRCR). Treat as directional only; run an in-house eval on ~200 labelled transcripts before SFT."]:
    c = wsM.cell(row=r, column=2, value="• " + t); c.font = f_note; r += 1

# =====================================================================
# 5. SLM BENCHMARKS
# =====================================================================
wsK = sheet("SLM_Benchmarks", "Accuracy Proxies & Benchmarks for Candidate SLMs (pre-fine-tuning)",
            "No summarization/NER-specific benchmark was published for these models; proxies are flagged. Source URL per row. Recommendation: build an in-house eval on real transcripts.", "70AD47")
hdr(wsK, 4, ["Model", "Benchmark", "Score", "Proxy for summarization/extraction?", "Note", "Source URL"], height=30)
widths(wsK, [38, 46, 44, 14, 90, 60])
for i, b in enumerate(bench):
    r = 5 + i
    put(wsK, r, 1, b["model"], al=wrap); put(wsK, r, 2, b["benchmark"], al=wrap); put(wsK, r, 3, str(b["score"]), al=wrap, font=f_in)
    put(wsK, r, 4, "Proxy" if b["is_proxy"] else "Direct"); put(wsK, r, 5, b.get("note", ""), al=wrap); url_cell(wsK, r, 6, b["source_url"])
wsK.freeze_panes = "A5"; wsK.auto_filter.ref = f"A4:F{4+len(bench)}"

# =====================================================================
# 6. GPU PRICING (catalog) + raw
# =====================================================================
wsG = sheet("GPU_Pricing", "GPU Catalog & Hosting Prices (per GPU-hour, us-central1)",
            f"Sources: https://cloud.google.com/products/compute/pricing/accelerator-optimized (GCE on-demand/Spot/CUD) and https://www.databricks.com/product/pricing/model-serving (DBU/hour) accessed {ACCESSED}. Full 235-record dump on 'GPU_Pricing_Raw'.", "ED7D31")
GH = ["GPU", "VRAM (GB)", "GCP instance (bundled vCPU/RAM)", "GPUs per instance", "Instance $/h On-demand", "Instance $/h Spot", "Instance $/h 1-yr CUD", "Instance $/h 3-yr CUD",
      "GCP On-demand", "GCP Spot", "GCP 1-yr CUD", "GCP 3-yr CUD", "Databricks DBU/h per GPU", "Databricks serving (DBU)", "SELECTED $/GPU-h (Assumptions)", "Throughput vs H100", "Databricks SKU used / notes", "Source"]
hdr(wsG, 5, GH, height=56)
widths(wsG, [20, 9, 22, 9, 12, 12, 12, 12, 11, 11, 11, 11, 12, 13, 14, 11, 60, 60])
GPUS = [
    ("L4 24GB", 24, "g2-standard-4", 1, 0.706832276, 0.423016, 0.445304335, 0.318074524, 20.0, 0.08, "Databricks has no L4 SKU: priced as A10G x1 (20 DBU/h), same 24GB class. Throughput ~0.08x H100 is a bandwidth-ratio estimate (0.3 vs 3.35 TB/s) - UNVERIFIED."),
    ("A100 80GB", 80, "a2-ultragpu-1g", 1, 5.06879789, 2.976774795, 4.199690411, 3.499997559, 628 / 8, 0.65, "Databricks A100-80GB is sold as x8 only (628 DBU/h) -> /8 per GPU. Throughput 0.59-0.73x H100 measured (Qwen3-14B/32B, GPUStack)."),
    ("RTX PRO 6000 96GB", 96, "g4-standard-48", 1, 4.49993, 1.6808, 3.105, 1.97945, 80.0, 0.55, "Databricks RTX PRO 6000 x1 = 80 DBU/h. Throughput 0.55x ASSUMED from memory-bandwidth ratio (~1.8 vs 3.35 TB/s) - UNVERIFIED. g4-standard-48 includes 48 vCPUs."),
    ("H100 80GB", 80, "a3-highgpu-8g", 8, 88.490000119, 52.700460712, 61.383674231, 38.864383195, 100.0, 1.0, "Reference machine. Databricks H100 x1 = 100 DBU/h. Per-GPU = 8-GPU instance / 8. Spot can be pre-empted (risky for long SFT)."),
]
G0 = 6
GR = {}
for i, g in enumerate(GPUS):
    r = G0 + i; GR[g[0]] = r
    put(wsG, r, 1, g[0], font=f_bold); put(wsG, r, 2, g[1], font=f_in); put(wsG, r, 3, g[2]); put(wsG, r, 4, g[3], font=f_in)
    for j in range(4): put(wsG, r, 5 + j, g[4 + j], fmt='$#,##0.000', font=f_in)
    for j in range(4): put(wsG, r, 9 + j, f"={L(5+j)}{r}/$D{r}", fmt='$#,##0.00')
    put(wsG, r, 13, g[8], fmt='0.0', font=f_in)
    put(wsG, r, 14, f"=M{r}*{A['dbu']}", fmt='$#,##0.00')
    h = A["host"]
    put(wsG, r, 15, f'=IF({h}="GCP On-demand",I{r},IF({h}="GCP Spot",J{r},IF({h}="GCP 1-yr CUD",K{r},IF({h}="GCP 3-yr CUD",L{r},N{r}))))', fmt='$#,##0.00', fill=fill_key)
    put(wsG, r, 16, g[9], fmt='0.00', font=f_in); put(wsG, r, 17, g[10], al=wrap)
    url_cell(wsG, r, 18, "https://cloud.google.com/products/compute/pricing/accelerator-optimized ; https://www.databricks.com/product/pricing/model-serving")
    wsG.row_dimensions[r].height = 52
G_last = G0 + len(GPUS) - 1
r = G_last + 2
for t in ["Databricks $/DBU ($0.07) is UNVERIFIED (third-party) - official tiles are JS-rendered. Databricks serving prices include the VM; GCE prices include bundled vCPU/RAM/SSD. Databricks H100 ($7.00/h) looking cheaper than GCE ($11.06/h) should be confirmed with Databricks before relying on it.",
          "Other GPUs on GCP (us-central1, per instance-hour on-demand): A100 40GB a2-highgpu-1g $3.673 | H200 a3-ultragpu-8g $84.81 (8 GPUs) | B200 a4-highgpu-8g no on-demand price (Spot $39.63; Flex-start $64.44; 1-yr CUD $88.93) - see GPU_Pricing_Raw.",
          "Databricks Foundation Model Fine-tuning is END OF LIFE; replacement 'AI Runtime' (serverless A10/H100) is listed for AWS and Azure only. On Databricks-on-GCP, fine-tune on classic GPU clusters (a2/a3/g2) = GCE VM price + DBU surcharge (DBU/h for GPU instances not retrievable - use Databricks' calculator).",
          "Vertex AI endpoints add a management fee: e.g. g2-standard-4 $0.813/h (vs $0.707 GCE), a2-ultragpu-1g $5.78/h (vs $5.07), a3-highgpu-8g custom training $101.01/h (vs $88.49)."]:
    c = wsG.cell(row=r, column=1, value="• " + t); c.font = f_note; r += 1
wsG.freeze_panes = "B6"

wsR = sheet("GPU_Pricing_Raw", "GPU / Platform Pricing - Full Research Dump (235 records)", f"Every record carries its source URL; accessed {ACCESSED}.", "ED7D31")
RH = ["Provider", "Platform", "GPU", "GPU count", "VRAM", "Instance type", "Price", "Pricing model", "Region", "Unit", "Source URL", "Accessed", "Notes"]
hdr(wsR, 4, RH, height=30)
widths(wsR, [14, 44, 22, 8, 12, 20, 12, 30, 12, 26, 60, 11, 70])
for i, g in enumerate(gpu_raw):
    r = 5 + i
    vals = [g["provider"], g["platform"], g["gpu"], g["gpu_count"], g["vram"], g["instance_type"], g["price_per_hour"], g["pricing_model"], g["region"], g["unit"], None, g["accessed"], g["notes"]]
    for j, v in enumerate(vals):
        if j == 10: url_cell(wsR, r, 11, g["source_url"])
        else: put(wsR, r, 1 + j, v, fmt='#,##0.000' if j == 6 else None, al=wrap if j in (1, 12) else None)
wsR.freeze_panes = "A5"; wsR.auto_filter.ref = f"A4:M{4+len(gpu_raw)}"

# =====================================================================
# 7. SLM SERVING FIT (feasibility + cost @ ref vol per GPU)
# =====================================================================
wsF = sheet("SLM_Serving_Fit", "SLM Serving - GPU Fit, Throughput & Cost at Reference Volume",
            "Per model: does it fit each GPU at the chosen precision, and what does one day of batch processing cost at the reference volume? Green = link to other sheet. Cheapest feasible GPU feeds SLM_Daily_Cost.", "7030A0")
GN = [g[0] for g in GPUS]
FH = ["Model", "Total params (B)", "Active params (B)", "Serving weights GB (chosen precision)", "Output tok/s on H100", "GPU-seconds per call on H100"] + \
     [f"Fits {g}?" for g in GN] + [f"$/day @ ref vol: {g}" for g in GN] + ["Cheapest feasible $/day", "Cheapest feasible GPU", "H100 $/day (reference machine)", "H100 / cheapest", "Always-on 24h $/day (cheapest GPU, 1 replica min.)"]
hdr(wsF, 5, FH, height=70)
widths(wsF, [30, 9, 9, 13, 12, 12] + [11] * 4 + [13] * 4 + [14, 18, 14, 11, 18])
F0 = 6
for i in range(SLM_N):
    r = F0 + i; mr = M0 + 1 + i
    put(wsF, r, 1, f"=SLM_Models!C{mr}", font=f_link, al=wrap)
    put(wsF, r, 2, f"=SLM_Models!K{mr}", fmt=NUM1, font=f_link); put(wsF, r, 3, f"=SLM_Models!M{mr}", fmt=NUM1, font=f_link)
    put(wsF, r, 4, f'=IF({A["prec"]}="BF16",SLM_Models!P{mr},IF({A["prec"]}="FP8",SLM_Models!Q{mr},SLM_Models!R{mr}))', fmt=NUM1, font=f_link)
    put(wsF, r, 5, f"=SLM_Models!U{mr}", fmt=NUM, font=f_link)
    put(wsF, r, 6, f"={A['out_slm']}/E{r}+{A['tot_in']}/(E{r}*{A['prefill']})", fmt='0.000')
    for k, g in enumerate(GN):
        gr = GR[g]
        put(wsF, r, 7 + k, f"=IF($D{r}<=GPU_Pricing!$B${gr}*{A['fit']},1,0)", fmt='0')
        put(wsF, r, 11 + k, f'=IF({L(7+k)}{r}=1,({A["ref_vol"]}*$F{r}/GPU_Pricing!$P${gr}/3600/{A["util"]}+{A["startup"]})*GPU_Pricing!$O${gr},"n/a")', fmt=USD)
    put(wsF, r, 15, f'=IF(COUNT(K{r}:N{r})=0,"n/a",MIN(K{r}:N{r}))', fmt=USD, fill=fill_sub, font=f_bold)
    put(wsF, r, 16, f'=IF(ISNUMBER(O{r}),SUBSTITUTE(INDEX($K$5:$N$5,MATCH(O{r},K{r}:N{r},0)),"$/day @ ref vol: ",""),"No single-GPU fit")', al=wrap)
    hr = f"GPU_Pricing!$P${GR['H100 80GB']}"
    put(wsF, r, 17, f'=IF(J{r}=1,N{r},"n/a")', fmt=USD)
    put(wsF, r, 18, f'=IF(AND(ISNUMBER(Q{r}),ISNUMBER(O{r})),Q{r}/O{r},"")', fmt='0.0"x"')
    # always-on: 24h x price of cheapest GPU x replicas needed
    put(wsF, r, 19, f'=IF(ISNUMBER(O{r}),24*INDEX(GPU_Pricing!$O${G0}:$O${G_last},MATCH(P{r},GPU_Pricing!$A${G0}:$A${G_last},0))*MAX(1,ROUNDUP({A["ref_vol"]}*F{r}/INDEX(GPU_Pricing!$P${G0}:$P${G_last},MATCH(P{r},GPU_Pricing!$A${G0}:$A${G_last},0))/3600/{A["util"]}/24,0)),"n/a")', fmt=USD)
F_last = F0 + SLM_N - 1
put(wsF, F_last + 2, 1, "Method: GPU-seconds/call = output tokens / tok/s + input tokens / (tok/s x prefill multiple). $/day = (volume x GPU-s per call / GPU throughput factor / 3600 / utilisation + start-up hours) x $/GPU-h. Elastic on-demand billing: scales out for as many GPU-hours as needed, billed only while running. 'Always-on' = 24h x replicas needed for sustained load. Fit check = weights <= usable VRAM share (KV cache not separately sized).", font=f_note, bd=False)
wsF.freeze_panes = "B6"
wsF.conditional_formatting.add(f"G{F0}:J{F_last}", ColorScaleRule(start_type="num", start_value=0, start_color="F8CBAD", end_type="num", end_value=1, end_color="C6EFCE"))

# =====================================================================
# 8. SLM DAILY COST (1K..15K) - two blocks
# =====================================================================
wsD = sheet("SLM_Daily_Cost", "Self-hosted SLM - Daily Inference Cost, 1K-15K calls/day (excludes one-time SFT)",
            "Block A: every model on the SAME reference machine (1xH100) - like-for-like. Block B: cheapest GPU each model fits on (lower-GPU option). Elastic on-demand billing at the hosting option chosen on Assumptions.", "7030A0")
def vol_header(ws, row, first_cols):
    hdr(ws, row, first_cols + [None] * NV + ["Monthly @ ref volume", "Annual @ ref volume"], height=44)
    for k in range(NV):
        c = ws.cell(row=row, column=len(first_cols) + 1 + k)
        c.value = f"={A['vol_step']}" if k == 0 else f"={L(len(first_cols)+k)}{row}+{A['vol_step']}"
        c.number_format = '#,##0" calls/day"'; c.font = f_hdr
widths(wsD, [30, 20, 11, 12, 13] + [11] * NV + [14, 14])
wsD["A4"] = "BLOCK A - Reference machine: 1 x H100 80GB"; wsD["A4"].font = f_bold
FC = ["Model", "GPU", "$/GPU-h", "Variable $/call", "Fixed $/day (start-up)"]
vol_header(wsD, 5, FC)
D0 = 6
nc = len(FC)
def daily_block(start, gpu_expr, label):
    for i in range(SLM_N):
        r = start + i; fr = F0 + i
        put(wsD, r, 1, f"=SLM_Serving_Fit!A{fr}", font=f_link, al=wrap)
        if label == "A":
            put(wsD, r, 2, "H100 80GB")
            gi = f"MATCH($B{r},GPU_Pricing!$A${G0}:$A${G_last},0)"
        else:
            put(wsD, r, 2, f'=IF(ISNUMBER(SLM_Serving_Fit!O{fr}),SLM_Serving_Fit!P{fr},"n/a")', font=f_link)
            gi = f"MATCH($B{r},GPU_Pricing!$A${G0}:$A${G_last},0)"
        put(wsD, r, 3, f'=IFERROR(INDEX(GPU_Pricing!$O${G0}:$O${G_last},{gi}),"n/a")', fmt=USD, font=f_link)
        put(wsD, r, 4, f'=IFERROR(SLM_Serving_Fit!F{fr}/INDEX(GPU_Pricing!$P${G0}:$P${G_last},{gi})/3600/{A["util"]}*C{r},"n/a")', fmt='$0.00000')
        put(wsD, r, 5, f'=IFERROR({A["startup"]}*C{r},"n/a")', fmt=USD)
        for k in range(NV):
            put(wsD, r, nc + 1 + k, f'=IF(ISNUMBER($D{r}),{L(nc+1+k)}${start-1}*$D{r}+$E{r},"n/a")', fmt=USD0)
        put(wsD, r, nc + NV + 1, f'=IF(ISNUMBER($D{r}),({A["ref_vol"]}*$D{r}+$E{r})*{A["days"]},"n/a")', fmt=USD0)
        put(wsD, r, nc + NV + 2, f'=IF(ISNUMBER($D{r}),({A["ref_vol"]}*$D{r}+$E{r})*365,"n/a")', fmt=USD0)
    return start + SLM_N - 1
# Block A fit gating: H100 fits all in list? (all fit at FP8; BF16 may not) -> use n/a when not fit
dA_last = daily_block(D0, None, "A")
# gate by fit on H100
for i in range(SLM_N):
    r = D0 + i
    wsD.cell(row=r, column=4).value = f'=IF(SLM_Serving_Fit!J{F0+i}=1,' + wsD.cell(row=r, column=4).value[1:] + ',"n/a")'
dB_hdr = dA_last + 3
wsD.cell(row=dB_hdr - 1, column=1, value="BLOCK B - Cheapest feasible single GPU per model (lower-GPU option)").font = f_bold
vol_header(wsD, dB_hdr, FC)
dB0 = dB_hdr + 1
dB_last = daily_block(dB0, None, "B")
wsD.freeze_panes = "B6"
for a, b in ((D0, dA_last), (dB0, dB_last)):
    wsD.conditional_formatting.add(f"{L(nc+1)}{a}:{L(nc+NV)}{b}", ColorScaleRule(start_type="min", start_color="C6EFCE", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="max", end_color="F8CBAD"))

# =====================================================================
# 9. SFT COST
# =====================================================================
wsT = sheet("SFT_Cost", "One-time Supervised Fine-Tuning (SFT) Cost per Model",
            "LoRA SFT on 1xH100 (full fine-tune multi-GPU shown for comparison). GPU-hours = training tokens / (tok/s per GPU) / 3600. All tok/s are ESTIMATES (+/-50%) anchored on NVIDIA NeMo AutoModel H100 figures and a VESSL Gemma-4-31B LoRA run - validate with a 100-step pilot.", "C00000")
TH = ["Model", "Total params (B)", "LoRA tok/s per H100 (est.)", "Training tokens (M)", "LoRA GPU-hours (1 run)", "Full-FT multiplier", "Full-FT GPU-hours (1 run)",
      "LoRA $ 1 run - GCP On-demand", "LoRA $ 1 run - GCP Spot", "LoRA $ 1 run - GCP 3-yr CUD", "LoRA $ 1 run - SELECTED hosting", "SFT budget (LoRA x runs, selected hosting)", "Full-FT $ 1 run (selected hosting)", "Basis"]
hdr(wsT, 5, TH, height=62)
widths(wsT, [30, 9, 13, 12, 12, 11, 12, 14, 14, 14, 14, 16, 14, 60])
T0 = 6
hr = GR["H100 80GB"]
tokens_M = f"{A['sft_n']}*{A['sft_tok']}*{A['sft_ep']}/1000000"
for i in range(SLM_N):
    r = T0 + i; mr = M0 + 1 + i
    n = NUMS[i]
    put(wsT, r, 1, f"=SLM_Models!C{mr}", font=f_link, al=wrap); put(wsT, r, 2, f"=SLM_Models!K{mr}", fmt=NUM1, font=f_link)
    put(wsT, r, 3, n[8], fmt=NUM, font=f_in, fill=fill_key)
    put(wsT, r, 4, f"={tokens_M}", fmt=NUM1)
    put(wsT, r, 5, f"=D{r}*1000000/C{r}/3600", fmt=NUM1)
    mult = 1.5 if n[0] <= 9 else 1.8
    put(wsT, r, 6, mult, fmt='0.0', font=f_in)
    put(wsT, r, 7, f"=E{r}*F{r}", fmt=NUM1)
    put(wsT, r, 8, f"=E{r}*GPU_Pricing!$I${hr}", fmt=USD0); put(wsT, r, 9, f"=E{r}*GPU_Pricing!$J${hr}", fmt=USD0)
    put(wsT, r, 10, f"=E{r}*GPU_Pricing!$L${hr}", fmt=USD0); put(wsT, r, 11, f"=E{r}*GPU_Pricing!$O${hr}", fmt=USD0, fill=fill_sub)
    put(wsT, r, 12, f"=K{r}*{A['sft_iter']}", fmt=USD0, font=f_bold, fill=fill_sub)
    put(wsT, r, 13, f"=G{r}*GPU_Pricing!$O${hr}", fmt=USD0)
    put(wsT, r, 14, "LoRA tok/s: NeMo AutoModel H100 (8B LoRA 12.5k best case, derated to 8k), VESSL Gemma-4-31B LoRA ~1.2k/GPU, MoE ~4k derated; others interpolated by size. Full-FT multiplier: 1.5x (<=9B) / 1.8x (larger) = own assumption; needs 8xH100 FSDP for >=8B.", al=wrap, font=f_note)
    wsT.row_dimensions[r].height = 40
T_last = T0 + SLM_N - 1
r = T_last + 3
wsT.cell(row=r - 1, column=1, value="Reference: managed per-token SFT on Vertex AI (open models; no 30B-A3B / Granite / Nemotron / Gemma 4 entries) - priced on the same training tokens").font = f_bold
hdr(wsT, r, ["Vertex managed SFT model", "$ per 1M training tokens", "Training tokens (M)", "Cost 1 run", "Source"], height=30)
VSFT = [("Llama 3.1 8B", 0.67), ("Gemma 3 12B", 1.82), ("Qwen3 8B", 4.18), ("Qwen3 32B", 6.57), ("Gemma 3 27B", 6.83)]
for j, (n_, p_) in enumerate(VSFT):
    rr = r + 1 + j
    put(wsT, rr, 1, n_); put(wsT, rr, 2, p_, fmt=USD, font=f_in); put(wsT, rr, 3, f"={tokens_M}", fmt=NUM1); put(wsT, rr, 4, f"=B{rr}*C{rr}", fmt=USD0)
    url_cell(wsT, rr, 5, "https://cloud.google.com/vertex-ai/generative-ai/pricing (supervised tuning, open models) - accessed 2026-11-08")
rr = r + len(VSFT) + 2
for t in ["Databricks Foundation Model Fine-tuning (Mosaic AI Model Training) is END OF LIFE; last published example Llama 3.1 8B: 4,400 DBU / 500M words (~$2,860 at $0.65/DBU) - https://www.databricks.com/product/pricing/foundation-model-training. Replacement AI Runtime is AWS/Azure only. On Databricks-on-GCP use classic GPU clusters (GCE VM + DBU surcharge).",
          "Spot GPUs are cheaper but can be pre-empted mid-run; checkpoint every ~15 min. 'SELECTED hosting' follows the Assumptions dropdown (Databricks DBU option applies the serving rate - for training use classic-cluster pricing instead).",
          "Run count: SFT budget multiplies one run by the 'Training runs' assumption (default 3) to cover re-runs and hyper-parameter sweeps; data curation/labelling cost is NOT included."]:
    wsT.cell(row=rr, column=1, value="• " + t).font = f_note; rr += 1
wsT.freeze_panes = "B6"

# =====================================================================
# 10. COMPARISON
# =====================================================================
wsC = sheet("Comparison", "Closed LLM vs Fine-tuned SLM - Cost Comparison & Break-even",
            "At the reference volume from Assumptions. Closed = Standard API unless stated; 'Batch' = provider batch API (50% off). SLM = cheapest feasible single GPU, elastic on-demand billing. Accuracy is NOT equalised - see README caveats.", "00B050")
widths(wsC, [44, 22, 16, 16, 16, 16, 16, 16, 16, 16, 18])
wsC["A4"] = "Closed-model reference points (monthly cost at reference volume)"; wsC["A4"].font = f_bold
hdr(wsC, 5, ["Reference", "Model", "Monthly $", "Annual $"], height=26)
sref = f"Closed_Std!$B$6:$B${S_last}"; smon = f"Closed_Std!${L(6+NV)}$6:${L(6+NV)}${S_last}"
bref = f"Closed_Batch!$B$6:$B${B_last}"; bmon = f"Closed_Batch!${L(6+NV)}$6:${L(6+NV)}${B_last}"
qref = f"Closed_BigQuery!$B$6:$B${Q_last}"; qmon = f"Closed_BigQuery!${L(6+NV)}$6:${L(6+NV)}${Q_last}"
refs = [("Cheapest closed model - Standard API", f"=INDEX({sref},MATCH(MIN({smon}),{smon},0))", f"=MIN({smon})"),
        ("Cheapest closed model - Batch API", f"=INDEX({bref},MATCH(MIN({bmon}),{bmon},0))", f"=MIN({bmon})"),
        ("Cheapest Gemini via BigQuery batch", f"=INDEX({qref},MATCH(MIN({qmon}),{qmon},0))", f"=MIN({qmon})"),
        ("Selected baseline - Standard API", f"={A['base']}", f"=INDEX({smon},MATCH({A['base']},{sref},0))"),
        ("Selected baseline - Batch API", f"={A['base']}", f"=IFERROR(INDEX({bmon},MATCH({A['base']},{bref},0)),\"n/a\")"),
        ("Selected baseline - BigQuery batch (Gemini only)", f"={A['base']}", f"=IFERROR(INDEX({qmon},MATCH({A['base']},{qref},0)),\"n/a\")")]
for i, (a, b, c) in enumerate(refs):
    r = 6 + i
    put(wsC, r, 1, a); put(wsC, r, 2, b, font=f_link); put(wsC, r, 3, c, fmt=USD0, font=f_link); put(wsC, r, 4, f'=IF(ISNUMBER(C{r}),C{r}*12/{A["days"]}*{A["days"]}/{A["days"]}*(365/12)/({A["days"]}) ,"n/a")' if False else f'=IF(ISNUMBER(C{r}),C{r}*365/{A["days"]},"n/a")', fmt=USD0)
c0 = 6 + len(refs) + 2
wsC.cell(row=c0 - 1, column=1, value="SLM economics vs baselines (selected baseline = Assumptions dropdown; break-even = months of operation to repay SFT)").font = f_bold
CH = ["Model", "Cheapest feasible GPU", "SLM monthly inference $", "SFT one-time budget $", "SFT amortised $/month", "SLM all-in $/month",
      "Baseline monthly $ (Std API)", "Saving vs baseline $/mo (all-in)", "Break-even months vs baseline (Std)", "Break-even months vs cheapest closed (Std)", "Verdict vs baseline (Std)"]
hdr(wsC, c0, CH, height=62)
for i in range(SLM_N):
    r = c0 + 1 + i; fr = F0 + i; dr = dB0 + i; tr_ = T0 + i
    put(wsC, r, 1, f"=SLM_Serving_Fit!A{fr}", font=f_link, al=wrap)
    put(wsC, r, 2, f"=SLM_Daily_Cost!B{dr}", font=f_link)
    put(wsC, r, 3, f"=SLM_Daily_Cost!{L(nc+NV+1)}{dr}", fmt=USD0, font=f_link)
    put(wsC, r, 4, f"=SFT_Cost!L{tr_}", fmt=USD0, font=f_link)
    put(wsC, r, 5, f"=D{r}/{A['amort']}", fmt=USD0)
    put(wsC, r, 6, f'=IF(ISNUMBER(C{r}),C{r}+E{r},"n/a")', fmt=USD0, font=f_bold)
    put(wsC, r, 7, "=$C$9", fmt=USD0)
    put(wsC, r, 8, f'=IF(ISNUMBER(F{r}),G{r}-F{r},"n/a")', fmt=USD0)
    put(wsC, r, 9, f'=IF(ISNUMBER(C{r}),IF(G{r}>C{r},D{r}/(G{r}-C{r}),"never (SLM run cost >= baseline)"),"n/a")', fmt='0.0')
    put(wsC, r, 10, f'=IF(ISNUMBER(C{r}),IF($C$6>C{r},D{r}/($C$6-C{r}),"never (SLM run cost >= cheapest closed)"),"n/a")', fmt='0.0')
    put(wsC, r, 11, f'=IF(NOT(ISNUMBER(H{r})),"No single-GPU fit",IF(H{r}>0,"SLM cheaper","Closed API cheaper"))', font=f_bold)
cl = c0 + SLM_N
for t in ["READ THIS: break-even compares cost only. Closed frontier models are zero-shot capable; SLMs need the SFT in this model and an in-house eval to confirm summarization/extraction quality and hallucination rate. Data-residency/privacy value of self-hosting (call transcripts contain PII/account data) is not priced here.",
          "SLM costs assume elastic on-demand GPUs used only for the batch window; real-time (sub-second) serving needs an always-on endpoint - see 'Always-on 24h' column on SLM_Serving_Fit.",
          "Add engineering/MLOps effort, eval/labelling, monitoring and storage - not included on any sheet."]:
    wsC.cell(row=cl + 2, column=1, value="• " + t).font = f_note; cl += 1
wsC.freeze_panes = f"B{c0+1}"

# =====================================================================
# 11. DEPLOYMENT OPTIONS
# =====================================================================
wsO = sheet("Deployment_Options", "GCP vs Databricks - Options to Host & Fine-tune the SLMs On-Demand",
            f"Prices verified on official pages unless marked UNVERIFIED. Accessed {ACCESSED}.", "ED7D31")
OH = ["Platform", "Option", "Used for", "GPU / unit", "Billing & price", "Pros", "Cons / caveats", "Source"]
hdr(wsO, 4, OH, height=30)
widths(wsO, [14, 34, 14, 26, 46, 44, 54, 56])
OPT = [
 ("GCP", "Compute Engine / GKE - GPU VMs (self-managed vLLM, Unsloth/Axolotl/NeMo)", "SFT + serving", "L4 (g2), A100 80GB (a2-ultra), RTX PRO 6000 (g4), H100/H200 (a3), B200 (a4)",
  "Per instance-hour, us-central1. On-demand: L4 $0.707, A100-80 $5.069, RTX PRO 6000 $4.50, 8xH100 $88.49 ($11.06/GPU). Spot ~40-50% less (8xH100 $52.70); 1-yr CUD 8xH100 $61.38, 3-yr $38.86.",
  "Lowest raw $/GPU-h; full control; Spot/CUD discounts; any model/framework.", "You operate the stack (drivers, vLLM, autoscaling, security). H100 only sold in 8-GPU a3-highgpu-8g in the table used (smaller a3 shapes may exist). B200 has no on-demand price (Spot/Flex/Calendar only).",
  "https://cloud.google.com/products/compute/pricing/accelerator-optimized"),
 ("GCP", "Vertex AI custom training (managed jobs)", "SFT", "Same GPUs as GCE + Vertex management fee",
  "a3-highgpu-8g $101.01/h (vs $88.49 GCE); g4-standard-48 $5.17/h; Flexible Savings Plans 1-yr $90.91, 3-yr $80.81 (a3).",
  "Managed orchestration, checkpoints, experiment tracking, IAM/VPC-SC for PII data.", "~14% premium over raw GCE; Vertex now branded 'Gemini Enterprise Agent Platform'.",
  "https://cloud.google.com/vertex-ai/pricing"),
 ("GCP", "Vertex AI managed supervised fine-tuning (per token)", "SFT", "Managed; per 1M training tokens",
  "Llama 3.1 8B $0.67; Gemma 3 12B $1.82; Qwen3 8B $4.18; Qwen3 32B $6.57; Gemma 3 27B $6.83 per 1M tokens (x tokens x epochs).",
  "Zero infrastructure; cheap for 8B-12B.", "Limited model catalog - none of Gemma 4, Granite 4.2, Nemotron 3.5, gpt-oss, Muse Glimmer, 30B-A3B MoEs listed.",
  "https://cloud.google.com/vertex-ai/generative-ai/pricing"),
 ("GCP", "Vertex AI endpoints / Model Garden self-deploy / batch prediction", "Serving", "L4, A100, RTX PRO 6000, H100 node-hours",
  "g2-standard-4 (1xL4) $0.813/h; a2-ultragpu-1g $5.78/h; g4-standard-48 $5.17/h (Flexible Savings Plan 1-yr/3-yr ~10-20% lower).",
  "Autoscaling, private endpoints, scale-to-zero batch jobs.", "Management fee vs GCE; cold-start for batch jobs.", "https://cloud.google.com/vertex-ai/pricing"),
 ("GCP", "Vertex AI MaaS (managed open-model API)", "Serving (no fine-tune)", "Per-token",
  "Gemma 4 26B $0.15 in / $0.60 out per 1M (official); gpt-oss-120b $0.09/$0.36 (UNVERIFIED third-party). Batch inference only for gpt-oss-120b (global).",
  "No GPUs to manage.", "Serves base models only - NOT your fine-tuned weights; not applicable after SFT.", "https://cloud.google.com/vertex-ai/generative-ai/pricing"),
 ("Databricks", "Model Serving - GPU serverless endpoint (custom/fine-tuned model)", "Serving", "T4 10.48 | A10G x1 20 | L40S x1 44.86 | RTX PRO 6000 x1 80 | H100 x1 100 | A100-80 x8 628 DBU/h",
  "DBU/h x $/DBU. $0.07/DBU (UNVERIFIED, third-party, GCP US regions) -> H100 x1 $7.00/h, A10G $1.40/h, RTX PRO 6000 $5.60/h.",
  "Serverless, autoscale to zero, Unity Catalog governance, VM included in price.", "Rates JS-rendered on official page; availability of each SKU on Databricks-on-GCP not confirmed; $/DBU needs confirming.",
  "https://www.databricks.com/product/pricing/model-serving"),
 ("Databricks", "Provisioned Throughput / batch inference (ai_query)", "Serving", "Per 'model unit' DBU/h",
  "Llama 3.1 8B & GPT-OSS-20B 53.571 DBU/h (~$3.75/h @ $0.07); Gemma 3 12B & GPT-OSS-120B 71.429 (~$5.00/h); Llama 3.3 70B 85.714. Batch (ai_query) billed at same DBU/h.",
  "Good for nightly batch summarization with SQL ai_query.", "Only supported foundation models / fine-tuned variants of them; tokens/s per unit NOT verified.", "https://www.databricks.com/product/pricing/foundation-model-serving"),
 ("Databricks", "Foundation Model Fine-tuning (Mosaic AI Model Training)", "SFT", "Per-token",
  "END OF LIFE. Last example: Llama 3.1 8B 4,400 DBU per 500M words (~$2,860 @ $0.65/DBU).", "-", "No longer available; do not plan on it.", "https://docs.databricks.com/aws/en/large-language-models/foundation-model-training/index.html"),
 ("Databricks", "AI Runtime (serverless GPU compute)", "SFT", "A10, 1xH100, 8xH100 (B300 waitlist)",
  "H100 ~$7.00/GPU-h; A10 $2.50 (AWS)/$4.90 (Azure) - UNVERIFIED (Flexera blog).", "Notebook-based serverless fine-tuning.", "Listed for AWS and Azure ONLY - not on Databricks-on-GCP.", "https://docs.databricks.com/aws/en/machine-learning/ai-runtime/overview"),
 ("Databricks", "Classic GPU cluster on GCP (a2/a3/g2 instances)", "SFT + serving", "g2 (L4), a2-highgpu (A100-40), a2-ultragpu (A100-80), a3-highgpu-8g (H100)",
  "GCE VM price (see GCP row) + DBU/h surcharge for the instance (DBU/h and $/DBU NOT retrievable - use Databricks calculator).", "Full-control SFT inside Databricks workspace next to the data.", "Double billing (VM + DBU); no H200/B200/G4 on GCP classic.", "https://docs.databricks.com/gcp/en/compute/gpu"),
]
for i, o in enumerate(OPT):
    r = 5 + i
    for j, v in enumerate(o):
        if j == 7: url_cell(wsO, r, 8, v)
        else: put(wsO, r, 1 + j, v, al=wrap, font=f_bold if j == 1 else None)
    wsO.row_dimensions[r].height = 95
wsO.freeze_panes = "C5"

# =====================================================================
# 12. SOURCES
# =====================================================================
wsU = sheet("Sources", "Source Register (all URLs referenced, with access date)", None, "7F7F7F")
hdr(wsU, 3, ["#", "URL", "Used for"], height=22); widths(wsU, [5, 100, 70])
urls = {}
def collect(txt, use):
    for u in re.findall(r"https?://[^\s;,)'\"]+", txt or ""):
        u = u.rstrip(".")
        urls.setdefault(u, set()).add(use)
for g in load("gemini_pricing"): collect(g.get("source_url", ""), "Gemini pricing")
for c in claude: collect(c.get("source_url", ""), "Claude pricing")
for o in openai: collect(o.get("source_url", ""), "OpenAI pricing")
for m in slm:
    collect(m["hf_url"], "SLM specs"); collect(m["license_url"], "SLM licence")
    su = m["source_urls"] if isinstance(m["source_urls"], str) else " ; ".join(m["source_urls"]); collect(su, "SLM specs")
for b in bench: collect(b["source_url"], "SLM benchmarks")
for g in gpu_raw: collect(g["source_url"], "GPU pricing")
for t in ["https://cloud.google.com/bigquery/pricing", "https://cloud.google.com/bigquery/docs/choose-text-generation-function"]: collect(t, "BigQuery batch discount")
for j, (u, use) in enumerate(sorted(urls.items())):
    r = 4 + j
    put(wsU, r, 1, j + 1); url_cell(wsU, r, 2, u); put(wsU, r, 3, ", ".join(sorted(use)))
put(wsU, 5 + len(urls), 2, f"All pages accessed {ACCESSED}.", font=f_note, bd=False)

# =====================================================================
# 0. README (first)
# =====================================================================
ws0 = sheet("README", "LLM vs SLM Cost Analysis - Energy Retail Call Transcript AI", f"Branch dev-v1 · built {ACCESSED} · all calculations are live formulas driven by the Assumptions sheet", "1F3864")
ws0.column_dimensions["A"].width = 30; ws0.column_dimensions["B"].width = 140
rows0 = [
 ("PURPOSE", "Compare the cost of running call-transcript summarization / entity extraction / classification on closed LLM APIs (Gemini 2.5 Flash+, Claude Sonnet 4.5+, GPT-5+) versus fine-tuned open small language models (SLMs) hosted on on-demand GPUs (GCP or Databricks), at 1K-15K calls/day."),
 ("HOW TO USE", "Edit blue cells on 'Assumptions' (call lengths, mix, task sizes, thinking tokens, hosting option, SFT size). Yellow cells are the most influential levers. Every other sheet recalculates."),
 ("SHEET GUIDE", ""),
 ("  Assumptions", "Token profile per call (residential vs C&I), tasks, thinking/cache, BigQuery, GPU hosting, throughput, SFT, baseline selector."),
 ("  Closed_Pricing", "Per-model $/1M tokens (input, output, cached, batch, post-promo) for 12 Gemini + 16 Claude + 22 OpenAI models, each with source URL + date."),
 ("  Closed_Std / _Batch / _BigQuery", "Daily cost 1K..15K calls/day for every closed model: real-time API; provider Batch API (-50%); Gemini via BigQuery (batch-rated + BigQuery compute)."),
 ("  SLM_Models", "17 SLMs: licence, commercial use, origin, params, MoE/active params, context, weights size, min GPU, fine-tuning support, planning tok/s on 1xH100, accuracy proxies."),
 ("  SLM_Benchmarks", "All accuracy proxy rows with source (no summarization/NER-specific scores exist publicly)."),
 ("  GPU_Pricing / _Raw", "GPU catalog used in the calculations (GCP + Databricks) and the 235-record research dump with sources."),
 ("  Deployment_Options", "GCP vs Databricks options for SFT and serving, pros/cons, prices."),
 ("  SLM_Serving_Fit", "Which GPU each model fits on (incl. lower-GPU options such as L4) and cost/day at reference volume."),
 ("  SLM_Daily_Cost", "Daily cost 1K..15K calls: Block A same-machine (1xH100), Block B cheapest feasible GPU."),
 ("  SFT_Cost", "One-time fine-tuning cost per model (50K samples), LoRA vs full; Vertex managed-SFT reference."),
 ("  Comparison", "Closed vs SLM monthly cost, SFT amortisation and break-even months."),
 ("  Sources", "Register of every URL used."),
 ("COLOUR CODE", "Blue text = hard-coded input/source value · Black = formula · Green = cross-sheet link · Yellow fill = key lever · Orange fill = promo / unverified value · Blue-grey fill = effective (selected) value."),
 ("KEY CAVEATS", "1) 'Qwen 2.5-25B-A3B' does not exist -> Qwen3.6-35B-A3B used (flagged). 2) No summarization/entity-extraction benchmarks exist for these models; proxies only -> run an in-house eval. 3) SLM tok/s are planning estimates (+/-50%), only 3 measured H100 points exist - replace with your own vLLM benchmark. 4) Transcript lengths / task output sizes / thinking tokens are assumptions (not vendor data). 5) Databricks $/DBU ($0.07) and AI Runtime rates are third-party/unverified; Databricks FM fine-tuning is end-of-life and AI Runtime is not on GCP. 6) Several Gemini/OpenAI prices are promotional (switch Assumptions > price basis to 'Post-promo'). 7) Many baseline models are deprecated/retiring (Claude Sonnet 4.5 2026-11-30, gpt-5 family 2026-12-11, Gemini 2.5 Flash/Pro 2026-12-01). 8) OpenAI batch support for gpt-5.2-pro/5.4-pro conflicts between OpenAI pages."),
 ("BIGQUERY BATCH DISCOUNT", "Verified: Gemini (2.0+) invoked from BigQuery (ML.GENERATE_TEXT, AI.GENERATE*, AI.IF/SCORE/CLASSIFY) is billed at Vertex AI batch rates = 50% below standard. Claude/open models in BigQuery are NOT discounted. Not confirmed: AI.GENERATE_TABLE is not named in the pricing table (likely same); verify in billing export."),
]
for i, (a, b) in enumerate(rows0):
    r = 4 + i
    c1 = ws0.cell(row=r, column=1, value=a); c1.font = f_bold if not a.startswith("  ") else f_norm; c1.alignment = wrap
    c2 = ws0.cell(row=r, column=2, value=b); c2.font = f_norm; c2.alignment = wrap
    ws0.row_dimensions[r].height = max(16, 15 * (len(b) // 135 + 1))
wb.move_sheet("README", offset=-(len(wb.sheetnames) - 1))
order = ["README", "Assumptions", "Closed_Pricing", "Closed_Std", "Closed_Batch", "Closed_BigQuery", "SLM_Models", "SLM_Benchmarks", "GPU_Pricing", "Deployment_Options",
         "SLM_Serving_Fit", "SLM_Daily_Cost", "SFT_Cost", "Comparison", "GPU_Pricing_Raw", "Sources"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    for rw in ws.iter_rows():
        for c in rw:
            if c.font and c.font.name != FONT: c.font = Font(name=FONT, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color, underline=c.font.underline)
out = os.path.join(HERE, "LLM_vs_SLM_Cost_Analysis.xlsx")
wb.save(out)
print("saved", out)
