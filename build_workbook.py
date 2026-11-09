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
add("host", "GPU hosting price option", "GCP On-demand", "list", "Choose: GCP On-demand (per-second GCE VM) | GCP Spot | GCP 1-yr CUD | GCP 3-yr CUD | GCP Cloud Run GPU (per-second; L4 and RTX PRO 6000 only) | Databricks serving (DBU). Drives every SLM cost sheet. All are billed ONLY while a run is executing.", key_lever=True)
add("dbu", "Databricks $/DBU for GPU Model Serving", 0.07, "$/DBU", "UNVERIFIED - third-party (dbuwatch.com) for GCP US regions; official Databricks tiles are JS-rendered. Confirm with your Databricks account team.", fmt='$0.000')
add("prec", "Serving weight precision used for GPU-fit check", "FP8", "list", "BF16 | FP8 | INT4. FP8 is the production default on FP8-capable GPUs (L4, RTX PRO 6000, H100). A100 (Ampere) has NO native FP8, so on A100 the fit check automatically falls back to BF16 weights (~2x memory). gpt-oss-20b uses its native MXFP4 size under 'FP8' (A100 dequantises to BF16).", key_lever=True)
add("fit", "Max share of VRAM usable for weights", 0.80, "%", "ASSUMPTION - remainder reserved for KV-cache, activations, CUDA graphs.", fmt=PCT)
add("prefill_tf", "Effective prefill compute, 1xH100 (FP8)", 400, "TFLOPs", "ASSUMPTION - H100 SXM FP8 dense peak is 1,979 TFLOPS (NVIDIA datasheet); ~20% sustained in vLLM prefill = ~400. Prefill tok/s on H100 = TFLOPs / (2 x active params). Other GPUs scaled by 'Throughput vs H100' on GPU_Pricing. Replace with `vllm bench serve` results.", fmt=NUM, key_lever=True)
add("util", "Batch-run efficiency derate", 0.75, "%", "ASSUMPTION - real runs achieve ~75% of benchmark throughput (batch tail, uneven lengths). Methodology range 0.6-0.8.", fmt=PCT)
add("cold_min", "Cold start per run (instance boot + model load)", 5, "minutes", "ASSUMPTION - billed on every on-demand run; Cloud Run GPU / GCE job loading weights from a bucket. 2-10 min typical.", fmt=NUM1, key_lever=True)
add("runs_day", "Batch runs per day (each pays a cold start)", 1, "runs/day", "1 = one nightly batch of all that day's calls. Use 24 for hourly runs.", fmt=NUM)
add("cr_cpu", "Cloud Run price per vCPU-second", 0.000018, "$/vCPU-s", "https://cloud.google.com/run/pricing - us-central1 default, instance-based billing, accessed 2026-11-08.", fmt='$0.000000')
add("cr_mem", "Cloud Run price per GiB-second", 0.000002, "$/GiB-s", "https://cloud.google.com/run/pricing - us-central1 default, accessed 2026-11-08.", fmt='$0.000000')
add("slm_think", "Thinking tokens per task call for fine-tuned SLM", 0, "tokens", "Assumption: SFT teaches direct JSON/summary output with no reasoning trace.", fmt=NUM)

sec("5. One-time SFT (fine-tuning) assumptions")
add("sft_h_low", "SFT wall-clock hours - low scenario", 12, "hours", "User guidance: SFT is roughly 12 or 24 hours of GPU time.", fmt=NUM, key_lever=True)
add("sft_h_high", "SFT wall-clock hours - high scenario", 24, "hours", "User guidance: 12 or 24 hours.", fmt=NUM, key_lever=True)
add("sft_sel", "SFT scenario used in Combined_Compare (hours)", 24, "hours", "Pick 12 or 24 (dropdown). GPUs per model are sized on SFT_Cost so that GPUs x 24h covers the estimated LoRA GPU-hours.", fmt=NUM)
add("a100_slow", "SFT slowdown on A100 80GB vs H100", 2.2, "x", "VESSL measured ~2.0x (Gemma-4-31B LoRA); NVIDIA/other sources cite up to 2.5x. Used for the A100 comparison columns on SFT_Cost and H100_vs_A100.", fmt='0.0')
add("sft_n", "SFT training samples (for GPU sizing cross-check)", 50000, "samples", "User requirement: ~50K samples, one-time.", fmt=NUM)
add("sft_tok", "Avg tokens per SFT sample", 3000, "tokens", "ASSUMPTION - matches the 3K-token call profile.", fmt=NUM)
add("sft_ep", "Training epochs (cross-check)", 3, "epochs", "Typical 2-3 for LoRA SFT.", fmt=NUM)

sec("6. Comparison baseline")
add("base", "Closed model used as baseline on Combined_Compare sheet", "Gemini 3.6 Flash", "list", "Pick any model from the Closed_Pricing sheet (dropdown). The 'cheapest closed model' is computed automatically as well.", key_lever=True)

# --- task table
sec("7. LLM tasks run per call (every call is processed by each enabled task)")
tr = row[0]
hdr(wsA, tr, ["Task", "Enabled (1/0)", "Prompt/instr. tokens", "Output tokens - Residential", "Output tokens - C&I",
              "Input tokens/call (blended)", "Output tokens/call (blended)"], height=44)
wsA.column_dimensions["E"].width = 18; wsA.column_dimensions["F"].width = 18; wsA.column_dimensions["G"].width = 18
tasks = [("ONE combined call per transcript: summary + entity extraction + disposition (single prompt) [DEFAULT]", 1, 1000, 450, 750),
         ("Optional - separate call: summarization only", 0, 350, 220, 380),
         ("Optional - separate call: entity extraction (JSON)", 0, 800, 350, 600),
         ("Optional - separate call: disposition / intent classification", 0, 300, 40, 60)]
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
put(wsA, row[0], 1, "DEFAULT = ONE call per transcript with ~3,000 input tokens (transcript ~2.0K blended + ~1K instructions/schema) and ~500 output tokens (summary + JSON). These are ASSUMPTIONS - edit blue cells. Enable the optional rows to model separate calls per task (input is re-sent for each).", font=f_note, bd=False)
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
dv = DataValidation(type="list", formula1='"GCP On-demand,GCP Spot,GCP 1-yr CUD,GCP 3-yr CUD,GCP Cloud Run GPU (per-second),Databricks serving (DBU)"'); wsA.add_data_validation(dv); dv.add(A["host"].split("!")[1].replace("$", ""))
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
      "EFFECTIVE input", "EFFECTIVE output", "EFFECTIVE cache read", "EFFECTIVE batch input", "EFFECTIVE batch output", "Comparison tier (Flash-level+ used for headline)"]
P0 = 5
hdr(wsP, P0, PH, height=56)
widths(wsP, [11, 24, 24, 34, 10, 10, 11, 10, 10, 11, 11, 11, 11, 11, 11, 12, 12, 11, 70, 60, 11, 11, 11, 11, 11, 11, 16])

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
TIER_FLASH = {"Gemini 3.5 Flash", "Gemini 3.6 Flash", "Gemini 3.7 Flash", "Gemini 3.8 Flash", "Claude Sonnet 4.5", "Claude Sonnet 4.6", "Claude Sonnet 5", "Claude Sonnet 5.5",
              "gpt-5.4-mini", "gpt-5.6-luna", "gpt-5.6-terra", "gpt-6-luna", "gpt-6.1-terra"}
TIER_LITE = {"Gemini 2.5 Flash-Lite", "Gemini 3.1 Flash-Lite", "Gemini 3.5 Flash-Lite", "Claude Haiku 4.5", "Claude Haiku 5.5", "gpt-5-nano", "gpt-5.4-nano"}
TIER_OLD = {"Gemini 2.5 Flash", "Gemini 2.5 Pro", "Gemini 3 Flash Preview", "gpt-5", "gpt-5-mini", "gpt-5-pro", "gpt-5.1", "gpt-5.2", "gpt-5.2-pro", "gpt-5.3-codex"}
for _p, _m, _r in rows_meta:
    t_ = "Flash-level" if _m in TIER_FLASH else "Lite (excluded)" if _m in TIER_LITE else "Older (excluded)" if _m in TIER_OLD else "Pro-level"
    put(wsP, _r, 27, t_, font=f_in)
put(wsP, P_last + 10, 1, "Comparison tier (editable): 'Flash-level' = same class as Gemini 3.6 Flash (Gemini 3.5-3.8 Flash, Claude Sonnet, OpenAI mini/terra/luna class); 'Pro-level' = larger/more expensive models; 'Lite' and 'Older' are EXCLUDED from the headline comparison (e.g. Gemini 2.5 Flash-Lite, gpt-5-nano, Claude Haiku).", font=f_note, bd=False)
wsP.auto_filter.ref = f"A{P0}:{L(27)}{P_last}"
dv = DataValidation(type="list", formula1=f"=Closed_Pricing!$B${P0+1}:$B${P_last}"); wsA.add_data_validation(dv); dv.add(A["base"].split("!")[1].replace("$", ""))

# =====================================================================
# 3. CLOSED COST SHEETS
# =====================================================================
NV = 15
def cost_sheet(name, title, sub, mode, tab):
    ws = sheet(name, title, sub, tab)
    heads = ["Provider", "Model", "Status", "$ per call", "$ per 1K calls"] + [None] * NV + ["Monthly @ ref volume", "Annual @ ref volume", "Rank (cheapest=1)", "Flag", "Monthly @ ref (excl. deprecated)", "Monthly @ ref (Flash-level+ only, not deprecated)"]
    hdr(ws, 5, heads, height=44)
    # volume header: row 5 cols F..T, numeric formulas
    for k in range(NV):
        c = ws.cell(row=5, column=6 + k)
        c.value = f"={A['vol_step']}" if k == 0 else f"={L(5 + k)}5+{A['vol_step']}"
        c.number_format = '#,##0" calls/day"'; c.font = f_hdr
    ws["F4"] = "Daily cost (USD) at calls/day  →"; ws["F4"].font = f_bold
    widths(ws, [16, 26, 30, 12, 12] + [11] * NV + [15, 15, 11, 38, 16, 18])
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
        put(ws, i, 10 + NV, f'=IF(AND(ISNUMBER({L(6+NV)}{i}),LEFT({L(9+NV)}{i},10)<>"Deprecated"),{L(6+NV)}{i},"n/a")', fmt=USD0)
        put(ws, i, 11 + NV, f'=IF(AND(ISNUMBER({L(10+NV)}{i}),OR(Closed_Pricing!$AA${pr}="Flash-level",Closed_Pricing!$AA${pr}="Pro-level")),{L(10+NV)}{i},"n/a")', fmt=USD0)
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
    ws.auto_filter.ref = f"A5:{L(11+NV)}{last}"
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
      "GCP On-demand", "GCP Spot", "GCP 1-yr CUD", "GCP 3-yr CUD", "Databricks DBU/h per GPU", "Databricks serving (DBU)", "SELECTED $/GPU-h (Assumptions)", "Throughput vs H100", "Databricks SKU used / notes", "Source",
      "Cloud Run GPU $/second", "Cloud Run vCPUs", "Cloud Run GiB RAM", "GCP Cloud Run GPU (per-second) $/h", "Native FP8 weights? (1/0)"]
hdr(wsG, 5, GH, height=56)
widths(wsG, [20, 9, 22, 9, 12, 12, 12, 12, 11, 11, 11, 11, 12, 13, 14, 11, 60, 60, 13, 9, 9, 14, 12])
GPUS = [
    ("L4 24GB", 24, "g2-standard-4", 1, 0.706832276, 0.423016, 0.445304335, 0.318074524, 20.0, 0.08, "Databricks has no L4 SKU: priced as A10G x1 (20 DBU/h), same 24GB class. Throughput ~0.08x H100 is a bandwidth-ratio estimate (0.3 vs 3.35 TB/s) - UNVERIFIED."),
    ("A100 80GB", 80, "a2-ultragpu-1g", 1, 5.06879789, 2.976774795, 4.199690411, 3.499997559, 628 / 8, 0.55, "Databricks A100-80GB is sold as x8 only (628 DBU/h) -> /8 per GPU. NO native FP8 (Ampere) -> serves BF16. Throughput 0.55x of an FP8 H100: measured 0.59-0.73x vs BF16 H100 (Qwen3-14B/32B, GPUStack) and H100 FP8 adds ~1.3-1.8x on top."),
    ("RTX PRO 6000 96GB", 96, "g4-standard-48", 1, 4.49993, 1.6808, 3.105, 1.97945, 80.0, 0.55, "Databricks RTX PRO 6000 x1 = 80 DBU/h. Throughput 0.55x ASSUMED from memory-bandwidth ratio (~1.8 vs 3.35 TB/s) - UNVERIFIED. g4-standard-48 includes 48 vCPUs."),
    ("H100 80GB", 80, "a3-highgpu-8g", 8, 88.490000119, 52.700460712, 61.383674231, 38.864383195, 100.0, 1.0, "Reference machine. Databricks H100 x1 = 100 DBU/h. Per-GPU = 8-GPU instance / 8. Spot can be pre-empted (risky for long SFT)."),
]
CR = {"L4 24GB": (0.0001867, 4, 16), "RTX PRO 6000 96GB": (0.00036522, 20, 80)}   # Cloud Run: L4 4vCPU/16GiB per pricing-page example; RTX min config ASSUMED
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
    put(wsG, r, 15, f'=IF({h}="GCP On-demand",I{r},IF({h}="GCP Spot",J{r},IF({h}="GCP 1-yr CUD",K{r},IF({h}="GCP 3-yr CUD",L{r},IF({h}="GCP Cloud Run GPU (per-second)",V{r},N{r})))))', fmt='$#,##0.00', fill=fill_key)
    put(wsG, r, 16, g[9], fmt='0.00', font=f_in); put(wsG, r, 17, g[10], al=wrap)
    url_cell(wsG, r, 18, "https://cloud.google.com/products/compute/pricing/accelerator-optimized ; https://www.databricks.com/product/pricing/model-serving")
    crp = CR.get(g[0])
    if crp:
        put(wsG, r, 19, crp[0], fmt='$0.00000000', font=f_in); put(wsG, r, 20, crp[1], font=f_in); put(wsG, r, 21, crp[2], font=f_in)
        put(wsG, r, 22, f"=(S{r}+T{r}*{A['cr_cpu']}+U{r}*{A['cr_mem']})*3600", fmt='$#,##0.00')
    else:
        for cc in (19, 20, 21): put(wsG, r, cc, "-")
        put(wsG, r, 22, "n/a")
    put(wsG, r, 23, 0 if g[0] == 'A100 80GB' else 1, font=f_in)
    wsG.row_dimensions[r].height = 52
G_last = G0 + len(GPUS) - 1
r = G_last + 2
for t in ["Databricks $/DBU ($0.07) is UNVERIFIED (third-party) - official tiles are JS-rendered. Databricks serving prices include the VM; GCE prices include bundled vCPU/RAM/SSD. Databricks H100 ($7.00/h) looking cheaper than GCE ($11.06/h) should be confirmed with Databricks before relying on it.",
          "Other GPUs on GCP (us-central1, per instance-hour on-demand): A100 40GB a2-highgpu-1g $3.673 | H200 a3-ultragpu-8g $84.81 (8 GPUs) | B200 a4-highgpu-8g no on-demand price (Spot $39.63; Flex-start $64.44; 1-yr CUD $88.93) - see GPU_Pricing_Raw.",
          "Databricks Foundation Model Fine-tuning is END OF LIFE; replacement 'AI Runtime' (serverless A10/H100) is listed for AWS and Azure only. On Databricks-on-GCP, fine-tune on classic GPU clusters (a2/a3/g2) = GCE VM price + DBU surcharge (DBU/h for GPU instances not retrievable - use Databricks' calculator).",
          "Cloud Run GPU (per-second billing, only while running; https://cloud.google.com/run/pricing, accessed 2026-11-08): L4 $0.0001867/s + CPU/RAM; RTX PRO 6000 $0.00036522/s + CPU/RAM. L4 config 4 vCPU/16 GiB is the pricing-page example; the RTX PRO 6000 minimum (20 vCPU/80 GiB) is ASSUMED. A100/H100 are not offered on Cloud Run (n/a).",
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
# 7. SLM SERVING FIT  (run-time model + feasibility + cost per GPU)
# =====================================================================
wsF = sheet("SLM_Serving_Fit", "SLM Batch Run - Time & Cost per GPU at Reference Volume",
            "Run time = calls x (input tokens / prefill tok/s + output tokens / batched decode tok/s) / efficiency. Cost = (run time + cold start) x $/GPU-h, billed ONLY while the on-demand instance runs. Green = link to other sheet.", "7030A0")
GN = [g[0] for g in GPUS]
FH = ["Model", "Total params (B)", "Active params (B)", "Serving weights GB (chosen precision; A100 uses BF16 - no FP8)", "Decode tok/s on H100 (batched)", "Prefill tok/s on H100 (= TFLOPs / 2 x active params)",
      "GPU-seconds per call on H100 (ideal)", "H100 minutes per 1K calls (incl. efficiency derate)"] + \
     [f"Fits {g}?" for g in GN] + [f"$/day @ ref vol: {g}" for g in GN] + \
     ["Cheapest feasible $/day", "Cheapest feasible GPU", "H100 $/day (reference machine)", "H100 / cheapest", "Always-on 24h $/day (cheapest GPU) - for comparison only", "H100 verdict"]
hdr(wsF, 5, FH, height=84)
widths(wsF, [30, 9, 9, 13, 12, 14, 12, 13] + [10] * 4 + [13] * 4 + [14, 18, 14, 10, 20, 46])
F0 = 6
CF1, CF2 = 13, 16       # cost columns M..P
for i in range(SLM_N):
    r = F0 + i; mr = M0 + 1 + i
    put(wsF, r, 1, f"=SLM_Models!C{mr}", font=f_link, al=wrap)
    put(wsF, r, 2, f"=SLM_Models!K{mr}", fmt=NUM1, font=f_link); put(wsF, r, 3, f"=SLM_Models!M{mr}", fmt=NUM1, font=f_link)
    put(wsF, r, 4, f'=IF({A["prec"]}="BF16",SLM_Models!P{mr},IF({A["prec"]}="FP8",SLM_Models!Q{mr},SLM_Models!R{mr}))', fmt=NUM1, font=f_link)
    put(wsF, r, 5, f"=SLM_Models!U{mr}", fmt=NUM, font=f_link)
    put(wsF, r, 6, f"={A['prefill_tf']}*1000000000000/(2*C{r}*1000000000)", fmt=NUM)
    put(wsF, r, 7, f"={A['tot_in']}/F{r}+{A['out_slm']}/E{r}", fmt='0.000')
    put(wsF, r, 8, f"=G{r}*1000/{A['util']}/60", fmt='0.0', fill=fill_sub)
    for k, g in enumerate(GN):
        gr = GR[g]
        wexp = f'IF(AND({A["prec"]}="FP8",GPU_Pricing!$W${gr}=0),SLM_Models!$P{mr},$D{r})'
        put(wsF, r, 9 + k, f"=IF({wexp}<=GPU_Pricing!$B${gr}*{A['fit']},1,0)", fmt='0')
        put(wsF, r, CF1 + k, f'=IF(AND({L(9+k)}{r}=1,ISNUMBER(GPU_Pricing!$O${gr})),({A["ref_vol"]}*$G{r}/GPU_Pricing!$P${gr}/{A["util"]}/3600+{A["runs_day"]}*{A["cold_min"]}/60)*GPU_Pricing!$O${gr},"n/a")', fmt=USD)
    put(wsF, r, 17, f'=IF(COUNT(M{r}:P{r})=0,"n/a",MIN(M{r}:P{r}))', fmt=USD, fill=fill_sub, font=f_bold)
    put(wsF, r, 18, f'=IF(ISNUMBER(Q{r}),SUBSTITUTE(INDEX($M$5:$P$5,MATCH(Q{r},M{r}:P{r},0)),"$/day @ ref vol: ",""),"No single-GPU fit / price n/a")', al=wrap)
    put(wsF, r, 19, f'=IF(L{r}=1,P{r},"n/a")', fmt=USD)
    put(wsF, r, 20, f'=IF(AND(ISNUMBER(S{r}),ISNUMBER(Q{r})),S{r}/Q{r},"")', fmt='0.0"x"')
    put(wsF, r, 21, f'=IF(ISNUMBER(Q{r}),24*INDEX(GPU_Pricing!$O${G0}:$O${G_last},MATCH(R{r},GPU_Pricing!$A${G0}:$A${G_last},0))*MAX(1,ROUNDUP({A["ref_vol"]}*G{r}/INDEX(GPU_Pricing!$P${G0}:$P${G_last},MATCH(R{r},GPU_Pricing!$A${G0}:$A${G_last},0))/{A["util"]}/3600/24,0)),"n/a")', fmt=USD)
    put(wsF, r, 22, f'=IF(L{r}=0,"H100 does not fit",IF(SUM(I{r}:K{r})=0,"H100 needed (no other GPU fits)",IF(R{r}="H100 80GB","H100 is the cheapest feasible","H100 optional: "&TEXT(S{r}/Q{r}-1,"0%")&" more than "&R{r}&"; ~"&TEXT(1/(INDEX(GPU_Pricing!$P${G0}:$P${G_last},MATCH(R{r},GPU_Pricing!$A${G0}:$A${G_last},0))),"0.0")&"x faster")))', al=wrap)
F_last = F0 + SLM_N - 1
put(wsF, F_last + 2, 1, "Method: prefill (reading the 3K-token transcript) is compute-bound and fast; decode (writing ~500 tokens) is memory-bound and batched across many calls at once. GPU time for N calls = N x (input / prefill tok/s + output / decode tok/s) / derate. The instance is started for the run, billed per second, then shut down (Cloud Run GPU / GCE job / Databricks job) - so cost scales with run time, not with 24h x 30 days. Fit check = weights <= usable VRAM share (KV cache not separately sized). 'Always-on 24h' is shown only to illustrate why on-demand matters.", font=f_note, bd=False)
wsF.freeze_panes = "B6"
wsF.conditional_formatting.add(f"I{F0}:L{F_last}", ColorScaleRule(start_type="num", start_value=0, start_color="F8CBAD", end_type="num", end_value=1, end_color="C6EFCE"))

# =====================================================================
# 8. SLM DAILY COST (1K..15K) - two blocks
# =====================================================================
wsD = sheet("SLM_Daily_Cost", "Self-hosted SLM - Daily Run Cost for 1K-15K calls/day (on-demand GPU, billed only while running; excludes one-time SFT)",
            "Block A: every model on the SAME reference machine (1xH100) - like-for-like. Block B: cheapest GPU each model fits on (lower-GPU option). Daily cost = calls x variable $/call + cold-start $ per run x runs/day.", "7030A0")
FC = ["Model", "GPU", "$/GPU-h (selected hosting)", "Run minutes per 1K calls (excl. cold start)", "Variable $ per call", "Cold-start $ per day"]
nc = len(FC)
def vol_header(ws, row):
    hdr(ws, row, FC + [None] * NV + ["Monthly @ ref volume", "Annual @ ref volume"], height=56)
    for k in range(NV):
        c = ws.cell(row=row, column=nc + 1 + k)
        c.value = f"={A['vol_step']}" if k == 0 else f"={L(nc+k)}{row}+{A['vol_step']}"
        c.number_format = '#,##0" calls/day"'; c.font = f_hdr
widths(wsD, [30, 18, 12, 13, 12, 12] + [10] * NV + [14, 14])
wsD["A4"] = "BLOCK A - Reference machine: 1 x H100 80GB (like-for-like TPS)"; wsD["A4"].font = f_bold
vol_header(wsD, 5)
D0 = 6
def daily_block(start, label):
    for i in range(SLM_N):
        r = start + i; fr = F0 + i
        put(wsD, r, 1, f"=SLM_Serving_Fit!A{fr}", font=f_link, al=wrap)
        if label == "A":
            put(wsD, r, 2, "H100 80GB")
        else:
            put(wsD, r, 2, f'=IF(ISNUMBER(SLM_Serving_Fit!Q{fr}),SLM_Serving_Fit!R{fr},"n/a")', font=f_link)
        gi = f"MATCH($B{r},GPU_Pricing!$A${G0}:$A${G_last},0)"
        put(wsD, r, 3, f'=IFERROR(INDEX(GPU_Pricing!$O${G0}:$O${G_last},{gi})*1,"n/a")', fmt=USD, font=f_link)
        fitcond = f"SLM_Serving_Fit!L{fr}=1," if label == "A" else ""
        base = f'SLM_Serving_Fit!G{fr}/INDEX(GPU_Pricing!$P${G0}:$P${G_last},{gi})/{A["util"]}'
        put(wsD, r, 4, f'=IFERROR({base}*1000/60,"n/a")' if label == "B" else f'=IFERROR(IF(SLM_Serving_Fit!L{fr}=1,{base}*1000/60,"n/a"),"n/a")', fmt='0.0')
        put(wsD, r, 5, f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(C{r})),D{r}/1000*60/3600*C{r},"n/a")', fmt='$0.00000')
        put(wsD, r, 6, f'=IF(ISNUMBER(C{r}),{A["runs_day"]}*{A["cold_min"]}/60*C{r},"n/a")', fmt=USD)
        for k in range(NV):
            put(wsD, r, nc + 1 + k, f'=IF(ISNUMBER($E{r}),{L(nc+1+k)}${start-1}*$E{r}+$F{r},"n/a")', fmt=USD)
        put(wsD, r, nc + NV + 1, f'=IF(ISNUMBER($E{r}),({A["ref_vol"]}*$E{r}+$F{r})*{A["days"]},"n/a")', fmt=USD0)
        put(wsD, r, nc + NV + 2, f'=IF(ISNUMBER($E{r}),({A["ref_vol"]}*$E{r}+$F{r})*365,"n/a")', fmt=USD0)
    return start + SLM_N - 1
dA_last = daily_block(D0, "A")
dB_hdr = dA_last + 3
wsD.cell(row=dB_hdr - 1, column=1, value="BLOCK B - Cheapest feasible single GPU per model (lower-GPU option, e.g. L4 for small models)").font = f_bold
vol_header(wsD, dB_hdr)
dB0 = dB_hdr + 1
dB_last = daily_block(dB0, "B")
wsD.freeze_panes = "B6"
for a_, b_ in ((D0, dA_last), (dB0, dB_last)):
    wsD.conditional_formatting.add(f"{L(nc+1)}{a_}:{L(nc+NV)}{b_}", ColorScaleRule(start_type="min", start_color="C6EFCE", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="max", end_color="F8CBAD"))

# =====================================================================
# 9. SFT COST (one-time)
# =====================================================================
wsT = sheet("SFT_Cost", "One-time Supervised Fine-Tuning (SFT) Cost per Model - 12h / 24h GPU scenarios",
            "Cost = H100 GPUs x wall-clock hours x $/GPU-h (GCP H100, per-GPU = a3-highgpu-8g / 8). GPUs per model are sized so GPUs x 24h covers the estimated LoRA GPU-hours (cross-check from token throughput; +/-50%, validate with a 100-step pilot).", "C00000")
TH = ["Model", "Total params (B)", "LoRA tok/s per H100 (est.)", "Training tokens (M) = samples x tokens x epochs", "Est. LoRA GPU-hours needed (cross-check)", "H100 GPUs used (sized for high scenario)",
      "SFT cost - LOW scenario (12h) $ On-demand", "SFT cost - HIGH scenario (24h) $ On-demand", "SFT cost - SELECTED hours $ On-demand", "SELECTED hours $ on Spot (pre-emptible)", "SELECTED hours $ on 3-yr CUD",
      "Share of est. GPU-hours covered by LOW scenario", "Basis", "Same job on A100 80GB: GPUs for 24h window", "Same job on A100 80GB: cost (selected hours basis) $"]
hdr(wsT, 5, TH, height=72)
widths(wsT, [30, 9, 13, 14, 14, 12, 15, 15, 15, 15, 15, 14, 60, 16, 18])
T0 = 6
hr = GR["H100 80GB"]
tokens_M = f"{A['sft_n']}*{A['sft_tok']}*{A['sft_ep']}/1000000"
for i in range(SLM_N):
    r = T0 + i; mr = M0 + 1 + i
    put(wsT, r, 1, f"=SLM_Models!C{mr}", font=f_link, al=wrap); put(wsT, r, 2, f"=SLM_Models!K{mr}", fmt=NUM1, font=f_link)
    put(wsT, r, 3, NUMS[i][8], fmt=NUM, font=f_in)
    put(wsT, r, 4, f"={tokens_M}", fmt=NUM1)
    put(wsT, r, 5, f"=D{r}*1000000/C{r}/3600", fmt=NUM1)
    put(wsT, r, 6, f"=IF(E{r}<={A['sft_h_high']},1,IF(E{r}<=2*{A['sft_h_high']},2,IF(E{r}<=4*{A['sft_h_high']},4,8)))", fmt='0', fill=fill_sub)
    put(wsT, r, 7, f"=F{r}*{A['sft_h_low']}*GPU_Pricing!$I${hr}", fmt=USD0)
    put(wsT, r, 8, f"=F{r}*{A['sft_h_high']}*GPU_Pricing!$I${hr}", fmt=USD0)
    put(wsT, r, 9, f"=F{r}*{A['sft_sel']}*GPU_Pricing!$I${hr}", fmt=USD0, font=f_bold, fill=fill_sub)
    put(wsT, r, 10, f"=F{r}*{A['sft_sel']}*GPU_Pricing!$J${hr}", fmt=USD0)
    put(wsT, r, 11, f"=F{r}*{A['sft_sel']}*GPU_Pricing!$L${hr}", fmt=USD0)
    put(wsT, r, 12, f"=MIN(1,F{r}*{A['sft_h_low']}/E{r})", fmt='0%')
    put(wsT, r, 13, "LoRA tok/s per H100: NeMo AutoModel (8B LoRA 12.5k best case, derated to 8k); VESSL Gemma-4-31B LoRA ~1.2k/GPU; MoE ~3.5-4.5k derated; others interpolated by size.", al=wrap, font=f_note)
    put(wsT, r, 14, f"=ROUNDUP(E{r}*{A['a100_slow']}/{A['sft_h_high']},0)", fmt='0')
    put(wsT, r, 15, f"=E{r}*{A['a100_slow']}*GPU_Pricing!$I${GR['A100 80GB']}", fmt=USD0)
    wsT.row_dimensions[r].height = 40
T_last = T0 + SLM_N - 1
r = T_last + 3
wsT.cell(row=r - 1, column=1, value="Reference: managed per-token SFT on Vertex AI (open models; no Gemma 4 / Granite / Nemotron / gpt-oss / 30B-A3B entries) - priced on the same training tokens").font = f_bold
hdr(wsT, r, ["Vertex managed SFT model", "$ per 1M training tokens", "Training tokens (M)", "Cost 1 run", "Source"], height=30)
VSFT = [("Llama 3.1 8B", 0.67), ("Gemma 3 12B", 1.82), ("Qwen3 8B", 4.18), ("Qwen3 32B", 6.57), ("Gemma 3 27B", 6.83)]
for j, (n_, p_) in enumerate(VSFT):
    rr = r + 1 + j
    put(wsT, rr, 1, n_); put(wsT, rr, 2, p_, fmt=USD, font=f_in); put(wsT, rr, 3, f"={tokens_M}", fmt=NUM1); put(wsT, rr, 4, f"=B{rr}*C{rr}", fmt=USD0)
    url_cell(wsT, rr, 5, "https://cloud.google.com/vertex-ai/generative-ai/pricing (supervised tuning, open models) - accessed 2026-11-08")
rr = r + len(VSFT) + 2
for t in ["12h/24h scenarios are the user's guidance. The 12h column assumes the SAME GPU count as 24h (i.e. roughly half the epochs/data) - check 'share covered'. SFT is one-time; budget a re-run if the first eval misses target.",
          "GCP sells H100 as a3-highgpu-1g/2g/4g/8g shapes; per-GPU price here = 8-GPU instance / 8 (smaller shapes may carry a small premium). Spot can be pre-empted mid-run - checkpoint every ~15 min.",
          "Databricks Foundation Model Fine-tuning is END OF LIFE (last example Llama 3.1 8B ~ $2,860 / 500M words); AI Runtime is AWS/Azure only. On Databricks-on-GCP use classic GPU clusters (GCE VM + DBU surcharge).",
          "Data curation/labelling, evaluation and engineering time are NOT included."]:
    wsT.cell(row=rr, column=1, value="• " + t).font = f_note; rr += 1
wsT.freeze_panes = "B6"

# =====================================================================
# 9b. H100 vs A100
# =====================================================================
wsH = sheet("H100_vs_A100", "H100 vs A100 80GB - What You Get, and When H100 Is Worth It",
            "Prices link to GPU_Pricing; spec rows are NVIDIA datasheet values quoted from memory (NOT re-fetched this session - verify); throughput/SFT ratios are measured (GPUStack, VESSL) or planning assumptions as labelled.", "ED7D31")
widths(wsH, [46, 20, 20, 18, 80])
hdr(wsH, 4, ["Property", "H100 80GB (SXM)", "A100 80GB", "H100 / A100", "Why it matters for these workloads"], height=30)
spec = [("GPU memory", "80 GB HBM3", "80 GB HBM2e", "1.0x", "Same capacity: model-size limits are identical IF the weights are the same precision. A100 cannot use FP8, so big models take ~2x the memory there."),
        ("Memory bandwidth", "3.35 TB/s", "~2.0 TB/s", "~1.6x", "Decode (writing the ~500 output tokens) is memory-bandwidth bound -> H100 generates tokens ~1.6x faster per stream."),
        ("BF16 dense compute", "~989 TFLOPS", "~312 TFLOPS", "~3.2x", "Prefill (reading the 3K-token transcript) and SFT training are compute-bound -> big H100 advantage."),
        ("FP8 support (Transformer Engine)", "Yes: ~1,979 TFLOPS dense", "No (Ampere)", "n/a", "FP8 halves weight memory and adds ~1.3-1.8x speed; lets 27-35B models fit comfortably with KV-cache room. A100 must run BF16 (or INT8/INT4 weight-only)."),
        ("NVLink bandwidth", "900 GB/s", "600 GB/s", "1.5x", "Matters only for multi-GPU tensor-parallel serving or multi-GPU SFT (8-GPU nodes)."),
        ("GCP shape / availability", "a3-highgpu-8g (8 GPUs)", "a2-ultragpu-1g (1 GPU)", "-", "A100 80GB comes in 1-GPU shapes (finer billing granularity); H100 is sold in the table used as an 8-GPU instance (per-GPU price = /8)."),
        ("Not on Cloud Run", "No", "No", "-", "Cloud Run GPU offers only L4 and RTX PRO 6000; H100/A100 need GCE/GKE/Vertex/Databricks.")]
for i, row_ in enumerate(spec):
    r = 5 + i
    for j, v in enumerate(row_): put(wsH, r, 1 + j, v, al=wrap, font=f_bold if j == 0 else None)
    wsH.row_dimensions[r].height = 44
r = 5 + len(spec) + 1
wsH.cell(row=r, column=1, value="Measured / assumed performance ratios").font = f_bold; r += 1
hdr(wsH, r, ["Metric", "H100", "A100 80GB", "H100 / A100", "Source / basis"], height=26); r += 1
meas = [("vLLM output tok/s, Qwen3-14B (BF16)", 3174, 1887, "GPUStack Inference Performance Lab, vLLM 0.11.0, ShareGPT"),
        ("vLLM output tok/s, Qwen3-32B (BF16)", 1132, 832, "same"),
        ("LoRA SFT speed, Gemma-4-31B (relative)", 1, f"=1/{A['a100_slow']}", "VESSL measured ~2.0x faster on H100; Assumptions 'A100 SFT slowdown' (2.2x) used in cost columns")]
m0 = r
for i, (a_, h_, a2, src) in enumerate(meas):
    put(wsH, r, 1, a_); put(wsH, r, 2, h_, fmt=NUM if i < 2 else '0.00', font=f_in); put(wsH, r, 3, a2, fmt=NUM if i < 2 else '0.00', font=f_in if i < 2 else f_norm)
    put(wsH, r, 4, f"=B{r}/C{r}", fmt='0.00"x"'); put(wsH, r, 5, src, al=wrap, font=f_note); r += 1
r += 1
wsH.cell(row=r, column=1, value="Cost per unit of work by hosting option (A100 throughput = 'Throughput vs H100' on GPU_Pricing; SFT slowdown from Assumptions)").font = f_bold; r += 1
hdr(wsH, r, ["Hosting option", "H100 $/GPU-h", "A100 80GB $/GPU-h", "Inference cost per token: A100 vs H100", "Reading"], height=40)
wsH.cell(row=r, column=6, value="SFT job cost: A100 vs H100").font = f_hdr; wsH.cell(row=r, column=6).fill = fill_hdr; wsH.cell(row=r, column=6).alignment = center
wsH.cell(row=r, column=7, value="SFT wall-clock: A100 vs H100").font = f_hdr; wsH.cell(row=r, column=7).fill = fill_hdr; wsH.cell(row=r, column=7).alignment = center
wsH.column_dimensions["F"].width = 18; wsH.column_dimensions["G"].width = 18
r += 1
hH, hA = GR["H100 80GB"], GR["A100 80GB"]
tm = f"GPU_Pricing!$P${hA}/GPU_Pricing!$P${hH}"
for lab, col in (("GCP On-demand", "I"), ("GCP Spot", "J"), ("GCP 1-yr CUD", "K"), ("GCP 3-yr CUD", "L"), ("Databricks serving (DBU)", "N")):
    put(wsH, r, 1, lab)
    put(wsH, r, 2, f"=GPU_Pricing!{col}{hH}", fmt=USD, font=f_link); put(wsH, r, 3, f"=GPU_Pricing!{col}{hA}", fmt=USD, font=f_link)
    put(wsH, r, 4, f"=(C{r}/GPU_Pricing!$P${hA})/(B{r}/GPU_Pricing!$P${hH})", fmt='0.00"x"', fill=fill_sub, font=f_bold)
    put(wsH, r, 5, f'=IF(D{r}<1,"A100 is "&TEXT(1-D{r},"0%")&" cheaper per token (slower, but cheaper)","H100 is "&TEXT(1-1/D{r},"0%")&" cheaper per token AND faster")', al=wrap)
    put(wsH, r, 6, f"=({A['a100_slow']}*C{r})/B{r}", fmt='0.00"x"', fill=fill_sub, font=f_bold)
    put(wsH, r, 7, f"={A['a100_slow']}", fmt='0.0"x slower"')
    r += 1
r += 1
wsH.cell(row=r, column=1, value="Takeaways").font = f_bold; r += 1
tk = ["INFERENCE: at GCP list prices H100 costs ~2.2x per GPU-hour but is only ~1.5-1.8x faster, so A100 is slightly cheaper per token for batch jobs - IF the model fits in BF16 and you can wait. On Databricks serving (H100 100 DBU/h vs A100 78.5 DBU/GPU-h) H100 is cheaper per token. Ratios above recompute if you change prices/throughput.",
      "H100 IS NEEDED / PREFERRED WHEN: (1) the model is ~30B+ dense and BF16 weights leave no KV-cache room on 80GB (e.g. Gemma 4 31B at BF16 = 69.9 GB; Qwen3.6-35B = ~70 GB) - FP8 on H100 (35 GB) fits easily; (2) you want the nightly batch to finish fast (15K calls on a 30B dense model take ~4.5 h on H100 vs ~8 h on A100 - see run minutes on SLM_GPU_Matrix; small/MoE models take well under an hour either way); (3) SFT with a fixed 12h/24h window - H100 needs ~half the GPUs; (4) you serve with FP8 for quality-neutral 1.3-1.8x speed-up.",
      "A100 IS FINE WHEN: the model is <=14B or a ~3B-active MoE that fits BF16 comfortably, the batch window is flexible, and cost per token matters more than finish time. Even cheaper: L4 (<=12B at FP8) or RTX PRO 6000 96GB (native FP8, per-second billing on Cloud Run).",
      "SFT: A100 and H100 cost about the same per training job (A100 is ~2.2x slower but ~2.2x cheaper per hour) - H100 simply finishes in less than half the time, which is why the 12h/24h SFT scenarios on SFT_Cost use H100.",
      "Practical note: A100 has no FP8, so the fit check on SLM_Serving_Fit uses BF16 weights for A100 and FP8 weights for L4/RTX PRO 6000/H100 (Assumptions > serving precision)."]
for t in tk:
    c = wsH.cell(row=r, column=1, value="• " + t); c.font = f_norm; r += 1

# =====================================================================
# 9c. SLM x GPU MATRIX (precision, tok/s, run time, cost for every GPU incl. H100)
# =====================================================================
wsMX = sheet("SLM_GPU_Matrix", "Every SLM on Every GPU - Precision Needed, Throughput (tok/s), Run Time and Cost (L4 / A100 / RTX PRO 6000 / H100)",
             "One row per model x GPU. Precision: FP8 on FP8-capable GPUs (L4, RTX PRO 6000, H100); A100 has no FP8 -> BF16 (2x memory). tok/s = H100 planning figure x 'Throughput vs H100' of the GPU. Use the filter to isolate a model or a GPU. Costs: on-demand, billed only while running, hosting option from Assumptions.", "7030A0")
MXH = ["Model", "GPU", "Weight precision needed", "Weights GB at that precision", "GPU VRAM GB", "Fits (<= usable VRAM)?", "Decode tok/s (batched output)", "Prefill tok/s (input)",
       "Run minutes per 1K calls (incl. derate)", "$/GPU-h (selected hosting)", "Marginal $ per 1K calls", "Cold-start $ per run", "Cost/day @ 1K calls", "Cost/day @ ref volume", "Monthly @ ref volume", "Cheapest GPU for this model?", "Throughput basis"]
hdr(wsMX, 5, MXH, height=60)
widths(wsMX, [34, 18, 12, 12, 9, 11, 13, 13, 13, 11, 12, 11, 12, 12, 13, 12, 40])
rr_ = 6
for i in range(SLM_N):
    mr = M0 + 1 + i; fr = F0 + i
    r_first = rr_
    for k, g in enumerate(GN):
        gr = GR[g]
        put(wsMX, rr_, 1, f"=SLM_Models!C{mr}", font=f_link, al=Alignment(wrap_text=False)); put(wsMX, rr_, 2, g, font=f_bold)
        put(wsMX, rr_, 3, f'=IF({A["prec"]}="FP8",IF(GPU_Pricing!$W${gr}=1,"FP8","BF16"),{A["prec"]})')
        put(wsMX, rr_, 4, f'=IF(C{rr_}="BF16",SLM_Models!$P{mr},IF(C{rr_}="FP8",SLM_Models!$Q{mr},SLM_Models!$R{mr}))', fmt=NUM1)
        put(wsMX, rr_, 5, f"=GPU_Pricing!$B${gr}", fmt=NUM, font=f_link)
        put(wsMX, rr_, 6, f'=IF(D{rr_}<=E{rr_}*{A["fit"]},"Yes","No")')
        put(wsMX, rr_, 7, f"=SLM_Models!$U{mr}*GPU_Pricing!$P${gr}", fmt=NUM)
        put(wsMX, rr_, 8, f"=SLM_Serving_Fit!$F{fr}*GPU_Pricing!$P${gr}", fmt=NUM)
        put(wsMX, rr_, 9, f"=1000*({A['tot_in']}/H{rr_}+{A['out_slm']}/G{rr_})/{A['util']}/60", fmt='0.0', fill=fill_sub)
        put(wsMX, rr_, 10, f"=GPU_Pricing!$O${gr}", fmt=USD, font=f_link)
        put(wsMX, rr_, 11, f'=IF(AND(F{rr_}="Yes",ISNUMBER(J{rr_})),I{rr_}/60*J{rr_},"n/a")', fmt=USD)
        put(wsMX, rr_, 12, f'=IF(AND(F{rr_}="Yes",ISNUMBER(J{rr_})),{A["runs_day"]}*{A["cold_min"]}/60*J{rr_},"n/a")', fmt=USD)
        put(wsMX, rr_, 13, f'=IF(ISNUMBER(K{rr_}),K{rr_}+L{rr_},"n/a")', fmt=USD)
        put(wsMX, rr_, 14, f'=IF(ISNUMBER(K{rr_}),K{rr_}*{A["ref_vol"]}/1000+L{rr_},"n/a")', fmt=USD, font=f_bold)
        put(wsMX, rr_, 15, f'=IF(ISNUMBER(N{rr_}),N{rr_}*{A["days"]},"n/a")', fmt=USD0)
        rr_ += 1
    for r2_ in range(r_first, rr_):
        put(wsMX, r2_, 16, f'=IF(ISNUMBER(N{r2_}),IF(N{r2_}=MIN(N${r_first}:N${rr_-1}),"CHEAPEST",""),"")', font=f_bold)
        put(wsMX, r2_, 17, "H100 figure: planning estimate (SLM_Models col U). GPU factor: GPU_Pricing 'Throughput vs H100'.", font=f_note, al=Alignment(wrap_text=False))
mx_last = rr_ - 1
wsMX.freeze_panes = "C6"; wsMX.auto_filter.ref = f"A5:Q{mx_last}"
wsMX.conditional_formatting.add(f"F6:F{mx_last}", ColorScaleRule(start_type="num", start_value=0, start_color="F8CBAD", end_type="num", end_value=1, end_color="C6EFCE")) if False else None
from openpyxl.formatting.rule import CellIsRule
wsMX.conditional_formatting.add(f"F6:F{mx_last}", CellIsRule(operator="equal", formula=['"No"'], fill=PatternFill("solid", bgColor="F8CBAD")))
wsMX.conditional_formatting.add(f"P6:P{mx_last}", CellIsRule(operator="equal", formula=['"CHEAPEST"'], fill=PatternFill("solid", bgColor="C6EFCE")))

# =====================================================================
# 10. COMBINED COMPARE
# =====================================================================
from openpyxl.chart import LineChart, Reference
wsC = sheet("Combined_Compare", "Closed LLMs vs Fine-tuned SLMs - Combined Cost Comparison (1K-15K calls/day)",
            "One table for every option: closed Standard API, closed Batch API, Gemini via BigQuery batch, and each SLM on an on-demand GPU. Filter/sort with the header drop-downs. SLM rows include cold start; Year-1 adds the one-time SFT.", "00B050")
vc = 6      # first volume column index (F)
def cc(col): return L(col)
# --- headline block
wsC["A4"] = "HEADLINE - cheapest option per category. Closed = Gemini 3.6 Flash-level and above only (Lite/older tiers such as Gemini 2.5 Flash-Lite, gpt-5-nano, Haiku, and deprecated models are excluded)"; wsC["A4"].font = f_bold
HH = ["Category", "Model", "$ per 1K calls (marginal)", None, None] + [None] * NV + ["Monthly @ ref volume", "Annual @ ref volume", "One-time SFT $", "Year-1 total $"]
hdr(wsC, 5, ["Category", "Model", "Marginal $ per 1K calls"] + [None] * 2 + [None] * NV + ["Monthly @ ref volume", "Annual @ ref volume", "One-time SFT $", "Year-1 total $"], height=48)
for k in range(NV):
    c = wsC.cell(row=5, column=vc + k)
    c.value = f"={A['vol_step']}" if k == 0 else f"={L(vc+k-1)}5+{A['vol_step']}"
    c.number_format = '#,##0" calls/day"'; c.font = f_hdr
wsC.cell(row=5, column=4, value="Run/billing mode").font = f_hdr
wsC.cell(row=5, column=5, value="-").font = f_hdr
widths(wsC, [30, 34, 13, 26, 12] + [10] * NV + [14, 14, 13, 14, 10, 15, 16])
vS = f"Closed_Std!${L(11+NV)}$6:${L(11+NV)}${S_last}"; vB = f"Closed_Batch!${L(11+NV)}$6:${L(11+NV)}${B_last}"; vQ = f"Closed_BigQuery!${L(11+NV)}$6:${L(11+NV)}${Q_last}"
def head_closed(r, label, sheetname, last, helper, mode):
    rngB = f"{sheetname}!$B$6:$B${last}"
    m = f"MATCH(MIN({helper}),{helper},0)"
    put(wsC, r, 1, label, font=f_bold)
    put(wsC, r, 2, f"=INDEX({rngB},{m})", font=f_link)
    put(wsC, r, 3, f"=INDEX({sheetname}!$E$6:$E${last},{m})", fmt=USD)
    put(wsC, r, 4, mode); put(wsC, r, 5, "")
    for k in range(NV):
        put(wsC, r, vc + k, f"=INDEX({sheetname}!${L(6+k)}$6:${L(6+k)}${last},{m})", fmt=USD0)
    put(wsC, r, vc + NV, f"=MIN({helper})", fmt=USD0, font=f_bold)
    put(wsC, r, vc + NV + 1, f"={L(vc+NV)}{r}*365/{A['days']}", fmt=USD0)
    put(wsC, r, vc + NV + 2, 0, fmt=USD0); put(wsC, r, vc + NV + 3, f"={L(vc+NV+1)}{r}+{L(vc+NV+2)}{r}", fmt=USD0)
head_closed(6, "Closed (Flash-level+) - cheapest Standard API", "Closed_Std", S_last, vS, "Real-time API, pay per token")
head_closed(7, "Closed (Flash-level+) - cheapest Batch API", "Closed_Batch", B_last, vB, "Provider Batch API (-50%)")
head_closed(8, "Closed (Flash-level+) - cheapest Gemini via BigQuery", "Closed_BigQuery", Q_last, vQ, "BigQuery AI.GENERATE (batch-rated -50%)")
# cheapest SLM (Block B monthly)
sm = f"SLM_Daily_Cost!${L(nc+NV+1)}${dB0}:${L(nc+NV+1)}${dB_last}"
mS = f"MATCH(MIN({sm}),{sm},0)"
put(wsC, 9, 1, "SLM - cheapest on-demand GPU", font=f_bold)
put(wsC, 9, 2, f"=INDEX(SLM_Daily_Cost!$A${dB0}:$A${dB_last},{mS})", font=f_link)
put(wsC, 9, 3, f"=INDEX(SLM_Daily_Cost!$E${dB0}:$E${dB_last},{mS})*1000", fmt=USD)
put(wsC, 9, 4, f'="On-demand "&INDEX(SLM_Daily_Cost!$B${dB0}:$B${dB_last},{mS})&" (billed while running)"'); put(wsC, 9, 5, "")
for k in range(NV):
    put(wsC, 9, vc + k, f"=INDEX(SLM_Daily_Cost!${L(nc+1+k)}${dB0}:${L(nc+1+k)}${dB_last},{mS})", fmt=USD0)
put(wsC, 9, vc + NV, f"=MIN({sm})", fmt=USD0, font=f_bold)
put(wsC, 9, vc + NV + 1, f"={L(vc+NV)}9*365/{A['days']}", fmt=USD0)
put(wsC, 9, vc + NV + 2, f"=INDEX(SFT_Cost!$I${T0}:$I${T_last},{mS})", fmt=USD0); put(wsC, 9, vc + NV + 3, f"={L(vc+NV+1)}9+{L(vc+NV+2)}9", fmt=USD0)
# selected baseline
mb = f"MATCH({A['base']},Closed_Std!$B$6:$B${S_last},0)"
put(wsC, 10, 1, "Closed - selected baseline (Assumptions)", font=f_bold); put(wsC, 10, 2, f"={A['base']}", font=f_link)
put(wsC, 10, 3, f"=INDEX(Closed_Std!$E$6:$E${S_last},{mb})", fmt=USD); put(wsC, 10, 4, "Real-time API, pay per token"); put(wsC, 10, 5, "")
for k in range(NV): put(wsC, 10, vc + k, f"=INDEX(Closed_Std!${L(6+k)}$6:${L(6+k)}${S_last},{mb})", fmt=USD0)
put(wsC, 10, vc + NV, f"=INDEX(Closed_Std!${L(6+NV)}$6:${L(6+NV)}${S_last},{mb})", fmt=USD0, font=f_bold)
put(wsC, 10, vc + NV + 1, f"={L(vc+NV)}10*365/{A['days']}", fmt=USD0); put(wsC, 10, vc + NV + 2, 0, fmt=USD0); put(wsC, 10, vc + NV + 3, f"={L(vc+NV+1)}10+{L(vc+NV+2)}10", fmt=USD0)
BASE_MONTH = f"${L(vc+NV)}$10"
# cheapest SLM on H100 (Block A)
smA = f"SLM_Daily_Cost!${L(nc+NV+1)}${D0}:${L(nc+NV+1)}${dA_last}"
mA = f"MATCH(MIN({smA}),{smA},0)"
put(wsC, 14, 1, "SLM - cheapest on 1xH100 (reference machine)", font=f_bold)
put(wsC, 14, 2, f"=INDEX(SLM_Daily_Cost!$A${D0}:$A${dA_last},{mA})", font=f_link)
put(wsC, 14, 3, f"=INDEX(SLM_Daily_Cost!$E${D0}:$E${dA_last},{mA})*1000", fmt=USD)
put(wsC, 14, 4, "On-demand H100 (billed while running)"); put(wsC, 14, 5, "")
for k in range(NV): put(wsC, 14, vc + k, f"=INDEX(SLM_Daily_Cost!${L(nc+1+k)}${D0}:${L(nc+1+k)}${dA_last},{mA})", fmt=USD0)
put(wsC, 14, vc + NV, f"=MIN({smA})", fmt=USD0, font=f_bold)
put(wsC, 14, vc + NV + 1, f"={L(vc+NV)}14*365/{A['days']}", fmt=USD0)
put(wsC, 14, vc + NV + 2, f"=INDEX(SFT_Cost!$I${T0}:$I${T_last},{mA})", fmt=USD0); put(wsC, 14, vc + NV + 3, f"={L(vc+NV+1)}14+{L(vc+NV+2)}14", fmt=USD0)
# per-provider cheapest (Standard, Flash-level+)
def head_provider(r, label, pfx):
    provA = f"Closed_Std!$A$6:$A${S_last}"
    mn = f"_xlfn.MINIFS({vS},{provA},\"{pfx}*\")"
    m = f"MATCH(1,INDEX((LEFT({provA},{len(pfx)})=\"{pfx}\")*({vS}={mn}),0),0)"
    put(wsC, r, 1, label, font=f_bold)
    put(wsC, r, 2, f"=INDEX(Closed_Std!$B$6:$B${S_last},{m})", font=f_link)
    put(wsC, r, 3, f"=INDEX(Closed_Std!$E$6:$E${S_last},{m})", fmt=USD)
    put(wsC, r, 4, "Real-time API, pay per token"); put(wsC, r, 5, "")
    for k in range(NV): put(wsC, r, vc + k, f"=INDEX(Closed_Std!${L(6+k)}$6:${L(6+k)}${S_last},{m})", fmt=USD0)
    put(wsC, r, vc + NV, f"={mn}", fmt=USD0, font=f_bold)
    put(wsC, r, vc + NV + 1, f"={L(vc+NV)}{r}*365/{A['days']}", fmt=USD0); put(wsC, r, vc + NV + 2, 0, fmt=USD0); put(wsC, r, vc + NV + 3, f"={L(vc+NV+1)}{r}+{L(vc+NV+2)}{r}", fmt=USD0)
head_provider(11, "  by provider - cheapest Gemini (Flash-level+)", "Google")
head_provider(12, "  by provider - cheapest Claude (Flash-level+)", "Anthropic")
head_provider(13, "  by provider - cheapest OpenAI (Flash-level+)", "OpenAI")
# chart
ch = LineChart(); ch.title = "Daily cost vs calls/day - cheapest option per category"; ch.height = 9; ch.width = 24
ch.y_axis.title = "USD per day"; ch.x_axis.title = "Calls per day"
data = Reference(wsC, min_col=vc - 5 + 1, max_col=vc + NV - 1, min_row=6, max_row=9)
for rrow in (6, 7, 9, 14, 10, 11, 12, 13):
    from openpyxl.chart import Series
    sr = Series(Reference(wsC, min_col=vc, max_col=vc + NV - 1, min_row=rrow, max_row=rrow), title_from_data=False, title=wsC.cell(row=rrow, column=1).value)
    ch.series.append(sr)
ch.set_categories(Reference(wsC, min_col=vc, max_col=vc + NV - 1, min_row=5, max_row=5))
wsC.add_chart(ch, "A15")
# --- full table
FT = 37
wsC.cell(row=FT - 1, column=1, value="FULL TABLE - every option (use the header filters; sort by Year-1 total to rank)").font = f_bold
hdr(wsC, FT, ["Category", "Provider / developer", "Model", "Flag / GPU used", "Marginal $ per 1K calls"] + [None] * NV +
    ["Monthly @ ref volume", "Annual @ ref volume", "One-time SFT $", "Year-1 total $ (annual + SFT)", "Rank (Year-1)", "Saving vs selected baseline $/month", "SFT break-even months vs baseline", "Comparison tier"], height=60)
for k in range(NV):
    c = wsC.cell(row=FT, column=vc + k)
    c.value = f"={A['vol_step']}" if k == 0 else f"={L(vc+k-1)}{FT}+{A['vol_step']}"
    c.number_format = '#,##0" calls/day"'; c.font = f_hdr
cM, cA, cS, cY, cR, cV, cB = [vc + NV + j for j in range(7)]
row_i = FT + 1
first = row_i
def closed_rows(label, sheetname, last):
    global row_i
    for src in range(6, last + 1):
        r = row_i
        put(wsC, r, 1, label); put(wsC, r, 2, f"={sheetname}!A{src}", font=f_link); put(wsC, r, 3, f"={sheetname}!B{src}", font=f_link, al=Alignment(wrap_text=False))
        put(wsC, r, 4, f"={sheetname}!{L(9+NV)}{src}", font=f_link)
        put(wsC, r, 5, f"={sheetname}!E{src}", fmt=USD, font=f_link)
        for k in range(NV): put(wsC, r, vc + k, f"={sheetname}!{L(6+k)}{src}", fmt=USD0, font=f_link)
        put(wsC, r, cM, f"={sheetname}!{L(6+NV)}{src}", fmt=USD0, font=f_link); put(wsC, r, cA, f"={sheetname}!{L(7+NV)}{src}", fmt=USD0, font=f_link)
        put(wsC, r, cS, 0, fmt=USD0)
        put(wsC, r, cY, f'=IF(ISNUMBER({L(cA)}{r}),{L(cA)}{r}+{L(cS)}{r},"n/a")', fmt=USD0, font=f_bold)
        put(wsC, r, cV, f'=IF(ISNUMBER({L(cM)}{r}),{BASE_MONTH}-{L(cM)}{r},"n/a")', fmt=USD0)
        put(wsC, r, cB, "-")
        put(wsC, r, cB + 1, f"=INDEX(Closed_Pricing!$AA${P0+1}:$AA${P_last},MATCH(C{r},Closed_Pricing!$B${P0+1}:$B${P_last},0))", font=f_link)
        row_i += 1
closed_rows("Closed - Standard API", "Closed_Std", S_last)
closed_rows("Closed - Batch API", "Closed_Batch", B_last)
closed_rows("Closed - Gemini via BigQuery (batch)", "Closed_BigQuery", Q_last)
for blk_label, blk0 in (("SLM - fine-tuned, on-demand GPU (cheapest that fits)", dB0), ("SLM - fine-tuned, on 1xH100 (reference)", D0)):
    for i in range(SLM_N):
        r = row_i; dr = blk0 + i; fr = F0 + i; mr = M0 + 1 + i
        put(wsC, r, 1, blk_label); put(wsC, r, 2, f"=SLM_Models!E{mr}", font=f_link, al=Alignment(wrap_text=False)); put(wsC, r, 3, f"=SLM_Daily_Cost!A{dr}", font=f_link, al=Alignment(wrap_text=False))
        put(wsC, r, 4, f'=IF(ISNUMBER(SLM_Daily_Cost!E{dr}),"Runs on "&SLM_Daily_Cost!B{dr}&"; "&TEXT(SLM_Serving_Fit!H{fr},"0.0")&" min per 1K calls on H100","No fit / price n/a")', font=f_link, al=Alignment(wrap_text=False))
        put(wsC, r, 5, f'=IF(ISNUMBER(SLM_Daily_Cost!E{dr}),SLM_Daily_Cost!E{dr}*1000,"n/a")', fmt=USD, font=f_link)
        for k in range(NV): put(wsC, r, vc + k, f"=SLM_Daily_Cost!{L(nc+1+k)}{dr}", fmt=USD0, font=f_link)
        put(wsC, r, cM, f"=SLM_Daily_Cost!{L(nc+NV+1)}{dr}", fmt=USD0, font=f_link); put(wsC, r, cA, f"=SLM_Daily_Cost!{L(nc+NV+2)}{dr}", fmt=USD0, font=f_link)
        put(wsC, r, cS, f"=SFT_Cost!I{T0+i}", fmt=USD0, font=f_link)
        put(wsC, r, cY, f'=IF(ISNUMBER({L(cA)}{r}),{L(cA)}{r}+{L(cS)}{r},"n/a")', fmt=USD0, font=f_bold)
        put(wsC, r, cV, f'=IF(ISNUMBER({L(cM)}{r}),{BASE_MONTH}-{L(cM)}{r},"n/a")', fmt=USD0)
        put(wsC, r, cB, f'=IF(NOT(ISNUMBER({L(cV)}{r})),"n/a",IF({L(cV)}{r}>0,{L(cS)}{r}/{L(cV)}{r},"never - costs more than baseline"))', fmt='0.0')
        put(wsC, r, cB + 1, "SLM")
        row_i += 1
last_c = row_i - 1
for r in range(first, last_c + 1):
    put(wsC, r, cR, f'=IFERROR(RANK({L(cY)}{r},${L(cY)}${first}:${L(cY)}${last_c},1),"")', fmt='0')
wsC.conditional_formatting.add(f"{L(vc)}{first}:{L(vc+NV-1)}{last_c}", ColorScaleRule(start_type="percentile", start_value=5, start_color="C6EFCE", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="percentile", end_value=95, end_color="F8CBAD"))
wsC.auto_filter.ref = f"A{FT}:{L(cB+1)}{last_c}"
wsC.freeze_panes = wsC.cell(row=FT + 1, column=4)
r = last_c + 2
for t in ["Cost only: closed frontier models work zero-shot, SLMs need SFT + an in-house eval to confirm summary/extraction quality. PII/data-residency value of self-hosting is not priced.",
          "SLM rows = elastic on-demand GPU billed only during the run (cold start + compute). Daily cost = calls x variable $/call + cold-start $; no 24x7 cost. If you need real-time (sub-second) answers an always-on endpoint is required - see 'Always-on 24h' on SLM_Serving_Fit.",
          "Year-1 total = annual run cost at the reference volume + one-time SFT (hours selected on Assumptions). Break-even months = SFT / monthly saving vs the selected baseline closed model (Standard API).",
          "Not included: engineering/MLOps, monitoring, storage/egress, labelling for SFT data, retraining cadence."]:
    wsC.cell(row=r, column=1, value="• " + t).font = f_note; r += 1

# =====================================================================
# 10b. FOCUS COMPARE  (Nemotron 3.5 Lightning + Gemma 4 12B vs Flash-tier-and-above closed models)
# =====================================================================
wsX = sheet("Focus_Compare", "Focus: Nemotron 3.5 Lightning & Gemma 4 12B (SLM) vs Flash-tier-and-above closed models",
            "Closed comparators = Gemini 3.5 Flash and up, Claude Sonnet/Opus/Fable, OpenAI gpt-5.4+/gpt-6 (luna, terra, sol, astra) - NOT the cheapest lite/nano/Haiku tier. SLMs on the cheapest GPU they fit (no H100); STRESS case halves throughput and lengthens cold start.", "00B050")
wsX["A4"] = "Stress-test inputs (edit)"; wsX["A4"].font = f_bold
put(wsX, 5, 1, "Throughput factor in stress case (0.5 = SLM runs at half the assumed speed)"); put(wsX, 5, 2, 0.5, fmt='0.00', font=f_in, fill=fill_key)
put(wsX, 6, 1, "Cold start in stress case (minutes per run)"); put(wsX, 6, 2, 15, fmt='0.0', font=f_in, fill=fill_key)
put(wsX, 7, 1, "Closed prices follow the 'price basis' switch on Assumptions (Current/Post-promo); two Gemini rows below always show the post-promo list price. Thinking tokens per call: Assumptions.", font=f_note, bd=False)
XH = 9; xv = 4
RC = [f"Monthly Std @ ref", "Monthly Batch @ ref", "Std: x vs Nemotron Lightning (base)", "Std: x vs Nemotron Lightning (stress)", "Std: x vs Gemma 4 12B (base)", "Std: x vs Gemma 4 12B (stress)", "Batch: x vs Nemotron Lightning (base)", "Batch: x vs Gemma 4 12B (base)"]
hdr(wsX, XH, ["Option", "Type", "Marginal $ per 1K calls (Standard)"] + [None] * NV + RC, height=70)
for k in range(NV):
    c = wsX.cell(row=XH, column=xv + k)
    c.value = f"={A['vol_step']}" if k == 0 else f"={L(xv+k-1)}{XH}+{A['vol_step']}"
    c.number_format = '#,##0" calls/day"'; c.font = f_hdr
widths(wsX, [62, 22, 13] + [10] * NV + [14, 14] + [13] * 6)
cm = xv + NV          # monthly std col
cb = cm + 1           # monthly batch col
ti, to = A["tot_in"], A["out_closed"]
rX = XH + 1
# --- SLM rows first (fixed positions)
def slm_row(label, idx, mode):
    global rX
    rb = dB0 + idx
    put(wsX, rX, 1, label, font=f_bold, fill=fill_sub); put(wsX, rX, 2, "SLM (self-hosted)" if mode in ("base", "h100") else "SLM (stress test)", fill=fill_sub)
    ra = D0 + idx
    if mode == "base": var = f"SLM_Daily_Cost!$E${rb}"; fix = f"SLM_Daily_Cost!$F${rb}"
    elif mode == "h100": var = f"SLM_Daily_Cost!$E${ra}"; fix = f"SLM_Daily_Cost!$F${ra}"
    else: var = f"SLM_Daily_Cost!$E${rb}/$B$5"; fix = f"{A['runs_day']}*$B$6/60*SLM_Daily_Cost!$C${rb}"
    put(wsX, rX, 3, f'=IFERROR({var}*1000,"n/a")', fmt=USD)
    for k in range(NV): put(wsX, rX, xv + k, f'=IFERROR({L(xv+k)}${XH}*({var})+{fix},"n/a")', fmt=USD)
    put(wsX, rX, cm, f'=IFERROR(({A["ref_vol"]}*({var})+{fix})*{A["days"]},"n/a")', fmt=USD0, font=f_bold)
    put(wsX, rX, cb, f"=${L(cm)}{rX}", fmt=USD0)
    rX += 1; return rX - 1
S_NB = slm_row("Nemotron 3.5 Lightning 30B-A3B - base (cheapest feasible GPU)", 7, "base")
S_NS = slm_row("Nemotron 3.5 Lightning 30B-A3B - STRESS (half speed, longer cold start)", 7, "stress")
S_GB = slm_row("Gemma 4 12B - base (cheapest feasible GPU)", 8, "base")
S_GS = slm_row("Gemma 4 12B - STRESS (half speed, longer cold start)", 8, "stress")
S_NH = slm_row("Nemotron 3.5 Lightning 30B-A3B - on 1xH100 (FP8, fastest; shown for comparison)", 7, "h100")
S_GH = slm_row("Gemma 4 12B - on 1xH100 (FP8, fastest; shown for comparison)", 8, "h100")
c_first = rX
COMP = [("Google", ["Gemini 3.5 Flash", "Gemini 3.6 Flash", "Gemini 3.7 Flash", "Gemini 3.8 Flash", "Gemini 3.1 Pro Preview", "Gemini 4 Argon"]),
        ("Anthropic", ["Claude Sonnet 4.6", "Claude Sonnet 5", "Claude Sonnet 5.5", "Claude Opus 5", "Claude Opus 5.5", "Claude Fable 5.1"]),
        ("OpenAI", ["gpt-5.4", "gpt-5.5", "gpt-5.6-terra", "gpt-5.6-sol", "gpt-6-luna", "gpt-6-sol", "gpt-6-astra", "gpt-6.1-terra", "gpt-6.1-sol"])]
def ratios(r):
    for j, (num, den) in enumerate([(cm, S_NB), (cm, S_NS), (cm, S_GB), (cm, S_GS), (cb, S_NB), (cb, S_GB)]):
        put(wsX, r, cb + 1 + j, f'=IFERROR(IF(ISNUMBER({L(num)}{r}),{L(num)}{r}/{L(cm)}${den},"n/a"),"n/a")', fmt='0.0"x"', fill=fill_key if j in (0, 2) else None)
def closed_row(name, prov, label=None, cols=None):
    global rX
    pr = prow[name]
    P = lambda c: f"Closed_Pricing!${c}${pr}"
    ci, co, bi, bo = cols or ("V", "W", "Y", "Z")
    put(wsX, rX, 1, label or name, font=f_bold); put(wsX, rX, 2, prov)
    put(wsX, rX, 3, f"=({ti}*{P(ci)}+{to}*{P(co)})/1000000*1000", fmt=USD)
    for k in range(NV): put(wsX, rX, xv + k, f"=$C{rX}*{L(xv+k)}${XH}/1000", fmt=USD)
    put(wsX, rX, cm, f"=$C{rX}*{A['ref_vol']}/1000*{A['days']}", fmt=USD0, font=f_bold)
    put(wsX, rX, cb, f'=IF({P("P")}="No","n/a",({ti}*{P(bi)}+{to}*{P(bo)})/1000000*{A["ref_vol"]}*{A["days"]})', fmt=USD0)
    ratios(rX); rX += 1; return rX - 1
XR = {}
for prov, names in COMP:
    for nme in names: XR[nme] = closed_row(nme, prov)
XR["g38_post"] = closed_row("Gemini 3.8 Flash", "Google", "Gemini 3.6/3.7/3.8 Flash - POST-PROMO list price (from 2027-01-01)", ("J", "K", "M", "N"))
XR["argon_post"] = closed_row("Gemini 4 Argon", "Google", "Gemini 4 Argon - POST-PROMO list price (from 2026-12-01)", ("J", "K", "M", "N"))
c_last = rX - 1
wsX.conditional_formatting.add(f"{L(xv)}{XH+1}:{L(xv+NV-1)}{c_last}", ColorScaleRule(start_type="percentile", start_value=5, start_color="C6EFCE", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="percentile", end_value=95, end_color="F8CBAD"))
# summary of range
rX += 1
wsX.cell(row=rX, column=1, value="Range across the closed comparators above (Standard API, base-case SLM)").font = f_bold; rX += 1
for lab, fn in (("Cheapest comparator: SLM is this many x cheaper", "MIN"), ("Median comparator", "MEDIAN"), ("Most expensive comparator", "MAX")):
    put(wsX, rX, 1, lab)
    put(wsX, rX, 2, "vs Nemotron Lightning"); put(wsX, rX, 3, f"={fn}({L(cb+1)}{c_first}:{L(cb+1)}{c_last})", fmt='0.0"x"', font=f_bold)
    put(wsX, rX, 4, "vs Gemma 4 12B"); put(wsX, rX, 5, f"={fn}({L(cb+3)}{c_first}:{L(cb+3)}{c_last})", fmt='0.0"x"', font=f_bold)
    put(wsX, rX, 6, "vs Nemotron (stress)"); put(wsX, rX, 7, f"={fn}({L(cb+2)}{c_first}:{L(cb+2)}{c_last})", fmt='0.0"x"', font=f_bold)
    put(wsX, rX, 8, "vs Gemma (stress)"); put(wsX, rX, 9, f"={fn}({L(cb+4)}{c_first}:{L(cb+4)}{c_last})", fmt='0.0"x"', font=f_bold)
    rX += 1
rX += 1
wsX.cell(row=rX, column=1, value="One-time SFT cost and months to repay vs Gemini 3.8 Flash Standard at ref volume (base-case SLM)").font = f_bold; rX += 1
hdr(wsX, rX, ["Model", "H100 GPUs for SFT", "SFT $ (12h)", "SFT $ (24h)", "Monthly saving vs Gemini 3.8 Flash Std", "Break-even months (24h SFT)"], height=56)
for j, (nm, idx, srow) in enumerate((("Nemotron 3.5 Lightning 30B-A3B", 7, S_NB), ("Gemma 4 12B", 8, S_GB))):
    r_ = rX + 1 + j
    put(wsX, r_, 1, nm, font=f_bold); put(wsX, r_, 2, f"=SFT_Cost!F{T0+idx}", fmt='0', font=f_link)
    put(wsX, r_, 3, f"=SFT_Cost!G{T0+idx}", fmt=USD0, font=f_link); put(wsX, r_, 4, f"=SFT_Cost!H{T0+idx}", fmt=USD0, font=f_link)
    put(wsX, r_, 5, f"={L(cm)}{XR['Gemini 3.8 Flash']}-{L(cm)}{srow}", fmt=USD0)
    put(wsX, r_, 6, f'=IF(E{r_}>0,D{r_}/E{r_},"never")', fmt='0.00')
rX += 4
wsX.cell(row=rX, column=1, value="MARKET CROSS-CHECK - what managed (per-token) APIs charge for comparable open models, same call profile").font = f_bold; rX += 1
hdr(wsX, rX, ["Managed open-model API", "Source basis", "Input $/1M", "Output $/1M", "$ per 1K calls", "Monthly @ ref volume"], height=40)
MK = [("Databricks pay-per-token: Gemma 3 12B (2.143 / 7.143 DBU per 1M)", "Databricks FM serving page x $/DBU (unverified $0.07)", f"=2.143*{A['dbu']}", f"=7.143*{A['dbu']}"),
      ("Databricks pay-per-token: GPT-OSS-20B (1.000 / 4.286 DBU)", "same", f"=1.0*{A['dbu']}", f"=4.286*{A['dbu']}"),
      ("Databricks pay-per-token: Llama 3.1 8B (2.143 / 6.429 DBU)", "same", f"=2.143*{A['dbu']}", f"=6.429*{A['dbu']}"),
      ("Vertex AI MaaS: Gemma 4 26B (official)", "cloud.google.com/vertex-ai/generative-ai/pricing", 0.15, 0.60)]
mk0 = rX + 1
for j, (nm, src, pi, po) in enumerate(MK):
    r_ = mk0 + j
    put(wsX, r_, 1, nm); put(wsX, r_, 2, src, font=f_note, al=wrap)
    put(wsX, r_, 3, pi, fmt=USD, font=f_in if not isinstance(pi, str) else f_norm); put(wsX, r_, 4, po, fmt=USD, font=f_in if not isinstance(po, str) else f_norm)
    put(wsX, r_, 5, f"=({A['tot_in']}*C{r_}+{A['out_slm']}*D{r_})/1000000*1000", fmt=USD)
    put(wsX, r_, 6, f"=E{r_}/1000*{A['ref_vol']}*{A['days']}", fmt=USD0)
rX = mk0 + len(MK) + 1
for t in ["Cross-check logic: managed providers sell 8B-26B open models at roughly $0.35-0.75 per 1K of these calls INCLUDING margin and idle capacity; self-hosted Nemotron/Gemma on an on-demand A100/L4 lands in the same range. Gemini 3.5+ Flash, Sonnet, Opus and gpt-6 models cost several dollars per 1K calls, so the gap is structural (per-token prices of frontier-tier closed models vs commodity open-model hosting), not an artefact of optimistic SLM speed.",
          "Earlier 'about the same' result was against Gemini 2.5 Flash-Lite / gpt-5-nano - the cheapest closed tier. This sheet uses Flash-tier-and-above comparators. gpt-6-luna ($0.10/$0.50) and Claude Haiku 5.5 are the closed models that stay close to SLM cost.",
          "STRESS case: SLM throughput x the factor above and a longer cold start, on the same cheap GPU. If the SLM is still several times cheaper, the conclusion does not hinge on the planning tok/s estimates. Claude Opus 4.5-4.8 share Opus 5 pricing ($5/$25); Sonnet 4.5 equals Sonnet 4.6.",
          "Quality caveat: closed frontier models are zero-shot capable; the SLM must pass an in-house eval after SFT. Cost advantage only matters if accuracy is acceptable."]:
    wsX.cell(row=rX, column=1, value="• " + t).font = f_note; rX += 1
wsX.freeze_panes = wsX.cell(row=XH + 1, column=2)

# =====================================================================
# 10c. ACCURACY EXPECTATIONS + EVIDENCE
# =====================================================================
ZS = load("accuracy_zeroshot"); FT = load("accuracy_finetune")
wsE = sheet("Accuracy_Expectations", "Expected Accuracy: Open SLMs vs Closed Flash-tier, and Gain from SFT / Distillation (teacher = Gemini 3.6 Flash)",
            "PLANNING ESTIMATES derived from published evidence (sheets Acc_Evidence_*) - NOT measured on your transcripts. Blue cells are editable; replace with your own eval results as soon as you have them.", "FF0000")
widths(wsE, [58, 14, 14, 14, 16, 90])
put(wsE, 4, 1, "IMPORTANT: no public study covers current models (Gemma 4, Nemotron 3.5, Gemini 3.6) on real energy-retail call transcripts. Ranges below extrapolate from older/adjacent studies. Run an in-house eval on 100-200 human-labelled transcripts before any decision.", font=Font(name=FONT, size=10, bold=True, color="C00000"), bd=False)
wsE["A6"] = "Inputs (editable)"; wsE["A6"].font = f_bold
hdr(wsE, 7, ["Parameter", "Value", "", "", "", "Basis"], height=22)
EI = {}
def einput(r, key, label, val, fmt, basis):
    put(wsE, r, 1, label, al=wrap); put(wsE, r, 2, val, fmt=fmt, font=f_in, fill=fill_key); put(wsE, r, 6, basis, al=wrap, font=f_note)
    EI[key] = f"$B${r}"
einput(8, "t_ext", "Teacher (Gemini 3.6 Flash, good prompt) field-level extraction accuracy vs human gold", 0.90, PCT, "ASSUMPTION - unknown until you measure. Frontier-class LLMs typically reach ~85-95% field accuracy on well-specified schemas (clinical/receipt studies in Acc_Evidence_FineTune). Everything below scales from this.")
einput(9, "gap_lo", "Zero-shot gap, open 7-35B vs flash-tier closed - BEST case (pts)", 5, '0.0', "Planning estimate from evidence: LLMStructBench Gemma3-12B 0.72 vs GPT-4o 0.74 (2 pts); StructEval 67 vs 76 (9 pts); LingVarBench call transcripts 0.60-0.61 vs 0.82-0.90 (20-30 pts, old models).")
einput(10, "gap_hi", "Zero-shot gap - WORST case (pts)", 20, '0.0', "High end = noisy multi-speaker call transcripts with older open models (LingVarBench). Newer 2026 open models (Gemma 4, Qwen3.8, Nemotron 3.5) are likely better but NOT measured in any study found.")
einput(11, "rec_lo", "Extraction: share of teacher quality recovered after SFT/distillation - LOW", 0.90, PCT, "Published: distilled/SFT students reach ~90-105% of teacher-as-prompted quality in-distribution (UniNER, LoRA Land, clinical JSON extraction, Extract-0).")
einput(12, "rec_hi", "Extraction: share recovered - HIGH (capped at 100% of gold)", 1.00, PCT, "Students can match/slightly exceed a zero-shot teacher on narrow schemas, but do not exceed human gold labels; capped at 100%.")
einput(13, "t_sum", "Teacher summary faithfulness (consistency with source)", 0.92, PCT, "PROXY: Vectara HHEM summarization hallucination for Gemini 2.5 Flash = 7.8% -> 92.2% consistent (leaderboard 2026-10-22). Gemini 3.6 Flash not measured.")
einput(14, "sgap_lo", "Summarization zero-shot faithfulness gap - BEST (pts)", 0, '0.0', "Vectara HHEM: Gemma-3-12B 4.4%, Qwen3-8B 4.8%, Phi-4 3.7% hallucination - equal or better than Gemini 2.5 Flash 7.8% on short documents.")
einput(15, "sgap_hi", "Summarization zero-shot faithfulness gap - WORST (pts)", 10, '0.0', "FaithJudge (harder mixed tasks): Mistral-Small-24B/Phi-4 17.0%, Llama-3.1-8B 28.4% vs Gemini-2.0-Flash 10.2% (7-18 pts). Long noisy transcripts likely nearer this end.")
einput(16, "srec_lo", "Summarization: share of teacher quality recovered after distillation - LOW", 0.90, PCT, "Published: distilled students ~100% of teacher on LLM-judge scores (Llama-2-7B from GPT-4: 4.12 vs ~4.1); faithfulness improves with data size.")
einput(17, "srec_hi", "Summarization: share recovered - HIGH", 1.00, PCT, "Capped at 100% of teacher.")
# --- Table 1
r0 = 19
wsE.cell(row=r0, column=1, value="1. ENTITY EXTRACTION - expected field-level accuracy vs human gold (planning range)").font = f_bold
hdr(wsE, r0 + 1, ["Scenario", "Low", "High", "Mid", "Gap vs teacher (pts, mid)", "Evidence / reasoning"], height=40)
rows1 = [
 ("Gemini 3.6 Flash - few-shot prompt (TEACHER / closed reference)", f"={EI['t_ext']}", f"={EI['t_ext']}", "Assumed reference level (see input). Closed flash-tier models are also not perfect: expect residual errors on ambiguous ASR (account numbers, meter IDs, dates)."),
 ("Gemma 4 12B - zero-shot, un-tuned (open, dense)", f"=MAX(0,{EI['t_ext']}-{EI['gap_hi']}/100)", f"={EI['t_ext']}-{EI['gap_lo']}/100", "Gap band applied to the teacher. JSON-validity failures (23-39% in LingVarBench for 7-8B models) are the main driver at the low end; constrained decoding removes most of them."),
 ("Nemotron 3.5 Lightning 30B-A3B - zero-shot, un-tuned (open, MoE ~3B active)", f"=MAX(0,{EI['t_ext']}-{EI['gap_hi']}/100)", f"={EI['t_ext']}-{EI['gap_lo']}/100", "Same band; no public evidence separates 3B-active MoE from 12B dense on extraction (MultiWOZ slot-filling: open 8-32B scored 45-53 vs Claude Opus 4.6 61 - a 'top-tier' closed model, so overstated vs flash)."),
 ("Gemma 4 12B - AFTER SFT distilled from Gemini 3.6 Flash labels", f"={EI['t_ext']}*{EI['rec_lo']}", f"=MIN(1,{EI['t_ext']}*{EI['rec_hi']})", "Published: +15 to +50 pts over un-tuned student on narrow schemas (LoRA Land CoNLL++ 0.73->0.99; clinical JSON 29-77% -> 87-92%; LingVarBench Llama3.1-8B 0.61->0.87 with 300 transcripts)."),
 ("Nemotron 3.5 Lightning 30B-A3B - AFTER SFT distilled from Gemini 3.6 Flash labels", f"={EI['t_ext']}*{EI['rec_lo']}", f"=MIN(1,{EI['t_ext']}*{EI['rec_hi']})", "Same recovery band; larger/MoE students tend to land toward the top of the band but this is not measured for Nemotron."),
]
for i, (lab, lo, hi, ev) in enumerate(rows1):
    r = r0 + 2 + i
    put(wsE, r, 1, lab, font=f_bold, al=wrap); put(wsE, r, 2, lo, fmt=PCT); put(wsE, r, 3, hi, fmt=PCT)
    put(wsE, r, 4, f"=AVERAGE(B{r}:C{r})", fmt=PCT, fill=fill_sub); put(wsE, r, 5, f"=({EI['t_ext']}-D{r})*100", fmt='0.0'); put(wsE, r, 6, ev, al=wrap, font=f_note)
    wsE.row_dimensions[r].height = 52
# --- Table 2
r1 = r0 + 2 + len(rows1) + 1
wsE.cell(row=r1, column=1, value="2. CALL SUMMARIZATION - expected faithfulness (share of summary content supported by the transcript)").font = f_bold
hdr(wsE, r1 + 1, ["Scenario", "Low", "High", "Mid", "Gap vs teacher (pts, mid)", "Evidence / reasoning"], height=40)
rows2 = [
 ("Gemini 3.6 Flash (teacher / closed reference)", f"={EI['t_sum']}", f"={EI['t_sum']}", "Proxy from Vectara HHEM (Gemini 2.5 Flash 92.2% consistent)."),
 ("Gemma 4 12B - zero-shot (open)", f"={EI['t_sum']}-{EI['sgap_hi']}/100", f"={EI['t_sum']}-{EI['sgap_lo']}/100", "Small open models are roughly on par on short-document faithfulness (Vectara) but trail on harder/longer inputs (FaithJudge)."),
 ("Nemotron 3.5 Lightning - zero-shot (open)", f"={EI['t_sum']}-{EI['sgap_hi']}/100", f"={EI['t_sum']}-{EI['sgap_lo']}/100", "Nemotron-3-Nano (predecessor) scored 9.6% hallucination on Vectara (weaker than Gemma/Qwen/Phi-4); Lightning not on leaderboard."),
 ("Either model - AFTER distillation from Gemini 3.6 Flash summaries", f"={EI['t_sum']}*{EI['srec_lo']}", f"=MIN(1,{EI['t_sum']}*{EI['srec_hi']})", "Published: ~100% of teacher on LLM-judge scores with 1K-10K examples; contrastive/faithfulness-aware distillation raised AlignScore 0.82->0.91 (ROUGE can fall). SFT on facts the model lacks can INCREASE hallucination."),
]
for i, (lab, lo, hi, ev) in enumerate(rows2):
    r = r1 + 2 + i
    put(wsE, r, 1, lab, font=f_bold, al=wrap); put(wsE, r, 2, lo, fmt=PCT); put(wsE, r, 3, hi, fmt=PCT)
    put(wsE, r, 4, f"=AVERAGE(B{r}:C{r})", fmt=PCT, fill=fill_sub); put(wsE, r, 5, f"=({EI['t_sum']}-D{r})*100", fmt='0.0'); put(wsE, r, 6, ev, al=wrap, font=f_note)
    wsE.row_dimensions[r].height = 46
# --- Table 3 text
r2 = r1 + 2 + len(rows2) + 1
wsE.cell(row=r2, column=1, value="3. OUTPUT FORMAT (valid JSON) - evidence, not a point estimate").font = f_bold
t3 = ["Un-tuned open 7-8B models (Llama3.1-8B, Qwen2.5-7B) had 23-39% JSON decode errors in LingVarBench (2025, call transcripts) vs ~0% for GPT-4o / Gemini-1.5-Flash. Newer 2026 models are likely much better - measure it.",
      "With constrained decoding (vLLM/xgrammar/guidance), even a 1B model reached 96% schema coverage on JSONSchemaBench - better than OpenAI/Gemini API structured outputs. Format validity should be ~100% after SFT + constrained decoding, so remaining errors are VALUE errors, not format errors.",
      "IFEval (instruction/format following): Gemma 3 12B 88.9 vs GPT-4o 88.6 - no gap on general format-following."]
for i, t in enumerate(t3): wsE.cell(row=r2 + 1 + i, column=1, value="• " + t).font = f_note
# --- Table 4 teacher labeling cost
r3 = r2 + 1 + len(t3) + 1
wsE.cell(row=r3, column=1, value="4. COST OF DISTILLATION LABELS - Gemini 3.6 Flash as teacher (1 call per transcript, same token profile as Assumptions)").font = f_bold
hdr(wsE, r3 + 1, ["Labelling scenario", "Transcripts labelled", "$ per call (Batch / BigQuery)", "Labelling cost - Batch / BigQuery $", "Labelling cost - Standard API $", "Note"], height=48)
pr36 = prow["Gemini 3.6 Flash"]
sc = [("Pilot / minimum (narrow schema plateaus ~100-500 examples)", 500, "Evidence: human-level at 43-100 reports (clinical); LingVarBench 300 transcripts."),
      ("Comfortable extraction set", 5000, "Evidence: 1-5K comfortable; diminishing returns after (NuNER 1K->1M = +4 F1)."),
      ("Summarization set", 10000, "Evidence: summarization needs ~1K-10K examples."),
      ("Full planned SFT set (user plan)", None, "Assumptions: SFT samples. Far more than evidence says is needed - consider labelling 5K first, then extend only if the learning curve is still rising.")]
for i, (lab, n_, note) in enumerate(sc):
    r = r3 + 2 + i
    put(wsE, r, 1, lab, al=wrap)
    put(wsE, r, 2, n_ if n_ else f"={A['sft_n']}", fmt=NUM, font=f_in if n_ else f_link)
    put(wsE, r, 3, f"=({A['tot_in']}*Closed_Pricing!$Y${pr36}+{A['out_closed']}*Closed_Pricing!$Z${pr36})/1000000", fmt=USD4)
    put(wsE, r, 4, f"=B{r}*C{r}", fmt=USD0, font=f_bold)
    put(wsE, r, 5, f"=B{r}*({A['tot_in']}*Closed_Pricing!$V${pr36}+{A['out_closed']}*Closed_Pricing!$W${pr36})/1000000", fmt=USD0)
    put(wsE, r, 6, note, al=wrap, font=f_note)
    wsE.row_dimensions[r].height = 32
r4 = r3 + 2 + len(sc) + 1
wsE.cell(row=r4, column=1, value="Labelling cost uses the price basis on Assumptions (intro price now; doubles from 2027-01-01). Cost is one-time and small versus the GPU SFT cost; human-labelled gold set (100-200 transcripts) is separate and the real quality lever.").font = f_note
# --- Table 5 recipe
r5 = r4 + 2
wsE.cell(row=r5, column=1, value="5. RECOMMENDED DISTILLATION RECIPE & RISKS").font = f_bold
rec = ["1) Hold out 100-200 real transcripts (mix residential/C&I), labelled by humans = GOLD. Never train on them. This is the only way to know the true teacher accuracy and the student's accuracy.",
       "2) Teacher labelling: run Gemini 3.6 Flash (batch/BigQuery) with schema + few-shot examples on 1-5K transcripts; validate labels against gold; filter/correct teacher errors (teacher-label noise sets the student's ceiling). LLM labels are high-precision / low-recall for entities - include negative examples (UniNER without negative sampling: 31.5 vs 53.4 F1).",
       "3) SFT (LoRA) Gemma 4 12B and Nemotron 3.5 Lightning on the labelled set with constrained JSON decoding at inference; track field-level F1, exact match per field, JSON validity, and hallucinated-entity rate on GOLD. Plot a learning curve (500 / 1K / 5K) before buying more labels.",
       "4) Summarization: distil with faithfulness-aware filtering (drop teacher summaries that fail a groundedness check); evaluate with an LLM judge + human spot-check, not ROUGE alone.",
       "5) Risks: student cannot beat the teacher's systematic errors; SFT on facts the base lacks can raise hallucination; adapters trained on one domain can hurt another; ASR noise shifts over time (monitor drift); PII in training data (keep inside GCP project/VPC).",
       "6) LEGAL: check the Gemini / Vertex AI terms before using model outputs to train another model - several provider terms restrict using outputs to develop competing models. Have legal confirm that internal-use distillation is permitted (not verified here).",
       "7) Decision rule: adopt the SLM per field/task when student >= ~95% of teacher on GOLD with 0 critical-field regressions; route low-confidence cases to the closed model (hybrid) to cap risk."]
for i, t in enumerate(rec):
    c = wsE.cell(row=r5 + 1 + i, column=1, value=t); c.font = f_norm
wsE.freeze_panes = "A6"

def evidence_sheet(name, title, hdrs, rows, widths_, tab):
    ws = sheet(name, title, "Collected by research agents on 2026-11-08 from search results / paper pages. Many figures come from summaries and snippets - VERIFY against the source before citing externally.", tab)
    hdr(ws, 4, hdrs, height=36); widths(ws, widths_)
    for i, rec_ in enumerate(rows):
        r = 5 + i
        for j, v in enumerate(rec_):
            if j == len(rec_) - 2: url_cell(ws, r, 1 + j, str(v))
            else: put(ws, r, 1 + j, "" if v is None else str(v), al=wrap)
    ws.freeze_panes = "A5"; ws.auto_filter.ref = f"A4:{L(len(hdrs))}{4+len(rows)}"
    return ws
evidence_sheet("Acc_Evidence_ZeroShot", "Evidence: Open vs Closed LLMs, zero-shot / as-published (extraction, slot-filling, summarization)",
    ["Model", "Task", "Benchmark", "Metric", "Score", "Setting", "Year", "Source URL", "Note"],
    [[z.get("model"), z.get("task"), z.get("benchmark"), z.get("metric"), z.get("score"), z.get("setting"), z.get("year"), z.get("source_url"), z.get("note")] for z in ZS],
    [30, 22, 34, 24, 28, 22, 7, 55, 70], "FF0000")
evidence_sheet("Acc_Evidence_FineTune", "Evidence: Accuracy gain from SFT / distillation of small models",
    ["Student", "Teacher / labels", "Task", "Metric", "Before", "After", "Teacher score", "Train samples", "Setting", "Year", "Source URL", "Note"],
    [[f.get("student"), f.get("teacher"), f.get("task"), f.get("metric"), f.get("before"), f.get("after"), f.get("teacher_score"), f.get("n_train"), f.get("setting"), f.get("year"), f.get("source_url"), f.get("note")] for f in FT],
    [30, 24, 30, 22, 14, 18, 16, 14, 22, 7, 55, 70], "FF0000")

# =====================================================================
# 10d. ACCURACY - ALL SLM CANDIDATES
# =====================================================================
wsY = sheet("Accuracy_All_Models", "Expected Accuracy for ALL 17 Candidate SLMs - zero-shot vs after SFT / distillation from Gemini 3.6 Flash (planning estimates)",
            "Same method as Accuracy_Expectations, per model. Per-model 'position in gap band' is ANALYST JUDGEMENT from public proxies (IFBench, Vectara hallucination, size/class, release) - NOT measured on extraction. Edit blue cells; replace with your own eval results.", "FF0000")
AE = lambda k: "Accuracy_Expectations!" + EI[k]
put(wsY, 4, 1, "No benchmark of these models on call-transcript extraction exists. Use this sheet to RANK candidates and decide whom to evaluate first; the absolute numbers are scenario estimates.", font=Font(name=FONT, size=10, bold=True, color="C00000"), bd=False)
wsY["A6"] = "Inputs shared with Accuracy_Expectations (teacher accuracy, gap bands) are edited THERE. Local inputs:"; wsY["A6"].font = f_bold
put(wsY, 7, 1, "Uncertainty around each zero-shot estimate (+/- pts)"); put(wsY, 7, 2, 5, fmt='0.0', font=f_in, fill=fill_key)
hdr(wsY, 9, ["Size class", "Extraction: recovery LOW", "Extraction: recovery HIGH", "Summarization: recovery LOW", "Summarization: recovery HIGH", "Basis"], height=44)
CLS = [("Small (<=5B)", 0.85, 0.97, 0.85, 0.97, "ASSUMPTION - published SFT results are mostly 7-8B; <=5B students have a lower ceiling on messy multi-speaker transcripts (small encoders like GLiNER do well on plain NER but not on free-form JSON fields)."),
       ("Mid (8-14B)", 0.90, 1.00, 0.90, 1.00, "Published range ~90-105% of teacher (LoRA Land, clinical extraction, UniNER, LingVarBench 0.61->0.87)."),
       ("Large (20-35B)", 0.92, 1.00, 0.92, 1.00, "ASSUMPTION - larger students expected at the top of the band; MoE with ~3B active treated as Large for capacity but could behave closer to Mid.")]
for i, c in enumerate(CLS):
    r = 10 + i
    put(wsY, r, 1, c[0], font=f_bold)
    for j in range(4): put(wsY, r, 2 + j, c[1 + j], fmt=PCT, font=f_in)
    put(wsY, r, 6, c[5], al=wrap, font=f_note); wsY.row_dimensions[r].height = 40
CLR = "$A$10:$A$12"
widths(wsY, [34, 15, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 13, 12, 60, 14])
YH = 15
hdr(wsY, YH, ["Model", "Size class (edit)", "Position in extraction gap band (0=best, 1=worst)", "Position in summarization gap band (0=best, 1=worst)",
              "EXTRACTION zero-shot LOW", "zero-shot HIGH", "zero-shot MID", "EXTRACTION after SFT LOW", "after SFT HIGH", "after SFT MID", "Lift from SFT (pts, mid)",
              "SUMMARY faithfulness zero-shot MID", "SUMMARY after SFT MID", "Lift (pts)", "Monthly $ @ ref vol (cheapest GPU)", "One-time SFT $ (selected hours)", "Origin", "Public proxy evidence (from SLM_Models)", "Rank by post-SFT extraction MID"], height=84)
PRIOR = [ # class, p_ext, p_sum, rationale in order of slm list
 ("Large (20-35B)", 0.35, 0.30), ("Large (20-35B)", 0.30, 0.40), ("Large (20-35B)", 0.30, 0.10), ("Large (20-35B)", 0.40, 0.40),
 ("Large (20-35B)", 0.40, 0.00), ("Large (20-35B)", 0.50, 0.40), ("Large (20-35B)", 0.40, 0.50), ("Large (20-35B)", 0.50, 0.50),
 ("Mid (8-14B)", 0.50, 0.20), ("Mid (8-14B)", 0.60, 0.50), ("Mid (8-14B)", 0.65, 0.55), ("Small (<=5B)", 0.85, 0.70),
 ("Small (<=5B)", 0.80, 0.60), ("Mid (8-14B)", 0.55, 0.10), ("Large (20-35B)", 0.55, 0.55), ("Large (20-35B)", 0.50, 0.50), ("Large (20-35B)", 0.50, 0.40)]
Y0 = YH + 1
for i in range(SLM_N):
    r = Y0 + i; mr = M0 + 1 + i
    put(wsY, r, 1, f"=SLM_Models!C{mr}", font=f_link, al=wrap)
    put(wsY, r, 2, PRIOR[i][0], font=f_in); put(wsY, r, 3, PRIOR[i][1], fmt='0.00', font=f_in, fill=fill_key); put(wsY, r, 4, PRIOR[i][2], fmt='0.00', font=f_in, fill=fill_key)
    gapm = f"({AE('gap_lo')}+C{r}*({AE('gap_hi')}-{AE('gap_lo')}))/100"
    put(wsY, r, 7, f"={AE('t_ext')}-{gapm}", fmt=PCT, fill=fill_sub)
    put(wsY, r, 5, f"=MAX(0,G{r}-$B$7/100)", fmt=PCT); put(wsY, r, 6, f"=MIN({AE('t_ext')},G{r}+$B$7/100)", fmt=PCT)
    rl = f"INDEX($B$10:$B$12,MATCH($B{r},{CLR},0))"; rh = f"INDEX($C$10:$C$12,MATCH($B{r},{CLR},0))"
    put(wsY, r, 8, f"={AE('t_ext')}*{rl}", fmt=PCT); put(wsY, r, 9, f"=MIN(1,{AE('t_ext')}*{rh})", fmt=PCT)
    put(wsY, r, 10, f"=MAX(G{r},{AE('t_ext')}*({rl}+({rh}-{rl})*(1-C{r})))", fmt=PCT, fill=fill_sub, font=f_bold)
    put(wsY, r, 11, f"=(J{r}-G{r})*100", fmt='0.0')
    sgap = f"({AE('sgap_lo')}+D{r}*({AE('sgap_hi')}-{AE('sgap_lo')}))/100"
    put(wsY, r, 12, f"={AE('t_sum')}-{sgap}", fmt=PCT)
    sl = f"INDEX($D$10:$D$12,MATCH($B{r},{CLR},0))"; sh = f"INDEX($E$10:$E$12,MATCH($B{r},{CLR},0))"
    put(wsY, r, 13, f"=MAX(L{r},{AE('t_sum')}*({sl}+({sh}-{sl})*(1-D{r})))", fmt=PCT, font=f_bold)
    put(wsY, r, 14, f"=(M{r}-L{r})*100", fmt='0.0')
    put(wsY, r, 15, f"=SLM_Daily_Cost!{L(nc+NV+1)}{dB0+i}", fmt=USD0, font=f_link); put(wsY, r, 16, f"=SFT_Cost!I{T0+i}", fmt=USD0, font=f_link)
    put(wsY, r, 17, f"=SLM_Models!F{mr}", font=f_link, al=wrap); put(wsY, r, 18, f"=SLM_Models!W{mr}", font=f_link, al=wrap)
    put(wsY, r, 19, f"=RANK(J{r},$J${Y0}:$J${Y0+SLM_N-1},0)", fmt='0')
    wsY.row_dimensions[r].height = 62
Y_last = Y0 + SLM_N - 1
wsY.freeze_panes = wsY.cell(row=Y0, column=2)
wsY.auto_filter.ref = f"A{YH}:S{Y_last}"
wsY.conditional_formatting.add(f"J{Y0}:J{Y_last}", ColorScaleRule(start_type="min", start_color="F8CBAD", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="max", end_color="C6EFCE"))
r = Y_last + 2
for t in ["How to read: zero-shot MID = teacher accuracy minus a gap that depends on where the model sits in the evidence band (position 0 = small gap, 1 = large gap). After-SFT MID = teacher accuracy x recovery, where recovery sits between the size-class LOW and HIGH bounds depending on base-model strength (strong base = near HIGH; ASSUMPTION - not measured), never below zero-shot. Ranges are scenario estimates, not predictions.",
          "Position judgements (editable) use: IFBench/IFEval (format following), Vectara hallucination for summarization where available (Gemma 4 26B-A4B 5.2% and Phi-4 3.7% best; Nemotron-3-Nano 9.6% weaker), model size/class and release recency. Models with no public proxy (Olmo 3, Mistral Small 3.2, Qwen3.6-35B-A3B) sit mid-band.",
          "Chinese-origin models (Qwen) may need a US-corporate security/provenance review; Phi-4 has only a 16K context (fine for ~3K-token calls). Qwen3.6-35B-A3B substitutes the non-existent 'Qwen 2.5-25B-A3B'.",
          "Decision use: shortlist 3-4 models spanning classes (e.g. one Large dense, one MoE, one Mid, one Small) for the in-house eval, since differences between neighbours here are inside the uncertainty."]:
    wsY.cell(row=r, column=1, value="• " + t).font = f_note; r += 1
wsE.cell(row=wsE.max_row + 2, column=1, value="Per-model estimates for ALL 17 candidates: see sheet 'Accuracy_All_Models'.").font = f_bold

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
 ("GCP", "Cloud Run GPU (jobs / services, per-second billing)", "Serving (batch or real-time)", "L4 24GB, RTX PRO 6000 96GB (1 GPU per instance)",
  "L4 $0.0001867/s (= $0.67/h) + vCPU $0.000018/s + GiB $0.000002/s -> ~ $1.05/h with 4 vCPU/16 GiB; RTX PRO 6000 $0.00036522/s (= $1.31/h) + CPU/RAM -> ~ $3.2/h with 20 vCPU/80 GiB (min config assumed). Billed only while running.",
  "Best fit for nightly/occasional batch: scales to zero, no idle cost, fast cold start, serverless ops. 96GB RTX PRO 6000 fits up to ~30B models at FP8.",
  "No A100/H100 on Cloud Run; one GPU per instance (scale out with parallel tasks); RTX PRO 6000 min CPU/RAM unverified.",
  "https://cloud.google.com/run/pricing"),
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
for t in ["https://cloud.google.com/run/pricing", "https://cloud.google.com/bigquery/pricing", "https://cloud.google.com/bigquery/docs/choose-text-generation-function"]: collect(t, "BigQuery batch discount")
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
 ("HOW TO USE", "Edit blue cells on 'Assumptions' (call lengths, mix, task sizes, thinking tokens, hosting option, cold start, SFT hours). Default = ONE call per transcript: ~3K input tokens, ~500 output tokens. SLM cost = on-demand GPU run time (prefill + batched decode throughput) x per-second GPU price - NOT 24x7. Yellow cells are the most influential levers. Every other sheet recalculates."),
 ("SHEET GUIDE", ""),
 ("  Assumptions", "Token profile per call (residential vs C&I), tasks, thinking/cache, BigQuery, GPU hosting, throughput, SFT, baseline selector."),
 ("  Closed_Pricing", "Per-model $/1M tokens (input, output, cached, batch, post-promo) for 12 Gemini + 16 Claude + 22 OpenAI models, each with source URL + date."),
 ("  Closed_Std / _Batch / _BigQuery", "Daily cost 1K..15K calls/day for every closed model: real-time API; provider Batch API (-50%); Gemini via BigQuery (batch-rated + BigQuery compute)."),
 ("  SLM_Models", "17 SLMs: licence, commercial use, origin, params, MoE/active params, context, weights size, min GPU, fine-tuning support, planning tok/s on 1xH100, accuracy proxies."),
 ("  SLM_Benchmarks", "All accuracy proxy rows with source (no summarization/NER-specific scores exist publicly)."),
 ("  GPU_Pricing / _Raw", "GPU catalog used in the calculations (GCP + Databricks) and the 235-record research dump with sources."),
 ("  Deployment_Options", "GCP vs Databricks options for SFT and serving, pros/cons, prices."),
 ("  Combined_Compare", "START HERE: one table + chart comparing every closed option (Standard, Batch, BigQuery) with every SLM, daily cost at 1K-15K calls, monthly/annual, one-time SFT, Year-1 total and break-even."),
 ("  Focus_Compare", "Nemotron 3.5 Lightning + Gemma 4 12B vs Flash-tier-and-above closed models (Gemini 3.5+, Claude Sonnet/Opus, gpt-6 luna/sol/astra), with stress test and managed-API cross-check."),
 ("  Accuracy_Expectations", "Expected accuracy ranges: open SLM zero-shot vs closed flash-tier vs after SFT/distillation from Gemini 3.6 Flash (extraction + summarization), label cost, recipe, risks. Evidence on Acc_Evidence_* sheets."),
 ("  Accuracy_All_Models", "Per-model expected accuracy for all 17 SLMs: zero-shot vs after SFT/distillation (extraction + summarization), with monthly cost and SFT cost alongside."),
 ("  SLM_GPU_Matrix", "Every model on every GPU (L4, A100, RTX PRO 6000, H100): precision needed, tok/s, run minutes, cost per 1K calls, per day and per month; cheapest flagged."),
 ("  H100_vs_A100", "What H100 adds over A100 (FP8, bandwidth, compute), measured ratios, cost per token / per SFT job by hosting option, and when each is the right choice."),
 ("  SLM_Serving_Fit", "How long a batch run takes per model (prefill + batched decode throughput), which GPU it fits on (incl. L4 for small models) and cost per run at the reference volume."),
 ("  SLM_Daily_Cost", "Daily cost 1K..15K calls: Block A same-machine (1xH100), Block B cheapest feasible GPU."),
 ("  SFT_Cost", "One-time fine-tuning cost per model = GPUs x 12h / 24h x H100 $/h (with token-throughput cross-check); Vertex managed-SFT reference."),
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
         "SLM_Serving_Fit", "SLM_GPU_Matrix", "H100_vs_A100", "SLM_Daily_Cost", "SFT_Cost", "Acc_Evidence_ZeroShot", "Acc_Evidence_FineTune", "GPU_Pricing_Raw", "Sources"]
order.insert(1, "Combined_Compare"); order.insert(2, "Focus_Compare"); order.insert(3, "Accuracy_Expectations"); order.insert(4, "Accuracy_All_Models")
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    for rw in ws.iter_rows():
        for c in rw:
            if c.font and c.font.name != FONT: c.font = Font(name=FONT, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color, underline=c.font.underline)
out = os.path.join(HERE, "LLM_vs_SLM_Cost_Analysis.xlsx")
wb.save(out)
print("saved", out)
