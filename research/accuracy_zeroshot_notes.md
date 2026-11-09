# Zero-shot accuracy gap: open SLMs (3-35B) vs closed flash-tier, extraction and summarization

Researched 2026-11-08. Records: accuracy_zeroshot.json (62 rows). Many figures came via WebFetch summaries and search snippets, not hand-checked against the PDFs (exceptions: LingVarBench table, read from the PDF text). Treat every number as "verify before citing externally".

## Evidence base is thin
- No single study compares current open SLMs (Qwen3.5, Gemma 4, Granite 4, Nemotron 3, gpt-oss-20b) with current flash-tier models (Gemini 3.x Flash, GPT-5.x mini) on call-transcript extraction or summarization. This was not found.
- Not found: any energy/utilities study with open-vs-closed numbers (only Oracle Utilities call-summarization product docs and an AHFE conference paper on utility call analysis, no scores retrieved).
- Not found: zero-shot numbers on CrossNER, MIT Movie/Restaurant, DocILE/SROIE/CORD, QMSum, MeetingBank, TweetSumm, ECTSum for 3-35B models. SAMSum/DialogSum: only qualitative statements, no usable numbers.
- Many closed baselines are old (GPT-4o, Gemini 1.5/2.0 Flash, GPT-3), which tends to understate today's flash-tier strength.

## Findings by task

### Structured / JSON extraction
- Call transcripts (closest to the use case), LingVarBench, 2025, zero-shot, N=20 questions per transcript: judgment accuracy GPT-4o 0.90, Gemini-1.5-Flash 0.82, GPT-4o-mini 0.81 vs Llama3.1-8B 0.61, Qwen2.5-7B 0.60. Gap to flash-tier about 20 points. Navigation F1 gap is larger (0.47 flash vs 0.20-0.29 open). JSON decode errors: 0% closed vs 23-39% open. Fine-tuning on 300 transcripts closed the gap (Llama3.1-8B 0.87). Caveats: the reference answers are GPT-4o's own, which favours GPT-4o; models are old.
- LLMStructBench 2026, zero-shot: Gemma3-12B 0.72 vs GPT-4o 0.74 (about 0.02 gap); Gemma3-27B ties GPT-4o. Optimistic result, document-parsing flavour.
- StructEval 2025: best open (Qwen3-4B) 67.0 vs GPT-4o 76.0, a 9-point gap (open set was older/smaller models).
- JSONSchemaBench 2025: with constrained decoding (Guidance, llama.cpp) even a 1B model beats OpenAI/Gemini API structured output on schema coverage (96/95 vs 89/86 on GlaiveAI). Format validity is a solved problem when a grammar engine is used; content accuracy is separate.
- NER: zero-shot open-vs-closed numbers were not found. Fine-tuned (LoRA) 3-9B models reach 89-94 F1 on CoNLL/OntoNotes (2026), beating a GPT-3 few-shot baseline; this is not a zero-shot gap. Extract-0 (7B, SFT+RL) beat GPT-4.1/o3 on its own benchmark (0.573 vs 0.46), vendor-style evidence.
- Slot filling / DST, 2026, zero-shot with agentic scaffold: MultiWOZ 2.1 JGA Qwen3-8B 47.3, Gemma3-12B 45.1, gpt-oss-20B 52.7, Qwen3-32B 51.5 vs Claude Opus 4.6 61.3. SGD: 56-64 vs 80.7. Closed comparator is a top-tier model, not flash, so this is an upper bound of the gap (9-16 pts MultiWOZ, 17-25 pts SGD).

### Summarization and faithfulness
- Vectara HHEM (snapshot 2026-10-22, via fetch summary): hallucination rate Gemma-3-12B 4.4%, Qwen3-8B 4.8%, Qwen3-14B 5.4%, Qwen3-4B 5.7%, Phi-4 3.7% (answer rate 81%), Mistral-Small 5.1% vs Gemini-2.5-Flash 7.8%, Gemini-2.5-Flash-Lite 3.3%, GPT-4o 9.6%, Claude Sonnet 4 10.3%. Open SLMs match or beat closed models on short-document faithfulness. Outliers: Granite-3.3-8B 10.6%, Nemotron-3-Nano-30B 9.6%, Gemma-3-4B has 67% answer rate. HHEM measures only consistency on short news-like passages, not completeness or call-style content.
- FaithJudge 2025 (harder, mixed tasks): Llama-3.1-8B 28.4%, Mistral-Small-24B 17.0%, Phi-4 17.0% vs Gemini-2.0-Flash 10.2%, GPT-4o 15.9%, Claude-3.7 16.1%. Here open SLMs are 7-18 pts worse than the flash model but similar to GPT-4o at 17-24B.
- Call summarization (Thulke 2024): fine-tuned Llama-2-7B faithfulness 84.4% vs GPT-4 88.9%, Prometheus 4.10 vs 4.19; fine-tuned, old.
- Dialogue summarization SAMSum/DialogSum and FACTS Grounding for open SLMs: no numbers retrieved. FACTS: Gemini 2.5 Flash 85.3%, Flash-Lite 84.1%.
- IFEval (format following proxy): Gemma 3 12B 88.9 vs GPT-4o 88.6 and Gemini 2.5 Flash 84.6, from an aggregator with mixed metric variants; no clear gap.

## Honest gap table (open 7-14B SLM vs flash-tier closed, zero-shot, closed minus open)

| Task | Gap range (points on native metric) | Confidence | Basis |
|---|---|---|---|
| Valid JSON / schema format, unconstrained | 0 to 40 pts error-rate difference (open 23-39% decode errors on old Llama/Qwen) | Low-medium | LingVarBench; near zero with constrained decoding |
| Valid JSON with constrained decoding | about 0 (open at or above closed) | Medium | JSONSchemaBench |
| Multi-field extraction from call transcripts (accuracy) | roughly 10-25 pts (observed 20 on old models) | Low; one study, old models | LingVarBench |
| Generic structured extraction composite | 0-10 pts | Low-medium | LLMStructBench 2, StructEval 9 |
| Slot filling / DST (JGA) | 9-25 pts vs top-tier; likely smaller vs flash | Low | ReacTOD |
| NER zero-shot | not found | none | n/a |
| Summary faithfulness (short docs) | about -3 to +3 pts (open on par) | Medium | HHEM |
| Faithfulness (harder, mixed) | 0-18 pts hallucination rate worse for open | Medium-low | FaithJudge |
| Summary quality/completeness on calls | unknown; fine-tuned ~4-5 pts behind GPT-4 | Low | Thulke 2024 |

Suggested planning range (not a measured value): expect open 7-14B models to trail flash-tier by about 5-20 points on zero-shot extraction accuracy and 0-10 on summarization faithfulness, with extraction on noisy multi-speaker transcripts toward the upper end. Newer models (Qwen3.5, Gemma 4) probably narrow this, but no published evidence retrieved. Closing gap: few-shot + constrained decoding + light fine-tuning (LingVarBench, 300 transcripts) is the best-supported route; zero-shot alone should not be the expected operating mode.

## Recommended next step
Run an in-house eval of 100-200 labelled utility call transcripts (our own field-level F1 and a human-checked faithfulness rubric) on the target models; the literature cannot substitute for this.
