# Accuracy lift from SFT / distillation of small open models (extraction + call summarization)

Compiled 2026-11-08. Records: accuracy_finetune.json (23 rows). All numbers are from the cited pages as read in this session; "null" = not found. Several rows are vendor/community anecdotes and are labelled as such. The NeurIPS-2026 and arXiv 2608.* entries were read only via listing/abstract pages; treat as lower confidence.

## 1. Structured / entity extraction

| Evidence | Before (base/zero-shot) | After SFT | Frontier ref | n_train |
|---|---|---|---|---|
| LoRA Land, CoNLL++ (7B, human labels) | 0.733 | 0.989 | GPT-4 0.742 | 14K |
| LoRA Land, BC5CDR | 0.703 | 0.972 | GPT-4 0.890 | 5.2K |
| Clinical reports -> JSON, Llama-3.1 8B | 29-77% EM | 87-92% EM | GPT-4 zero-shot 86% | 95-391 |
| Extract-0 (7B, doc extraction) | 0.232 | 0.507 SFT / 0.573 +RL | GPT-4.1 0.457, o3 0.464 | 280K synthetic |
| UniNER-7B (ChatGPT-distilled, zero-shot NER) | Alpaca/Vicuna >30 F1 lower | 41.7 F1 | ChatGPT 34.9 | 45.9K |
| GLiNER-L (0.3B, same ChatGPT labels) | - | 47.8 F1 | ChatGPT 36.5 | ~45K |
| NuNER (125M encoder, GPT-3.5 labels, k=8) | 52.6 | 64.8 | GPT-4 higher on some sets | 1M sentences |

Take-aways
- Lift over the un-tuned small model is large: typically +15 to +60 points absolute on narrow, fixed-schema extraction (zero-shot 7-8B models are often 30-75%).
- Fine-tuned small models usually match or beat zero-shot/few-shot frontier models on a narrow schema (by +0 to +25 pts in LoRA Land; +11 in Extract-0; +4-7 in clinical). Caveat: most comparators are zero-shot GPT-4/ChatGPT-class, not a well-prompted current frontier model with few-shot examples; gains shrink against strong prompting. No head-to-head against Gemini 2.5/3-class or current Claude was found.
- Distillation from an LLM annotator can make the student exceed the teacher *as a prompted model* (UniNER +6.8 F1, GLiNER +11 F1 over ChatGPT). This is not "student > teacher labels"; the teacher is evaluated zero-shot while the student is trained in-distribution on many entity types and negatives.
- Generative/loosely-defined tasks gain less (E2E_NLG +7 pts, only +4 over GPT-4).
- Practitioner reports (OpenPipe: ~$80K->$15K/mo; "1/50th cost") are vendor claims without protocol; Predibase LoRA Land says <$8 per fine-tune. Extract-0 gives a rare auditable cost: $196 total (self-reported). The clinical study: $0.80-$2.40 training per dataset.

## 2. Dialogue / call summarization

- Llama-2-7B distilled from GPT-4 on ~11K synthetic summaries: Prometheus factual validity 4.12 vs 3.84 for chat baseline, equal to GPT-4 (about 100% of teacher; LLM-judge, simulated calls). Without task-specific data: 2.60.
- Llama-3.1-8B on 10K GPT-3.5 summaries per channel (customer service): 4.33-4.41 vs teacher 4.05-4.29 on a 1-5 GPT-4-Turbo judge (+0.1 to +0.3), after targeted error correction of teacher outputs.
- Small seq2seq (Flan-T5, 770M) with contrastive distillation from gpt-3.5 on 14.7K SAMSum: AlignScore 0.82 -> 0.91 while ROUGE falls. Faithfulness gains come at the cost of reference overlap; do not use ROUGE alone.
- Data scaling (LoRA, Llama-3.1-8B): MiniCheck faithfulness 0.496 -> 0.680 from zero-shot to full data; rank 8 vs 16 irrelevant. An XSum-trained adapter hurt SAMSum (20.5 vs 23.3 ROUGE-1): domain mismatch can hurt.
- I found no public paper with call-centre transcripts, real (not simulated) calls, and a current frontier teacher reporting before/after numbers; this is a gap. Treat our expected values as extrapolation.

## 3. How much data, ceiling, noise, LoRA vs full

- Narrow-schema extraction plateaus fast: human-level at ~43-100 labelled reports (clinical); NuNER few-shot gain from 1K to 1M examples is only 60.3 -> 64.8 F1 (diminishing). Summarization needs more: faithfulness kept rising up to full-train-set size; 1K examples gave ROUGE-1 ~28.6 vs ~37.8 on full data (XSum).
- Practical bands (my synthesis, not measured in one study): ~100-500 examples = format/schema compliance and most of the lift for fixed-schema JSON; 1-5K = reliable for moderately variable extraction and summaries; 10K+ = summaries faithfulness/style robustness; 50K-280K only for open-schema/open-NER generality.
- Ceiling vs teacher: published students reach about 95-110% of teacher-as-prompted quality on in-distribution tasks (summaries 100-110% on judge scales; NER > 100% vs zero-shot ChatGPT). Against human gold labels, a student cannot systematically exceed teacher label quality; "student > teacher" cases come from filtering/correction, in-distribution training, or weak teacher prompting.
- Teacher noise: LLM annotators are high-precision/low-recall (NuNER); noise-aware losses recover roughly 1-2 F1; students partly denoise (Student-as-denoiser, NeurIPS 2023). Missing negatives/uniform sampling dropped UniNER from 53.4 to 31.5 F1: label-recipe errors cost far more than random noise.
- LoRA vs full FT: LoRA matches full FT on small/medium SFT sets (Thinking Machines 2025; LoRA Land used r=8 4-bit throughout); gaps appear for large data/hard new skills (Biderman 2024). LR ~10x higher for LoRA.

## 4. Risks
- Hallucination: SFT on facts the model does not know raises hallucination (Gekhman 2024). Use SFT for format/skill. Fine-tuned extractors in the community example default to null more often, but this is anecdotal.
- OOD transcripts and schema drift: adapter transferred negatively (SAMSum); the clinical models are dataset-specific. Expect re-tuning or monitoring when schema or domain changes.
- Judge bias: summary results rely on LLM judges (often from the teacher family); ROUGE disagrees with faithfulness.
- Generalization of vendor numbers: LoRA Land numbers are mean over many easy classification tasks; do not extrapolate to long-document extraction.

## 5. Honest expected range for our use case (7-35B student, LoRA, teacher = frontier LLM labels, in-domain)

- Extraction (fixed JSON schema): absolute lift over zero-shot student +15 to +50 pts (field-level EM/F1), typical post-FT 85-95% field accuracy; recovers ~90-105% of teacher-as-prompted quality (low end for long/noisy documents, high end for short, repetitive forms). Samples: 100-500 minimum viable, 1-5K comfortable; >5K rarely helps unless schema is open.
- Call summarization: lift on judge/faithfulness metrics modest but real: +0.3 to +1.0 on a 5-point rubric versus un-tuned small model, or +10-25% relative faithfulness; recovers ~90-100% of teacher quality (some studies >100% with error-corrected labels). Samples: 1K-10K; ROUGE may drop.
- Unknown/uncertain: behaviour on real noisy ASR call transcripts, and against current (2026) frontier teachers with strong few-shot prompts. Plan a held-out in-domain eval of 200+ items before trusting these ranges.
