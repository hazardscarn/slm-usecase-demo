# Gemini via BigQuery: pricing notes (accessed 2026-11-08)

## Bottom line
For Google Gemini models called from BigQuery generative AI functions, Vertex/Agent Platform token charges are billed at the **Batch/Flex rate (50% off standard)**, plus normal BigQuery compute. No extra per-call BigQuery surcharge is listed.

## Sources (verbatim)
- https://cloud.google.com/bigquery/pricing, "BigQuery ML remote model" table: Google models hosted on Gemini platform with ML.GENERATE_TEXT, ML.GENERATE_EMBEDDING, AI.GENERATE, AI.GENERATE_BOOL/DOUBLE/INT, AI.IF, AI.SCORE, AI.CLASSIFY -> "Generative AI in Gemini platform batch pricing". Anthropic Claude via ML.GENERATE_TEXT and open models are NOT batch: standard Vertex pricing.
- https://cloud.google.com/bigquery/docs/choose-text-generation-function: "If you are using a Gemini 2.0 or greater model, the call is billed at the batch API rate."
- https://cloud.google.com/vertex-ai/generative-ai/pricing: "Flex/Batch" tab for each family = 50% of standard (e.g. 3.1 Flash-Lite $0.125 in / $0.75 out vs $0.25 / $1.50; 2.5 Pro $0.625/$5 vs $1.25/$10).
- AI.GENERATE_TABLE is not named in the BigQuery pricing table (verified list above omits it); by sibling behavior it is likely batch-rated too, but UNVERIFIED. Check actual SKUs/billing export.

## How billing works
1. BigQuery compute: bytes processed ($6.25/TiB on-demand, first 1 TiB/month free) or slot-hours under Editions. BigQuery ML charges not itemized separately.
2. Vertex token charges billed directly by Vertex (billing_service='bigquery_ml' label + bigquery_job_id_prefix label let you filter in billing export). Thinking tokens billed as output.
3. Context cache storage applies only if you pass a cachedContent via model_params; cached input price rate applies.
4. Provisioned Throughput (request_type = DEDICATED) is billed under PT, not batch. Default request_type UNSPECIFIED/SHARED uses dynamic shared quota.
5. Non-global regional endpoints: +10% on Gemini 3.x+ (Vertex page, effective 2026-07-01). Use global endpoint to avoid.
6. Long-context: if prompt >200K, all tokens billed at long-context rate (Pro).
7. Supervised tuning: extra bytes-processed + GPU/TPU charges.

## Caveats
- Intro pricing of 3.6/3.7/3.8 Flash ends 2026-12-31; Gemini 4 Argon promo ends 2026-11-30 (Vertex promo = credits back).
- Whether batch rate applies to a given newest model not confirmed on a per-model basis; docs say "2.0 or greater".
- Cost control docs: AI.GENERATE page warns actual rows processed can exceed expected (reruns/retries); materialize inputs first; use token quotas.
- Gemini 2.5 Flash/Pro are deprecated; 2026-12-01 auto-route to 3.5 Flash-Lite / 3.8 Flash.
- Gemini Developer API (ai.google.dev) "Batch API (50% cost reduction)" is a separate product from the BigQuery path but same discount.
