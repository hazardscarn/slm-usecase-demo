# Inference engines for batch call-transcript processing (accessed 2026-11-09)

Workload: 1K-15K transcripts/day, ~3K in / ~500 out, JSON-constrained, fine-tuned (LoRA) open models, on-demand L4 24GB / A100 80GB / RTX PRO 6000 96GB / H100 80GB.
Raw data: inference_engines.json, inference_engines_support_matrix.json, inference_engines_benchmarks.json. Anything marked UNKNOWN/UNVERIFIED there is unverified here too.

## Sizing the load first
15K x 3.5K tokens ~ 52M tokens/day (45M prefill, 7.5M decode). Over 24 h that is ~87 output tok/s average; even at 6 h it is ~350 tok/s. Per the only comparable data (H100, 8B-27B class models), vLLM/SGLang deliver ~1,600-3,700 tok/s at 50 concurrent requests, so a single GPU is plenty for the batch and run-time is hours, not days. Engine choice therefore matters for (a) whether your model loads and works correctly, (b) constrained decoding correctness/overhead, (c) cost per run (GPU-hours billed on demand), not for peak tok/s. Throughput at 1K/day is almost any engine; at 15K/day Ollama (33 tok/s flat on H100 for 27B) would take ~2.6 days (7.5M / 33 tok/s) for 7.5M output tokens, i.e. unusable.

## Key verified facts
- vLLM v0.33.0 (2026-11-02), Apache-2.0, 96.6K stars, 2-week releases. SGLang v0.5.23 (2026-11-02), Apache-2.0, 38.6K stars. llama.cpp v0.9.0 / b12224 (2026-11), MIT, 133.8K stars. Ollama v0.40.15 (2026-11-03), MIT, 186K stars. TensorRT-LLM stable v1.2.1 (2026-04-20) but v1.4.0rc3 (2026-11-06) is where new models land; Apache-2.0 + notices. LMDeploy v0.18.0 (2026-09-28). Dynamo v1.6.2 (2026-11-05).
- **TGI is in maintenance mode AND the GitHub repo is archived** (last push 2026-03-21, last release v3.3.7 2025-12-19). README tells users to move to vLLM, SGLang, llama.cpp or MLX. Do not use.
- Granite 4.2 is a **dense** transformer (3B/8B/30B, Apache-2.0, 2026-08-25), not Mamba-hybrid. Easy for every engine except TRT-LLM (not listed) and LMDeploy (not listed). IBM and the card recommend vLLM >=0.20 / SGLang >=0.5.18.
- Gemma 4 includes a 12B Unified dense model (HF blog) in addition to E2B/E4B/26B-A4B/31B; Apache-2.0. vLLM recipes cover all. Ollama ships 12b/26b/31b/E4B.
- Nemotron 3.5 Lightning: OpenMDW-1.1 license; validated on H100/A100/Blackwell; NVFP4 checkpoint runs as W4A16 on Hopper/Ampere (no L4/Ada statement found); vLLM is NVIDIA's primary recipe. 30B-A3B BF16 = ~60GB so L4 24GB cannot host it except heavily quantized (GGUF Q4 ~19-25GB per Ollama tag sizes is borderline with KV).
- Qwen3.8 DeltaNet hybrid: Qwen's card lists SGLang, vLLM, TokenSpeed only. llama.cpp support is probably via the shared Qwen3.5 loader but conflicting/unconfirmed. TRT-LLM failed to load Qwen3.8-27B in a Sep 2026 test (v1.2.1).
- Ollama security: default bind 127.0.0.1 but no auth; CVE-2026-7482 "Bleeding Llama" (unauth heap read, ~300K exposed servers) and CVE-2026-5757. Keep it off the network.
- Cloud Run GPU offers only L4 and RTX PRO 6000 (no A100/H100), ~5 s instance start with drivers, scale to zero but min-instances billed full. Vertex Model Garden's documented vLLM container is v0.9.2 (Jul 2025), which cannot serve Gemma 4/Nemotron 3.5/Qwen3.8; plan a custom container. Databricks: engine and model coverage not documented in retrieved pages (UNKNOWN); batch via ai_query.

## Caveats on evidence
- Benchmarks are scarce and mostly H100 with generic prompts. Orderings conflict: Jarvislabs (vLLM > SGLang ~ TRT-LLM), Spheron (TRT-LLM > SGLang > vLLM), Winder (SGLang +7% vs vLLM at 50 conc). Differences within ~10% are noise; pick on support and ops.
- No published L4/RTX PRO 6000 multi-engine benchmark; none for Gemma 4 or Nemotron 3.5 Lightning cross-engine; none at 3K/500 JSON-constrained. Benchmark yourself with `vllm bench serve` / GuideLLM on your own transcripts.
- Constrained decoding: SqueezeBits (older versions) found vLLM losing throughput at batch >=8 with XGrammar and SGLang closer to baseline; re-test on v0.33.0. Unconstrained models produced valid-per-schema output only 61-94% of the time, so constrained decoding is worth its overhead. LLGuidance better when schema changes per request; XGrammar when fixed.
- Hybrid-model rough edges: Mamba/DeltaNet prefix caching needs special flags in vLLM; SM120 NVFP4 had bugs (Feb-Apr 2026). Validate with your fine-tuned checkpoint on the actual GPU. LoRA on hybrid/MoE models: vLLM table says yes; SGLang only just added quantized-MoE LoRA (0.5.23). Safest path for MoE/hybrid fine-tunes: merge LoRA then serve (loses multi-LoRA).

## Ranking: THIS batch use case (offline, high volume, JSON-constrained, multi-model, on-demand GPUs)
1. **vLLM** - widest, day-0 coverage of every target model (Gemma 4, NemotronH, Granite, gpt-oss, Qwen3Next/3.8 per cards), every vendor card recommends it, offline `LLM` + `vllm run-batch` with per-request structured outputs, xgrammar/llguidance, multi-LoRA, prefix caching on by default, runs on L4/A100/H100/SM120. Strongest or tied-strongest throughput in most benchmarks. Pin the version; test SM120 FP4.
2. **SGLang** - same throughput class, best if the long shared instruction prefix dominates (RadixAttention) and for constrained decoding overhead; day-0 for Gemma 4/Qwen3.8; slightly rougher docs, pre-1.0 churn, SM120 unverified, MoE LoRA new.
3. **TensorRT-LLM (via trtllm-serve; NIM if you want support)** - only worth it on H100/Blackwell for a frozen model after you have everything else working; lags on Granite (not listed), Nemotron 3.5 needs rc, Qwen3.8-27B load failed; long compile; rc-heavy releases. NIM adds enterprise support/licensing cost (UNVERIFIED pricing).
4. **llama.cpp (llama-server)** - good on L4 with GGUF Q4 for small models and for 1K/day where throughput is irrelevant; collapses under concurrency (44x slower than vLLM at 64 users on H200). Hybrid-model support exists but perf lags.
5. **LMDeploy** - fast historically but does not list Gemma 4, NemotronH, Granite; skip.
6. **Ollama** - dev only; 33 tok/s flat on H100, security exposure history.
7. **TGI** - archived/maintenance; do not adopt.
Managed: **Cloud Run GPU + your own vLLM image** is the best managed fit for bursty daily batch (L4/RTX PRO 6000, scale to zero, ~5 s start) but no A100/H100. Vertex Model Garden container (vLLM v0.9.2 documented) and Databricks are only viable if model coverage is confirmed. For H100/A100 use a VM or GKE/K8s Job running `vllm run-batch` and terminate after the run.

## Ranking: REAL-TIME use case (low latency, interactive, mixed concurrency)
1. **vLLM** (async scheduling, spec decode incl. DFlash/MTP/EAGLE for Nemotron 3.5, Gemma 4 MTP) - safest default.
2. **SGLang** - lowest TTFT in some tests, best prefix reuse for chat/agents.
3. **TensorRT-LLM / NIM** - best peak per-GPU numbers on H100/Blackwell after tuning; TTFT can degrade under load (Jarvislabs).
4. **llama.cpp** - best single-user latency (185 tok/s vs 150 vLLM on 8B) and edge; poor concurrent scaling.
5. **Ollama** - local prototyping only.
6. LMDeploy, TGI (archived) last.

## Practical recommendation
Use vLLM (container pinned to v0.33.x) with `vllm run-batch` or the offline `LLM` class and `structured_outputs` JSON schema; merge LoRA for MoE/hybrid models, use adapters only for dense models if you need several variants; benchmark SGLang as a challenger on the same transcripts; run on whichever GPU gives the lowest on-demand cost per run, and note that 30B-A3B BF16 (~60GB) needs A100/H100/RTX PRO 6000, while L4 24GB suits only <=9B BF16 or 4-bit quantized larger models (L4 NVFP4/W4A16 for Lightning is undocumented).
