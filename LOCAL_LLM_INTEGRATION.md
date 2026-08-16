# Local LLM Integration — Context & Plan

Carried over from a separate chat session that tuned the local LLM server and scoped
this integration. Read this first before starting work so we don't re-derive it.

## Local LLM server (already running, standalone process)

- **Endpoint:** `http://127.0.0.1:1234` (OpenAI-compatible API — LM Studio server)
- **Model id:** `qwen3-4b-thinking` (Qwen3 4B, thinking/reasoning model, Q4_K_M quant)
- **Chat endpoint:** `POST http://127.0.0.1:1234/v1/chat/completions` (standard OpenAI
  chat-completions schema — `messages`, `max_tokens`, etc.)
- **Launcher:** `D:\Local AI\start-gpu.bat` — loads the model with GPU offload
  (`--gpu max --parallel 1`) via LM Studio's Vulkan backend on the machine's Intel
  Iris Xe iGPU, then starts the server. Run this before testing if the server isn't
  already up (check with `curl http://127.0.0.1:1234/v1/models`).
- This is a separate OS process, independent of any coding session — it stays running
  regardless of which chat/task is open.

## Hardware reality (why this model, why this shape of integration)

- Machine: i7-1185G7 (4C/8T), Iris Xe iGPU, 32GB RAM, no dedicated VRAM.
- Achieved speed: **~8-10 tokens/sec** via GPU offload (Vulkan), stable. CPU-only
  fallback (llamafile) was slower and highly inconsistent (2-8 t/s) due to contention
  with other running processes — GPU path is the one to use.
- At this speed, a 2-3 sentence answer takes roughly 10-20 seconds. Fine for: async
  report narration, one-off staff questions in chat. **Not** fine for snappy
  multi-user live chat — the model serves one request at a decent speed at a time.
- Other local models available on this machine (gemma-4-12B, Qwen3.5-9B, gemma-4-E4B,
  Nemotron-3-Nano-4B) are all slower — 4B is the speed sweet spot here. Don't suggest
  swapping models for speed reasons without re-benchmarking.

## Codebase findings (from inspecting this repo)

- Django app, SQLite (`db.sqlite3`), Django REST Framework already installed
  (`djangorestframework==3.17.1`).
- `bot/` module: scheduled email reports. Flow today is
  `report_adapter.py` (fetch data) → `template_engine.py` (render) → `services.py`
  (`run_bot()`: render email, send via SMTP, log a `BotExecution` row). **No AI
  involved currently** — this is the first integration point.
- `chat/` module: real-time internal staff chat/presence over Django Channels
  (`consumers.py`, `presence.py`, `realtime.py`, websockets). No AI involved
  currently — this is the second integration point (an "AI Assistant" bot user
  staff could message).
- No LLM/AI libraries in `requirements.txt` yet (no `openai`, `requests`-based LLM
  client wrapper, no `langchain`). Will need a small client module — a plain
  `requests.post()` to the local endpoint is enough, no SDK required.

## Design principle — non-negotiable

**The LLM must never be the source of factual numbers** (job costs, stock counts,
machine hours, dates). It only phrases/summarizes numbers the Django ORM already
fetched exactly. Pattern:

```
Django ORM query → structured/exact data → LLM turns it into readable prose
```

This avoids hallucinated figures in a business-critical system, and matches how a
small 4B model should be used — narrating known facts, not "knowing" the business.

## Proposed integration points, in order of easiest win

1. **`bot/` report narration** — insert an LLM step between `report_adapter.py`'s
   fetched data and `template_engine.py`'s render, to prepend/append a plain-English
   summary paragraph to the existing report email. Smallest, most contained change.
2. **`chat/` AI assistant** — a bot user in the existing chat system. Staff message
   it in natural language (e.g. "what's the status of JC-07-26-PP-0701"); a view
   resolves the query to an ORM lookup, then asks the LLM to phrase the answer.
3. **Standalone internal API endpoint** — a DRF view: natural-language question in,
   ORM query resolved server-side, LLM-phrased answer out. Useful as a building block
   for either of the above, or a future dashboard "ask a question" box.

## Deployment plan — local first, then Oracle VM

Test and validate entirely on this local machine first (local LLM server + local
Django dev server). Only once the integration works and is worth keeping do we
port it to the Oracle VM production server. Implications:

- Don't hardcode `http://127.0.0.1:1234` deep in business logic — put the LLM
  endpoint URL in settings/env (it already uses `django-environ`/`.env`, so this
  fits the existing pattern), so switching targets later is a config change, not
  a code change.
- The Oracle VM almost certainly won't have this machine's iGPU. Before porting,
  check what's available there (CPU-only? a GPU? how much RAM?) — the achievable
  t/s and possibly the model choice may need to be re-evaluated for that hardware;
  don't assume the same ~8-10 t/s carries over.
- Whether the LLM server itself later runs *on* the Oracle VM, or the VM calls out
  to an LLM server running elsewhere, is an open decision — revisit once the local
  proof-of-concept is working and we know what the feature actually needs
  (latency tolerance, concurrent users, etc.).

## Suggested first step (proof of concept)

Build a small Django management command or API endpoint that:
1. Pulls one existing model's data (e.g. a single job card or a week of machine
   planning data).
2. Sends it + a short instruction to `http://127.0.0.1:1234/v1/chat/completions`.
3. Prints/returns the LLM's plain-English summary.

Confirms the plumbing end-to-end before wiring into `bot/` or `chat/` for real.

## Privacy/offline note (relevant context, not a rule to enforce)

This ERP's own docs (`DISASTER_RECOVERY.md`, `DEPLOYMENT.md`) indicate an offline/
on-prem posture. A local LLM fits that: job costing and supplier pricing data never
leaves the machine, no per-token cost, no internet dependency for this feature.
