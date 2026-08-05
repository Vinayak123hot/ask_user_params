# ask_user_params generator (one-time / local tool)

Generates `ask_user_params.json` entries with an LLM (**GPT-5 mini**) so you don't have to
hand-write them. For each KB it pairs the KB's article (`.docx`) with its PowerShell script
(`.ps1`/`.txt`), sends both to the model in one **stateless** structured call, and writes an
entry in the exact shape the validation service expects.

## Folder layout
```
ask_user_params/
  config.yaml                  ← endpoint, GPT-5 mini deployment, folder paths
  generate_ask_user_params.py  ← the generator
  requirements.txt
  kb_articles/                 ← put your <KB>.docx article files here      (you create/fill)
    KB001234.docx
  scripts/                     ← put your <KB>.ps1 (or <KB>.txt) scripts here (you create/fill)
    KB001234.ps1
  ask_user_params.json         ← generated output (merged)                  (created by the tool)
```
Files are **paired by the KB number** in the filename (the stem becomes the kb_id key, which
must match `validation.json`). Create the `kb_articles/` and `scripts/` folders and drop your
files in.

## What the LLM produces
For each parameter it extracts `name`, `required`, and `description` from the script's
`param()` block + the article, plus a few **example values** and worked **examples**.
`allowed_values` are **format examples only** — never an enforced or complete list (a value
set may be huge, e.g. time zones or add-ins). The validation service does not reject a value
for being outside them; instead every value is safety-checked (bounded, no metacharacters)
and must be grounded in what the user actually said (provenance), then confirmed.

```json
"KB001234": {
  "script": "KB001234.ps1",
  "params": [
    { "name": "AddinName", "required": true, "description": "The add-in ProgID to enable" },
    { "name": "Scope", "required": false, "default": "CurrentUser", "description": "Registry scope" }
  ],
  "allowed_values": {
    "AddinName": ["TeamsAddin.FastConnect", "AdobePDFMakerOfficeAddin"],   // format examples (NOT the full set)
    "Scope": ["CurrentUser", "LocalMachine"]                              // format examples (not enforced)
  },
  "examples": [ { "prompt": "enable the Teams add-in", "params": { "AddinName": "TeamsAddin.FastConnect", "Scope": "CurrentUser" } } ]
}
```

## Setup
1. Install deps from your **private index**:
   ```
   pip install -r requirements.txt --index-url https://<your-private-index>/simple
   ```
2. `az login` (no keys). Your user needs **Cognitive Services OpenAI User** on the AI Services resource.
3. Edit **`config.yaml`** — set `azure_openai.deployment` to your **GPT-5 mini** deployment
   (and `endpoint` if it's on a different resource).

## Run
```
# all KB pairs:
python generate_ask_user_params.py

# just one KB:
python generate_ask_user_params.py --kb KB001234

# override config values ad hoc:
python generate_ask_user_params.py --deployment my-gpt5mini --endpoint https://....services.ai.azure.com/
```
Output: **`ask_user_params.json`** (merged — running per-KB accumulates). The tool prints a
summary and ⚠ warnings (e.g. a param with no example values, or an example using an undeclared param).

## After generating
1. **Review** the JSON — confirm the params/examples look right and `script` is the correct
   `.ps1` filename.
2. **Copy** `ask_user_params.json` into `clasification_agent_validation_user_input/` and deploy.
3. Ensure each kb_id is also in `validation.json` as `["automatic","ask_user"]`.

## Notes
- **Stateless**, one call per KB, **no `temperature`** (GPT-5 family).
- **Auth:** Entra ID via `az login` — no API keys, no App Settings.
- **Images** in the `.docx` are ignored (only text is needed to derive parameters).
- **Values aren't enforced against a fixed list** in the service — a wrong-but-safe value is
  caught by the confirm turn and by the downstream executor validating the value (e.g. the
  script's own `ValidateSet`, or that an add-in ProgID exists).
