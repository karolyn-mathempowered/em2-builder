# Render redeploy — fix for `/app/assets/logo.png` missing

Replace the contents of your Render-connected repo with everything in this zip
(keep the same root layout — `server.py`, `Dockerfile`, `assets/` all sit at
the repo root). Commit and push; Render will rebuild automatically.

## Files
- `Dockerfile` — now copies `assets/` and `EM2_Lesson_Template.pptx` explicitly.
- `server.py` — `/health` now reports `assets_loaded` and `has_logo` so you can verify the fix from the browser.
- `build_module.py` — unchanged from your latest version.
- `requirements.txt` — unchanged.
- `EM2_Lesson_Template.pptx` — required by the builder.
- `assets/` — every image the builder needs (logo, DARE icons, schedule icons, timer pills, etc.).

## Verify after deploy
Open `https://em2-builder.onrender.com/health`. You should see something like:
```json
{"ok": true, "service": "em2-module-builder", "assets_loaded": 21, "has_logo": true}
```
If `has_logo` is `true`, click **Generate Deck** in the app and it will work.
