# PaymentSave Arrears Portal

A searchable portal for arrears across three lease providers: **FDGL**, **Paytek** and **PS Lease**, plus **PS Team** and **New in Arrears** views.

Split rule: in the Del Report, any *Lease No.* containing `PSAVE` = **Paytek**; all others = **FDGL**. The PS Lease file feeds the **PS Lease** tab.

## Files

| File | Purpose |
|---|---|
| `index.html` | The portal (dashboard + tabs + search) |
| `data.json` | Generated data the portal reads |
| `data/` | Drop the weekly xlsx files here |
| `scripts/convert.py` | Converts xlsx → data.json |
| `.github/workflows/update-data.yml` | Auto-runs the converter on every push |

## One-time hosting setup (GitHub Pages)

1. Create a **private-name but public repo** on github.com (e.g. `arrears-portal`). *Note: GitHub Pages on a free account is publicly reachable via the URL — the passcode gate is the access control.*
2. Push this whole folder to the repo:
   ```
   git init
   git add .
   git commit -m "Arrears portal"
   git branch -M main
   git remote add origin https://github.com/YOUR-USER/arrears-portal.git
   git push -u origin main
   ```
3. In the repo: **Settings → Pages → Source: Deploy from a branch → Branch: main / (root) → Save**.
4. After a minute the portal is live at `https://YOUR-USER.github.io/arrears-portal/`. Share this link + passcode with staff.

## Weekly update

1. Drop the new files into `data/` (delete or keep old ones — the converter always picks the **newest** file of each type by the date in the filename, e.g. `Del Report (20-07-2026).xlsx`, `PS Lease Del report -20-07-2026.xlsx`).
2. Push:
   ```
   git add data/
   git commit -m "Weekly update"
   git push
   ```
3. The GitHub Action regenerates `data.json` automatically; the portal updates within ~2 minutes. (You can also update via the GitHub website: repo → `data/` folder → **Add file → Upload files**.)

To run the converter locally instead: `pip install openpyxl` then `python scripts/convert.py`.

## Passcode

Current passcode: `PSave2026`

To change it: generate the SHA-256 hash of the new passcode (e.g. at any "sha256 online" tool, or `echo -n "NewPass" | sha256sum`), then replace the `PASS_HASH` value near the top of the `<script>` in `index.html`.

> The passcode is a deterrent, not strong security — the data itself is in `data.json` in the repo. Don't share the repo/portal URL outside staff.
