# PR/PO Proxy — deploy guide

Small Azure Function (Node 22) that reads D365 F&O OData and returns dashboard JSON with CORS.
Endpoints: `GET /api/pr` (working), `GET /api/po` (stub, added after PR is validated).

## Prerequisites (already done)
- Function App `pr-po-dashboard-proxy` (Flex Consumption, Node 22) created.
- App settings set on the Function App: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `FO_RESOURCE`, `ALLOWED_ORIGIN`.
- The Entra app (CLIENT_ID) is registered in F&O under System administration → Setup → Azure Active Directory applications.

## Deploy via GitHub Actions (no local tools — matches your GitHub Desktop flow)
1. Make **this folder** (`pr-po-proxy`) its own GitHub repo:
   - GitHub Desktop → File → **Add local repository** → select this `pr-po-proxy` folder → **create a repository** → **Publish repository** (private).
   - Important: `host.json` must sit at the **repo root** (it does here). Do NOT put this inside the dashboard repo.
2. Azure portal → Function App `pr-po-dashboard-proxy` → **Deployment** → **Deployment Center**:
   - Source: **GitHub** → authorize → pick your org, the `pr-po-proxy` repo, branch `main`.
   - Build provider: **GitHub Actions** → **Save**.
   - Azure adds a workflow file and runs it. Watch the **Actions** tab in the repo until it's green (~2–3 min).
3. Leave the Function App's **CORS** blade **empty** — this code already sets the CORS header. (Setting both can cause duplicate-header errors.)

## Get the URL the dashboard will call
Function App → **Functions** → `pr` → **Get Function URL**. It looks like:
`https://pr-po-dashboard-proxy.azurewebsites.net/api/pr?code=XXXXXXXX`
That whole URL (with `?code=`) is what the dashboard fetches. Send it to me and I'll wire the dashboard.

## Test
Open the URL above in a browser:
- ✅ JSON `{ type:"pr", count: …, rows:[…] }` → working.
- ❌ `{ "error": "token 401 …" }` → the Entra app isn't granted in F&O (recheck the AAD applications step) or a wrong secret.
- ❌ `{ "error": "odata 401 …" }` → token OK but the user account tied to the app lacks read access to purchase requisitions.

## After the first successful response — send me the JSON
Two fields are deliberately left to finalize from real data (so we get them right once):
- `stepElementId` on every row → I complete `stepMap.json` (GUID → readable Step name), you redeploy.
- `ledgerDimensionRaw` (+ `location`) → I finalize `department` / `location` / `contract` parsing.

Then PO gets added the same way.
