# Brand Audit Cloudflare

Cloudflare-native replacement for the former hosted brand-audit app.

## What changed

- API keys are read from Cloudflare Worker secrets (`env.OPENAI_API_KEY`), not legacy process variables.
- The workspace focuses on four core audit modules and has no external search-runner dependency.
- The useful modules remain: Brand Messaging Framework, Pain Point Questions, Home Page FAQs, and AI Info Page.
- The frontend uses the dark graphite / green ICP Simulator visual language.
- Results are downloaded as Markdown or JSON directly from the browser.

## Local setup

```bash
npm install
cp .dev.vars.example .dev.vars
# Edit .dev.vars and add the real OpenAI key.
npm run dev
```

Open the local URL Wrangler prints. The app's status pill and `/api/health` endpoint confirm whether the Worker can see the secret.

## Cloudflare deployment

Authenticate Wrangler, then add the encrypted Worker secret:

```bash
npx wrangler login
npx wrangler secret put OPENAI_API_KEY
npm run deploy
```

The secret is intentionally not stored in `wrangler.jsonc`, `.env`, or any frontend file. When you create the GitHub repository, connect the repository to this Worker in Cloudflare or deploy from the project directory with Wrangler.

## API routes

- `GET /api/health` - safe configuration check; returns booleans only.
- `POST /api/run` - runs one of the four audit modules.
