# Aster DEX — Transaction Reader

## Account
- Main wallet (ASTER_USER): `0xa916b71eebcc78f98032066B833D6EA9546F1769`
- `0x2a15651864C634555b2177F458F53203e725739b` is an API wallet (@aster-desktop, has trade permission) — not the main wallet; do not use it here.

## Rules
- **READ-ONLY.** This project only fetches data from Aster DEX.
- Only call `GET` endpoints (trades, income, orders history, account/balance).
- Never place, modify, or cancel orders; never transfer, withdraw, or change leverage/margin.
- Never add `POST`/`PUT`/`DELETE` request code. `aster_tx.py` and `gas/Code.gs` enforce GET-only.
- V3 auth: `ASTER_USER`, `ASTER_SIGNER`, `ASTER_SIGNER_PRIVATE_KEY` from env vars or `.env` — never commit them.
  `ASTER_SIGNER_PRIVATE_KEY` is the API wallet (agent) key, never the main wallet key.

## Google Apps Script (`gas/Code.gs`)
- Read credentials only from Script Properties (`ASTER_USER`, `ASTER_SIGNER`, `ASTER_SIGNER_PRIVATE_KEY`); never hardcode them.
- Logging goes to `console` (Executions) and `toast` only — never write logs to a sheet.
- Never log or display the private key; mask addresses in logs.

## Language
- Always ask and answer in Thai (ภาษาไทย). Code, identifiers, and commit messages may stay in English.
