"""Read-only fetcher for Aster DEX futures history: orders, trades, transactions (income).

Usage:
    pip install eth-account
    export ASTER_USER=0x... ASTER_SIGNER=0x... ASTER_SIGNER_PRIVATE_KEY=0x...
    python3 aster_tx.py orders --days 7
    python3 aster_tx.py positions
    python3 aster_tx.py trades --days 7
    python3 aster_tx.py income --days 30 --csv income.csv
"""
import argparse, csv, json, os, sys, time, urllib.parse, urllib.request

BASE_URL = "https://fapi.asterdex.com"

# EIP-712 envelope from Aster V3 docs; the urlencoded query string is signed as `msg`.
TYPED_DATA = {
    "types": {
        "EIP712Domain": [
            {"name": "name", "type": "string"},
            {"name": "version", "type": "string"},
            {"name": "chainId", "type": "uint256"},
            {"name": "verifyingContract", "type": "address"},
        ],
        "Message": [{"name": "msg", "type": "string"}],
    },
    "primaryType": "Message",
    "domain": {"name": "AsterSignTransaction", "version": "1", "chainId": 1666,
               "verifyingContract": "0x0000000000000000000000000000000000000000"},
}

_last_nonce = 0


def next_nonce():
    """Microsecond timestamp, strictly increasing (server rejects reused nonces)."""
    global _last_nonce
    _last_nonce = max(int(time.time() * 1_000_000), _last_nonce + 1)
    return _last_nonce


def sign(query, private_key):
    from eth_account import Account
    try:
        from eth_account.messages import encode_typed_data
    except ImportError:  # eth-account < 0.10
        from eth_account.messages import encode_structured_data as encode_typed_data
    msg = encode_typed_data(full_message={**TYPED_DATA, "message": {"msg": query}})
    sig = Account.sign_message(msg, private_key=private_key).signature.hex()
    return sig if sig.startswith("0x") else "0x" + sig


def load_env(path=".env"):
    if os.path.exists(path):
        for line in open(path):
            if "=" in line and not line.lstrip().startswith("#"):
                k, v = line.strip().split("=", 1)
                os.environ.setdefault(k, v)


def signed_get(path, params):
    """GET-only V3 signed request. There is intentionally no POST/PUT/DELETE helper."""
    user, signer, pk = (os.environ.get(k) for k in
                        ("ASTER_USER", "ASTER_SIGNER", "ASTER_SIGNER_PRIVATE_KEY"))
    if not user or not signer or not pk:
        sys.exit("Missing ASTER_USER / ASTER_SIGNER / ASTER_SIGNER_PRIVATE_KEY")
    params = {k: v for k, v in params.items() if v is not None}
    params.update(user=user, signer=signer, nonce=next_nonce())
    query = urllib.parse.urlencode(params)
    req = urllib.request.Request(f"{BASE_URL}{path}?{query}&signature={sign(query, pk)}",
                                 method="GET", headers={"User-Agent": "aster-tx-reader/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        sys.exit(f"HTTP {e.code}: {e.read().decode()}")


def fetch_range(path, params, start_ms, end_ms):
    """One window; if it hits the 1000-row cap, split it in half so nothing is dropped."""
    rows = signed_get(path, {**params, "startTime": start_ms, "endTime": end_ms, "limit": 1000})
    if len(rows) < 1000 or end_ms - start_ms < 1000:
        return rows
    mid = (start_ms + end_ms) // 2
    return fetch_range(path, params, start_ms, mid) + fetch_range(path, params, mid + 1, end_ms)


def paginate(path, params, start_ms, end_ms, window_ms):
    """API limits each query to a time window, so walk the range in chunks."""
    rows, t = [], start_ms
    while t < end_ms:
        chunk_end = min(t + window_ms, end_ms)
        rows += fetch_range(path, params, t, chunk_end)
        t = chunk_end + 1
    return rows


def main():
    load_env()
    p = argparse.ArgumentParser()
    p.add_argument("kind", choices=["orders", "trades", "income", "positions"])
    p.add_argument("--symbol", help="filter to one symbol, e.g. BTCUSDT (default: all)")
    p.add_argument("--days", type=int, default=7)
    p.add_argument("--csv", help="write results to this CSV file")
    a = p.parse_args()

    end = int(time.time() * 1000)
    start = end - a.days * 86_400_000
    week = 7 * 86_400_000 - 1
    if a.kind == "positions":
        # Snapshot, not history: only open positions (API also returns every symbol with positionAmt 0).
        rows = [r for r in signed_get("/fapi/v3/positionRisk", {"symbol": a.symbol})
                if float(r["positionAmt"]) != 0]
    elif a.kind == "orders":
        rows = paginate("/fapi/v3/allOrders", {"symbol": a.symbol}, start, end, week)
    elif a.kind == "trades":
        if a.symbol:
            symbols = [a.symbol]
        else:
            # userTrades needs a symbol: find every symbol touched in the range via orders + income.
            seen = paginate("/fapi/v3/allOrders", {}, start, end, week) + \
                   paginate("/fapi/v3/income", {}, start, end, week)
            symbols = sorted({r["symbol"] for r in seen if r.get("symbol")})
            print(f"Symbols: {', '.join(symbols) or '(none)'}", file=sys.stderr)
        rows = [r for sym in symbols
                for r in paginate("/fapi/v3/userTrades", {"symbol": sym}, start, end, week)]
        rows.sort(key=lambda r: r["time"])
    else:
        rows = paginate("/fapi/v3/income", {"symbol": a.symbol}, start, end, week)

    if a.csv and rows:
        with open(a.csv, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys())
            w.writeheader(); w.writerows(rows)
        print(f"Wrote {len(rows)} rows to {a.csv}")
    else:
        print(json.dumps(rows, indent=2))


if __name__ == "__main__":
    main()
