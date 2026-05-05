from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable

from flask import jsonify, request, session


def register_api_routes(
    app,
    *,
    login_required: Callable,
    load_stock_list: Callable[[], list[dict[str, Any]]],
    search_stocks: Callable[[str], list[dict[str, Any]]],
    load_user_state: Callable[[str | None], dict[str, Any] | None],
    save_user_state: Callable[[str, dict[str, Any]], None],
    clean_requested_fields: Callable[[Any], set[str]],
    fetch_stock_snapshot: Callable[[str, set[str]], dict[str, str]],
    failed_snapshot: Callable[[set[str]], dict[str, str]],
    max_update_workers: int,
) -> None:
    @app.get("/api/stocks")
    @login_required
    def list_stocks():
        return jsonify({"stocks": load_stock_list()})

    @app.get("/api/stocks/search")
    @login_required
    def search_stock_list():
        return jsonify({"stocks": search_stocks(request.args.get("q", ""))})

    @app.get("/api/user-data")
    @login_required
    def get_user_data():
        return jsonify({"state": load_user_state(session.get("username"))})

    @app.put("/api/user-data")
    @login_required
    def put_user_data():
        payload = request.get_json(silent=True) or {}
        state_payload = payload.get("state", payload)

        if not isinstance(state_payload, dict):
            return jsonify({"error": "state must be an object"}), 400

        save_user_state(session["username"], state_payload)
        return jsonify({"ok": True})

    @app.post("/api/stocks/update")
    @login_required
    def update_stocks():
        data = request.get_json(silent=True) or {}
        tickers = data.get("tickers", [])
        requested_fields = clean_requested_fields(data.get("fields"))

        if not isinstance(tickers, list):
            return jsonify({"error": "tickers must be a list"}), 400

        clean_tickers = []
        seen_tickers = set()
        for ticker in tickers:
            if not isinstance(ticker, str):
                continue

            clean_ticker = ticker.strip()
            if clean_ticker and clean_ticker not in seen_tickers:
                clean_tickers.append(clean_ticker)
                seen_tickers.add(clean_ticker)

        results: dict[str, dict[str, str]] = {}
        if clean_tickers and requested_fields:
            worker_count = min(max_update_workers, len(clean_tickers))
            with ThreadPoolExecutor(max_workers=worker_count) as executor:
                futures = {
                    executor.submit(fetch_stock_snapshot, ticker, requested_fields): ticker
                    for ticker in clean_tickers
                }
                for future in as_completed(futures):
                    ticker = futures[future]
                    try:
                        results[ticker] = future.result()
                    except Exception:
                        results[ticker] = failed_snapshot(requested_fields)

        return jsonify({"stocks": results})
