from __future__ import annotations

import json
import math
import os
import time
import unicodedata
from ipaddress import ip_address
from functools import wraps
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Callable
from urllib.error import URLError
from urllib.request import Request, urlopen

from flask import Flask, jsonify, redirect, render_template, request, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

print("choice_app.py is running ...", flush=True)

app = Flask(__name__)
app.secret_key = os.environ.get("CHOICE_SECRET_KEY", "choice-dev-secret-change-me")
app.config["TEMPLATES_AUTO_RELOAD"] = True
MAX_UPDATE_WORKERS = 8
STOCK_LIST_PATH = Path(__file__).with_name("stock_list.json")
USER_STORE_PATH = Path(__file__).with_name("users.json")
DEFAULT_STOCK_FIELDS = {
    "price": "-",
    "change_amount": "-",
    "change": "-",
    "volume": "-",
    "close": "-",
    "open": "-",
    "high": "-",
    "low": "-",
    "per": "-",
    "roic": "-",
    "operating_income_growth": "-",
    "market_cap": "-",
    "performance_1d": "-",
    "performance_1w": "-",
    "performance_1m": "-",
    "performance_ytd": "-",
    "performance_1y": "-",
    "performance_3y": "-",
    "performance_5y": "-",
}
STATIC_FIELDS = {"name", "ticker", "category", "market"}
QUOTE_FIELDS = {"price", "change_amount", "change", "volume", "close", "open", "high", "low"}
FUNDAMENTAL_FIELDS = {"per", "roic", "operating_income_growth", "market_cap"}
PERFORMANCE_FIELDS = {
    "performance_1d",
    "performance_1w",
    "performance_1m",
    "performance_ytd",
    "performance_1y",
    "performance_3y",
    "performance_5y",
}
FETCHABLE_FIELDS = QUOTE_FIELDS | FUNDAMENTAL_FIELDS | PERFORMANCE_FIELDS


def load_users() -> dict[str, dict[str, str]]:
    if not USER_STORE_PATH.exists():
        return {}

    try:
        with USER_STORE_PATH.open(encoding="utf-8") as file:
            payload = json.load(file)
    except (OSError, json.JSONDecodeError):
        return {}

    return payload if isinstance(payload, dict) else {}


def save_users(users: dict[str, dict[str, str]]) -> None:
    with USER_STORE_PATH.open("w", encoding="utf-8") as file:
        json.dump(users, file, ensure_ascii=False, indent=2)


def login_required(view: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(view)
    def wrapped(*args: Any, **kwargs: Any) -> Any:
        if session.get("username"):
            return view(*args, **kwargs)

        if request.path.startswith("/api/"):
            return jsonify({"error": "login required"}), 401

        return redirect(url_for("login", next=request.path))

    return wrapped


@app.after_request
def disable_browser_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


def load_stock_list() -> list[dict[str, Any]]:
    with STOCK_LIST_PATH.open(encoding="utf-8") as file:
        payload = json.load(file)

    stocks: list[dict[str, Any]] = []
    for category in payload.get("categories", []):
        category_name = category.get("name", "")
        for item in category.get("items", []):
            stocks.append({
                "category": category_name,
                **DEFAULT_STOCK_FIELDS,
                **item,
            })

    return stocks


def normalize_search_text(value: Any) -> str:
    return unicodedata.normalize("NFC", str(value or "")).strip().casefold()


def search_stocks(query: str) -> list[dict[str, Any]]:
    normalized_query = normalize_search_text(query)
    if not normalized_query:
        return []

    matches: list[dict[str, Any]] = []
    for stock in load_stock_list():
        searchable_text = " ".join(
            normalize_search_text(stock.get(field))
            for field in ("name", "ticker", "market", "category")
        )
        if normalized_query in searchable_text:
            matches.append(stock)

    return matches


def _clean_requested_fields(fields: Any) -> set[str]:
    if not isinstance(fields, list):
        return set(FETCHABLE_FIELDS)

    return {
        field.strip()
        for field in fields
        if isinstance(field, str) and field.strip() in FETCHABLE_FIELDS
    }


def _filter_snapshot(snapshot: dict[str, str], requested_fields: set[str]) -> dict[str, str]:
    return {
        field: snapshot.get(field, "-")
        for field in requested_fields
    }


def _failed_snapshot(requested_fields: set[str] | None = None) -> dict[str, str]:
    fields = requested_fields or FETCHABLE_FIELDS
    snapshot = {field: "-" for field in fields}
    if "price" in snapshot:
        snapshot["price"] = "\uc870\ud68c \uc2e4\ud328"
    return snapshot


def _format_number(value: Any, decimals: int = 2) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "-"

    if math.isnan(number):
        return "-"

    return f"{number:,.{decimals}f}"


def _format_volume(value: Any) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return "-"


def _format_market_cap(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "-"

    if math.isnan(number):
        return "-"

    units = [
        (1_000_000_000_000, "T"),
        (1_000_000_000, "B"),
        (1_000_000, "M"),
    ]
    for divisor, suffix in units:
        if abs(number) >= divisor:
            return f"{number / divisor:,.2f}{suffix}"

    return f"{number:,.0f}"


def _format_ratio(value: Any, decimals: int = 2) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "-"

    if math.isnan(number):
        return "-"

    return f"{number:,.{decimals}f}"


def _format_percent_value(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "-"

    if math.isnan(number):
        return "-"

    if abs(number) <= 1:
        number *= 100

    return f"{number:+.2f}%"


def _format_change_amount(current_price: Any, previous_close: Any) -> str:
    try:
        change_amount = float(current_price) - float(previous_close)
    except (TypeError, ValueError):
        return "-"

    return f"{change_amount:+,.2f}"


def _format_rate(current_value: Any, base_value: Any) -> str:
    try:
        current = float(current_value)
        base = float(base_value)
        if base == 0:
            return "-"
        return f"{((current - base) / base) * 100:+.2f}%"
    except (TypeError, ValueError):
        return "-"


def _close_from_offset(history: Any, offset: int) -> Any:
    if len(history) <= offset:
        return None

    return history.iloc[-1 - offset].get("Close")


def _value_from_offset(values: list[Any], offset: int) -> Any:
    if len(values) <= offset:
        return None

    return values[-1 - offset]


def _value_at_index(values: list[Any], index: int) -> Any:
    if index < 0 or index >= len(values):
        return None

    return values[index]


def _ytd_base_close(history: Any) -> Any:
    if history.empty:
        return None

    latest_index = history.index[-1]
    year_start = latest_index.replace(month=1, day=1)
    ytd_history = history[history.index >= year_start]
    if ytd_history.empty:
        return history.iloc[0].get("Close")

    return ytd_history.iloc[0].get("Close")


def _market_cap_from_ticker(ticker_info: Any) -> str:
    try:
        return _format_market_cap(ticker_info.fast_info.get("market_cap"))
    except Exception:
        return "-"


def _first_numeric_value(payload: dict[str, Any], keys: tuple[str, ...]) -> Any:
    for key in keys:
        value = payload.get(key)
        if value is None:
            continue

        try:
            number = float(value)
        except (TypeError, ValueError):
            continue

        if not math.isnan(number):
            return number

    return None


def _statement_value(statement: Any, names: tuple[str, ...], column_index: int = 0) -> Any:
    try:
        if statement is None or statement.empty:
            return None

        normalized_rows = {str(index).lower().replace(" ", "").replace("_", ""): index for index in statement.index}
        for name in names:
            row = normalized_rows.get(name.lower().replace(" ", "").replace("_", ""))
            if row is None or len(statement.columns) <= column_index:
                continue

            value = statement.loc[row].iloc[column_index]
            if value is not None and not (isinstance(value, float) and math.isnan(value)):
                return value
    except Exception:
        return None

    return None


def _safe_info(ticker_info: Any) -> dict[str, Any]:
    try:
        info = ticker_info.get_info()
        return info if isinstance(info, dict) else {}
    except Exception:
        try:
            info = ticker_info.info
            return info if isinstance(info, dict) else {}
        except Exception:
            return {}


def _format_per_from_info(info: dict[str, Any]) -> str:
    return _format_ratio(_first_numeric_value(info, ("trailingPE", "forwardPE", "trailingPegRatio")))


def _format_operating_income_growth(ticker_info: Any, info: dict[str, Any]) -> str:
    info_value = _first_numeric_value(info, ("earningsGrowth", "revenueGrowth"))
    if info_value is not None:
        return _format_percent_value(info_value)

    try:
        financials = ticker_info.financials
    except Exception:
        financials = None

    latest = _statement_value(financials, ("Operating Income", "OperatingIncome"), 0)
    previous = _statement_value(financials, ("Operating Income", "OperatingIncome"), 1)
    if latest is None or previous in (None, 0):
        return "-"

    return _format_rate(latest, previous)


def _format_roic(ticker_info: Any, info: dict[str, Any]) -> str:
    info_value = _first_numeric_value(info, ("returnOnCapital", "returnOnInvestedCapital"))
    if info_value is not None:
        return _format_percent_value(info_value)

    try:
        financials = ticker_info.financials
    except Exception:
        financials = None

    try:
        balance_sheet = ticker_info.balance_sheet
    except Exception:
        balance_sheet = None

    operating_income = _statement_value(financials, ("Operating Income", "OperatingIncome"))
    tax_rate = _first_numeric_value(info, ("effectiveTaxRate",))
    if tax_rate is None or tax_rate < 0 or tax_rate > 1:
        tax_rate = 0.21

    total_debt = _statement_value(balance_sheet, ("Total Debt", "TotalDebt"))
    stockholder_equity = _statement_value(balance_sheet, ("Stockholders Equity", "Total Stockholder Equity", "StockholdersEquity"))
    cash = _statement_value(balance_sheet, ("Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments", "CashAndCashEquivalents"))

    if operating_income is None or total_debt is None or stockholder_equity is None:
        fallback = _first_numeric_value(info, ("returnOnAssets", "returnOnEquity"))
        return _format_percent_value(fallback)

    invested_capital = float(total_debt) + float(stockholder_equity) - float(cash or 0)
    if invested_capital == 0:
        return "-"

    roic = (float(operating_income) * (1 - tax_rate)) / invested_capital
    return _format_percent_value(roic)


def _fundamentals_from_ticker(ticker_info: Any, requested_fields: set[str]) -> dict[str, str]:
    if not requested_fields & FUNDAMENTAL_FIELDS:
        return {}

    info = _safe_info(ticker_info)
    snapshot: dict[str, str] = {}

    if "per" in requested_fields:
        snapshot["per"] = _format_per_from_info(info)
    if "roic" in requested_fields:
        snapshot["roic"] = _format_roic(ticker_info, info)
    if "operating_income_growth" in requested_fields:
        snapshot["operating_income_growth"] = _format_operating_income_growth(ticker_info, info)
    if "market_cap" in requested_fields:
        market_cap = _market_cap_from_ticker(ticker_info)
        if market_cap == "-":
            market_cap = _format_market_cap(_first_numeric_value(info, ("marketCap",)))
        snapshot["market_cap"] = market_cap

    return snapshot


def _reported_raw_values(item: dict[str, Any], key: str) -> list[float]:
    values = []
    for entry in item.get(key, []):
        try:
            value = entry.get("reportedValue", {}).get("raw")
            number = float(value)
        except (AttributeError, TypeError, ValueError):
            continue

        if not math.isnan(number):
            values.append(number)

    return values


def _latest_timeseries_values(payload: dict[str, Any]) -> dict[str, list[float]]:
    values: dict[str, list[float]] = {}
    for item in payload.get("timeseries", {}).get("result", []):
        if not isinstance(item, dict):
            continue

        for key in item:
            if key in {"meta", "timestamp"}:
                continue

            raw_values = _reported_raw_values(item, key)
            if raw_values:
                values[key] = raw_values

    return values


def _snapshot_from_yahoo_fundamentals(ticker: str, requested_fields: set[str]) -> dict[str, str] | None:
    if not requested_fields & FUNDAMENTAL_FIELDS:
        return {}

    now = int(time.time())
    six_years_ago = now - (60 * 60 * 24 * 365 * 6)
    types = [
        "trailingPeRatio",
        "trailingMarketCap",
        "annualOperatingIncome",
        "annualPretaxIncome",
        "annualTaxProvision",
        "annualInvestedCapital",
        "annualTotalDebt",
        "annualStockholdersEquity",
        "annualCashAndCashEquivalents",
    ]
    url = (
        "https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/"
        f"{ticker}?symbol={ticker}&type={','.join(types)}&period1={six_years_ago}&period2={now}"
    )
    req = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 Choice/1.0",
            "Accept": "application/json",
        },
    )

    try:
        with urlopen(req, timeout=10) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (OSError, URLError, json.JSONDecodeError):
        return None

    values = _latest_timeseries_values(payload)
    snapshot: dict[str, str] = {}

    if "per" in requested_fields:
        per_values = values.get("trailingPeRatio", [])
        snapshot["per"] = _format_ratio(per_values[-1] if per_values else None)

    if "market_cap" in requested_fields:
        market_cap_values = values.get("trailingMarketCap", [])
        snapshot["market_cap"] = _format_market_cap(market_cap_values[-1] if market_cap_values else None)

    operating_income_values = values.get("annualOperatingIncome", [])
    if "operating_income_growth" in requested_fields:
        if len(operating_income_values) >= 2:
            snapshot["operating_income_growth"] = _format_rate(operating_income_values[-1], operating_income_values[-2])
        else:
            snapshot["operating_income_growth"] = "-"

    if "roic" in requested_fields:
        operating_income = operating_income_values[-1] if operating_income_values else None
        invested_capital_values = values.get("annualInvestedCapital", [])
        invested_capital = invested_capital_values[-1] if invested_capital_values else None

        if invested_capital is None:
            total_debt_values = values.get("annualTotalDebt", [])
            equity_values = values.get("annualStockholdersEquity", [])
            cash_values = values.get("annualCashAndCashEquivalents", [])
            if total_debt_values and equity_values:
                invested_capital = total_debt_values[-1] + equity_values[-1] - (cash_values[-1] if cash_values else 0)

        pretax_values = values.get("annualPretaxIncome", [])
        tax_values = values.get("annualTaxProvision", [])
        tax_rate = 0.21
        if pretax_values and tax_values and pretax_values[-1] not in (0, None):
            tax_rate = max(0, min(1, tax_values[-1] / pretax_values[-1]))

        if operating_income is not None and invested_capital not in (None, 0):
            snapshot["roic"] = _format_percent_value((operating_income * (1 - tax_rate)) / invested_capital)
        else:
            snapshot["roic"] = "-"

    return _filter_snapshot(snapshot, requested_fields)


def _snapshot_from_yfinance(ticker: str, requested_fields: set[str]) -> dict[str, str] | None:
    """Use yfinance when it is installed."""
    try:
        import yfinance as yf
    except Exception:
        return None

    try:
        ticker_info = yf.Ticker(ticker)
        needs_quote = bool(requested_fields & QUOTE_FIELDS)
        needs_performance = bool(requested_fields & PERFORMANCE_FIELDS)
        needs_fundamentals = bool(requested_fields & FUNDAMENTAL_FIELDS)
        period = "5y" if needs_performance else "5d"

        history = ticker_info.history(period=period, interval="1d", auto_adjust=False) if needs_quote or needs_performance else None
        fundamentals = _fundamentals_from_ticker(ticker_info, requested_fields) if needs_fundamentals else {}

        if history is None:
            if fundamentals:
                return _filter_snapshot(fundamentals, requested_fields)
            return {}

        if history.empty:
            return None

        latest = history.iloc[-1]
        previous = history.iloc[-2] if len(history) >= 2 else latest
        current_price = latest.get("Close")
        previous_close = previous.get("Close")
        volume = latest.get("Volume")

        snapshot = _build_snapshot(
            current_price=current_price,
            previous_close=previous_close,
            volume=volume,
            open_price=latest.get("Open"),
            high_price=latest.get("High"),
            low_price=latest.get("Low"),
            close_price=latest.get("Close"),
            market_cap=fundamentals.get("market_cap", "-"),
            performance={
                "performance_1d": _format_rate(current_price, _close_from_offset(history, 1)),
                "performance_1w": _format_rate(current_price, _close_from_offset(history, 5)),
                "performance_1m": _format_rate(current_price, _close_from_offset(history, 21)),
                "performance_ytd": _format_rate(current_price, _ytd_base_close(history)),
                "performance_1y": _format_rate(current_price, _close_from_offset(history, 252)),
                "performance_3y": _format_rate(current_price, _close_from_offset(history, 756)),
                "performance_5y": _format_rate(current_price, _close_from_offset(history, 1260)),
            },
        )
        snapshot.update(fundamentals)
        return _filter_snapshot(snapshot, requested_fields)
    except Exception:
        return None


def _snapshot_from_yahoo_chart(ticker: str, requested_fields: set[str]) -> dict[str, str] | None:
    """Fallback that does not require extra packages."""
    needs_quote = bool(requested_fields & QUOTE_FIELDS)
    needs_performance = bool(requested_fields & PERFORMANCE_FIELDS)
    if not needs_quote and not needs_performance:
        return {}

    period = "5y" if needs_performance else "5d"
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?range={period}&interval=1d"
    req = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 Choice/1.0",
            "Accept": "application/json",
        },
    )

    try:
        with urlopen(req, timeout=8) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (OSError, URLError, json.JSONDecodeError):
        return None

    try:
        result = payload["chart"]["result"][0]
        quote = result["indicators"]["quote"][0]
        timestamps = result.get("timestamp", [])
        raw_closes = quote.get("close", [])
        closes = [
            (index, timestamps[index], value)
            for index, value in enumerate(raw_closes)
            if value is not None and index < len(timestamps)
        ]

        if not closes:
            return None

        latest_index, latest_timestamp, current_price = closes[-1]
        previous_close = closes[-2][2] if len(closes) >= 2 else result.get("meta", {}).get("chartPreviousClose", current_price)
        volume = _value_at_index(quote.get("volume", []), latest_index)

        close_values = [value for _, _, value in closes]
        latest_year = datetime.fromtimestamp(latest_timestamp).year
        ytd_base = next(
            (value for _, timestamp, value in closes if datetime.fromtimestamp(timestamp).year == latest_year),
            close_values[0],
        )

        snapshot = _build_snapshot(
            current_price=current_price,
            previous_close=previous_close,
            volume=volume,
            open_price=_value_at_index(quote.get("open", []), latest_index),
            high_price=_value_at_index(quote.get("high", []), latest_index),
            low_price=_value_at_index(quote.get("low", []), latest_index),
            close_price=current_price,
            performance={
                "performance_1d": _format_rate(current_price, _value_from_offset(close_values, 1)),
                "performance_1w": _format_rate(current_price, _value_from_offset(close_values, 5)),
                "performance_1m": _format_rate(current_price, _value_from_offset(close_values, 21)),
                "performance_ytd": _format_rate(current_price, ytd_base),
                "performance_1y": _format_rate(current_price, _value_from_offset(close_values, 252)),
                "performance_3y": _format_rate(current_price, _value_from_offset(close_values, 756)),
                "performance_5y": _format_rate(current_price, _value_from_offset(close_values, 1260)),
            },
        )
        return _filter_snapshot(snapshot, requested_fields)
    except (KeyError, IndexError, TypeError):
        return None


def _build_snapshot(
    current_price: Any,
    previous_close: Any,
    volume: Any,
    open_price: Any = None,
    high_price: Any = None,
    low_price: Any = None,
    close_price: Any = None,
    market_cap: str = "-",
    performance: dict[str, str] | None = None,
) -> dict[str, str]:
    return {
        "price": _format_number(current_price),
        "change_amount": _format_change_amount(current_price, previous_close),
        "change": _format_rate(current_price, previous_close),
        "volume": _format_volume(volume),
        "close": _format_number(close_price if close_price is not None else current_price),
        "open": _format_number(open_price),
        "high": _format_number(high_price),
        "low": _format_number(low_price),
        "market_cap": market_cap,
        **(performance or {}),
    }


def _merge_snapshots(*snapshots: dict[str, str] | None) -> dict[str, str] | None:
    merged: dict[str, str] = {}
    saw_snapshot = False

    for snapshot in snapshots:
        if snapshot is None:
            continue

        saw_snapshot = True
        for key, value in snapshot.items():
            if value != "-" or key not in merged:
                merged[key] = value

    return merged if saw_snapshot else None


def _complete_snapshot(snapshot: dict[str, str] | None, requested_fields: set[str]) -> dict[str, str] | None:
    if snapshot is None:
        return None

    return {
        field: snapshot.get(field, "-") or "-"
        for field in requested_fields
    }


def fetch_stock_snapshot(ticker: str, requested_fields: set[str]) -> dict[str, str]:
    if not requested_fields:
        return {}

    yahoo_snapshot = _merge_snapshots(
        _snapshot_from_yahoo_chart(ticker, requested_fields),
        _snapshot_from_yahoo_fundamentals(ticker, requested_fields),
    )
    snapshot = _merge_snapshots(
        _snapshot_from_yfinance(ticker, requested_fields),
        yahoo_snapshot,
    )
    return _complete_snapshot(snapshot, requested_fields) if snapshot is not None else _failed_snapshot(requested_fields)


def get_page_title() -> str:
    host = request.host.split(":", 1)[0].lower()
    if host in {"localhost", "127.0.0.1", "::1"}:
        return "로컬 Choice"
    if host == "158.247.209.218":
        return "웹서버 Choice"

    try:
        parsed_host = ip_address(host)
    except ValueError:
        return "Choice"

    return "로컬 Choice" if parsed_host.is_private else "Choice"


@app.route("/")
@login_required
def index() -> str:
    return render_template(
        "index.html",
        stocks=load_stock_list(),
        username=session.get("username"),
        page_title=get_page_title(),
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("username"):
        return redirect(url_for("index"))

    error = ""
    mode = request.form.get("mode", "login")
    next_url = request.args.get("next") or url_for("index")

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        users = load_users()

        if not username or not password:
            error = "아이디와 비밀번호를 입력하세요."
        elif mode == "register":
            if username in users:
                error = "이미 존재하는 아이디입니다."
            else:
                users[username] = {"password_hash": generate_password_hash(password)}
                save_users(users)
                session["username"] = username
                return redirect(next_url)
        else:
            user = users.get(username)
            if not user or not check_password_hash(user.get("password_hash", ""), password):
                error = "아이디 또는 비밀번호가 올바르지 않습니다."
            else:
                session["username"] = username
                return redirect(next_url)

    return render_template("login.html", error=error, mode=mode, next_url=next_url, page_title=get_page_title())


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/api/stocks")
@login_required
def list_stocks():
    return jsonify({"stocks": load_stock_list()})


@app.get("/api/stocks/search")
@login_required
def search_stock_list():
    return jsonify({"stocks": search_stocks(request.args.get("q", ""))})


@app.post("/api/stocks/update")
@login_required
def update_stocks():
    data = request.get_json(silent=True) or {}
    tickers = data.get("tickers", [])
    requested_fields = _clean_requested_fields(data.get("fields"))

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
        worker_count = min(MAX_UPDATE_WORKERS, len(clean_tickers))
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
                    results[ticker] = _failed_snapshot(requested_fields)

    return jsonify({"stocks": results})


if __name__ == "__main__":
    app.run(debug=True)
