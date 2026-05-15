let STOCKS = Array.isArray(window.CHOICE_STOCKS) ? window.CHOICE_STOCKS : [];
const STORAGE_KEY = "choice.state.v1";
const STOCK_PRICES_URL = "/public/stock_prices.json";
const COLUMN_SCHEMA_VERSION = 12;
const inlineChartCache = new Map();
const PRICE_LINK_OVERRIDES = new Map([
    ["^GSPC", "https://finviz.com/map?t=sec"],
    ["^IXIC", "https://finviz.com/map?t=sec_ndx"],
    ["^KS11", "https://markets.hankyung.com/marketmap/kospi"],
    ["^KQ11", "https://markets.hankyung.com/marketmap/kosdaq"],
]);
const YAHOO_SVG_CHART_TICKERS = new Set([
    "^VIX",
    "KRW=X",
    "DX-Y.NYB",
    "KR2Y",
    "KR10Y",
    "^TNX",
    "GC=F",
    "NG=F",
    "SI=F",
    "HG=F",
    "ZW=F",
    "CL=F",
    "PENTAGON_PIZZA_INDEX",
    "^DJI",
    "^SOX",
    "000001.SS",
    "^N225",
    "^HSI",
    "^FTSE",
    "^FCHI",
    "^GDAXI",
    "JPYKRW=X",
    "EURKRW=X",
    "YM=F",
    "NKD=F",
    "ES=F",
    "6E=F",
    "NQ=F",
    "FDAX.DE",
    "RTY=F",
    "VX=F",
    "KR_FUTURE_ENERGY_CHEMICAL",
    "KR_FUTURE_HEAVY_INDUSTRY",
    "KR_FUTURE_IT",
    "KR_FUTURE_HEALTHCARE",
    "KR_FUTURE_FINANCE",
    "KR_FUTURE_CONSUMER_STAPLES",
    "KR_FUTURE_CONSUMER_DISCRETIONARY",
    "KR_FUTURE_STEEL_MATERIALS",
    "KR_FUTURE_CONSTRUCTION",
    "KR_FUTURE_INDUSTRIALS",
]);

const COLUMN_DEFS = [
    { key: "name", label: "종목명", className: "stock-name", width: "var(--stock-name-column-width)" },
    { key: "ticker", label: "티커", width: "92px" },
    { key: "price", label: "현재가", width: "92px" },
    { key: "average_price", label: "평단가", className: "average-price-cell", width: "92px" },
    { key: "quantity", label: "수량", className: "average-price-cell", width: "76px" },
    { key: "holding_amount", label: "매입금액", width: "108px" },
    { key: "profit_rate", label: "수익%", width: "92px", valueClass: getChangeClass },
    { key: "profit_amount", label: "손익", width: "108px", valueClass: getChangeClass },
    { key: "change_amount", label: "변동", width: "96px", valueClass: getChangeClass },
    { key: "change", label: "오늘변동률", width: "104px", valueClass: getChangeClass },
    { key: "volume", label: "거래량", width: "112px" },
    { key: "close", label: "종가", width: "92px" },
    { key: "open", label: "시가", width: "92px" },
    { key: "high", label: "고가", width: "92px" },
    { key: "low", label: "저가", width: "92px" },
    { key: "per", label: "PER", width: "minmax(74px, 0.85fr)" },
    { key: "roic", label: "ROIC", width: "minmax(74px, 0.85fr)" },
    { key: "operating_income_growth", label: "영업이익증가율", width: "minmax(110px, 1fr)" },
    { key: "market_cap", label: "시가총액", width: "minmax(86px, 0.9fr)" },
    { key: "performance_1d", label: "1D", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "performance_1w", label: "1W", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "performance_1m", label: "1M", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "performance_ytd", label: "YTD", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "performance_1y", label: "1Y", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "performance_3y", label: "3Y", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "performance_5y", label: "5Y", className: "performance-cell", width: "minmax(48px, 0.75fr)", valueClass: getChangeClass },
    { key: "chart_d", label: "1일", className: "chart-cell", width: "220px", chartType: "stock", chartPeriod: "d" },
    { key: "chart_candle_d", label: "일봉", className: "chart-cell", width: "220px", chartType: "candle", chartPeriod: "d" },
    { key: "chart_candle_w", label: "주봉", className: "chart-cell", width: "220px", chartType: "candle", chartPeriod: "w" },
    { key: "chart_y", label: "1년", className: "chart-cell", width: "220px", chartType: "stock", chartPeriod: "y" },
    { key: "chart_y3", label: "3년", className: "chart-cell", width: "220px", chartType: "stock", chartPeriod: "y3" },
    { key: "chart_y10", label: "10년", className: "chart-cell", width: "220px", chartType: "stock", chartPeriod: "y10" },
];

const COLUMN_GROUPS = [
    {
        key: "price",
        label: "가격",
        columns: ["name", "ticker", "price", "average_price", "quantity", "holding_amount", "profit_rate", "profit_amount", "change", "volume", "close", "open", "high", "low"],
    },
    {
        key: "fundamental",
        label: "펀더멘털",
        columns: ["name", "per", "roic", "operating_income_growth", "market_cap"],
    },
    {
        key: "performance",
        label: "성과 상승률",
        columns: ["name", "performance_1d", "performance_1w", "performance_1m", "performance_1y", "performance_ytd", "performance_3y", "performance_5y"],
    },
    {
        key: "chart",
        label: "차트",
        columns: ["name", "ticker", "price", "chart_d", "chart_candle_d", "chart_candle_w", "chart_y", "chart_y3", "chart_y10"],
    },
];

const DEFAULT_VISIBLE_COLUMNS = COLUMN_DEFS.map((column) => column.key);
const DEFAULT_COLUMN_GROUP = COLUMN_GROUPS[0].key;
const FIXED_COLUMN_KEYS = new Set(["name"]);
const FETCHABLE_COLUMN_KEYS = new Set([
    "price",
    "change_amount",
    "change",
    "volume",
    "close",
    "open",
    "high",
    "low",
    "per",
    "roic",
    "operating_income_growth",
    "market_cap",
    "performance_1d",
    "performance_1w",
    "performance_1m",
    "performance_ytd",
    "performance_1y",
    "performance_3y",
    "performance_5y",
]);

const tabsEl = document.getElementById("tabs");
const addTabButton = document.getElementById("addTabButton");
const editTabButton = document.getElementById("editTabButton");
const stockInput = document.getElementById("stockInput");
const suggestionsEl = document.getElementById("suggestions");
const suggestionStatusEl = document.getElementById("suggestionStatus");
const topHorizontalScrollEl = document.getElementById("topHorizontalScroll");
const topHorizontalScrollInnerEl = document.getElementById("topHorizontalScrollInner");
const bottomHorizontalScrollEl = document.getElementById("bottomHorizontalScroll");
const bottomHorizontalScrollInnerEl = document.getElementById("bottomHorizontalScrollInner");
const stockTableScrollEl = document.getElementById("stockTableScroll");
const fixedColumnCurtainEl = document.getElementById("fixedColumnCurtain");
const tableHeadEl = document.getElementById("tableHead");
const stockListEl = document.getElementById("stockList");
const emptyStateEl = document.getElementById("emptyState");
const updateButton = document.getElementById("updateButton");
const layoutToggleButton = document.getElementById("layoutToggleButton");
const columnGroupTabsEl = document.getElementById("columnGroupTabs");
const columnSettingsButton = document.getElementById("columnSettingsButton");
const columnSettingsMenu = document.getElementById("columnSettingsMenu");
const contextMenu = document.getElementById("contextMenu");
const deleteTabButton = document.getElementById("deleteTabButton");
const renameTabButton = document.getElementById("renameTabButton");
const stockContextMenu = document.getElementById("stockContextMenu");
const writeMemoButton = document.getElementById("writeMemoButton");
const tabDialogOverlay = document.getElementById("tabDialogOverlay");
const tabDialogTitle = document.getElementById("tabDialogTitle");
const tabNameInput = document.getElementById("tabNameInput");
const tabNameError = document.getElementById("tabNameError");
const cancelTabDialogButton = document.getElementById("cancelTabDialogButton");
const confirmTabDialogButton = document.getElementById("confirmTabDialogButton");
const deleteTabDialogOverlay = document.getElementById("deleteTabDialogOverlay");
const cancelDeleteTabDialogButton = document.getElementById("cancelDeleteTabDialogButton");
const confirmDeleteTabDialogButton = document.getElementById("confirmDeleteTabDialogButton");
const memoDialogOverlay = document.getElementById("memoDialogOverlay");
const memoInput = document.getElementById("memoInput");
const memoError = document.getElementById("memoError");
const cancelMemoDialogButton = document.getElementById("cancelMemoDialogButton");
const confirmMemoDialogButton = document.getElementById("confirmMemoDialogButton");
const averagePriceDialogOverlay = document.getElementById("averagePriceDialogOverlay");
const averagePriceInput = document.getElementById("averagePriceInput");
const averagePriceError = document.getElementById("averagePriceError");
const cancelAveragePriceDialogButton = document.getElementById("cancelAveragePriceDialogButton");
const confirmAveragePriceDialogButton = document.getElementById("confirmAveragePriceDialogButton");
const memoTooltipEl = document.createElement("div");
memoTooltipEl.className = "memo-tooltip";
document.body.appendChild(memoTooltipEl);

let state = loadState();
let contextTabId = null;
let contextStockTicker = null;
let editingAveragePriceTicker = null;
let editingNumberField = "average_price";
let activeSuggestionIndex = -1;
let draggedStockRowId = null;
let dragHandleArmedStockRowId = null;
let draggedColumnKey = null;
let recentlyDraggedHeaderColumn = false;
let draggedTabId = null;
let resizingColumn = null;
let isUpdating = false;
let tabDialogMode = "create";
let editingTabId = null;
let pendingDeleteTabId = null;
let currentSuggestions = [];
let suggestionRequestId = 0;
let lastSuggestionQuery = "";
let pendingSuggestionPromise = null;
let layoutToggleStage = 0;
let saveStateTimer = null;
let entryImagePopoutTopZ = 2147483000;
const headerSortState = {
    tabId: null,
    groupKey: null,
    columnKey: null,
    direction: null,
    baseOrder: [],
};

function createId() {
    return `tab-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createStockRowId() {
    return `stock-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeSearchText(value) {
    return String(value || "").normalize("NFC").trim().toLowerCase();
}

function createDefaultTab(index = 1) {
    return {
        id: createId(),
        title: index === 1 ? "기본" : `new_${String(index - 1).padStart(2, "0")}`,
        stocks: [],
    };
}

function normalizeVisibleColumns(columns) {
    if (!Array.isArray(columns)) return [...DEFAULT_VISIBLE_COLUMNS];

    const validColumns = columns.filter((column) => DEFAULT_VISIBLE_COLUMNS.includes(column));
    return validColumns.length > 0 ? validColumns : [...DEFAULT_VISIBLE_COLUMNS];
}

function normalizeColumnGroup(groupKey) {
    return COLUMN_GROUPS.some((group) => group.key === groupKey) ? groupKey : DEFAULT_COLUMN_GROUP;
}

function createDefaultGroupVisibleColumns() {
    return Object.fromEntries(COLUMN_GROUPS.map((group) => [group.key, [...group.columns]]));
}

function createDefaultGroupColumnOrder() {
    return Object.fromEntries(COLUMN_GROUPS.map((group) => [group.key, [...group.columns]]));
}

function createDefaultGroupColumnWidths() {
    return Object.fromEntries(COLUMN_GROUPS.map((group) => [group.key, {}]));
}

function normalizeGroupVisibleColumns(groupColumns) {
    const defaults = createDefaultGroupVisibleColumns();
    if (!groupColumns || typeof groupColumns !== "object") return defaults;

    return Object.fromEntries(COLUMN_GROUPS.map((group) => {
        if (!Array.isArray(groupColumns[group.key])) return [group.key, defaults[group.key]];

        const validColumns = groupColumns[group.key].filter((key) => group.columns.includes(key));
        return [group.key, validColumns];
    }));
}

function normalizeGroupColumnOrder(groupOrder) {
    const defaults = createDefaultGroupColumnOrder();
    if (!groupOrder || typeof groupOrder !== "object") return defaults;

    return Object.fromEntries(COLUMN_GROUPS.map((group) => {
        const saved = Array.isArray(groupOrder[group.key]) ? groupOrder[group.key] : [];
        const ordered = saved.filter((key) => group.columns.includes(key));
        const missing = group.columns.filter((key) => !ordered.includes(key));
        return [group.key, withFixedColumns([...ordered, ...missing])];
    }));
}

function normalizeGroupColumnWidths(widths) {
    const defaults = createDefaultGroupColumnWidths();
    if (!widths || typeof widths !== "object") return defaults;

    return Object.fromEntries(COLUMN_GROUPS.map((group) => {
        const groupWidths = widths[group.key];
        if (!groupWidths || typeof groupWidths !== "object" || Array.isArray(groupWidths)) {
            return [group.key, {}];
        }

        return [
            group.key,
            Object.fromEntries(
                Object.entries(groupWidths)
                    .filter(([key, value]) => group.columns.includes(key) && Number.isFinite(Number(value)))
                    .map(([key, value]) => [key, Math.max(72, Number(value))]),
            ),
        ];
    }));
}

function withFixedColumns(columns) {
    return ["name", ...columns.filter((key) => key !== "name")];
}

function getColumnGroup(groupKey = state.activeColumnGroup) {
    return COLUMN_GROUPS.find((group) => group.key === normalizeColumnGroup(groupKey)) || COLUMN_GROUPS[0];
}

function getColumnByKey(key) {
    return COLUMN_DEFS.find((column) => column.key === key);
}

function getPriceLinkUrl(stock) {
    const ticker = String(stock?.ticker || "").trim();
    if (!ticker) return "";
    const overrideUrl = PRICE_LINK_OVERRIDES.get(ticker.toUpperCase());
    if (overrideUrl) return overrideUrl;
    return `https://finance.yahoo.com/quote/${encodeURIComponent(ticker)}`;
}

function openPriceLinkWindow(stock) {
    const priceLinkUrl = getPriceLinkUrl(stock);
    if (!priceLinkUrl) return;
    const ticker = String(stock?.ticker || "stock").replace(/[^\w.-]/g, "_");
    window.open(
        priceLinkUrl,
        `choicePriceLink_${ticker}_${Date.now()}`,
        "popup=yes,width=1200,height=820,left=80,top=60,noopener"
    );
}

function readLocalActiveTabId() {
    try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
        return Array.isArray(saved?.tabs) ? saved.activeTabId : null;
    } catch (error) {
        return null;
    }
}

function resolveInitialActiveTabId(saved) {
    const localActiveTabId = readLocalActiveTabId();
    const savedTabIds = new Set(saved.tabs.map((tab) => tab.id).filter(Boolean));
    if (localActiveTabId && savedTabIds.has(localActiveTabId)) return localActiveTabId;
    return saved.activeTabId || saved.tabs[0].id;
}

function loadState() {
    if (window.CHOICE_USER_STATE && Array.isArray(window.CHOICE_USER_STATE.tabs) && window.CHOICE_USER_STATE.tabs.length > 0) {
        const saved = window.CHOICE_USER_STATE;
        const shouldResetColumns = saved.columnSchemaVersion !== COLUMN_SCHEMA_VERSION;
        return {
            tabs: saved.tabs.map((tab, index) => ({
                id: tab.id || createId(),
                title: tab.title || `관심종목 ${index + 1}`,
                stocks: Array.isArray(tab.stocks) ? tab.stocks : [],
            })),
            activeTabId: resolveInitialActiveTabId(saved),
            visibleColumns: shouldResetColumns ? [...DEFAULT_VISIBLE_COLUMNS] : normalizeVisibleColumns(saved.visibleColumns),
            activeColumnGroup: normalizeColumnGroup(saved.activeColumnGroup),
            groupVisibleColumns: shouldResetColumns ? createDefaultGroupVisibleColumns() : normalizeGroupVisibleColumns(saved.groupVisibleColumns),
            groupColumnOrder: shouldResetColumns ? createDefaultGroupColumnOrder() : normalizeGroupColumnOrder(saved.groupColumnOrder),
            groupColumnWidths: shouldResetColumns ? createDefaultGroupColumnWidths() : normalizeGroupColumnWidths(saved.groupColumnWidths),
            columnSchemaVersion: COLUMN_SCHEMA_VERSION,
        };
    }

    try {
        const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
        if (saved && Array.isArray(saved.tabs) && saved.tabs.length > 0) {
            const shouldResetColumns = saved.columnSchemaVersion !== COLUMN_SCHEMA_VERSION;
            return {
                tabs: saved.tabs.map((tab, index) => ({
                    id: tab.id || createId(),
                    title: tab.title || `관심종목 ${index + 1}`,
                    stocks: Array.isArray(tab.stocks) ? tab.stocks : [],
                })),
                activeTabId: resolveInitialActiveTabId(saved),
                visibleColumns: shouldResetColumns ? [...DEFAULT_VISIBLE_COLUMNS] : normalizeVisibleColumns(saved.visibleColumns),
                activeColumnGroup: normalizeColumnGroup(saved.activeColumnGroup),
                groupVisibleColumns: shouldResetColumns ? createDefaultGroupVisibleColumns() : normalizeGroupVisibleColumns(saved.groupVisibleColumns),
                groupColumnOrder: shouldResetColumns ? createDefaultGroupColumnOrder() : normalizeGroupColumnOrder(saved.groupColumnOrder),
                groupColumnWidths: shouldResetColumns ? createDefaultGroupColumnWidths() : normalizeGroupColumnWidths(saved.groupColumnWidths),
                columnSchemaVersion: COLUMN_SCHEMA_VERSION,
            };
        }
    } catch (error) {
        console.warn("저장된 상태를 불러오지 못했습니다.", error);
    }

    const firstTab = createDefaultTab();
    return {
        tabs: [firstTab],
        activeTabId: firstTab.id,
        visibleColumns: [...DEFAULT_VISIBLE_COLUMNS],
        activeColumnGroup: DEFAULT_COLUMN_GROUP,
        groupVisibleColumns: createDefaultGroupVisibleColumns(),
        groupColumnOrder: createDefaultGroupColumnOrder(),
        groupColumnWidths: createDefaultGroupColumnWidths(),
        columnSchemaVersion: COLUMN_SCHEMA_VERSION,
    };
}

function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    scheduleServerStateSave();
}

function scheduleServerStateSave() {
    window.clearTimeout(saveStateTimer);
    saveStateTimer = window.setTimeout(() => {
        fetch("/api/user-data", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ state }),
        }).catch((error) => {
            console.warn("Failed to save user state", error);
        });
    }, 400);
}

function applyLayoutToggleStage() {
    document.body.classList.toggle("layout-hide-tools", layoutToggleStage === 1 || layoutToggleStage === 2);
    document.body.classList.toggle("layout-hide-navigation", layoutToggleStage === 2);
    if (layoutToggleButton) {
        layoutToggleButton.dataset.stage = String(layoutToggleStage + 1);
        layoutToggleButton.setAttribute("aria-pressed", String(layoutToggleStage !== 0));
        layoutToggleButton.title = `레이아웃 토글 ${layoutToggleStage + 1}/3`;
    }
    hideColumnSettings();
    suggestionsEl.classList.remove("open");
    updateTopHorizontalScroll();
}

function toggleLayoutStage() {
    layoutToggleStage = (layoutToggleStage + 1) % 3;
    applyLayoutToggleStage();
}

function getVisibleColumns() {
    state.groupVisibleColumns = normalizeGroupVisibleColumns(state.groupVisibleColumns);
    state.groupColumnOrder = normalizeGroupColumnOrder(state.groupColumnOrder);
    const activeGroup = getColumnGroup();
    const visibleKeys = new Set(state.groupVisibleColumns[activeGroup.key]);
    visibleKeys.add("name");
    return state.groupColumnOrder[activeGroup.key]
        .filter((key) => visibleKeys.has(key))
        .map(getColumnByKey)
        .filter(Boolean);
}

function applyGridColumns() {
    const activeGroup = getColumnGroup();
    state.groupColumnWidths = normalizeGroupColumnWidths(state.groupColumnWidths);
    const widths = state.groupColumnWidths[activeGroup.key] || {};
    const columnWidths = getVisibleColumns().map((column) => widths[column.key] ? `${widths[column.key]}px` : column.width).join(" ");
    document.documentElement.style.setProperty("--stock-grid-columns", `var(--drag-column-width) ${columnWidths} 52px`);
}

function getActiveTab() {
    let tab = state.tabs.find((item) => item.id === state.activeTabId);
    if (!tab) {
        tab = state.tabs[0];
        state.activeTabId = tab.id;
    }
    return tab;
}

function normalizeStockPriceSnapshot(snapshot) {
    if (!snapshot || typeof snapshot !== "object" || Array.isArray(snapshot)) return null;

    const fields = snapshot.fields && typeof snapshot.fields === "object" && !Array.isArray(snapshot.fields)
        ? snapshot.fields
        : snapshot;
    const normalized = {};

    FETCHABLE_COLUMN_KEYS.forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(fields, key)) {
            normalized[key] = fields[key];
        }
    });

    const updatedAt = snapshot.updated_at || fields.updated_at;
    if (updatedAt) normalized.updated_at = String(updatedAt);

    return Object.keys(normalized).length ? normalized : null;
}

function isNewerStockPriceSnapshot(snapshot, stock) {
    if (!snapshot?.updated_at) return false;
    if (!stock?.updated_at) return true;

    const snapshotTime = Date.parse(snapshot.updated_at);
    const stockTime = Date.parse(stock.updated_at);
    if (!Number.isFinite(snapshotTime)) return false;
    if (!Number.isFinite(stockTime)) return true;
    return snapshotTime > stockTime;
}

async function applyLatestStockPricesFromFile() {
    const activeTab = getActiveTab();
    if (!activeTab.stocks.length) return false;

    try {
        const response = await fetch(STOCK_PRICES_URL, { cache: "no-store" });
        if (!response.ok) return false;

        const payload = await response.json();
        if (!payload || typeof payload !== "object" || Array.isArray(payload)) return false;

        let changed = false;
        activeTab.stocks = activeTab.stocks.map((stock) => {
            const ticker = String(stock.ticker || "").trim();
            const snapshot = normalizeStockPriceSnapshot(payload[ticker]);
            if (!snapshot || !isNewerStockPriceSnapshot(snapshot, stock)) return stock;

            changed = true;
            return {
                ...stock,
                ...snapshot,
            };
        });

        if (changed) {
            clearInlineChartCache(activeTab.stocks.map((stock) => stock.ticker));
            render();
        }

        return changed;
    } catch (error) {
        console.warn("stock_prices.json load failed", error);
        return false;
    }
}

function refreshActiveTabFromStockPrices() {
    applyLatestStockPricesFromFile();
}

function getChartMarket(stock) {
    const ticker = String(stock?.ticker || "").toUpperCase();
    const market = String(stock?.market || "").toUpperCase();
    return ticker === "^KS11" || ticker === "^KQ11" || ticker.endsWith(".KS") || ticker.endsWith(".KQ") || market.includes("KOSPI") || market.includes("KOSDAQ") || market.includes("KOREA")
        ? "kr"
        : "us";
}

function getStockplusSymbol(stock) {
    const rawTicker = String(stock?.ticker || "").trim().toUpperCase();
    if (!rawTicker) return "";
    if (rawTicker === "^KS11") return "KGG01P";
    if (rawTicker === "^KQ11") return "QGG01P";
    if (rawTicker === "^IXIC") return "COMP";
    if (rawTicker === "^GSPC") return "SP500";
    if (rawTicker === "KRW=X") return "FRX.KRWUSD";
    if (getChartMarket(stock) === "kr") {
        const code = rawTicker.replace(/\.(KS|KQ)$/i, "").replace(/^A/i, "");
        return `A${code}`;
    }
    return rawTicker;
}

function getStockplusStockUrl(stock) {
    const symbol = getStockplusSymbol(stock);
    if (!symbol) return "#";
    return getChartMarket(stock) === "kr"
        ? `https://www.stockplus.com/m/stocks/KOREA-${symbol}`
        : `https://www.stockplus.com/m/stocks/USA-${symbol}`;
}

function getChartImageUrl(stock, column) {
    const symbol = getStockplusSymbol(stock);
    if (!symbol || !column?.chartType) return "";
    const market = getChartMarket(stock);
    if (symbol === "FRX.KRWUSD" && column.chartType === "candle") {
        return `https://quot-chart.stockplus.com/images/global/forexcandle/${column.chartPeriod}/${symbol}.png`;
    }
    if (column.chartType === "candle") {
        return `https://quot-chart.stockplus.com/images/${market}/candle/${column.chartPeriod}/${symbol}.png`;
    }
    return `https://quot-chart.stockplus.com/images/${market}/stock/${column.chartPeriod}/${symbol}.png`;
}

function shouldRenderInlineChart(stock) {
    const ticker = String(stock?.ticker || "").trim().toUpperCase();
    return YAHOO_SVG_CHART_TICKERS.has(ticker);
}

function getInlineChartCacheKey(stock, column) {
    return `${String(stock?.ticker || "").trim().toUpperCase()}|${column.chartType}|${column.chartPeriod}`;
}

function clearInlineChartCache(tickers = []) {
    const targets = new Set(tickers.map((ticker) => String(ticker || "").trim().toUpperCase()).filter(Boolean));
    if (!targets.size) {
        inlineChartCache.clear();
        return;
    }

    Array.from(inlineChartCache.keys()).forEach((key) => {
        const [ticker] = key.split("|");
        if (targets.has(ticker)) inlineChartCache.delete(key);
    });
}

function getInlineChartData(stock, column) {
    const key = getInlineChartCacheKey(stock, column);
    if (!inlineChartCache.has(key)) {
        const params = new URLSearchParams({
            ticker: stock.ticker,
            type: column.chartType,
            period: column.chartPeriod,
        });
        inlineChartCache.set(
            key,
            fetch(`/api/chart-data?${params.toString()}`, { cache: "no-store" })
                .then((response) => {
                    if (!response.ok) throw new Error("chart data unavailable");
                    return response.json();
                })
        );
    }
    return inlineChartCache.get(key);
}

function scaleChartY(value, minValue, maxValue, top, bottom) {
    if (maxValue === minValue) return (top + bottom) / 2;
    return top + ((maxValue - value) / (maxValue - minValue)) * (bottom - top);
}

function formatChartPrice(value) {
    if (!Number.isFinite(value)) return "";
    return value.toLocaleString("en-US", {
        minimumFractionDigits: value >= 100 ? 0 : 2,
        maximumFractionDigits: value >= 100 ? 0 : 2,
    });
}

function formatChartDate(timestamp) {
    const date = new Date(Number(timestamp) * 1000);
    if (Number.isNaN(date.getTime())) return "";
    return `${String(date.getFullYear()).slice(-2)}.${date.getMonth() + 1}.${date.getDate()}`;
}

function formatChartKstTime(timestamp) {
    const date = new Date(Number(timestamp) * 1000);
    if (Number.isNaN(date.getTime())) return "";
    return new Intl.DateTimeFormat("en-US", {
        timeZone: "Asia/Seoul",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
    }).format(date);
}

function getChartScale(values) {
    const minRaw = Math.min(...values);
    const maxRaw = Math.max(...values);
    const span = Math.max(maxRaw - minRaw, Math.abs(maxRaw) * 0.02, 1);
    return {
        minValue: minRaw - span * 0.08,
        maxValue: maxRaw + span * 0.08,
    };
}

function getRecentSixHourPoints(points) {
    const sortedPoints = points
        .filter((point) => Number.isFinite(Number(point.time)))
        .sort((left, right) => Number(left.time) - Number(right.time));
    if (sortedPoints.length < 2) return sortedPoints;

    const latestTime = Number(sortedPoints[sortedPoints.length - 1].time);
    const startTime = latestTime - 6 * 60 * 60;
    const recentPoints = sortedPoints.filter((point) => Number(point.time) >= startTime);
    return recentPoints.length >= 2 ? recentPoints : sortedPoints.slice(-2);
}

function getTimeScaleX(timestamp, startTime, endTime, plotLeft, plotRight) {
    if (endTime === startTime) return (plotLeft + plotRight) / 2;
    return plotLeft + ((Number(timestamp) - startTime) / (endTime - startTime)) * (plotRight - plotLeft);
}

function renderChartAxes(points, minValue, maxValue, width, height, plotLeft, plotRight, plotTop, plotBottom, chartPeriod, chartType) {
    const priceTicks = [maxValue, (maxValue + minValue) / 2, minValue];
    const isOneDayLineChart = chartPeriod === "d" && chartType === "stock";
    const dateRatios = isOneDayLineChart ? [1 / 6, 2 / 6, 3 / 6, 4 / 6, 5 / 6, 1] : [0.2, 0.4, 0.6, 0.8, 1];
    const dateStep = (plotRight - plotLeft) / Math.max(points.length - 1, 1);
    const latestTime = isOneDayLineChart ? Number(points[points.length - 1]?.time) : null;
    const startTime = isOneDayLineChart ? latestTime - 6 * 60 * 60 : null;
    const endTime = isOneDayLineChart ? latestTime : null;
    const grid = priceTicks.map((value) => {
        const y = scaleChartY(value, minValue, maxValue, plotTop, plotBottom);
        return `M${plotLeft} ${y.toFixed(1)}H${plotRight}`;
    }).join("");
    const priceLabels = priceTicks.map((value) => {
        const y = scaleChartY(value, minValue, maxValue, plotTop, plotBottom);
        return `<text class="inline-stock-chart-price" x="${width - 4}" y="${(y + 4).toFixed(1)}">${formatChartPrice(value)}</text>`;
    }).join("");
    const dateLabels = dateRatios.map((ratio) => {
        const tickTime = isOneDayLineChart ? startTime + 6 * 60 * 60 * ratio : null;
        const index = Math.round((points.length - 1) * ratio);
        if (!isOneDayLineChart && (index <= 0 || index >= points.length)) return "";
        const x = isOneDayLineChart
            ? getTimeScaleX(tickTime, startTime, endTime, plotLeft, plotRight)
            : plotLeft + index * dateStep;
        const label = isOneDayLineChart ? formatChartKstTime(tickTime) : formatChartDate(points[index].time);
        return `<text class="inline-stock-chart-date" x="${x.toFixed(1)}" y="${height - 5}">${label}</text>`;
    }).join("");
    return `
        <path class="inline-stock-chart-grid" d="${grid}"></path>
        <line class="inline-stock-chart-axis" x1="${plotLeft}" y1="${plotBottom}" x2="${plotRight}" y2="${plotBottom}"></line>
        ${priceLabels}
        ${dateLabels}
    `;
}

function renderLineChartSvg(points, chartPeriod) {
    const width = 420;
    const height = 130;
    const plotLeft = 8;
    const plotRight = 342;
    const plotTop = 8;
    const plotBottom = 104;
    const rawChartPoints = points
        .map((point) => ({ time: point.time, close: Number(point.close) }))
        .filter((point) => Number.isFinite(point.close));
    const chartPoints = chartPeriod === "d" ? getRecentSixHourPoints(rawChartPoints) : rawChartPoints;
    if (chartPoints.length < 2) return "";
    const { minValue, maxValue } = getChartScale(chartPoints.map((point) => point.close));
    const latestTime = Number(chartPoints[chartPoints.length - 1].time);
    const startTime = chartPeriod === "d" ? latestTime - 6 * 60 * 60 : Number(chartPoints[0].time);
    const endTime = chartPeriod === "d" ? latestTime : Number(chartPoints[chartPoints.length - 1].time);
    const step = (plotRight - plotLeft) / Math.max(chartPoints.length - 1, 1);
    const path = chartPoints.map((point, index) => {
        const x = chartPeriod === "d"
            ? getTimeScaleX(point.time, startTime, endTime, plotLeft, plotRight)
            : plotLeft + index * step;
        const y = scaleChartY(point.close, minValue, maxValue, plotTop, plotBottom);
        return `${index ? "L" : "M"}${x.toFixed(1)} ${y.toFixed(1)}`;
    }).join(" ");
    const rising = chartPoints[chartPoints.length - 1].close >= chartPoints[0].close;
    const stroke = rising ? "#dc2626" : "#2563eb";
    return `
        <svg class="inline-stock-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="chart">
            ${renderChartAxes(chartPoints, minValue, maxValue, width, height, plotLeft, plotRight, plotTop, plotBottom, chartPeriod, "stock")}
            <path class="inline-stock-chart-line" d="${path}" stroke="${stroke}"></path>
        </svg>
    `;
}

function renderCandleChartSvg(points, chartPeriod) {
    const width = 420;
    const height = 130;
    const plotLeft = 8;
    const plotRight = 342;
    const plotTop = 8;
    const plotBottom = 104;
    const candles = points
        .map((point) => ({
            time: point.time,
            open: Number(point.open),
            high: Number(point.high),
            low: Number(point.low),
            close: Number(point.close),
        }))
        .filter((point) => [point.open, point.high, point.low, point.close].every(Number.isFinite));
    if (candles.length < 2) return "";
    const values = candles.flatMap((point) => [point.high, point.low]);
    const { minValue, maxValue } = getChartScale(values);
    const step = (plotRight - plotLeft) / candles.length;
    const bodyWidth = Math.max(2, Math.min(8, step * 0.62));
    const shapes = candles.map((point, index) => {
        const x = plotLeft + index * step + step / 2;
        const highY = scaleChartY(point.high, minValue, maxValue, plotTop, plotBottom);
        const lowY = scaleChartY(point.low, minValue, maxValue, plotTop, plotBottom);
        const openY = scaleChartY(point.open, minValue, maxValue, plotTop, plotBottom);
        const closeY = scaleChartY(point.close, minValue, maxValue, plotTop, plotBottom);
        const top = Math.min(openY, closeY);
        const bodyHeight = Math.max(1, Math.abs(closeY - openY));
        const color = point.close >= point.open ? "#dc2626" : "#2563eb";
        return `<line x1="${x.toFixed(1)}" y1="${highY.toFixed(1)}" x2="${x.toFixed(1)}" y2="${lowY.toFixed(1)}" stroke="${color}" stroke-width="1"></line><rect x="${(x - bodyWidth / 2).toFixed(1)}" y="${top.toFixed(1)}" width="${bodyWidth.toFixed(1)}" height="${bodyHeight.toFixed(1)}" fill="${color}"></rect>`;
    }).join("");
    return `
        <svg class="inline-stock-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="chart">
            ${renderChartAxes(candles, minValue, maxValue, width, height, plotLeft, plotRight, plotTop, plotBottom, chartPeriod, "candle")}
            ${shapes}
        </svg>
    `;
}

function renderInlineChartSvg(payload) {
    const points = Array.isArray(payload?.points) ? payload.points : [];
    if (payload?.type === "candle") return renderCandleChartSvg(points, payload?.period);
    return renderLineChartSvg(points, payload?.period);
}

function mountInlineChart(thumb, stock, column) {
    thumb.classList.add("inline-stock-chart-thumb");
    thumb.textContent = "...";
    getInlineChartData(stock, column)
        .then((payload) => {
            const svg = renderInlineChartSvg(payload);
            thumb.innerHTML = svg || `<span class="inline-stock-chart-empty">-</span>`;
            if (svg) {
                const chartName = `${stock.name || stock.ticker || "stock"} ${column.label}`;
                thumb.addEventListener("click", (event) => openInlineChartPopout(event, svg, chartName));
            }
        })
        .catch(() => {
            thumb.innerHTML = `<span class="inline-stock-chart-empty">-</span>`;
        });
}

function clampEntryImagePopoutPosition(popout) {
    if (!popout) return;
    const rect = popout.getBoundingClientRect();
    const gap = 8;
    const maxLeft = Math.max(gap, window.innerWidth - rect.width - gap);
    const maxTop = Math.max(gap, window.innerHeight - rect.height - gap);
    const left = Math.min(Math.max(gap, Number.parseFloat(popout.style.left) || gap), maxLeft);
    const top = Math.min(Math.max(gap, Number.parseFloat(popout.style.top) || gap), maxTop);
    popout.style.left = `${left}px`;
    popout.style.top = `${top}px`;
}

function closeEntryImagePopout(id) {
    document.getElementById(id)?.remove();
}

function bringEntryImagePopoutToFront(popoutOrId) {
    const popout = typeof popoutOrId === "string"
        ? document.getElementById(popoutOrId)
        : popoutOrId;
    if (!popout) return;
    entryImagePopoutTopZ += 1;
    popout.style.zIndex = String(entryImagePopoutTopZ);
}

function setEntryImagePopoutScale(id, delta) {
    const popout = document.getElementById(id);
    const media = popout?.querySelector(".entry-image-popout-body img, .entry-image-popout-body svg");
    if (!popout || !media) return;
    bringEntryImagePopoutToFront(popout);
    const current = Number(popout.dataset.scale || "1") || 1;
    const next = Math.min(4, Math.max(0.25, Math.round((current + delta) * 100) / 100));
    popout.dataset.scale = String(next);
    media.style.width = `${next * 100}%`;
    popout.classList.toggle("is-zoomed", next > 1);
}

function startEntryImagePopoutDrag(event, id) {
    const popout = document.getElementById(id);
    if (!popout) return;
    event.preventDefault();
    bringEntryImagePopoutToFront(popout);
    const startX = event.clientX;
    const startY = event.clientY;
    const rect = popout.getBoundingClientRect();
    const startLeft = rect.left;
    const startTop = rect.top;
    const handleMove = (moveEvent) => {
        popout.style.left = `${startLeft + moveEvent.clientX - startX}px`;
        popout.style.top = `${startTop + moveEvent.clientY - startY}px`;
        clampEntryImagePopoutPosition(popout);
    };
    const handleUp = () => {
        document.removeEventListener("pointermove", handleMove);
        document.removeEventListener("pointerup", handleUp);
    };
    document.addEventListener("pointermove", handleMove);
    document.addEventListener("pointerup", handleUp, { once: true });
}

function startEntryImagePopoutPan(event, id) {
    const popout = document.getElementById(id);
    const body = popout?.querySelector(".entry-image-popout-body");
    if (!popout || !body || Number(popout.dataset.scale || "1") <= 1) return;
    if (event.button !== undefined && event.button !== 0) return;
    event.preventDefault();
    bringEntryImagePopoutToFront(popout);
    body.classList.add("is-panning");
    const startX = event.clientX;
    const startY = event.clientY;
    const startScrollLeft = body.scrollLeft;
    const startScrollTop = body.scrollTop;
    const handleMove = (moveEvent) => {
        body.scrollLeft = startScrollLeft - (moveEvent.clientX - startX);
        body.scrollTop = startScrollTop - (moveEvent.clientY - startY);
    };
    const handleUp = () => {
        body.classList.remove("is-panning");
        document.removeEventListener("pointermove", handleMove);
        document.removeEventListener("pointerup", handleUp);
    };
    document.addEventListener("pointermove", handleMove);
    document.addEventListener("pointerup", handleUp, { once: true });
}

function openChartImagePopout(event, imageSrc, imageName = "chart image") {
    if (!imageSrc) return;
    event?.stopPropagation?.();
    const existing = Array.from(document.querySelectorAll(".entry-image-popout"))
        .find((item) => item.dataset.imageSrc === imageSrc);
    if (existing) {
        bringEntryImagePopoutToFront(existing);
        clampEntryImagePopoutPosition(existing);
        return;
    }

    const id = `entryImagePopout-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const popout = document.createElement("div");
    popout.id = id;
    popout.className = "entry-image-popout";
    popout.dataset.scale = "1";
    popout.dataset.imageSrc = imageSrc;
    popout.innerHTML = `
        <div class="entry-image-popout-toolbar">
            <button class="entry-image-popout-drag" type="button" onpointerdown="startEntryImagePopoutDrag(event, '${id}')" title="창 이동">::</button>
            <button class="entry-image-popout-btn" type="button" onclick="setEntryImagePopoutScale('${id}', 0.15)" title="확대">+</button>
            <button class="entry-image-popout-btn" type="button" onclick="setEntryImagePopoutScale('${id}', -0.15)" title="축소">-</button>
            <button class="entry-image-popout-btn" type="button" onclick="closeEntryImagePopout('${id}')" title="닫기">x</button>
        </div>
        <div class="entry-image-popout-body" onpointerdown="startEntryImagePopoutPan(event, '${id}')">
            <div class="entry-image-popout-title">${escapeHtml(imageName)}</div>
            <img src="${escapeHtml(imageSrc)}" alt="${escapeHtml(imageName)}">
        </div>
    `;
    document.body.appendChild(popout);
    popout.addEventListener("pointerdown", () => bringEntryImagePopoutToFront(popout), { capture: true });
    bringEntryImagePopoutToFront(popout);
    const sourceRect = event?.target?.getBoundingClientRect?.();
    const left = sourceRect ? sourceRect.left + 16 : 80;
    const top = sourceRect ? sourceRect.bottom + 8 : 80;
    popout.style.left = `${left}px`;
    popout.style.top = `${top}px`;
    clampEntryImagePopoutPosition(popout);
}

function openInlineChartPopout(event, svgMarkup, chartName = "chart") {
    if (!svgMarkup) return;
    event?.stopPropagation?.();
    const popoutKey = `${chartName}|${svgMarkup.length}`;
    const existing = Array.from(document.querySelectorAll(".entry-image-popout"))
        .find((item) => item.dataset.inlineChart === popoutKey);
    if (existing) {
        bringEntryImagePopoutToFront(existing);
        clampEntryImagePopoutPosition(existing);
        return;
    }

    const id = `entryImagePopout-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const popout = document.createElement("div");
    popout.id = id;
    popout.className = "entry-image-popout inline-chart-popout";
    popout.dataset.scale = "1";
    popout.dataset.inlineChart = popoutKey;
    popout.innerHTML = `
        <div class="entry-image-popout-toolbar">
            <button class="entry-image-popout-drag" type="button" onpointerdown="startEntryImagePopoutDrag(event, '${id}')" title="창 이동">::</button>
            <button class="entry-image-popout-btn" type="button" onclick="setEntryImagePopoutScale('${id}', 0.15)" title="확대">+</button>
            <button class="entry-image-popout-btn" type="button" onclick="setEntryImagePopoutScale('${id}', -0.15)" title="축소">-</button>
            <button class="entry-image-popout-btn" type="button" onclick="closeEntryImagePopout('${id}')" title="닫기">x</button>
        </div>
        <div class="entry-image-popout-body" onpointerdown="startEntryImagePopoutPan(event, '${id}')">
            <div class="entry-image-popout-title">${escapeHtml(chartName)}</div>
            ${svgMarkup}
        </div>
    `;
    document.body.appendChild(popout);
    popout.addEventListener("pointerdown", () => bringEntryImagePopoutToFront(popout), { capture: true });
    bringEntryImagePopoutToFront(popout);
    const sourceRect = event?.target?.getBoundingClientRect?.();
    popout.style.left = `${sourceRect ? sourceRect.left + 16 : 80}px`;
    popout.style.top = `${sourceRect ? sourceRect.bottom + 8 : 80}px`;
    clampEntryImagePopoutPosition(popout);
}

function render() {
    applyGridColumns();
    renderTabs();
    renderColumnGroupTabs();
    renderColumnSettings();
    renderTableHead();
    renderStocks();
    updateTopHorizontalScroll();
    saveState();
}

function renderColumnGroupTabs() {
    columnGroupTabsEl.innerHTML = "";

    COLUMN_GROUPS.forEach((group) => {
        const button = document.createElement("button");
        button.className = `column-group-tab${group.key === state.activeColumnGroup ? " active" : ""}`;
        button.type = "button";
        button.textContent = group.label;

        button.addEventListener("click", () => {
            state.activeColumnGroup = group.key;
            render();
            refreshActiveTabFromStockPrices();
        });

        columnGroupTabsEl.appendChild(button);
    });
}

function renderTabs() {
    tabsEl.innerHTML = "";

    state.tabs.forEach((tab) => {
        const tabButton = document.createElement("button");
        tabButton.className = `tab${tab.id === state.activeTabId ? " active" : ""}`;
        tabButton.type = "button";
        tabButton.draggable = true;
        tabButton.dataset.tabId = tab.id;

        const title = document.createElement("span");
        title.className = "tab-title";
        title.textContent = tab.title;

        tabButton.addEventListener("click", () => {
            state.activeTabId = tab.id;
            hideContextMenu();
            render();
            refreshActiveTabFromStockPrices();
        });

        tabButton.addEventListener("contextmenu", (event) => {
            event.preventDefault();
            contextTabId = tab.id;
            showContextMenu(event.clientX, event.clientY);
        });

        tabButton.addEventListener("dragstart", (event) => {
            draggedTabId = tab.id;
            tabButton.classList.add("dragging");
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", tab.id);
        });

        tabButton.addEventListener("dragend", () => {
            draggedTabId = null;
            tabsEl.querySelectorAll(".tab").forEach((item) => {
                item.classList.remove("dragging", "drop-target");
            });
            saveState();
        });

        tabButton.addEventListener("dragover", (event) => {
            if (!draggedTabId || draggedTabId === tab.id) return;
            event.preventDefault();
            event.dataTransfer.dropEffect = "move";
            tabsEl.querySelectorAll(".tab.drop-target").forEach((item) => {
                item.classList.remove("drop-target");
            });
            tabButton.classList.add("drop-target");
        });

        tabButton.addEventListener("drop", (event) => {
            event.preventDefault();
            moveTab(draggedTabId, tab.id);
        });

        tabButton.append(title);
        tabsEl.appendChild(tabButton);
    });
}

function moveTab(sourceTabId, targetTabId) {
    if (!sourceTabId || !targetTabId || sourceTabId === targetTabId) return;

    const sourceIndex = state.tabs.findIndex((tab) => tab.id === sourceTabId);
    const targetIndex = state.tabs.findIndex((tab) => tab.id === targetTabId);
    if (sourceIndex < 0 || targetIndex < 0) return;

    const [moved] = state.tabs.splice(sourceIndex, 1);
    state.tabs.splice(targetIndex, 0, moved);
    render();
}

function renderColumnSettings() {
    columnSettingsMenu.innerHTML = "";
    const activeGroup = getColumnGroup();

    activeGroup.columns
        .map(getColumnByKey)
        .filter(Boolean)
        .forEach((column) => {
        const label = document.createElement("label");
        label.className = "column-setting-item";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = state.groupVisibleColumns[activeGroup.key].includes(column.key);

        checkbox.addEventListener("change", () => {
            if (checkbox.checked) {
                state.groupVisibleColumns[activeGroup.key] = [...new Set([...state.groupVisibleColumns[activeGroup.key], column.key])];
            } else {
                state.groupVisibleColumns[activeGroup.key] = state.groupVisibleColumns[activeGroup.key].filter((key) => key !== column.key);
            }

            render();
            columnSettingsMenu.classList.add("open");
            columnSettingsButton.setAttribute("aria-expanded", "true");
        });

        const text = document.createElement("span");
        text.textContent = column.label;

        label.append(checkbox, text);
        columnSettingsMenu.appendChild(label);
    });
}

function renderTableHead() {
    tableHeadEl.innerHTML = "";
    tableHeadEl.appendChild(document.createElement("span"));

    getVisibleColumns().forEach((column) => {
        const cell = document.createElement("button");
        cell.className = `stock-header-cell${FIXED_COLUMN_KEYS.has(column.key) ? " fixed" : ""}`;
        cell.type = "button";
        cell.draggable = !FIXED_COLUMN_KEYS.has(column.key);
        cell.dataset.columnKey = column.key;
        cell.title = FIXED_COLUMN_KEYS.has(column.key) ? "고정 열" : "열 위치 변경";

        const label = document.createElement("span");
        label.textContent = column.label;

        const resizeHandle = document.createElement("span");
        resizeHandle.className = "column-resize-handle";
        resizeHandle.dataset.columnKey = column.key;
        resizeHandle.title = "열 너비 조절";

        cell.append(label, resizeHandle);
        tableHeadEl.appendChild(cell);
    });

    tableHeadEl.appendChild(document.createElement("span"));
    bindHeaderColumnDrag();
    bindHeaderColumnSort();
    bindHeaderColumnResize();
}

function getStockSortValue(stock, columnKey) {
    if (columnKey === "holding_amount") {
        const averagePrice = parseStockNumber(stock.average_price);
        const quantity = parseStockNumber(stock.quantity);
        return averagePrice === null || quantity === null ? null : averagePrice * quantity;
    }
    if (columnKey === "profit_rate") {
        const currentPrice = parseStockNumber(stock.price);
        const averagePrice = parseStockNumber(stock.average_price);
        return currentPrice === null || averagePrice === null || averagePrice === 0
            ? null
            : ((currentPrice - averagePrice) / averagePrice) * 100;
    }
    if (columnKey === "profit_amount") {
        const currentPrice = parseStockNumber(stock.price);
        const averagePrice = parseStockNumber(stock.average_price);
        const quantity = parseStockNumber(stock.quantity);
        return currentPrice === null || averagePrice === null || quantity === null
            ? null
            : (currentPrice - averagePrice) * quantity;
    }

    const value = stock[columnKey];
    if (value === undefined || value === null || value === "" || value === "-") return null;
    const numericValue = parseSortNumber(value);
    return numericValue === null ? String(value || "") : numericValue;
}

function compareStockValues(a, b, columnKey, direction) {
    const aValue = getStockSortValue(a, columnKey);
    const bValue = getStockSortValue(b, columnKey);
    const multiplier = direction === "asc" ? 1 : -1;

    if (aValue === null && bValue === null) return 0;
    if (aValue === null) return 1;
    if (bValue === null) return -1;

    if (typeof aValue === "number" && typeof bValue === "number") {
        return (aValue - bValue) * multiplier;
    }

    return String(aValue).localeCompare(String(bValue), "ko-KR", { numeric: true }) * multiplier;
}

function restoreHeaderSortOrder(activeTab) {
    const order = new Map(headerSortState.baseOrder.map((rowId, index) => [rowId, index]));
    activeTab.stocks.sort((a, b) => {
        const aIndex = order.has(a._rowId) ? order.get(a._rowId) : Number.MAX_SAFE_INTEGER;
        const bIndex = order.has(b._rowId) ? order.get(b._rowId) : Number.MAX_SAFE_INTEGER;
        return aIndex - bIndex;
    });
    headerSortState.tabId = null;
    headerSortState.groupKey = null;
    headerSortState.columnKey = null;
    headerSortState.direction = null;
    headerSortState.baseOrder = [];
}

function sortActiveStocksByColumn(columnKey) {
    const activeTab = getActiveTab();
    activeTab.stocks = activeTab.stocks.map((stock) => (
        stock._rowId ? stock : { ...stock, _rowId: createStockRowId() }
    ));

    const groupKey = getColumnGroup().key;
    const isSameSort = headerSortState.tabId === activeTab.id
        && headerSortState.groupKey === groupKey
        && headerSortState.columnKey === columnKey;

    if (!isSameSort) {
        headerSortState.tabId = activeTab.id;
        headerSortState.groupKey = groupKey;
        headerSortState.columnKey = columnKey;
        headerSortState.direction = "desc";
        headerSortState.baseOrder = activeTab.stocks.map((stock) => stock._rowId);
    } else if (headerSortState.direction === "desc") {
        headerSortState.direction = "asc";
    } else {
        restoreHeaderSortOrder(activeTab);
        render();
        return;
    }

    activeTab.stocks.sort((a, b) => compareStockValues(a, b, columnKey, headerSortState.direction));
    render();
}

function bindHeaderColumnSort() {
    tableHeadEl.querySelectorAll(".stock-header-cell").forEach((cell) => {
        cell.addEventListener("click", (event) => {
            if (event.target.closest(".column-resize-handle")) return;
            if (recentlyDraggedHeaderColumn) return;

            const columnKey = cell.dataset.columnKey;
            if (!columnKey) return;
            sortActiveStocksByColumn(columnKey);
        });
    });
}

function bindHeaderColumnDrag() {
    tableHeadEl.querySelectorAll(".stock-header-cell").forEach((cell) => {
        cell.addEventListener("dragstart", (event) => {
            const columnKey = cell.dataset.columnKey;
            if (!columnKey || FIXED_COLUMN_KEYS.has(columnKey)) {
                event.preventDefault();
                return;
            }

            draggedColumnKey = columnKey;
            cell.classList.add("dragging");
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", columnKey);
        });

        cell.addEventListener("dragend", () => {
            draggedColumnKey = null;
            recentlyDraggedHeaderColumn = true;
            window.setTimeout(() => {
                recentlyDraggedHeaderColumn = false;
            }, 0);
            tableHeadEl.querySelectorAll(".stock-header-cell").forEach((item) => {
                item.classList.remove("dragging", "drop-target");
            });
        });

        cell.addEventListener("dragover", (event) => {
            if (!draggedColumnKey) return;
            event.preventDefault();
            const targetKey = cell.dataset.columnKey;
            if (!targetKey || targetKey === draggedColumnKey || FIXED_COLUMN_KEYS.has(targetKey)) return;

            tableHeadEl.querySelectorAll(".stock-header-cell.drop-target").forEach((item) => {
                item.classList.remove("drop-target");
            });
            cell.classList.add("drop-target");
        });

        cell.addEventListener("drop", (event) => {
            event.preventDefault();
            recentlyDraggedHeaderColumn = true;
            moveVisibleColumn(draggedColumnKey, cell.dataset.columnKey);
        });
    });
}

function moveVisibleColumn(sourceKey, targetKey) {
    if (!sourceKey || !targetKey || sourceKey === targetKey) return;
    if (FIXED_COLUMN_KEYS.has(sourceKey) || FIXED_COLUMN_KEYS.has(targetKey)) return;

    const activeGroup = getColumnGroup();
    const order = withFixedColumns([...(state.groupColumnOrder[activeGroup.key] || activeGroup.columns)]);
    const sourceIndex = order.indexOf(sourceKey);
    const targetIndex = order.indexOf(targetKey);
    if (sourceIndex < 0 || targetIndex < 0) return;

    const [moved] = order.splice(sourceIndex, 1);
    order.splice(targetIndex, 0, moved);
    state.groupColumnOrder[activeGroup.key] = withFixedColumns(order);
    render();
}

function bindHeaderColumnResize() {
    tableHeadEl.querySelectorAll(".column-resize-handle").forEach((handle) => {
        handle.addEventListener("pointerdown", (event) => {
            event.preventDefault();
            event.stopPropagation();

            const columnKey = handle.dataset.columnKey;
            const cell = handle.closest(".stock-header-cell");
            if (!columnKey || !cell) return;

            const activeGroup = getColumnGroup();
            state.groupColumnWidths = normalizeGroupColumnWidths(state.groupColumnWidths);
            resizingColumn = {
                groupKey: activeGroup.key,
                columnKey,
                startX: event.clientX,
                startWidth: cell.getBoundingClientRect().width,
                pointerId: event.pointerId,
                handle,
            };

            document.body.classList.add("is-resizing-column");
            cell.classList.add("resizing");
            if (handle.setPointerCapture) handle.setPointerCapture(event.pointerId);
        });
    });
}

function applyLiveColumnWidths() {
    applyGridColumns();
    tableHeadEl.style.gridTemplateColumns = "var(--stock-grid-columns)";
    stockListEl.querySelectorAll(".stock-row").forEach((row) => {
        row.style.gridTemplateColumns = "var(--stock-grid-columns)";
    });
    updateTopHorizontalScroll();
}

function updateTopHorizontalScroll() {
    if (!topHorizontalScrollEl || !topHorizontalScrollInnerEl || !bottomHorizontalScrollEl || !bottomHorizontalScrollInnerEl || !stockTableScrollEl) return;

    window.requestAnimationFrame(() => {
        updateFixedColumnMaskWidth();
        const scrollWidth = stockTableScrollEl.scrollWidth;
        const clientWidth = stockTableScrollEl.clientWidth;
        const fixedWidth = getFixedColumnMaskWidth();
        const scrollInnerWidth = `${Math.max(0, scrollWidth - fixedWidth)}px`;
        topHorizontalScrollInnerEl.style.width = scrollInnerWidth;
        bottomHorizontalScrollInnerEl.style.width = scrollInnerWidth;
        if (fixedColumnCurtainEl) {
            fixedColumnCurtainEl.style.setProperty("--fixed-column-curtain-height", `${stockTableScrollEl.scrollHeight}px`);
        }
        const shouldHide = scrollWidth <= clientWidth + 1;
        topHorizontalScrollEl.classList.toggle("hidden", shouldHide);
        bottomHorizontalScrollEl.classList.toggle("hidden", shouldHide);
        topHorizontalScrollEl.scrollLeft = stockTableScrollEl.scrollLeft;
        bottomHorizontalScrollEl.scrollLeft = stockTableScrollEl.scrollLeft;
    });
}

let isSyncingHorizontalScroll = false;

function syncHorizontalScroll(source) {
    if (!source || isSyncingHorizontalScroll) return;

    isSyncingHorizontalScroll = true;
    const nextScrollLeft = source.scrollLeft;
    [stockTableScrollEl, topHorizontalScrollEl, bottomHorizontalScrollEl]
        .filter((target) => target && target !== source)
        .forEach((target) => {
            target.scrollLeft = nextScrollLeft;
        });
    window.requestAnimationFrame(() => {
        isSyncingHorizontalScroll = false;
    });
}

function getFixedColumnMaskWidth() {
    const value = Number.parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--fixed-stock-columns-width"));
    return Number.isFinite(value) ? value : 190;
}

function updateFixedColumnMaskWidth() {
    const nameCells = [
        tableHeadEl.querySelector('[data-column-key="name"]'),
        ...stockListEl.querySelectorAll(".stock-name"),
    ].filter(Boolean);
    if (!nameCells.length || !stockTableScrollEl) return;

    const tableRect = stockTableScrollEl.getBoundingClientRect();
    const width = Math.max(
        0,
        ...nameCells.map((nameCell) => Math.ceil(nameCell.getBoundingClientRect().right - tableRect.left)),
    );
    document.documentElement.style.setProperty("--fixed-stock-columns-width", `${width}px`);
}

document.addEventListener("pointermove", (event) => {
    if (!resizingColumn || resizingColumn.pointerId !== event.pointerId) return;

    const isPerformanceColumn = resizingColumn.columnKey?.startsWith("performance_");
    const isChartColumn = Boolean(getColumnByKey(resizingColumn.columnKey)?.chartType);
    const minWidth = isChartColumn ? 180 : (isPerformanceColumn ? 44 : 72);
    const nextWidth = Math.max(minWidth, Math.round(resizingColumn.startWidth + event.clientX - resizingColumn.startX));
    state.groupColumnWidths[resizingColumn.groupKey][resizingColumn.columnKey] = nextWidth;
    applyLiveColumnWidths();
});

function finishColumnResize(event) {
    if (!resizingColumn) return;
    if (event?.pointerId != null && resizingColumn.pointerId !== event.pointerId) return;

    try {
        resizingColumn.handle?.releasePointerCapture?.(resizingColumn.pointerId);
    } catch (error) {
        console.warn("컬럼 너비 조절 종료 처리 실패", error);
    }

    tableHeadEl.querySelectorAll(".stock-header-cell.resizing").forEach((cell) => cell.classList.remove("resizing"));
    resizingColumn = null;
    document.body.classList.remove("is-resizing-column");
    saveState();
}

document.addEventListener("pointerup", finishColumnResize);
document.addEventListener("pointercancel", finishColumnResize);

function renderStocks() {
    const activeTab = getActiveTab();
    stockListEl.innerHTML = "";
    activeTab.stocks = activeTab.stocks.map((stock) => (
        stock._rowId ? stock : { ...stock, _rowId: createStockRowId() }
    ));

    emptyStateEl.classList.toggle("hidden", activeTab.stocks.length > 0);

    activeTab.stocks.forEach((stock) => {
        const stockRowId = stock._rowId;
        const row = document.createElement("div");
        row.className = "stock-row";
        row.draggable = true;
        row.dataset.ticker = stock.ticker;
        row.dataset.stockRowId = stockRowId;

        const dragHandle = document.createElement("button");
        dragHandle.className = "drag-handle";
        dragHandle.type = "button";
        dragHandle.title = "순서 변경";
        dragHandle.draggable = true;
        dragHandle.textContent = "☰";
        row.appendChild(dragHandle);

        getVisibleColumns().forEach((column) => {
            const cell = document.createElement("span");
            const classValue = column.key === "profit_rate"
                ? formatProfitRate(stock)
                : (column.key === "profit_amount" ? formatProfitAmount(stock) : stock[column.key]);
            const valueClass = typeof column.valueClass === "function" ? column.valueClass(classValue) : "";
            cell.className = ["stock-cell", column.className || "", valueClass].filter(Boolean).join(" ");
            cell.dataset.label = column.label;

            if (column.key === "name") {
                const primary = document.createElement("span");
                primary.className = "stock-primary";
                primary.textContent = stock.name || "-";

                cell.append(primary);

                const memo = String(stock.memo || "").trim();
                cell.dataset.ticker = stock.ticker;
                cell.classList.toggle("has-memo", Boolean(memo));
                if (memo) {
                    cell.dataset.memo = memo;
                    const memoMark = document.createElement("span");
                    memoMark.className = "memo-mark";
                    memoMark.textContent = "★";
                    memoMark.setAttribute("aria-label", "메모 있음");
                    cell.prepend(memoMark);
                }
                cell.addEventListener("mouseenter", (event) => showMemoTooltip(memo, stock.market, event.clientX, event.clientY));
                cell.addEventListener("mousemove", (event) => moveMemoTooltip(event.clientX, event.clientY));
                cell.addEventListener("mouseleave", hideMemoTooltip);

                cell.addEventListener("contextmenu", (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    contextStockTicker = stock.ticker;
                    showStockContextMenu(event.clientX, event.clientY);
                });
            } else if (column.chartType) {
                const imageName = `${stock.name || stock.ticker || "stock"} ${column.label}`;
                const thumb = document.createElement("span");
                thumb.className = "voca-entry-thumb stock-chart-thumb";
                if (shouldRenderInlineChart(stock)) {
                    thumb.title = imageName;
                    mountInlineChart(thumb, stock, column);
                } else {
                    const imageSrc = getChartImageUrl(stock, column);
                    const image = document.createElement("img");
                    image.className = "stock-chart-image";
                    image.src = imageSrc;
                    image.alt = imageName;
                    image.loading = "lazy";
                    image.addEventListener("click", (event) => openChartImagePopout(event, imageSrc, imageName));
                    thumb.appendChild(image);
                }
                cell.appendChild(thumb);
            } else {
                if (column.key === "average_price" || column.key === "quantity") {
                    cell.textContent = formatAveragePrice(stock[column.key]);
                } else if (column.key === "holding_amount") {
                    cell.textContent = formatHoldingAmount(stock);
                } else if (column.key === "profit_rate") {
                    cell.textContent = formatProfitRate(stock);
                } else if (column.key === "profit_amount") {
                    cell.textContent = formatProfitAmount(stock);
                } else {
                    cell.textContent = stock[column.key] || "-";
                }
            }

            if (column.key === "price" && getColumnGroup().key === "chart") {
                const priceLinkUrl = getPriceLinkUrl(stock);
                if (priceLinkUrl) {
                    cell.classList.add("yahoo-price-link");
                    cell.title = "외부 사이트에서 확인";
                    cell.tabIndex = 0;
                    cell.addEventListener("click", () => openPriceLinkWindow(stock));
                    cell.addEventListener("keydown", (event) => {
                        if (event.key !== "Enter" && event.key !== " ") return;
                        event.preventDefault();
                        openPriceLinkWindow(stock);
                    });
                }
            }

            if (column.key === "average_price" || column.key === "quantity") {
                cell.tabIndex = 0;
                cell.title = "평단가 입력";
                cell.title = column.key === "quantity" ? "수량 입력" : "평단가 입력";
                cell.addEventListener("click", () => openAveragePriceDialog(stock.ticker, column.key));
                cell.addEventListener("keydown", (event) => {
                    if (event.key !== "Enter" && event.key !== " ") return;
                    event.preventDefault();
                    openAveragePriceDialog(stock.ticker, column.key);
                });
            }

            row.appendChild(cell);
        });

        const deleteButton = document.createElement("button");
        deleteButton.className = "delete-stock";
        deleteButton.type = "button";
        deleteButton.title = "삭제";
        deleteButton.textContent = "x";
        deleteButton.addEventListener("click", () => removeStock(stockRowId));
        row.appendChild(deleteButton);

        dragHandle.addEventListener("pointerdown", () => {
            dragHandleArmedStockRowId = stockRowId;
        });

        dragHandle.addEventListener("pointerup", () => {
            dragHandleArmedStockRowId = null;
        });

        dragHandle.addEventListener("dragstart", (event) => {
            draggedStockRowId = stockRowId;
            row.classList.add("dragging");
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", stockRowId);
        });

        row.addEventListener("dragstart", (event) => {
            if (dragHandleArmedStockRowId !== stockRowId && !event.target.closest(".drag-handle")) {
                event.preventDefault();
                return;
            }
            draggedStockRowId = stockRowId;
            row.classList.add("dragging");
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", stockRowId);
        });

        row.addEventListener("dragend", () => {
            draggedStockRowId = null;
            dragHandleArmedStockRowId = null;
            row.classList.remove("dragging");
            saveState();
        });

        row.addEventListener("dragover", (event) => {
            if (!draggedStockRowId) return;
            event.preventDefault();
            reorderStock(stockRowId);
        });

        stockListEl.appendChild(row);
    });
}

function getChangeClass(change) {
    if (!change || change === "-") return "";
    const value = String(change).trim();
    if (value.startsWith("+")) return "positive";
    if (value.startsWith("-")) return "negative";
    return "";
}

function showMemoTooltip(memo, market, x, y) {
    memoTooltipEl.innerHTML = "";
    if (!market && !memo) return;
    if (market) {
        const marketEl = document.createElement("span");
        marketEl.className = "memo-tooltip-market";
        marketEl.textContent = market;
        memoTooltipEl.appendChild(marketEl);
    }
    if (memo) {
        const memoEl = document.createElement("span");
        memoEl.className = "memo-tooltip-text";
        memoEl.textContent = memo;
        memoTooltipEl.appendChild(memoEl);
    }
    memoTooltipEl.classList.add("open");
    moveMemoTooltip(x, y);
}

function moveMemoTooltip(x, y) {
    if (!memoTooltipEl.classList.contains("open")) return;

    const margin = 12;
    const rect = memoTooltipEl.getBoundingClientRect();
    const left = Math.min(Math.max(margin, x + 12), window.innerWidth - rect.width - margin);
    const top = Math.min(Math.max(margin, y + 12), window.innerHeight - rect.height - margin);
    memoTooltipEl.style.left = `${left}px`;
    memoTooltipEl.style.top = `${top}px`;
}

function hideMemoTooltip() {
    memoTooltipEl.classList.remove("open");
}

function addTab() {
    openTabDialog("create");
}

function getNextTabTitle() {
    let index = 1;
    const existingTitles = new Set(state.tabs.map((tab) => normalizeSearchText(tab.title)));
    while (existingTitles.has(normalizeSearchText(`new tab ${index}`))) {
        index += 1;
    }
    return `new tab ${index}`;
}

function createTabWithTitle(title) {
    const tab = createDefaultTab(state.tabs.length + 1);
    tab.title = title;
    state.tabs.push(tab);
    state.activeTabId = tab.id;
    render();
}

function closeTab(tabId) {
    if (state.tabs.length === 1) return;

    const index = state.tabs.findIndex((tab) => tab.id === tabId);
    state.tabs = state.tabs.filter((tab) => tab.id !== tabId);

    if (state.activeTabId === tabId) {
        const nextTab = state.tabs[Math.max(0, index - 1)] || state.tabs[0];
        state.activeTabId = nextTab.id;
    }

    render();
}

function openDeleteTabDialog(tabId) {
    if (state.tabs.length === 1) return;
    pendingDeleteTabId = tabId;
    deleteTabDialogOverlay.classList.add("open");
    deleteTabDialogOverlay.setAttribute("aria-hidden", "false");
    hideContextMenu();
    hideColumnSettings();
    suggestionsEl.classList.remove("open");
    window.setTimeout(() => {
        confirmDeleteTabDialogButton.focus();
    }, 0);
}

function closeDeleteTabDialog() {
    deleteTabDialogOverlay.classList.remove("open");
    deleteTabDialogOverlay.setAttribute("aria-hidden", "true");
    pendingDeleteTabId = null;
}

function confirmDeleteTabDialog() {
    const tabId = pendingDeleteTabId;
    closeDeleteTabDialog();
    if (tabId) closeTab(tabId);
}

function renameTab() {
    const tab = state.tabs.find((item) => item.id === contextTabId);
    if (!tab) return;

    openTabDialog("rename", tab);
    hideContextMenu();
}

function deleteContextTab() {
    const tab = state.tabs.find((item) => item.id === contextTabId);
    if (!tab) return;
    if (state.tabs.length === 1) {
        hideContextMenu();
        return;
    }

    openDeleteTabDialog(tab.id);
}

function getContextStock() {
    const activeTab = getActiveTab();
    return activeTab.stocks.find((stock) => stock.ticker === contextStockTicker);
}

function openMemoDialog() {
    const stock = getContextStock();
    if (!stock) return;

    memoInput.value = stock.memo || "";
    memoError.textContent = "";
    memoDialogOverlay.classList.add("open");
    memoDialogOverlay.setAttribute("aria-hidden", "false");
    hideContextMenu();
    hideStockContextMenu();
    hideColumnSettings();
    suggestionsEl.classList.remove("open");
    window.setTimeout(() => {
        memoInput.focus();
        memoInput.select();
    }, 0);
}

function closeMemoDialog() {
    memoDialogOverlay.classList.remove("open");
    memoDialogOverlay.setAttribute("aria-hidden", "true");
    memoInput.value = "";
    memoError.textContent = "";
}

function confirmMemoDialog() {
    const stock = getContextStock();
    if (!stock) return;

    const memo = memoInput.value.trim();
    stock.memo = memo;
    closeMemoDialog();
    render();
}

function normalizeNumberInput(value) {
    const sanitized = String(value || "")
        .replace(/,/g, "")
        .replace(/[^\d.]/g, "");
    const [integerPart, ...decimalParts] = sanitized.split(".");
    const integerDigits = integerPart.replace(/[^\d]/g, "");
    const decimalDigits = decimalParts.join("").replace(/[^\d]/g, "").slice(0, 2);

    return decimalParts.length > 0
        ? `${integerDigits}.${decimalDigits}`
        : integerDigits;
}

function parseStockNumber(value) {
    const normalized = normalizeNumberInput(value);
    if (!normalized || normalized === ".") return null;

    const number = Number(normalized);
    return Number.isFinite(number) ? number : null;
}

function parseSortNumber(value) {
    if (value === undefined || value === null) return null;
    const normalized = String(value).trim().replace(/,/g, "").replace(/%/g, "");
    if (!normalized || normalized === "-") return null;

    const number = Number(normalized);
    return Number.isFinite(number) ? number : null;
}

function formatAveragePrice(value) {
    const normalized = normalizeNumberInput(value);
    if (!normalized || normalized === ".") return "-";

    const [integerPart, decimalPart] = normalized.split(".");
    const formattedInteger = integerPart ? Number(integerPart).toLocaleString("ko-KR") : "0";
    return decimalPart !== undefined && decimalPart.length > 0
        ? `${formattedInteger}.${decimalPart}`
        : formattedInteger;
}

function formatAveragePriceInput(value) {
    const normalized = normalizeNumberInput(value);
    if (!normalized) return "";

    const [integerPart, decimalPart] = normalized.split(".");
    const formattedInteger = integerPart ? Number(integerPart).toLocaleString("ko-KR") : "";
    return normalized.includes(".")
        ? `${formattedInteger}.${decimalPart || ""}`
        : formattedInteger;
}

function formatHoldingAmount(stock) {
    const averagePrice = parseStockNumber(stock.average_price);
    const quantity = parseStockNumber(stock.quantity);
    if (averagePrice === null || quantity === null) return "-";

    return (averagePrice * quantity).toLocaleString("ko-KR", {
        maximumFractionDigits: 2,
    });
}

function formatProfitRate(stock) {
    const currentPrice = parseStockNumber(stock.price);
    const averagePrice = parseStockNumber(stock.average_price);
    if (currentPrice === null || averagePrice === null || averagePrice === 0) return "-";

    const profitRate = ((currentPrice - averagePrice) / averagePrice) * 100;
    return `${profitRate >= 0 ? "+" : ""}${profitRate.toFixed(2)}%`;
}

function formatProfitAmount(stock) {
    const currentPrice = parseStockNumber(stock.price);
    const averagePrice = parseStockNumber(stock.average_price);
    const quantity = parseStockNumber(stock.quantity);
    if (currentPrice === null || averagePrice === null || quantity === null) return "-";

    const profitAmount = (currentPrice - averagePrice) * quantity;
    const formatted = Math.abs(profitAmount).toLocaleString("ko-KR", {
        maximumFractionDigits: 2,
    });
    return `${profitAmount >= 0 ? "+" : "-"}${formatted}`;
}

function getEditingAveragePriceStock() {
    const activeTab = getActiveTab();
    return activeTab.stocks.find((stock) => stock.ticker === editingAveragePriceTicker);
}

function openAveragePriceDialog(ticker, field = "average_price") {
    editingAveragePriceTicker = ticker;
    editingNumberField = field;
    const stock = getEditingAveragePriceStock();
    if (!stock) return;

    const title = field === "quantity" ? "수량 입력" : "평단가 입력";
    const label = field === "quantity" ? "수량" : "평단가";
    const titleEl = document.getElementById("averagePriceDialogTitle");
    if (titleEl) titleEl.textContent = title;
    averagePriceInput.setAttribute("aria-label", label);
    averagePriceInput.value = formatAveragePriceInput(stock[field]);
    averagePriceError.textContent = "";
    averagePriceDialogOverlay.classList.add("open");
    averagePriceDialogOverlay.setAttribute("aria-hidden", "false");
    hideContextMenu();
    hideStockContextMenu();
    hideColumnSettings();
    suggestionsEl.classList.remove("open");
    window.setTimeout(() => {
        averagePriceInput.focus();
        averagePriceInput.select();
    }, 0);
}

function closeAveragePriceDialog() {
    averagePriceDialogOverlay.classList.remove("open");
    averagePriceDialogOverlay.setAttribute("aria-hidden", "true");
    averagePriceInput.value = "";
    averagePriceError.textContent = "";
    editingAveragePriceTicker = null;
    editingNumberField = "average_price";
}

function confirmAveragePriceDialog() {
    const stock = getEditingAveragePriceStock();
    if (!stock) return;

    stock[editingNumberField] = normalizeNumberInput(averagePriceInput.value);
    closeAveragePriceDialog();
    render();
}

function openTabDialog(mode, tab = null) {
    tabDialogMode = mode;
    editingTabId = tab ? tab.id : null;
    tabDialogTitle.textContent = mode === "create" ? "새 탭 제목" : "탭 이름 변경";
    confirmTabDialogButton.textContent = mode === "create" ? "생성" : "변경";
    tabNameInput.value = tab ? tab.title : getNextTabTitle();
    tabNameError.textContent = "";
    tabDialogOverlay.classList.add("open");
    tabDialogOverlay.setAttribute("aria-hidden", "false");
    hideContextMenu();
    hideColumnSettings();
    suggestionsEl.classList.remove("open");
    window.setTimeout(() => {
        tabNameInput.focus();
        tabNameInput.select();
    }, 0);
}

function closeTabDialog() {
    tabDialogOverlay.classList.remove("open");
    tabDialogOverlay.setAttribute("aria-hidden", "true");
    tabNameInput.value = "";
    tabNameError.textContent = "";
    editingTabId = null;
}

function confirmTabDialog() {
    const title = tabNameInput.value.trim();
    if (!title) {
        tabNameError.textContent = "탭 제목을 입력하세요.";
        tabNameInput.focus();
        return;
    }

    const safeTitle = title.slice(0, 32);
    if (tabDialogMode === "create") {
        createTabWithTitle(safeTitle);
    } else {
        const tab = state.tabs.find((item) => item.id === editingTabId);
        if (tab) {
            tab.title = safeTitle;
            render();
        }
    }

    closeTabDialog();
}

function showContextMenu(x, y) {
    hideStockContextMenu();
    contextMenu.style.left = `${x}px`;
    contextMenu.style.top = `${y}px`;
    contextMenu.classList.add("open");
    deleteTabButton.disabled = state.tabs.length === 1;
    const rect = contextMenu.getBoundingClientRect();
    const maxLeft = Math.max(8, window.innerWidth - rect.width - 8);
    const maxTop = Math.max(8, window.innerHeight - rect.height - 8);
    const left = Math.min(Math.max(8, x), maxLeft);
    const top = Math.min(Math.max(8, y), maxTop);
    contextMenu.style.left = `${left}px`;
    contextMenu.style.top = `${top}px`;
    editTabButton?.setAttribute("aria-expanded", "true");
}

function hideContextMenu() {
    contextMenu.classList.remove("open");
    editTabButton?.setAttribute("aria-expanded", "false");
}

function toggleEditTabMenu() {
    if (contextMenu.classList.contains("open")) {
        hideContextMenu();
        return;
    }

    contextTabId = state.activeTabId;
    const rect = editTabButton.getBoundingClientRect();
    showContextMenu(rect.left, rect.bottom + 6);
    const maxLeft = Math.max(8, window.innerWidth - contextMenu.offsetWidth - 8);
    const left = Math.min(
        Math.max(8, rect.right - contextMenu.offsetWidth),
        maxLeft
    );
    contextMenu.style.left = `${left}px`;
}

function showStockContextMenu(x, y) {
    hideContextMenu();
    stockContextMenu.style.left = `${x}px`;
    stockContextMenu.style.top = `${y}px`;
    stockContextMenu.classList.add("open");
}

function hideStockContextMenu() {
    stockContextMenu.classList.remove("open");
}

function toggleColumnSettings() {
    const isOpen = columnSettingsMenu.classList.toggle("open");
    columnSettingsButton.setAttribute("aria-expanded", String(isOpen));
}

function hideColumnSettings() {
    columnSettingsMenu.classList.remove("open");
    columnSettingsButton.setAttribute("aria-expanded", "false");
}

function findMatches(query) {
    const normalized = normalizeSearchText(query);
    if (!normalized) return [];

    return STOCKS.filter((stock) => {
        const searchableText = [
            stock.name,
            stock.ticker,
            stock.market,
            stock.category,
        ].map(normalizeSearchText).join(" ");

        return searchableText.includes(normalized);
    });
}

async function fetchMatches(query) {
    const normalized = normalizeSearchText(query);
    if (!normalized) return [];

    const localMatches = findMatches(query);

    try {
        const params = new URLSearchParams({ q: query });
        const response = await fetch(`/api/stocks/search?${params.toString()}`, { cache: "no-store" });
        if (!response.ok) return localMatches;

        const data = await response.json();
        return Array.isArray(data.stocks) ? data.stocks : localMatches;
    } catch (error) {
        console.warn("서버 종목 검색 실패", error);
        return localMatches;
    }
}

async function refreshStockCatalog() {
    try {
        const response = await fetch("/api/stocks", { cache: "no-store" });
        if (!response.ok) return;

        const data = await response.json();
        if (Array.isArray(data.stocks)) {
            STOCKS = data.stocks;
            if (document.activeElement === stockInput && stockInput.value.trim()) {
                renderSuggestions();
            }
        }
    } catch (error) {
        console.warn("종목 목록 갱신 실패", error);
    }
}

function renderSuggestionList(matches) {
    suggestionsEl.innerHTML = "";
    currentSuggestions = matches;
    suggestionStatusEl.textContent = stockInput.value.trim()
        ? `${matches.length}개 종목 검색됨`
        : "";

    if (matches.length === 0) {
        suggestionsEl.classList.remove("open");
        return;
    }

    if (activeSuggestionIndex < 0) {
        activeSuggestionIndex = 0;
    } else {
        activeSuggestionIndex = Math.min(activeSuggestionIndex, matches.length - 1);
    }

    matches.forEach((stock, index) => {
        const item = document.createElement("button");
        item.className = `suggestion-item${index === activeSuggestionIndex ? " active" : ""}`;
        item.type = "button";
        item.setAttribute("role", "option");
        item.setAttribute("aria-selected", String(index === activeSuggestionIndex));
        item.innerHTML = `
            <strong>${escapeHtml(stock.name)}</strong>
            <span class="suggestion-meta">${escapeHtml(stock.category || "-")} · ${escapeHtml(stock.ticker)} · ${escapeHtml(stock.market)}</span>
        `;
        item.addEventListener("click", () => addStock(stock));
        suggestionsEl.appendChild(item);
    });

    suggestionsEl.classList.add("open");

    const activeItem = suggestionsEl.querySelector(".suggestion-item.active");
    if (activeItem) {
        activeItem.scrollIntoView({ block: "nearest" });
    }

    return matches;
}

async function renderSuggestions() {
    const query = stockInput.value;
    const normalized = normalizeSearchText(query);
    const requestId = suggestionRequestId + 1;
    suggestionRequestId = requestId;
    const isNewQuery = normalized !== lastSuggestionQuery;
    lastSuggestionQuery = normalized;
    if (isNewQuery) {
        activeSuggestionIndex = -1;
    }

    if (!normalized) {
        return renderSuggestionList([]);
    }

    suggestionStatusEl.textContent = "검색 중...";

    const promise = fetchMatches(query)
        .then((matches) => {
            if (requestId !== suggestionRequestId) return currentSuggestions;
            return renderSuggestionList(matches);
        })
        .catch((error) => {
            console.warn("종목 검색 실패", error);
            if (requestId !== suggestionRequestId) return currentSuggestions;
            return renderSuggestionList(findMatches(query));
        });

    pendingSuggestionPromise = promise;
    const result = await promise;
    if (pendingSuggestionPromise === promise) {
        pendingSuggestionPromise = null;
    }
    return result;
}

async function ensureSuggestionsReady() {
    if (pendingSuggestionPromise) {
        await pendingSuggestionPromise;
    } else if (currentSuggestions.length === 0 && stockInput.value.trim()) {
        await renderSuggestions();
    }
    return currentSuggestions;
}

function handleStockInputChange() {
    activeSuggestionIndex = -1;
    pendingSuggestionPromise = renderSuggestions();
}

function handleStockInputKeyup(event) {
    if (["ArrowDown", "ArrowUp", "Enter", "Escape"].includes(event.key)) return;
    handleStockInputChange();
}

function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;",
    }[char]));
}

function addStock(stock) {
    const activeTab = getActiveTab();
    activeTab.stocks.unshift({ ...stock, _rowId: createStockRowId() });

    stockInput.value = "";
    activeSuggestionIndex = -1;
    suggestionsEl.classList.remove("open");
    render();
}

function removeStock(stockRowId) {
    const activeTab = getActiveTab();
    activeTab.stocks = activeTab.stocks.filter((stock) => stock._rowId !== stockRowId);
    render();
}

function reorderStock(targetStockRowId) {
    if (!draggedStockRowId || draggedStockRowId === targetStockRowId) return;

    const activeTab = getActiveTab();
    const fromIndex = activeTab.stocks.findIndex((stock) => stock._rowId === draggedStockRowId);
    const toIndex = activeTab.stocks.findIndex((stock) => stock._rowId === targetStockRowId);

    if (fromIndex < 0 || toIndex < 0) return;

    const [moved] = activeTab.stocks.splice(fromIndex, 1);
    activeTab.stocks.splice(toIndex, 0, moved);
    renderStocks();
    saveState();
}

async function updateActiveStocks() {
    if (isUpdating) return;

    const activeTab = getActiveTab();
    const tickers = activeTab.stocks.map((stock) => stock.ticker);
    if (tickers.length === 0) return;

    const fields = getVisibleColumns()
        .map((column) => column.key)
        .filter((key) => FETCHABLE_COLUMN_KEYS.has(key));
    const hasVisibleInlineCharts = getVisibleColumns().some((column) => column.chartType)
        && activeTab.stocks.some((stock) => shouldRenderInlineChart(stock));
    if (fields.length === 0 && !hasVisibleInlineCharts) return;

    isUpdating = true;
    updateButton.disabled = true;
    updateButton.textContent = "업데이트 중...";

    try {
        if (fields.length) {
            const response = await fetch("/api/stocks/update", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ tickers, fields }),
            });

            if (!response.ok) throw new Error("서버 응답 오류");

            const data = await response.json();
            activeTab.stocks = activeTab.stocks.map((stock) => ({
                ...stock,
                ...(data.stocks && data.stocks[stock.ticker] ? {
                    ...Object.fromEntries(fields.map((field) => [field, "-"])),
                    ...data.stocks[stock.ticker],
                } : {
                    ...Object.fromEntries(fields.map((field) => [field, "-"])),
                    ...(fields.includes("price") ? { price: "조회 실패" } : {}),
                }),
            }));
        }
    } catch (error) {
        console.warn("업데이트 실패", error);
        if (fields.length) {
            activeTab.stocks = activeTab.stocks.map((stock) => ({
                ...stock,
                ...Object.fromEntries(fields.map((field) => [field, "-"])),
                ...(fields.includes("price") ? { price: "조회 실패" } : {}),
            }));
        }
    } finally {
        clearInlineChartCache(tickers);
        isUpdating = false;
        updateButton.disabled = false;
        updateButton.textContent = "업데이트";
        render();
    }
}

addTabButton.addEventListener("click", addTab);
editTabButton.addEventListener("click", (event) => {
    event.stopPropagation();
    toggleEditTabMenu();
});
deleteTabButton.addEventListener("click", deleteContextTab);
renameTabButton.addEventListener("click", renameTab);
writeMemoButton.addEventListener("click", openMemoDialog);
updateButton.addEventListener("click", updateActiveStocks);
layoutToggleButton?.addEventListener("click", toggleLayoutStage);
window.addEventListener("load", refreshActiveTabFromStockPrices);
columnSettingsButton.addEventListener("click", toggleColumnSettings);
stockTableScrollEl.addEventListener("scroll", () => syncHorizontalScroll(stockTableScrollEl));
topHorizontalScrollEl.addEventListener("scroll", () => syncHorizontalScroll(topHorizontalScrollEl));
bottomHorizontalScrollEl.addEventListener("scroll", () => syncHorizontalScroll(bottomHorizontalScrollEl));
cancelTabDialogButton.addEventListener("click", closeTabDialog);
confirmTabDialogButton.addEventListener("click", confirmTabDialog);
cancelDeleteTabDialogButton.addEventListener("click", closeDeleteTabDialog);
confirmDeleteTabDialogButton.addEventListener("click", confirmDeleteTabDialog);
cancelMemoDialogButton.addEventListener("click", closeMemoDialog);
confirmMemoDialogButton.addEventListener("click", confirmMemoDialog);
cancelAveragePriceDialogButton.addEventListener("click", closeAveragePriceDialog);
confirmAveragePriceDialogButton.addEventListener("click", confirmAveragePriceDialog);

tabDialogOverlay.addEventListener("click", (event) => {
    if (event.target === tabDialogOverlay) closeTabDialog();
});

deleteTabDialogOverlay.addEventListener("click", (event) => {
    if (event.target === deleteTabDialogOverlay) closeDeleteTabDialog();
});

memoDialogOverlay.addEventListener("click", (event) => {
    if (event.target === memoDialogOverlay) closeMemoDialog();
});

averagePriceDialogOverlay.addEventListener("click", (event) => {
    if (event.target === averagePriceDialogOverlay) closeAveragePriceDialog();
});

tabNameInput.addEventListener("input", () => {
    tabNameError.textContent = "";
});

tabNameInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
        event.preventDefault();
        confirmTabDialog();
    }

    if (event.key === "Escape") {
        event.preventDefault();
        closeTabDialog();
    }
});

deleteTabDialogOverlay.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
        event.preventDefault();
        confirmDeleteTabDialog();
    }

    if (event.key === "Escape") {
        event.preventDefault();
        closeDeleteTabDialog();
    }
});

memoInput.addEventListener("input", () => {
    memoError.textContent = "";
});

memoInput.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
        event.preventDefault();
        closeMemoDialog();
    }

    if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        confirmMemoDialog();
    }
});

averagePriceInput.addEventListener("input", () => {
    const cursorAtEnd = averagePriceInput.selectionStart === averagePriceInput.value.length;
    averagePriceInput.value = formatAveragePriceInput(averagePriceInput.value);
    if (cursorAtEnd) {
        averagePriceInput.setSelectionRange(averagePriceInput.value.length, averagePriceInput.value.length);
    }
    averagePriceError.textContent = "";
});

averagePriceInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
        event.preventDefault();
        confirmAveragePriceDialog();
    }

    if (event.key === "Escape") {
        event.preventDefault();
        closeAveragePriceDialog();
    }
});

stockInput.addEventListener("input", handleStockInputChange);

stockInput.addEventListener("keyup", handleStockInputKeyup);

stockInput.addEventListener("change", handleStockInputChange);

stockInput.addEventListener("compositionend", handleStockInputChange);

stockInput.addEventListener("focus", () => {
    if (stockInput.value.trim()) pendingSuggestionPromise = renderSuggestions();
});

stockInput.addEventListener("keydown", async (event) => {
    if (event.key === "ArrowDown") {
        event.preventDefault();
        await ensureSuggestionsReady();
        if (currentSuggestions.length === 0) return;
        activeSuggestionIndex = activeSuggestionIndex < 0
            ? 0
            : Math.min(activeSuggestionIndex + 1, currentSuggestions.length - 1);
        renderSuggestionList(currentSuggestions);
    }

    if (event.key === "ArrowUp") {
        event.preventDefault();
        await ensureSuggestionsReady();
        if (currentSuggestions.length === 0) return;
        activeSuggestionIndex = activeSuggestionIndex < 0
            ? 0
            : Math.max(activeSuggestionIndex - 1, 0);
        renderSuggestionList(currentSuggestions);
    }

    if (event.key === "Enter") {
        event.preventDefault();
        await ensureSuggestionsReady();
        if (currentSuggestions.length === 0) return;
        const selected = currentSuggestions[activeSuggestionIndex >= 0 ? activeSuggestionIndex : 0];
        if (selected) addStock(selected);
    }

    if (event.key === "Escape") {
        event.preventDefault();
        suggestionsEl.classList.remove("open");
        hideColumnSettings();
        stockInput.blur();
    }
});

document.addEventListener("click", (event) => {
    if (!contextMenu.contains(event.target) && !event.target.closest("#editTabButton")) hideContextMenu();
    if (!stockContextMenu.contains(event.target)) hideStockContextMenu();
    if (!event.target.closest(".search-wrap")) suggestionsEl.classList.remove("open");
    if (!event.target.closest(".column-settings")) hideColumnSettings();
});

document.addEventListener("keydown", (event) => {
    if (event.key !== "F2") return;

    event.preventDefault();
    closeTabDialog();
    closeMemoDialog();
    closeAveragePriceDialog();
    hideContextMenu();
    hideStockContextMenu();
    hideColumnSettings();
    stockInput.focus();
    stockInput.select();
    if (stockInput.value.trim()) renderSuggestions();
});

window.addEventListener("resize", () => {
    hideContextMenu();
    hideStockContextMenu();
    hideColumnSettings();
    updateTopHorizontalScroll();
});

refreshStockCatalog();
applyLayoutToggleStage();
render();
