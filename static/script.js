let STOCKS = Array.isArray(window.CHOICE_STOCKS) ? window.CHOICE_STOCKS : [];
const STORAGE_KEY = "choice.state.v1";
const COLUMN_SCHEMA_VERSION = 6;

const COLUMN_DEFS = [
    { key: "name", label: "종목명", className: "stock-name", width: "minmax(140px, 1.4fr)" },
    { key: "ticker", label: "티커", width: "minmax(82px, 0.9fr)" },
    { key: "price", label: "현재가", width: "minmax(74px, 0.85fr)" },
    { key: "average_price", label: "평단가", className: "average-price-cell", width: "minmax(78px, 0.85fr)" },
    { key: "change_amount", label: "변동", width: "minmax(74px, 0.85fr)", valueClass: getChangeClass },
    { key: "change", label: "오늘변동률", width: "minmax(84px, 0.9fr)", valueClass: getChangeClass },
    { key: "volume", label: "거래량", width: "minmax(74px, 0.85fr)" },
    { key: "close", label: "종가", width: "minmax(74px, 0.85fr)" },
    { key: "open", label: "시가", width: "minmax(74px, 0.85fr)" },
    { key: "high", label: "고가", width: "minmax(74px, 0.85fr)" },
    { key: "low", label: "저가", width: "minmax(74px, 0.85fr)" },
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
];

const COLUMN_GROUPS = [
    {
        key: "price",
        label: "가격",
        columns: ["name", "ticker", "price", "average_price", "change", "volume", "close", "open", "high", "low"],
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
let activeSuggestionIndex = -1;
let draggedTicker = null;
let dragHandleArmedTicker = null;
let draggedColumnKey = null;
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

function createId() {
    return `tab-${Date.now()}-${Math.random().toString(16).slice(2)}`;
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

function loadState() {
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
                activeTabId: saved.activeTabId || saved.tabs[0].id,
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

        const close = document.createElement("button");
        close.className = "tab-close";
        close.type = "button";
        close.textContent = "x";
        close.title = "닫기";

        close.addEventListener("click", (event) => {
            event.stopPropagation();
            openDeleteTabDialog(tab.id);
        });

        tabButton.addEventListener("click", () => {
            state.activeTabId = tab.id;
            hideContextMenu();
            render();
        });

        tabButton.addEventListener("contextmenu", (event) => {
            event.preventDefault();
            contextTabId = tab.id;
            showContextMenu(event.clientX, event.clientY);
        });

        tabButton.addEventListener("dragstart", (event) => {
            if (event.target.closest(".tab-close")) {
                event.preventDefault();
                return;
            }

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

        tabButton.append(title, close);
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
    bindHeaderColumnResize();
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
    const firstRow = stockListEl.querySelector(".stock-row") || tableHeadEl;
    const nameCell = firstRow.querySelector(".stock-name") || tableHeadEl.querySelector('[data-column-key="name"]');
    if (!nameCell || !stockTableScrollEl) return;

    const tableRect = stockTableScrollEl.getBoundingClientRect();
    const nameRect = nameCell.getBoundingClientRect();
    const width = Math.max(0, Math.ceil(nameRect.right - tableRect.left));
    document.documentElement.style.setProperty("--fixed-stock-columns-width", `${width}px`);
}

document.addEventListener("pointermove", (event) => {
    if (!resizingColumn || resizingColumn.pointerId !== event.pointerId) return;

    const isPerformanceColumn = resizingColumn.columnKey?.startsWith("performance_");
    const minWidth = isPerformanceColumn ? 44 : 72;
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

    emptyStateEl.classList.toggle("hidden", activeTab.stocks.length > 0);

    activeTab.stocks.forEach((stock) => {
        const row = document.createElement("div");
        row.className = "stock-row";
        row.draggable = true;
        row.dataset.ticker = stock.ticker;

        const dragHandle = document.createElement("button");
        dragHandle.className = "drag-handle";
        dragHandle.type = "button";
        dragHandle.title = "순서 변경";
        dragHandle.draggable = true;
        dragHandle.textContent = "☰";
        row.appendChild(dragHandle);

        getVisibleColumns().forEach((column) => {
            const cell = document.createElement("span");
            const valueClass = typeof column.valueClass === "function" ? column.valueClass(stock[column.key]) : "";
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
            } else {
                cell.textContent = column.key === "average_price"
                    ? formatAveragePrice(stock.average_price)
                    : stock[column.key] || "-";
            }

            if (column.key === "average_price") {
                cell.tabIndex = 0;
                cell.title = "평단가 입력";
                cell.addEventListener("click", () => openAveragePriceDialog(stock.ticker));
                cell.addEventListener("keydown", (event) => {
                    if (event.key !== "Enter" && event.key !== " ") return;
                    event.preventDefault();
                    openAveragePriceDialog(stock.ticker);
                });
            }

            row.appendChild(cell);
        });

        const deleteButton = document.createElement("button");
        deleteButton.className = "delete-stock";
        deleteButton.type = "button";
        deleteButton.title = "삭제";
        deleteButton.textContent = "x";
        deleteButton.addEventListener("click", () => removeStock(stock.ticker));
        row.appendChild(deleteButton);

        dragHandle.addEventListener("pointerdown", () => {
            dragHandleArmedTicker = stock.ticker;
        });

        dragHandle.addEventListener("pointerup", () => {
            dragHandleArmedTicker = null;
        });

        dragHandle.addEventListener("dragstart", (event) => {
            draggedTicker = stock.ticker;
            row.classList.add("dragging");
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", stock.ticker);
        });

        row.addEventListener("dragstart", (event) => {
            if (dragHandleArmedTicker !== stock.ticker && !event.target.closest(".drag-handle")) {
                event.preventDefault();
                return;
            }
            draggedTicker = stock.ticker;
            row.classList.add("dragging");
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", stock.ticker);
        });

        row.addEventListener("dragend", () => {
            draggedTicker = null;
            dragHandleArmedTicker = null;
            row.classList.remove("dragging");
            saveState();
        });

        row.addEventListener("dragover", (event) => {
            if (!draggedTicker) return;
            event.preventDefault();
            reorderStock(stock.ticker);
        });

        stockListEl.appendChild(row);
    });
}

function getChangeClass(change) {
    if (!change || change === "-") return "";
    if (change.startsWith("+")) return "positive";
    if (change.startsWith("-")) return "negative";
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

function getEditingAveragePriceStock() {
    const activeTab = getActiveTab();
    return activeTab.stocks.find((stock) => stock.ticker === editingAveragePriceTicker);
}

function openAveragePriceDialog(ticker) {
    editingAveragePriceTicker = ticker;
    const stock = getEditingAveragePriceStock();
    if (!stock) return;

    averagePriceInput.value = formatAveragePriceInput(stock.average_price);
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
}

function confirmAveragePriceDialog() {
    const stock = getEditingAveragePriceStock();
    if (!stock) return;

    stock.average_price = normalizeNumberInput(averagePriceInput.value);
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
}

function hideContextMenu() {
    contextMenu.classList.remove("open");
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
    const exists = activeTab.stocks.some((item) => item.ticker === stock.ticker);
    if (!exists) {
        activeTab.stocks.push({ ...stock });
    }

    stockInput.value = "";
    activeSuggestionIndex = -1;
    suggestionsEl.classList.remove("open");
    render();
}

function removeStock(ticker) {
    const activeTab = getActiveTab();
    activeTab.stocks = activeTab.stocks.filter((stock) => stock.ticker !== ticker);
    render();
}

function reorderStock(targetTicker) {
    if (!draggedTicker || draggedTicker === targetTicker) return;

    const activeTab = getActiveTab();
    const fromIndex = activeTab.stocks.findIndex((stock) => stock.ticker === draggedTicker);
    const toIndex = activeTab.stocks.findIndex((stock) => stock.ticker === targetTicker);

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
    if (fields.length === 0) return;

    isUpdating = true;
    updateButton.disabled = true;
    updateButton.textContent = "업데이트 중...";

    try {
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
    } catch (error) {
        console.warn("업데이트 실패", error);
        activeTab.stocks = activeTab.stocks.map((stock) => ({
            ...stock,
            ...Object.fromEntries(fields.map((field) => [field, "-"])),
            ...(fields.includes("price") ? { price: "조회 실패" } : {}),
        }));
    } finally {
        isUpdating = false;
        updateButton.disabled = false;
        updateButton.textContent = "업데이트";
        render();
    }
}

addTabButton.addEventListener("click", addTab);
renameTabButton.addEventListener("click", renameTab);
writeMemoButton.addEventListener("click", openMemoDialog);
updateButton.addEventListener("click", updateActiveStocks);
layoutToggleButton?.addEventListener("click", toggleLayoutStage);
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
    if (!contextMenu.contains(event.target)) hideContextMenu();
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
