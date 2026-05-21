import os
import json
from flask import Flask, render_template_string, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
import uuid
from datetime import datetime
import pytz

app = Flask(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_spreadsheet():
    creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_raw:
        raise RuntimeError("GOOGLE_CREDENTIALS environment variable is missing!")
    
    creds_dict = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open("Aeris Beaute - Stock Opname Master Template")

COUNTER_PREFIXES = (
    "counter:", "counter name:", "name:", "nama:", "nama petugas:",
    "petugas:", "id:", "badge:", "id badge:",
)
LOCATION_PREFIXES = (
    "loc:", "location:", "lokasi:", "kode lokasi:", "precise location:",
)

def normalize_scan_text(code, kind="any"):
    """Normalize QR payload or sheet cell. kind: 'counter', 'location', or 'any'."""
    s = str(code or "").strip().replace("\ufeff", "").replace("\u200b", "").replace("\r", "").replace("\n", "")
    if not s:
        return ""
    lower = s.lower()
    if lower.startswith("http://") or lower.startswith("https://") or "://" in s:
        s = s.rstrip("/").split("/")[-1]
    if "?" in s:
        s = s.split("?")[0]
    if "#" in s:
        s = s.split("#")[0]
    s = s.strip()
    if kind == "counter":
        prefixes = COUNTER_PREFIXES
    elif kind == "location":
        prefixes = LOCATION_PREFIXES
    else:
        prefixes = COUNTER_PREFIXES + LOCATION_PREFIXES
    for prefix in prefixes:
        if s.lower().startswith(prefix):
            s = s[len(prefix):].strip()
            break
    return s.strip()

def get_worksheet_by_names(wb, names):
    for name in names:
        try:
            return wb.worksheet(name)
        except Exception:
            continue
    return None

def read_codes_from_sheet(ws, header_aliases, default_col=0):
    """Read codes from the first matching header column, else column A."""
    try:
        rows = ws.get_all_values()
    except Exception:
        return []

    if not rows:
        return []

    header_aliases_lower = {h.lower() for h in header_aliases}
    col_idx = default_col
    data_start = 0

    header_cells = [str(c).strip().lower() for c in rows[0]]
    for i, cell in enumerate(header_cells):
        if cell in header_aliases_lower:
            col_idx = i
            data_start = 1
            break
    else:
        if header_cells and header_cells[default_col] in header_aliases_lower:
            data_start = 1

    codes = []
    for row in rows[data_start:]:
        if col_idx >= len(row):
            continue
        val = str(row[col_idx]).strip()
        if val:
            codes.append(val)
    return codes

def read_codes_from_column_a(ws, header_aliases):
    """Read counter/location names from column A; skip row 1 only if A1 is a header label."""
    try:
        rows = ws.get_all_values()
    except Exception:
        return []

    if not rows:
        return []

    header_aliases_lower = {h.lower() for h in header_aliases}
    start = 0
    if rows[0]:
        first = str(rows[0][0]).strip().lower()
        if first in header_aliases_lower:
            start = 1

    codes = []
    for row in rows[start:]:
        if not row:
            continue
        val = str(row[0]).strip()
        if val:
            codes.append(val)
    return codes

def build_lookup(codes, sheet_label="", normalize_kind="any"):
    """Map normalized lowercase key -> canonical value. Returns (lookup, warning_messages)."""
    lookup = {}
    warnings = []
    for raw in codes:
        canonical = str(raw).strip()
        if not canonical:
            continue
        key = normalize_scan_text(canonical, normalize_kind).lower()
        if not key:
            continue
        if key in lookup and lookup[key] != canonical:
            prefix = f"{sheet_label}: " if sheet_label else ""
            warnings.append(
                f"{prefix}'{lookup[key]}' bentrok dengan '{canonical}' (kode sama setelah normalisasi)"
            )
        lookup[key] = canonical
        plain = " ".join(canonical.lower().split())
        if plain and plain != key:
            if plain in lookup and lookup[plain] != canonical:
                prefix = f"{sheet_label}: " if sheet_label else ""
                warnings.append(
                    f"{prefix}'{lookup[plain]}' bentrok dengan '{canonical}' (varian huruf sama)"
                )
            lookup[plain] = canonical
    return lookup, warnings

def get_valid_locations(wb):
    ws = get_worksheet_by_names(wb, ("LOCATIONS", "Locations", "Location"))
    if not ws:
        return {}, []
    codes = read_codes_from_sheet(
        ws,
        ("location", "lokasi", "precise location", "locations", "kode lokasi", "code", "kode"),
    )
    return build_lookup(codes, "LOCATIONS", "location")

def resolve_location(code, location_lookup):
    if not code or not location_lookup:
        return None
    trimmed = str(code).strip()
    if trimmed in location_lookup.values():
        return trimmed
    key = normalize_scan_text(trimmed, "location").lower()
    return location_lookup.get(key)

def get_valid_counters(wb):
    ws = get_worksheet_by_names(wb, ("COUNTERS", "Counters", "Counter"))
    if not ws:
        return {}, []
    counter_aliases = (
        "counter", "counter name", "nama", "nama petugas", "petugas",
        "name", "counters", "id badge", "badge", "kode",
    )
    # COUNTERS names live in column A; header-in-column-B must not empty the list.
    codes = read_codes_from_column_a(ws, counter_aliases)
    if not codes:
        codes = read_codes_from_sheet(ws, counter_aliases)
    return build_lookup(codes, "COUNTERS", "counter")

def resolve_counter(name, counter_lookup):
    if not name or not counter_lookup:
        return None
    trimmed = str(name).strip()
    if trimmed in counter_lookup.values():
        return trimmed
    key = normalize_scan_text(trimmed, "counter").lower()
    return counter_lookup.get(key)

def load_sku_tree(wb):
    """Build nested SKU tree from SKU List worksheet."""
    sku_tree = {}
    try:
        sku_worksheet = wb.worksheet("SKU List")
        list_of_lists = sku_worksheet.get_all_values()
        if len(list_of_lists) > 1:
            headers = list_of_lists[0]
            sku_idx = headers.index("SKU Code") if "SKU Code" in headers else 0
            type_idx = headers.index("SKU Type") if "SKU Type" in headers else 1
            cat_idx = headers.index("SKU Category") if "SKU Category" in headers else 2
            for row in list_of_lists[1:]:
                if len(row) <= max(sku_idx, type_idx, cat_idx):
                    continue
                sku_code = str(row[sku_idx]).strip()
                goods_type = str(row[type_idx]).strip() or "General Goods"
                category = str(row[cat_idx]).strip() or "Unassigned"
                if not sku_code:
                    continue
                if goods_type not in sku_tree:
                    sku_tree[goods_type] = {}
                if category not in sku_tree[goods_type]:
                    sku_tree[goods_type][category] = []
                if sku_code not in sku_tree[goods_type][category]:
                    sku_tree[goods_type][category].append(sku_code)
    except Exception:
        pass
    return sku_tree

def resolve_sku(code, sku_tree):
    """Match scanned or submitted SKU to a code in the SKU tree."""
    if not code or not sku_tree:
        return None
    trimmed = str(code).strip()
    for goods_type in sku_tree.values():
        for skus in goods_type.values():
            if trimmed in skus:
                return trimmed
            key = normalize_scan_text(trimmed).lower()
            for sku in skus:
                if normalize_scan_text(sku).lower() == key:
                    return sku
    return None

def get_row_counter_name(row):
    """Read counter from new or legacy column header."""
    return str(row.get("Counter Name") or row.get("Counter Team") or "").strip()

def parse_physical_count(value):
    try:
        return int(value)
    except (ValueError, TypeError):
        return 0

def aggregate_counts_by_sku(records):
    """Sum Physical Count per SKU Code from Raw Counts rows."""
    by_sku = {}
    for row in records:
        sku = str(row.get("SKU Code", "")).strip()
        if not sku:
            continue
        count = parse_physical_count(row.get("Physical Count"))
        if sku not in by_sku:
            by_sku[sku] = {"total": 0, "locations": set()}
        by_sku[sku]["total"] += count
        loc = str(row.get("Precise Location", "")).strip()
        if loc:
            by_sku[sku]["locations"].add(loc)

    rows = [
        {
            "sku": sku,
            "total": data["total"],
            "location_count": len(data["locations"]),
        }
        for sku, data in sorted(by_sku.items())
    ]
    grand_total = sum(r["total"] for r in rows)
    return rows, grand_total

# --- HTML INTERFACE WITH DUPLICATE PROTECTION & AUTO-RESET ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
    <title>Aeris Opname 2026</title>
    <style>
        #scanModal { display: none !important; visibility: hidden !important; pointer-events: none !important; }
        #scanModal.is-open { display: block !important; visibility: visible !important; pointer-events: auto !important; }
        #toast { pointer-events: none !important; }
        #step1Card { position: relative; z-index: 50; pointer-events: auto !important; }
        #step1Card input, #step1Card button { pointer-events: auto !important; touch-action: manipulation; }
        .opname-field {
            font-family: inherit;
            font-size: 1rem;
            line-height: 1.5;
            font-weight: 600;
        }
        .opname-field::placeholder { font-weight: 500; color: rgb(251 191 36 / 0.75); }
        .step-card { border-radius: 0.75rem; border: 1px solid #e4e4e7; background: #fff; box-shadow: 0 1px 2px rgb(0 0 0 / 0.05); overflow: hidden; transition: opacity 0.2s, border-color 0.2s; }
        .step-card--active { border-color: #ddd6fe; border-width: 2px; }
        .step-card__head { display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem 1rem; border-bottom: 1px solid #f4f4f5; }
        .step-card__head--active { background: rgb(245 243 255 / 0.9); border-bottom-color: #ede9fe; }
        .step-card__head--idle { background: #fafafa; }
        .step-card__badge { font-size: 0.75rem; font-weight: 600; padding: 0.125rem 0.5rem; border-radius: 0.25rem; }
        .step-card__badge--active { color: #6d28d9; background: #ede9fe; }
        .step-card__badge--idle { color: #71717a; background: #f4f4f5; }
        .step-card__title { font-size: 0.875rem; font-weight: 600; }
        .step-card__title--active { color: #4c1d95; }
        .step-card__title--idle { color: #52525b; }
        .step-locked { opacity: 0.65; }
        .step-locked button, .step-locked select { pointer-events: none; }
        .step-locked input:not([readonly]) { pointer-events: none; }
        .location-sticky-bar {
            position: sticky;
            top: 0;
            z-index: 45;
            background: rgb(236 253 245 / 0.97);
            backdrop-filter: blur(8px);
            border-bottom: 2px solid #34d399;
            box-shadow: 0 2px 8px rgb(16 185 129 / 0.12);
        }
    </style>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
</head>
<body class="bg-zinc-50 font-sans text-zinc-900 antialiased min-h-screen flex flex-col">
    <script>
        (function () {
            function unblockPage() {
                document.body.style.overflow = '';
                document.body.style.pointerEvents = '';
                var modal = document.getElementById('scanModal');
                if (modal) {
                    modal.hidden = true;
                    modal.classList.remove('is-open');
                    modal.style.display = 'none';
                }
            }
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', unblockPage);
            } else {
                unblockPage();
            }
            window.addEventListener('pageshow', unblockPage);
        })();
    </script>

    <!-- Toast -->
    <div id="toast" class="fixed top-4 left-4 right-4 z-[60] mx-auto max-w-md translate-y-[-120%] opacity-0 transition-all duration-300 pointer-events-none">
        <div id="toastInner" class="rounded-xl px-4 py-3 text-sm font-medium shadow-lg border"></div>
    </div>

    <!-- Header: brand only (inputs live in main — avoids sticky-header tap/focus bugs on mobile) -->
    <header class="relative z-10 bg-white border-b border-zinc-200 shadow-sm">
        <div class="max-w-md lg:max-w-5xl mx-auto px-4 py-3">
            <div class="flex items-center justify-between gap-3">
                <div>
                    <h1 class="text-lg font-bold text-zinc-900 tracking-tight">Aeris Beaute</h1>
                    <p class="text-xs text-zinc-500">Stock Opname 2026</p>
                </div>
                <nav class="flex gap-3 text-xs font-semibold shrink-0">
                    <span class="text-violet-700">Count</span>
                    <a href="/summary" class="text-zinc-500 hover:text-violet-700">Summary</a>
                </nav>
            </div>
        </div>
    </header>

    <div id="locationStickyBar" class="location-sticky-bar hidden" aria-live="polite">
        <div class="max-w-md lg:max-w-5xl mx-auto px-4 py-2.5 flex items-center gap-2">
            <span class="text-xs font-semibold uppercase tracking-wide text-emerald-800 shrink-0">Lokasi</span>
            <span id="locationStickyValue" class="flex-1 min-w-0 text-base font-bold text-emerald-950 truncate"></span>
            <button type="button" id="changeLocationBtn" class="shrink-0 text-xs font-semibold text-emerald-800 bg-emerald-100 hover:bg-emerald-200 px-2.5 py-1 rounded-md">Ubah</button>
        </div>
    </div>

    <div id="lookupWarnings" class="hidden max-w-md lg:max-w-5xl mx-auto px-4">
        <div class="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-900 space-y-1"></div>
    </div>

    <main class="flex-1 max-w-md lg:max-w-5xl w-full mx-auto px-4 py-4 pb-28 lg:pb-6 relative z-0">

        <!-- Stepper: Petugas → Lokasi → Produk → Jumlah -->
        <nav class="flex items-center gap-1 sm:gap-2 mb-5" aria-label="Count progress">
            <div id="step1Indicator" class="flex-1 flex items-center gap-1 min-w-0">
                <span id="step1Dot" class="shrink-0 flex h-7 w-7 sm:h-8 sm:w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-violet-600 bg-violet-600 text-white">1</span>
                <span class="text-[10px] sm:text-xs font-medium text-zinc-700 truncate hidden sm:inline">Petugas</span>
            </div>
            <div class="h-px w-2 sm:w-4 bg-zinc-200 shrink-0" aria-hidden="true"></div>
            <div id="step2Indicator" class="flex-1 flex items-center gap-1 min-w-0">
                <span id="step2Dot" class="shrink-0 flex h-7 w-7 sm:h-8 sm:w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400">2</span>
                <span class="text-[10px] sm:text-xs font-medium text-zinc-400 truncate hidden sm:inline">Lokasi</span>
            </div>
            <div class="h-px w-2 sm:w-4 bg-zinc-200 shrink-0" aria-hidden="true"></div>
            <div id="step3Indicator" class="flex-1 flex items-center gap-1 min-w-0">
                <span id="step3Dot" class="shrink-0 flex h-7 w-7 sm:h-8 sm:w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400">3</span>
                <span class="text-[10px] sm:text-xs font-medium text-zinc-400 truncate hidden sm:inline">Produk</span>
            </div>
            <div class="h-px w-2 sm:w-4 bg-zinc-200 shrink-0" aria-hidden="true"></div>
            <div id="step4Indicator" class="flex-1 flex items-center gap-1 min-w-0">
                <span id="step4Dot" class="shrink-0 flex h-7 w-7 sm:h-8 sm:w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400">4</span>
                <span class="text-[10px] sm:text-xs font-medium text-zinc-400 truncate hidden sm:inline">Jumlah</span>
            </div>
        </nav>

        <!-- Mobile tabs -->
        <div class="lg:hidden flex rounded-lg bg-zinc-100 p-1 mb-4" role="tablist">
            <button id="tabCount" type="button" onclick="switchTab('count')" role="tab" aria-selected="true" class="flex-1 py-2 text-sm font-semibold rounded-md bg-white text-zinc-900 shadow-sm transition">Count</button>
            <button id="tabHistory" type="button" onclick="switchTab('history')" role="tab" aria-selected="false" class="flex-1 py-2 text-sm font-semibold rounded-md text-zinc-500 transition">Riwayat</button>
        </div>

        <div class="lg:grid lg:grid-cols-2 lg:gap-6 lg:items-start">

            <!-- Count panel -->
            <div id="panelCount" class="space-y-4 relative z-0">

                <!-- Step 1: Petugas -->
                <section id="step1Card" class="step-card step-card--active relative z-50">
                    <div class="step-card__head step-card__head--active">
                        <span class="step-card__badge step-card__badge--active">Step 1</span>
                        <h2 class="step-card__title step-card__title--active">Petugas</h2>
                    </div>
                    <div class="p-4">
                        <label for="counterName" class="block text-sm font-medium text-zinc-700 mb-1.5">Nama petugas</label>
                        <div class="flex gap-2">
                            <input type="text" id="counterName" name="counterName" autocomplete="off" inputmode="text" placeholder="Ketik nama atau scan ID badge" class="opname-field flex-1 min-w-0 border border-amber-200 p-3 rounded-lg bg-amber-50 text-amber-900 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
                            <button type="button" id="scanCounterBtn" class="shrink-0 px-4 py-3 rounded-lg bg-violet-600 hover:bg-violet-700 active:bg-violet-800 text-white text-sm font-semibold disabled:opacity-40">Scan</button>
                        </div>
                        <p class="text-xs text-zinc-500 mt-2">Scan QR pada ID badge, atau ketik nama.</p>
                    </div>
                </section>

                <!-- Step 2: Lokasi -->
                <section id="step2Card" class="step-card step-locked" aria-disabled="true">
                    <div id="step2CardHead" class="step-card__head step-card__head--idle">
                        <span class="step-card__badge step-card__badge--idle">Step 2</span>
                        <h2 class="step-card__title step-card__title--idle">Lokasi</h2>
                    </div>
                    <div class="p-4">
                        <label for="location" class="block text-sm font-medium text-zinc-700 mb-1.5">Kode lokasi</label>
                        <div class="flex gap-2">
                            <input type="text" id="location" readonly tabindex="-1" placeholder="Scan kode QR lokasi" aria-describedby="locationHint" class="opname-field flex-1 min-w-0 border border-amber-200 p-3 rounded-lg bg-amber-50 text-amber-900">
                            <button type="button" id="scanLocationBtn" disabled class="shrink-0 px-4 py-3 rounded-lg bg-violet-600 hover:bg-violet-700 active:bg-violet-800 text-white text-sm font-semibold disabled:opacity-40 disabled:cursor-not-allowed">Scan</button>
                        </div>
                        <p id="locationHint" class="text-xs text-zinc-500 mt-2">Pastikan lokasi terscan dahulu sebelum scan produk.</p>
                    </div>
                </section>

                <!-- Step 3: Produk -->
                <section id="step3Card" class="step-card step-locked" aria-disabled="true">
                    <div id="step3CardHead" class="step-card__head step-card__head--idle">
                        <span class="step-card__badge step-card__badge--idle">Step 3</span>
                        <h2 class="step-card__title step-card__title--idle">Produk</h2>
                    </div>
                    <div class="p-4 space-y-3">
                        <button type="button" id="scanSkuBtn" onclick="openScanModal('sku')" disabled class="w-full py-3 px-4 rounded-lg border-2 border-zinc-200 text-zinc-500 font-semibold text-sm transition disabled:opacity-50 disabled:cursor-not-allowed enabled:border-violet-300 enabled:text-violet-700 enabled:bg-violet-50 enabled:hover:bg-violet-100">
                            Scan SKU
                        </button>
                        <div>
                            <label for="skuType" class="block text-sm font-medium text-zinc-700 mb-1">Jenis barang</label>
                            <select id="skuType" onchange="updateCategories()" disabled class="w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 text-sm font-medium focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
                                <option value="">Scan lokasi dulu</option>
                            </select>
                        </div>
                        <div>
                            <label for="skuCategory" class="block text-sm font-medium text-zinc-700 mb-1">Kategori</label>
                            <select id="skuCategory" onchange="updateSkus()" disabled class="w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 text-sm font-medium focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
                                <option value="">Pilih jenis dulu</option>
                            </select>
                        </div>
                        <div>
                            <label for="skuSelector" class="block text-sm font-medium text-zinc-700 mb-1">Kode SKU</label>
                            <select id="skuSelector" onchange="updateStepperUI(); updateSubmitState();" disabled class="w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 text-sm font-medium focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
                                <option value="">Pilih kategori dulu</option>
                            </select>
                        </div>
                    </div>
                </section>

                <!-- Step 4: Jumlah -->
                <section id="step4Card" class="step-card step-locked" aria-disabled="true">
                    <div id="step4CardHead" class="step-card__head step-card__head--idle">
                        <span class="step-card__badge step-card__badge--idle">Step 4</span>
                        <h2 class="step-card__title step-card__title--idle">Jumlah</h2>
                    </div>
                    <div class="p-4 space-y-4">
                        <div class="flex items-center justify-center gap-2 sm:gap-3">
                            <button type="button" onclick="adjustCount(-10)" class="text-xs font-semibold text-zinc-600 bg-zinc-100 hover:bg-zinc-200 px-2.5 py-2 rounded-lg">−10</button>
                            <button type="button" onclick="adjustCount(-1)" class="flex h-11 w-11 sm:h-12 sm:w-12 items-center justify-center rounded-full bg-zinc-100 hover:bg-zinc-200 text-xl font-bold text-zinc-800">−</button>
                            <input type="number" id="count" oninput="updateSubmitState()" onfocus="this.select()" onclick="this.select()" class="w-24 sm:w-28 border border-zinc-200 rounded-xl text-center text-3xl sm:text-4xl font-bold tabular-nums text-zinc-900 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none" value="0" min="0">
                            <button type="button" onclick="adjustCount(1)" class="flex h-11 w-11 sm:h-12 sm:w-12 items-center justify-center rounded-full bg-violet-600 hover:bg-violet-700 text-xl font-bold text-white">+</button>
                            <button type="button" onclick="adjustCount(10)" class="text-xs font-semibold text-zinc-600 bg-zinc-100 hover:bg-zinc-200 px-2.5 py-2 rounded-lg">+10</button>
                        </div>
                        <div>
                            <label for="notes" class="block text-sm font-medium text-zinc-700 mb-1">Catatan <span class="text-zinc-400 font-normal">(opsional)</span></label>
                            <input type="text" id="notes" class="w-full border border-zinc-200 p-3 rounded-lg text-sm text-zinc-900 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none" placeholder="Contoh: kemasan rusak">
                        </div>
                    </div>
                </section>
            </div>

            <!-- History panel -->
            <div id="panelHistory" class="hidden lg:block">
                <div class="step-card lg:sticky lg:top-4">
                    <div class="step-card__head step-card__head--idle">
                        <h2 class="step-card__title step-card__title--idle flex-1">Riwayat</h2>
                        <button type="button" onclick="fetchHistory()" class="text-sm font-semibold text-violet-600 hover:text-violet-800">Refresh</button>
                    </div>
                    <div id="historyContainer" class="p-4 pt-2 space-y-2 max-h-[calc(100vh-12rem)] overflow-y-auto text-sm">
                        <p class="text-zinc-400 text-center py-8">Memuat riwayat…</p>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <div id="footerSubmit" class="fixed bottom-0 left-0 right-0 z-20 bg-white/95 backdrop-blur-md border-t border-zinc-200 p-4 pb-[max(1rem,env(safe-area-inset-bottom))] lg:max-w-5xl lg:mx-auto lg:left-0 lg:right-0">
        <button id="submitBtn" type="button" onclick="submitData()" disabled class="w-full max-w-md mx-auto block bg-emerald-600 hover:bg-emerald-700 disabled:bg-zinc-300 disabled:cursor-not-allowed text-white font-semibold text-base py-3.5 rounded-xl shadow-sm active:scale-[0.98] transition flex items-center justify-center gap-2">
                        <span id="submitBtnLabel">Kirim ke sheet</span>
            <svg id="submitSpinner" class="hidden h-5 w-5 animate-spin" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"></path></svg>
        </button>
    </div>

    <!-- Scanner modal (native hidden until opened — do not rely on Tailwind "hidden" alone) -->
    <div id="scanModal" hidden class="fixed inset-0 z-50" role="dialog" aria-modal="true" aria-labelledby="scanModalTitle">
        <div class="absolute inset-0 bg-black/70" id="scanModalBackdrop"></div>
        <div class="absolute inset-x-0 bottom-0 max-h-[92vh] flex flex-col bg-white rounded-t-2xl shadow-2xl lg:inset-auto lg:top-1/2 lg:left-1/2 lg:-translate-x-1/2 lg:-translate-y-1/2 lg:w-full lg:max-w-lg lg:rounded-2xl">
            <div class="flex items-center justify-between px-4 py-3 border-b border-zinc-100 shrink-0">
                <h3 id="scanModalTitle" class="text-base font-semibold text-zinc-900">Scan QR code</h3>
                <button type="button" onclick="closeScanModal()" class="text-sm font-medium text-zinc-500 hover:text-zinc-800 px-2 py-1">Cancel</button>
            </div>
            <div class="p-4 overflow-hidden flex-1 min-h-0">
                <div id="reader" class="w-full aspect-square max-h-[55vh] mx-auto rounded-xl overflow-hidden bg-black"></div>
                <p id="scanModalHint" class="text-center text-xs text-zinc-500 mt-3">Arahkan kamera ke QR code</p>
            </div>
        </div>
    </div>

    <script id="app-boot-data" type="application/json">{{ boot_data|tojson }}</script>
    <script src="https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js"></script>
    <script>
        const _boot = (function () {
            try {
                return JSON.parse(document.getElementById('app-boot-data').textContent || '{}');
            } catch (e) {
                console.error('Boot data parse failed', e);
                return {};
            }
        })();
        const skuTree = _boot.sku_tree || {};
        let locationLookup = _boot.location_lookup || {};
        let counterLookup = _boot.counter_lookup || {};
        let lookupWarnings = _boot.lookup_warnings || [];
        let validLocations = new Set(_boot.valid_locations || []);
        let validCounters = new Set(_boot.valid_counters || []);
        let currentTarget = '';
        let locationFrozen = false;
        let scannerRunning = false;
        let lastHandledScan = { key: '', at: 0, target: '' };
        let html5QrcodeScanner = null;
        const COUNTER_PREFIXES = [
            'counter:', 'counter name:', 'name:', 'nama:', 'nama petugas:',
            'petugas:', 'id:', 'badge:', 'id badge:',
        ];
        const LOCATION_PREFIXES = [
            'loc:', 'location:', 'lokasi:', 'kode lokasi:', 'precise location:',
        ];

        const FIELD_AMBER = "opname-field flex-1 min-w-0 border border-amber-200 p-3 rounded-lg bg-amber-50 text-amber-900 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none";
        const FIELD_EMERALD = "opname-field flex-1 min-w-0 border border-emerald-200 p-3 rounded-lg bg-emerald-50 text-emerald-900 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none";
        const CLS = {
            counterLocked: FIELD_AMBER,
            counterUnlocked: FIELD_EMERALD,
            locLocked: FIELD_AMBER,
            locUnlocked: FIELD_EMERALD,
            selLocked: "w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 text-sm font-medium",
            selUnlocked: "w-full border border-zinc-200 p-3 rounded-lg bg-white text-zinc-900 text-sm font-medium focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none",
        };

        const STEP_DOT = {
            pending: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400",
            active: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-violet-600 bg-violet-600 text-white",
            done: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-emerald-500 bg-emerald-500 text-white",
        };

        function showLookupWarnings(warnings) {
            const box = document.getElementById('lookupWarnings');
            const inner = box.querySelector('div');
            if (!warnings || !warnings.length) {
                box.classList.add('hidden');
                inner.innerHTML = '';
                return;
            }
            inner.innerHTML = '<p class="font-semibold">Periksa sheet master:</p>' +
                warnings.map(w => `<p>• ${w}</p>`).join('');
            box.classList.remove('hidden');
        }

        function applyLookupData(data) {
            locationLookup = data.locations || {};
            counterLookup = data.counters || {};
            validLocations = new Set(Object.values(locationLookup));
            validCounters = new Set(Object.values(counterLookup));
            if (data.warnings) {
                lookupWarnings = data.warnings;
                showLookupWarnings(lookupWarnings);
            }
        }

        function resolveSku(code) {
            const trimmed = String(code || '').trim();
            if (!trimmed) return null;
            for (const type in skuTree) {
                for (const cat in skuTree[type]) {
                    if (skuTree[type][cat].includes(trimmed)) return trimmed;
                    const key = normalizeScanText(trimmed).toLowerCase();
                    for (const sku of skuTree[type][cat]) {
                        if (normalizeScanText(sku).toLowerCase() === key) return sku;
                    }
                }
            }
            return null;
        }

        async function loadLookups() {
            try {
                const res = await fetch('/api/lookups');
                if (res.ok) {
                    applyLookupData(await res.json());
                }
            } catch (e) {}
        }

        function resetScanDebounce() {
            lastHandledScan = { key: '', at: 0, target: '' };
        }

        function normalizeScanText(code, kind = 'any') {
            let s = String(code || '').trim().replace(/\\ufeff/g, '').replace(/\\u200b/g, '').replace(/\\r/g, '').replace(/\\n/g, '');
            if (!s) return '';
            if (/^https?:\/\//i.test(s) || s.includes('://')) {
                s = s.replace(/\/+$/, '').split('/').pop();
            }
            if (s.includes('?')) s = s.split('?')[0];
            if (s.includes('#')) s = s.split('#')[0];
            s = s.trim();
            let prefixes = COUNTER_PREFIXES.concat(LOCATION_PREFIXES);
            if (kind === 'counter') prefixes = COUNTER_PREFIXES;
            else if (kind === 'location') prefixes = LOCATION_PREFIXES;
            for (const p of prefixes) {
                if (s.toLowerCase().startsWith(p)) {
                    s = s.slice(p.length).trim();
                    break;
                }
            }
            return s.trim();
        }

        function scanKindForTarget(target) {
            if (target === 'counter') return 'counter';
            if (target === 'location') return 'location';
            return 'any';
        }

        function shouldHandleScan(raw, target) {
            const kind = scanKindForTarget(target || currentTarget);
            const key = normalizeScanText(raw, kind).toLowerCase();
            if (!key) return false;
            const now = Date.now();
            if (lastHandledScan.key === key && lastHandledScan.target === target && now - lastHandledScan.at < 500) {
                return false;
            }
            return true;
        }

        function markScanHandled(raw, target) {
            const kind = scanKindForTarget(target || currentTarget);
            const key = normalizeScanText(raw, kind).toLowerCase();
            if (key) lastHandledScan = { key, at: Date.now(), target: target || currentTarget };
        }

        function updateLocationStickyBar() {
            const bar = document.getElementById('locationStickyBar');
            const valEl = document.getElementById('locationStickyValue');
            if (!bar || !valEl) return;
            const loc = resolveLocation(document.getElementById('location').value);
            if (locationFrozen && loc) {
                valEl.textContent = loc;
                bar.classList.remove('hidden');
            } else {
                bar.classList.add('hidden');
            }
        }

        function freezeLocation() {
            locationFrozen = true;
            document.getElementById('scanLocationBtn').disabled = true;
            updateLocationUI();
            updateLocationStickyBar();
        }

        function unfreezeLocation() {
            locationFrozen = false;
            const counterOk = isValidCounter(document.getElementById('counterName').value);
            document.getElementById('scanLocationBtn').disabled = !counterOk;
            updateLocationUI();
            updateLocationStickyBar();
        }

        function requestChangeLocation() {
            if (!locationFrozen) return;
            if (!confirm('Ganti lokasi? Anda perlu scan QR lokasi baru sebelum menghitung produk.')) return;
            const locInput = document.getElementById('location');
            locInput.value = '';
            locInput.className = CLS.locLocked;
            unfreezeLocation();
            lockSkuFields();
            updateStepperUI();
            updateSubmitState();
        }

        function updateLocationUI() {
            const locInput = document.getElementById('location');
            const hint = document.getElementById('locationHint');
            const counterOk = isValidCounter(document.getElementById('counterName').value);
            if (!counterOk) {
                if (hint) hint.textContent = 'Pastikan lokasi terscan dahulu sebelum scan produk.';
                locInput.placeholder = 'Scan kode QR lokasi';
            } else if (locInput.value.trim()) {
                if (locationFrozen) {
                    if (hint) hint.textContent = 'Lokasi terkunci di atas. Gunakan Ubah hanya jika pindah rak.';
                } else if (hint) {
                    hint.textContent = 'Lokasi terisi. Tekan Scan untuk mengganti.';
                }
                locInput.placeholder = locInput.value;
            } else {
                if (hint) hint.textContent = 'Scan kode QR lokasi sebelum scan produk.';
                locInput.placeholder = 'Scan kode QR lokasi';
            }
            updateLocationStickyBar();
        }

        function lockSkuFields() {
            document.getElementById('scanSkuBtn').disabled = true;
            ['skuType', 'skuCategory', 'skuSelector'].forEach(id => {
                const el = document.getElementById(id);
                el.disabled = true;
                el.innerHTML = id === 'skuType'
                    ? '<option value="">Scan lokasi dulu</option>'
                    : id === 'skuCategory'
                    ? '<option value="">Pilih jenis dulu</option>'
                    : '<option value="">Pilih kategori dulu</option>';
                el.className = CLS.selLocked;
            });
        }

        function resetLocationAndSku() {
            const locInput = document.getElementById('location');
            locInput.value = '';
            locInput.className = CLS.locLocked;
            unfreezeLocation();
            lockSkuFields();
            updateStepperUI();
            updateSubmitState();
        }

        function enableSkuFields() {
            document.getElementById('scanSkuBtn').disabled = false;
            const typeSelect = document.getElementById('skuType');
            typeSelect.disabled = false;
            typeSelect.className = CLS.selUnlocked;
            typeSelect.innerHTML = '<option value="">Pilih jenis barang</option>';
            Object.keys(skuTree).sort().forEach(type => {
                typeSelect.options[typeSelect.options.length] = new Option(type, type);
            });
            const catSelect = document.getElementById('skuCategory');
            catSelect.disabled = true;
            catSelect.innerHTML = '<option value="">Pilih jenis dulu</option>';
            catSelect.className = CLS.selLocked;
            const skuSelect = document.getElementById('skuSelector');
            skuSelect.disabled = true;
            skuSelect.innerHTML = '<option value="">Pilih kategori dulu</option>';
            skuSelect.className = CLS.selLocked;
        }

        function resetSkuAndCount() {
            document.getElementById('count').value = '0';
            document.getElementById('notes').value = '';
            enableSkuFields();
            updateStepperUI();
            updateSubmitState();
        }

        async function syncUIState() {
            const counterInput = document.getElementById('counterName');
            const locInput = document.getElementById('location');
            const historyContainer = document.getElementById('historyContainer');

            if (isValidCounter(counterInput.value)) {
                counterInput.className = CLS.counterUnlocked;
                counterInput.removeAttribute('readonly');
                if (!locationFrozen) {
                    document.getElementById('scanLocationBtn').disabled = false;
                }
                setStepCardEnabled('step2Card', true);
                updateLocationUI();
                if (locInput.value.trim() && isValidLocation(locInput.value)) {
                    locInput.className = CLS.locUnlocked;
                    if (!locationFrozen) {
                        unlockFormForLocation();
                    } else {
                        document.getElementById('scanSkuBtn').disabled = false;
                        updateLocationUI();
                    }
                } else {
                    if (locInput.value.trim() && !isValidLocation(locInput.value)) {
                        locInput.value = '';
                    }
                    unfreezeLocation();
                    lockSkuFields();
                    locInput.className = CLS.locLocked;
                }
                fetchHistory();
            } else {
                counterInput.className = CLS.counterLocked;
                counterInput.removeAttribute('readonly');
                document.getElementById('scanLocationBtn').disabled = true;
                setStepCardEnabled('step2Card', false);
                locInput.value = '';
                locInput.className = CLS.locLocked;
                unfreezeLocation();
                lockSkuFields();
                historyContainer.innerHTML = '<p class="text-zinc-400 text-center py-8">Scan badge atau ketik nama petugas.</p>';
            }
            updateStepperUI();
            updateSubmitState();
        }

        function getScanner() {
            if (!html5QrcodeScanner && typeof Html5Qrcode !== 'undefined') {
                html5QrcodeScanner = new Html5Qrcode('reader');
            }
            return html5QrcodeScanner;
        }

        function bindScanButtons() {
            const counterBtn = document.getElementById('scanCounterBtn');
            const locationBtn = document.getElementById('scanLocationBtn');
            const counterInput = document.getElementById('counterName');
            if (counterBtn) {
                counterBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    openScanModal('counter');
                });
            }
            if (locationBtn) {
                locationBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    openScanModal('location');
                });
            }
            if (counterInput) {
                counterInput.addEventListener('input', onCounterNameInput);
                counterInput.addEventListener('blur', onCounterNameInput);
                counterInput.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') onCounterNameInput(e);
                });
            }
            const changeLocBtn = document.getElementById('changeLocationBtn');
            if (changeLocBtn) {
                changeLocBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    requestChangeLocation();
                });
            }
            const backdrop = document.getElementById('scanModalBackdrop');
            if (backdrop) {
                backdrop.addEventListener('click', () => closeScanModal());
            }
        }

        function resetPageInteractionState() {
            document.body.style.overflow = '';
            document.body.style.pointerEvents = '';
            const modal = document.getElementById('scanModal');
            if (!modal) return;
            modal.hidden = true;
            modal.classList.remove('is-open');
            modal.style.display = 'none';
        }

        async function initApp() {
            resetPageInteractionState();
            bindScanButtons();
            showLookupWarnings(lookupWarnings);
            await loadLookups();
            await syncUIState();
            switchTab('count');
        }

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => { initApp().catch(console.error); });
        } else {
            initApp().catch(console.error);
        }

        function switchTab(tab) {
            const isLg = window.matchMedia('(min-width: 1024px)').matches;
            const panelCount = document.getElementById('panelCount');
            const panelHistory = document.getElementById('panelHistory');
            const footer = document.getElementById('footerSubmit');
            const tabCount = document.getElementById('tabCount');
            const tabHistory = document.getElementById('tabHistory');

            if (isLg) {
                panelCount.classList.remove('hidden');
                panelHistory.classList.remove('hidden');
                footer.classList.remove('hidden');
                return;
            }

            const showCount = tab === 'count';
            panelCount.classList.toggle('hidden', !showCount);
            panelHistory.classList.toggle('hidden', showCount);
            footer.classList.toggle('hidden', !showCount);

            tabCount.className = showCount
                ? 'flex-1 py-2 text-sm font-semibold rounded-md bg-white text-zinc-900 shadow-sm transition'
                : 'flex-1 py-2 text-sm font-semibold rounded-md text-zinc-500 transition';
            tabHistory.className = showCount
                ? 'flex-1 py-2 text-sm font-semibold rounded-md text-zinc-500 transition'
                : 'flex-1 py-2 text-sm font-semibold rounded-md bg-white text-zinc-900 shadow-sm transition';
            tabCount.setAttribute('aria-selected', showCount);
            tabHistory.setAttribute('aria-selected', !showCount);
        }

        window.addEventListener('resize', () => switchTab(document.getElementById('tabHistory').getAttribute('aria-selected') === 'true' ? 'history' : 'count'));

        function setStepCardVisual(cardId, active) {
            const card = document.getElementById(cardId);
            const head = document.getElementById(cardId + 'Head');
            if (!card) return;
            card.classList.toggle('step-card--active', active);
            if (head) {
                head.classList.toggle('step-card__head--active', active);
                head.classList.toggle('step-card__head--idle', !active);
                const badge = head.querySelector('.step-card__badge');
                const title = head.querySelector('.step-card__title');
                if (badge) {
                    badge.classList.toggle('step-card__badge--active', active);
                    badge.classList.toggle('step-card__badge--idle', !active);
                }
                if (title) {
                    title.classList.toggle('step-card__title--active', active);
                    title.classList.toggle('step-card__title--idle', !active);
                }
            }
        }

        function setStepCardEnabled(cardId, enabled) {
            if (cardId === 'step1Card') return;
            const card = document.getElementById(cardId);
            if (!card) return;
            card.classList.toggle('step-locked', !enabled);
            card.setAttribute('aria-disabled', String(!enabled));
            setStepCardVisual(cardId, enabled);
        }

        function setStepDot(num, state) {
            document.getElementById(`step${num}Dot`).className = STEP_DOT[state];
            const label = document.getElementById(`step${num}Indicator`).querySelector('span:last-child');
            if (label) {
                label.className = state === 'pending'
                    ? 'text-xs font-medium text-zinc-400 truncate hidden sm:inline'
                    : 'text-xs font-medium text-zinc-700 truncate hidden sm:inline';
            }
        }

        function updateStepperUI() {
            const counterOk = isValidCounter(document.getElementById('counterName').value);
            const loc = document.getElementById('location').value.trim();
            const sku = document.getElementById('skuSelector').value;

            setStepCardEnabled('step2Card', counterOk);

            if (!counterOk) {
                setStepDot(1, 'active');
                setStepDot(2, 'pending');
                setStepDot(3, 'pending');
                setStepDot(4, 'pending');
                setStepCardEnabled('step3Card', false);
                setStepCardEnabled('step4Card', false);
                return;
            }

            if (!loc) {
                setStepDot(1, 'done');
                setStepDot(2, 'active');
                setStepDot(3, 'pending');
                setStepDot(4, 'pending');
                setStepCardEnabled('step3Card', false);
                setStepCardEnabled('step4Card', false);
            } else if (!sku) {
                setStepDot(1, 'done');
                setStepDot(2, 'done');
                setStepDot(3, 'active');
                setStepDot(4, 'pending');
                setStepCardEnabled('step3Card', true);
                setStepCardEnabled('step4Card', false);
            } else {
                setStepDot(1, 'done');
                setStepDot(2, 'done');
                setStepDot(3, 'done');
                setStepDot(4, 'active');
                setStepCardEnabled('step3Card', true);
                setStepCardEnabled('step4Card', true);
            }
        }

        function updateSubmitState() {
            const ready = isValidCounter(document.getElementById('counterName').value)
                && resolveLocation(document.getElementById('location').value)
                && resolveSku(document.getElementById('skuSelector').value)
                && document.getElementById('count').value !== '';
            const btn = document.getElementById('submitBtn');
            btn.disabled = !ready || btn.dataset.loading === '1';
        }

        function resolveCounter(name) {
            const trimmed = String(name || '').trim();
            if (!trimmed) return null;
            const key = normalizeScanText(trimmed, 'counter').toLowerCase();
            if (key && counterLookup[key]) return counterLookup[key];
            const plain = trimmed.toLowerCase().replace(/\s+/g, ' ');
            if (plain && counterLookup[plain]) return counterLookup[plain];
            if (validCounters.has(trimmed)) return trimmed;
            for (const c of validCounters) {
                if (String(c).toLowerCase() === plain) return c;
            }
            return null;
        }

        function isValidCounter(name) {
            return !!resolveCounter(name);
        }

        async function applyCounterScan(text) {
            if (!Object.keys(counterLookup).length) {
                await loadLookups();
            }
            if (!Object.keys(counterLookup).length) {
                showToast('Daftar petugas belum dimuat. Hubungi admin.', 'error');
                return false;
            }
            const resolved = resolveCounter(text);
            if (!resolved) {
                const scanned = normalizeScanText(text, 'counter') || String(text).trim();
                showToast(`Petugas tidak dikenali: "${scanned.slice(0, 40)}". Ketik nama sesuai dengan ID badge.`, 'warning');
                return false;
            }
            document.getElementById('counterName').value = resolved;
            unlockAfterCounter();
            await syncUIState();
            return true;
        }

        let lastCounterNameToast = '';
        async function onCounterNameInput(ev) {
            const counterInput = document.getElementById('counterName');
            const trimmed = String(counterInput.value || '').trim();
            const isCommit = ev && (ev.type === 'blur' || (ev.type === 'keydown' && ev.key === 'Enter'));

            if (!trimmed) {
                lastCounterNameToast = '';
                syncUIState();
                return;
            }

            if (!Object.keys(counterLookup).length) {
                await loadLookups();
            }

            const resolved = resolveCounter(counterInput.value);
            if (!resolved) {
                if (isCommit) {
                    let msg;
                    let type = 'warning';
                    if (!Object.keys(counterLookup).length) {
                        msg = 'Daftar petugas belum dimuat. Hubungi admin.';
                        type = 'error';
                    } else {
                        const key = normalizeScanText(trimmed, 'counter') || trimmed;
                        msg = `Petugas tidak dikenali: "${String(key).slice(0, 40)}". Ketik nama sesuai dengan ID badge.`;
                    }
                    if (lastCounterNameToast !== msg) {
                        showToast(msg, type);
                        lastCounterNameToast = msg;
                    }
                }
            } else {
                lastCounterNameToast = '';
                if (isCommit && counterInput.value !== resolved) {
                    counterInput.value = resolved;
                }
            }
            syncUIState();
        }

        function unlockAfterCounter() {
            document.getElementById('counterName').className = CLS.counterUnlocked;
            document.getElementById('scanLocationBtn').disabled = false;
            updateLocationUI();
            updateStepperUI();
            updateSubmitState();
            fetchHistory();
        }

        let toastTimer;
        function showToast(message, type = 'success') {
            const toast = document.getElementById('toast');
            const inner = document.getElementById('toastInner');
            const styles = {
                success: 'bg-emerald-50 text-emerald-900 border-emerald-200',
                error: 'bg-rose-50 text-rose-900 border-rose-200',
                warning: 'bg-amber-50 text-amber-900 border-amber-200',
            };
            inner.className = `rounded-xl px-4 py-3 text-sm font-medium shadow-lg border ${styles[type] || styles.success}`;
            inner.textContent = message;
            toast.classList.remove('translate-y-[-120%]', 'opacity-0');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => {
                toast.classList.add('translate-y-[-120%]', 'opacity-0');
            }, type === 'warning' ? 6000 : 3500);
        }

        function resolveLocation(code) {
            const trimmed = String(code || '').trim();
            if (!trimmed) return null;
            if (validLocations.has(trimmed)) return trimmed;
            const key = normalizeScanText(trimmed, 'location').toLowerCase();
            return locationLookup[key] || null;
        }

        function isValidLocation(code) {
            return !!resolveLocation(code);
        }

        function applyLocationScan(text) {
            if (!isValidCounter(document.getElementById('counterName').value)) {
                showToast('Scan kode QR lokasi', 'warning');
                return false;
            }
            if (!Object.keys(locationLookup).length) {
                showToast('Daftar lokasi belum dimuat. Hubungi admin.', 'error');
                return false;
            }
            const resolved = resolveLocation(text);
            if (!resolved) {
                const scanned = normalizeScanText(text, 'location') || String(text).trim();
                showToast(`Lokasi tidak dikenali: "${scanned.slice(0, 40)}". Periksa tab LOCATIONS.`, 'warning');
                return false;
            }
            document.getElementById('location').value = resolved;
            unlockFormForLocation();
            return true;
        }

        function unlockFormForLocation() {
            const locInput = document.getElementById('location');
            locInput.className = CLS.locUnlocked;
            freezeLocation();
            enableSkuFields();
            updateStepperUI();
            updateSubmitState();
        }


        function updateCategories() {
            const typeVal = document.getElementById('skuType').value;
            const catSelect = document.getElementById('skuCategory');
            const skuSelect = document.getElementById('skuSelector');

            catSelect.innerHTML = '<option value="">Pilih kategori</option>';
            skuSelect.innerHTML = '<option value="">Pilih SKU</option>';
            skuSelect.disabled = true;
            skuSelect.className = CLS.selLocked;

            if (!typeVal || !skuTree[typeVal]) {
                catSelect.disabled = true;
                catSelect.className = CLS.selLocked;
                updateStepperUI();
                updateSubmitState();
                return;
            }

            catSelect.disabled = false;
            catSelect.className = CLS.selUnlocked;
            Object.keys(skuTree[typeVal]).sort().forEach(cat => {
                catSelect.options[catSelect.options.length] = new Option(cat, cat);
            });
            updateStepperUI();
            updateSubmitState();
        }

        function updateSkus() {
            const typeVal = document.getElementById('skuType').value;
            const catVal = document.getElementById('skuCategory').value;
            const skuSelect = document.getElementById('skuSelector');

            skuSelect.innerHTML = '<option value="">Pilih SKU</option>';

            if (!catVal || !skuTree[typeVal] || !skuTree[typeVal][catVal]) {
                skuSelect.disabled = true;
                skuSelect.className = CLS.selLocked;
                updateStepperUI();
                updateSubmitState();
                return;
            }

            skuSelect.disabled = false;
            skuSelect.className = CLS.selUnlocked;
            skuTree[typeVal][catVal].sort().forEach(sku => {
                skuSelect.options[skuSelect.options.length] = new Option(sku, sku);
            });
            updateStepperUI();
            updateSubmitState();
        }

        async function ensureScannerStopped() {
            const scanner = getScanner();
            if (!scanner) return;
            try {
                await scanner.stop();
            } catch (e) {}
            try {
                scanner.clear();
            } catch (e) {}
            scannerRunning = false;
        }

        async function openScanModal(target) {
            if (target === 'location' && locationFrozen) {
                showToast('Lokasi terkunci. Tekan Ubah di bar hijau atas untuk ganti lokasi.', 'warning');
                return;
            }
            if (target === 'location' && document.getElementById('scanLocationBtn').disabled) return;
            if (target === 'sku' && document.getElementById('scanSkuBtn').disabled) return;
            if (typeof Html5Qrcode === 'undefined') {
                showToast('Pemindai QR tidak termuat. Muat ulang halaman.', 'error');
                return;
            }

            currentTarget = target;
            const scanTitles = {
                counter: 'Scan ID badge',
                location: 'Scan lokasi',
                sku: 'Scan SKU',
            };
            const scanHints = {
                counter: 'Arahkan kamera ke QR pada ID badge petugas',
                location: 'Arahkan kamera ke QR lokasi',
                sku: 'Arahkan kamera ke QR SKU',
            };
            document.getElementById('scanModalTitle').textContent = scanTitles[target] || 'Scan QR';
            document.getElementById('scanModalHint').textContent = scanHints[target] || 'Arahkan kamera ke QR code';
            const modal = document.getElementById('scanModal');
            modal.hidden = false;
            modal.classList.add('is-open');
            modal.style.display = 'block';
            document.body.style.overflow = 'hidden';
            resetScanDebounce();

            await ensureScannerStopped();
            await new Promise(r => setTimeout(r, 200));

            const scanner = getScanner();
            if (!scanner) {
                showToast('Pemindai QR tidak siap. Muat ulang halaman.', 'error');
                closeScanModal();
                return;
            }

            try {
                scannerRunning = true;
                await scanner.start(
                    { facingMode: "environment" },
                    { fps: 15, qrbox: { width: 250, height: 250 } },
                    async (decodedText) => {
                        const text = decodedText.trim();
                        if (!text || !shouldHandleScan(text, currentTarget)) return;

                        let shouldClose = false;
                        if (currentTarget === 'counter') {
                            shouldClose = await applyCounterScan(text);
                        } else if (currentTarget === 'location') {
                            shouldClose = applyLocationScan(text);
                        } else if (currentTarget === 'sku') {
                            let found = false;
                            const skuText = normalizeScanText(text) || text;
                            for (const type in skuTree) {
                                for (const cat in skuTree[type]) {
                                    if (skuTree[type][cat].includes(skuText) || skuTree[type][cat].includes(text)) {
                                        document.getElementById('skuType').value = type;
                                        updateCategories();
                                        document.getElementById('skuCategory').value = cat;
                                        updateSkus();
                                        document.getElementById('skuSelector').value = skuTree[type][cat].includes(skuText) ? skuText : text;
                                        found = true;
                                        break;
                                    }
                                }
                                if (found) break;
                            }
                            if (!found) {
                                showToast('SKU tidak dikenali: ' + skuText, 'warning');
                            } else {
                                const resolved = resolveSku(document.getElementById('skuSelector').value);
                                if (resolved) {
                                    document.getElementById('skuSelector').value = resolved;
                                }
                                updateStepperUI();
                                updateSubmitState();
                                shouldClose = true;
                            }
                        }
                        if (shouldClose) {
                            markScanHandled(text, currentTarget);
                            await closeScanModal();
                        }
                    },
                    () => {}
                );
            } catch (err) {
                scannerRunning = false;
                showToast('Could not start camera. Check permissions.', 'error');
                closeScanModal();
            }
        }

        async function closeScanModal() {
            await ensureScannerStopped();
            const modal = document.getElementById('scanModal');
            modal.hidden = true;
            modal.classList.remove('is-open');
            modal.style.display = 'none';
            document.body.style.overflow = '';
        }

        window.openScanModal = openScanModal;
        window.closeScanModal = closeScanModal;
        window.onCounterNameInput = onCounterNameInput;
        window.switchTab = switchTab;
        window.submitData = submitData;
        window.fetchHistory = fetchHistory;
        window.updateCategories = updateCategories;
        window.updateSkus = updateSkus;
        window.adjustCount = adjustCount;

        function adjustCount(amount) {
            const countInput = document.getElementById('count');
            let currentVal = parseInt(countInput.value) || 0;
            currentVal += amount;
            if (currentVal < 0) currentVal = 0;
            countInput.value = currentVal;
            updateSubmitState();
        }

        function escapeHtml(str) {
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');
        }

        async function fetchHistory() {
            const counterName = resolveCounter(document.getElementById('counterName').value) || '';
            const container = document.getElementById('historyContainer');

            if (!counterName) {
                container.innerHTML = '<p class="text-zinc-400 text-center py-8">Scan badge atau ketik nama petugas.</p>';
                return;
            }

            container.innerHTML = '<p class="text-zinc-400 text-center py-8 animate-pulse">Memuat…</p>';

            try {
                const response = await fetch(`/history?name=${encodeURIComponent(counterName)}`);
                const data = await response.json();

                if (data.length === 0) {
                    container.innerHTML = '<p class="text-zinc-400 text-center py-8">Belum ada catatan untuk petugas ini.</p>';
                    return;
                }

                container.innerHTML = data.map(item => `
                    <div class="border border-zinc-100 rounded-lg p-3 hover:bg-zinc-50/80 transition">
                        <div class="flex justify-between items-start gap-2">
                            <div class="min-w-0 flex-1">
                                <div class="flex justify-between items-baseline gap-2">
                                    <span class="text-sm font-semibold text-violet-700 truncate">${escapeHtml(item.location)}</span>
                                    <span class="text-xl font-bold tabular-nums text-zinc-900 shrink-0">${escapeHtml(String(item.count))}</span>
                                </div>
                                <p class="text-sm font-medium text-zinc-700 truncate mt-0.5">${escapeHtml(item.sku)}</p>
                                ${item.timestamp ? `<p class="text-xs text-zinc-400 mt-0.5">${escapeHtml(item.timestamp)}</p>` : ''}
                                ${item.notes ? `<p class="text-xs text-zinc-500 mt-0.5 truncate">${escapeHtml(item.notes)}</p>` : ''}
                            </div>
                        </div>
                        <div class="flex gap-2 mt-2 pt-2 border-t border-zinc-100">
                            <button type="button" onclick="editItem('${escapeHtml(item.id)}', ${parseInt(item.count) || 0})" class="text-xs font-medium text-amber-700 hover:text-amber-900">Edit</button>
                            <button type="button" onclick="deleteItem('${escapeHtml(item.id)}')" class="text-xs font-medium text-rose-600 hover:text-rose-800">Delete</button>
                        </div>
                    </div>
                `).join('');
            } catch (err) {
                container.innerHTML = '<p class="text-rose-600 text-center py-8">Failed to load history.</p>';
            }
        }

        async function submitData() {
            const locInput = document.getElementById('location').value;
            const skuInput = document.getElementById('skuSelector').value;
            const countInput = document.getElementById('count').value;
            const btn = document.getElementById('submitBtn');
            const label = document.getElementById('submitBtnLabel');
            const spinner = document.getElementById('submitSpinner');

            const counterName = resolveCounter(document.getElementById('counterName').value);
            const resolvedLocation = resolveLocation(locInput);
            const resolvedSku = resolveSku(skuInput);

            if (!counterName) {
                showToast('Scan ID badge petugas dulu.', 'warning');
                return;
            }

            if (!resolvedLocation || !resolvedSku || countInput === '') {
                showToast('Lengkapi lokasi, SKU, dan jumlah sebelum submit.', 'warning');
                return;
            }

            btn.disabled = true;
            btn.dataset.loading = '1';
            label.textContent = 'Menyimpan…';
            spinner.classList.remove('hidden');

            const payload = {
                counter_name: counterName,
                location: resolvedLocation,
                sku: resolvedSku,
                count: countInput,
                notes: document.getElementById('notes').value.trim()
            };

            try {
                const response = await fetch('/submit', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const result = await response.json();

                if (response.status === 409) {
                    showToast(result.message || 'Data duplikat. Gunakan Edit di Riwayat.', 'warning');
                } else if (response.status === 400) {
                    showToast(result.message || 'Lokasi tidak valid.', 'warning');
                } else if (response.ok) {
                    showToast('Count saved successfully.', 'success');
                    resetSkuAndCount();
                    fetchHistory();
                } else {
                    showToast('Sync failed. Try again.', 'error');
                }
            } catch (err) {
                showToast('Network error. Try again.', 'error');
            } finally {
                btn.dataset.loading = '0';
                label.textContent = 'Kirim ke sheet';
                spinner.classList.add('hidden');
                updateSubmitState();
            }
        }

        async function editItem(logId, currentCount) {
            const newCount = prompt(`Enter new physical count for this row:`, currentCount);
            if (newCount === null || newCount.trim() === "" || isNaN(newCount)) return;
            
            try {
                const response = await fetch('/edit', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ id: logId, count: parseInt(newCount) })
                });
                if (response.ok) {
                    fetchHistory();
                } else {
                    alert('Failed to update record on sheet.');
                }
            } catch (err) {
                alert('Network error, update aborted.');
            }
        }

        async function deleteItem(logId) {
            if (!confirm("Are you sure you want to delete this specific count record from the master sheet?")) return;
            
            try {
                const response = await fetch('/delete', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ id: logId })
                });
                if (response.ok) {
                    fetchHistory();
                } else {
                    alert('Failed to delete record from sheet.');
                }
            } catch (err) {
                alert('Network error, delete aborted.');
            }
        }
    </script>
</body>
</html>
"""

SUMMARY_HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SKU Summary — Aeris Opname</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
</head>
<body class="bg-zinc-50 font-sans text-zinc-900 antialiased min-h-screen">

    <header class="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-zinc-200 shadow-sm">
        <div class="max-w-3xl mx-auto px-4 py-3">
            <div class="flex items-center justify-between gap-3">
                <div>
                    <h1 class="text-lg font-bold text-zinc-900 tracking-tight">SKU Summary</h1>
                    <p class="text-xs text-zinc-500">Running totals from Raw Counts</p>
                </div>
                <nav class="flex gap-3 text-xs font-semibold shrink-0">
                    <a href="/" class="text-zinc-500 hover:text-violet-700">Count</a>
                    <span class="text-violet-700">Summary</span>
                </nav>
            </div>
        </div>
    </header>

    <main class="max-w-3xl mx-auto px-4 py-4 space-y-4">
        <div class="flex flex-col sm:flex-row sm:items-center gap-3">
            <input type="search" id="searchSku" placeholder="Cari SKU…" oninput="filterTable()"
                class="flex-1 border border-zinc-200 rounded-lg px-3 py-2.5 text-sm focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
            <button type="button" onclick="loadSummary()" class="shrink-0 px-4 py-2.5 rounded-lg bg-violet-600 hover:bg-violet-700 text-white text-sm font-semibold transition">
                Refresh
            </button>
        </div>

        <div class="bg-white rounded-xl border border-zinc-200 shadow-sm overflow-hidden">
            <div class="overflow-x-auto">
                <table class="w-full text-sm">
                    <thead class="bg-zinc-50 border-b border-zinc-200">
                        <tr>
                            <th class="text-left px-4 py-3 font-semibold text-zinc-700">SKU Code</th>
                            <th class="text-right px-4 py-3 font-semibold text-zinc-700">Total qty</th>
                            <th class="text-right px-4 py-3 font-semibold text-zinc-700 hidden sm:table-cell">Lokasi</th>
                        </tr>
                    </thead>
                    <tbody id="summaryBody">
                        <tr><td colspan="3" class="px-4 py-8 text-center text-zinc-400">Loading…</td></tr>
                    </tbody>
                    <tfoot id="summaryFoot" class="bg-violet-50 border-t border-violet-100 hidden">
                        <tr>
                            <td class="px-4 py-3 font-bold text-zinc-800">Grand total</td>
                            <td id="grandTotal" class="px-4 py-3 text-right font-bold text-violet-800 tabular-nums">0</td>
                            <td id="skuCount" class="px-4 py-3 text-right text-sm text-zinc-600 hidden sm:table-cell"></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
        <p id="lastUpdated" class="text-xs text-zinc-400 text-center"></p>
    </main>

    <script>
        let summaryRows = [];

        function escapeHtml(str) {
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');
        }

        function renderTable(rows, grandTotal) {
            const body = document.getElementById('summaryBody');
            const foot = document.getElementById('summaryFoot');

            if (!rows.length) {
                body.innerHTML = '<tr><td colspan="3" class="px-4 py-8 text-center text-zinc-400">Belum ada data di Raw Counts.</td></tr>';
                foot.classList.add('hidden');
                return;
            }

            body.innerHTML = rows.map(r => `
                <tr class="summary-row border-b border-zinc-100 hover:bg-zinc-50/80" data-sku="${escapeHtml(r.sku).toLowerCase()}">
                    <td class="px-4 py-3 font-mono font-semibold text-violet-700">${escapeHtml(r.sku)}</td>
                    <td class="px-4 py-3 text-right font-bold tabular-nums text-zinc-900">${escapeHtml(String(r.total))}</td>
                    <td class="px-4 py-3 text-right text-zinc-500 hidden sm:table-cell">${escapeHtml(String(r.location_count))}</td>
                </tr>
            `).join('');

            document.getElementById('grandTotal').textContent = grandTotal;
            document.getElementById('skuCount').textContent = rows.length + ' SKU';
            foot.classList.remove('hidden');
        }

        function filterTable() {
            const q = document.getElementById('searchSku').value.trim().toLowerCase();
            const filtered = q
                ? summaryRows.filter(r => r.sku.toLowerCase().includes(q))
                : summaryRows;
            const grand = filtered.reduce((s, r) => s + r.total, 0);
            renderTable(filtered, grand);
        }

        async function loadSummary() {
            const body = document.getElementById('summaryBody');
            body.innerHTML = '<tr><td colspan="3" class="px-4 py-8 text-center text-zinc-400 animate-pulse">Loading…</td></tr>';

            try {
                const res = await fetch('/summary/data');
                const data = await res.json();
                summaryRows = data.rows || [];
                renderTable(summaryRows, data.grand_total || 0);
                document.getElementById('lastUpdated').textContent =
                    'Diperbarui: ' + new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
            } catch (e) {
                body.innerHTML = '<tr><td colspan="3" class="px-4 py-8 text-center text-rose-600">Gagal memuat data.</td></tr>';
            }
        }

        window.onload = loadSummary;
    </script>
</body>
</html>
"""

# --- ROUTES ---

@app.route('/')
def home():
    sku_tree = {}
    valid_locations = []
    location_lookup = {}
    counter_lookup = {}
    valid_counters = []
    lookup_warnings = []
    try:
        wb = get_spreadsheet()
        location_lookup, loc_warnings = get_valid_locations(wb)
        counter_lookup, counter_warnings = get_valid_counters(wb)
        lookup_warnings = loc_warnings + counter_warnings
        valid_locations = sorted(set(location_lookup.values()))
        valid_counters = sorted(set(counter_lookup.values()))
        sku_tree = load_sku_tree(wb)
        if not sku_tree:
            sku_tree = {
                "Finished": {
                    "Brushes": ["AR-BRSH-01", "AR-BRSH-02"]
                }
            }
    except Exception:
        sku_tree = {
            "Finished": {
                "Brushes": ["AR-BRSH-01", "AR-BRSH-02"]
            }
        }
        if not location_lookup:
            location_lookup = {}
        if not counter_lookup:
            counter_lookup = {}
        
    boot_data = {
        "sku_tree": sku_tree,
        "location_lookup": location_lookup,
        "counter_lookup": counter_lookup,
        "valid_locations": valid_locations,
        "valid_counters": valid_counters,
        "lookup_warnings": lookup_warnings,
    }
    return render_template_string(
        HTML_TEMPLATE,
        sku_tree=sku_tree,
        valid_locations=valid_locations,
        location_lookup=location_lookup,
        counter_lookup=counter_lookup,
        valid_counters=valid_counters,
        lookup_warnings=lookup_warnings,
        boot_data=boot_data,
    )

@app.route('/api/lookups')
def api_lookups():
    try:
        wb = get_spreadsheet()
        location_lookup, loc_warnings = get_valid_locations(wb)
        counter_lookup, counter_warnings = get_valid_counters(wb)
        return jsonify({
            "locations": location_lookup,
            "counters": counter_lookup,
            "location_count": len(location_lookup),
            "counter_count": len(counter_lookup),
            "warnings": loc_warnings + counter_warnings,
        }), 200
    except Exception as e:
        return jsonify({
            "locations": {},
            "counters": {},
            "location_count": 0,
            "counter_count": 0,
            "warnings": [],
            "error": str(e),
        }), 500

@app.route('/summary')
def summary_page():
    return render_template_string(SUMMARY_HTML_TEMPLATE)

@app.route('/summary/data')
def summary_data():
    try:
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        records = sheet.get_all_records()
        rows, grand_total = aggregate_counts_by_sku(records)
        return jsonify({
            "rows": rows,
            "grand_total": grand_total,
            "sku_count": len(rows),
        }), 200
    except Exception as e:
        return jsonify({"rows": [], "grand_total": 0, "sku_count": 0, "error": str(e)}), 500

@app.route('/history', methods=['GET'])
def history():
    try:
        raw_name = request.args.get('name', '').strip()
        wb = get_spreadsheet()
        counter_lookup, _ = get_valid_counters(wb)
        target_name = resolve_counter(raw_name, counter_lookup) or raw_name
        sheet = wb.worksheet("Raw Counts")
        all_records = sheet.get_all_records()
        
        def row_matches_counter(row):
            row_name = get_row_counter_name(row)
            if not row_name or not target_name:
                return False
            if row_name == target_name:
                return True
            return resolve_counter(row_name, counter_lookup) == target_name

        counter_data = [
            {
                "id": row.get("Log ID"),
                "location": row.get("Precise Location"),
                "sku": row.get("SKU Code"),
                "count": row.get("Physical Count"),
                "timestamp": row.get("Timestamp"),
                "notes": row.get("Notes") or ""
            }
            for row in all_records if row_matches_counter(row)
        ]
        return jsonify(list(reversed(counter_data))), 200
    except Exception:
        return jsonify([]), 500

@app.route('/submit', methods=['POST'])
def submit():
    try:
        data = request.json
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        
        counter_lookup, _ = get_valid_counters(wb)
        counter_name = resolve_counter(data.get('counter_name', ''), counter_lookup)

        if not counter_lookup:
            return jsonify({
                "status": "error",
                "message": "Daftar petugas tidak tersedia. Ketik nama sesuai dengan ID badge.",
            }), 400

        if not counter_name:
            return jsonify({
                "status": "invalid_counter",
                "message": "Nama petugas tidak valid. Scan ID badge atau ketik nama sesuai dengan ID badge.",
            }), 400

        location_lookup, _ = get_valid_locations(wb)
        loc_string = resolve_location(data.get('location', ''), location_lookup)

        if not location_lookup:
            return jsonify({
                "status": "error",
                "message": "Daftar lokasi tidak tersedia. Periksa tab LOCATIONS di sheet.",
            }), 400

        if not loc_string:
            raw = str(data.get('location', '')).strip()
            return jsonify({
                "status": "invalid_location",
                "message": f"Lokasi tidak valid: {raw}. Scan QR lokasi yang benar.",
            }), 400

        sku_tree = load_sku_tree(wb)
        if not sku_tree:
            return jsonify({
                "status": "error",
                "message": "Daftar SKU tidak tersedia. Periksa tab SKU List di sheet.",
            }), 400

        sku_code = resolve_sku(data.get('sku', ''), sku_tree)
        if not sku_code:
            raw_sku = str(data.get('sku', '')).strip()
            return jsonify({
                "status": "invalid_sku",
                "message": f"SKU tidak valid: {raw_sku}. Pilih atau scan SKU dari daftar.",
            }), 400

        # --- FEATURE 1: REAL-TIME DUPLICATE DETECTOR INTERCEPTOR ---
        all_records = sheet.get_all_records()
        for row in all_records:
            if (str(row.get("Precise Location")).strip() == loc_string and 
                str(row.get("SKU Code")).strip() == sku_code and 
                get_row_counter_name(row) == counter_name):
                
                return jsonify({
                    "status": "duplicate",
                    "message": (
                        f"Sudah tercatat: {sku_code} di {loc_string} "
                        f"(jumlah {row.get('Physical Count')}). "
                        f"Ubah lewat Edit di tab Riwayat."
                    ),
                }), 409

        # Process standard row generation if duplicate test passes
        log_id = str(uuid.uuid4())[:8] 
        parts = loc_string.split('-')
        
        zone = parts[0] if len(parts) > 0 else loc_string[:1]
        rack = parts[1] if len(parts) > 1 else ""
        shelf = parts[2] if len(parts) > 2 else ""
        
        jakarta_tz = pytz.timezone('Asia/Jakarta')
        now_wib = datetime.now(jakarta_tz)
        timestamp = now_wib.strftime("%d/%m/%Y %H:%M:%S")
        notes = str(data.get('notes', '')).strip()
        
        try:
            physical_count = int(data['count'])
        except (ValueError, TypeError):
            return jsonify({"status": "error", "message": "Jumlah tidak valid."}), 400
        if physical_count < 0:
            return jsonify({"status": "error", "message": "Jumlah tidak boleh negatif."}), 400
        
        row_to_append = [
            log_id,              # Column A: Log ID
            counter_name,        # Column B: Counter Name
            zone,                # Column C: Zone
            rack,                # Column D: Rack
            shelf,               # Column E: Shelf
            loc_string,          # Column F: Precise Location
            sku_code,            # Column G: SKU Code
            physical_count,    # Column H: Physical Count
            timestamp,           # Column I: Timestamp
            notes                # Column J: Notes
        ]
        
        sheet.append_row(row_to_append)
        return jsonify({"status": "success"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/edit', methods=['POST'])
def edit():
    try:
        data = request.json
        log_id = data['id']
        new_count = data['count']
        
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        cell = sheet.find(log_id)
        
        if cell:
            sheet.update_cell(cell.row, 8, new_count)  # Column H: Physical Count
            return jsonify({"status": "success"}), 200
        return jsonify({"status": "error", "message": "Record row not found"}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/delete', methods=['POST'])
def delete():
    try:
        data = request.json
        log_id = data['id']
        
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        cell = sheet.find(log_id)
        
        if cell:
            sheet.delete_rows(cell.row)
            return jsonify({"status": "success"}), 200
        return jsonify({"status": "error", "message": "Record row not found"}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/favicon.ico')
@app.route('/favicon.png')
def favicon():
    return '', 204

app = app