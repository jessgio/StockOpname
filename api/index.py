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

def get_valid_locations(wb):
    """Return a set of allowed location codes from LOCATIONS column A."""
    try:
        ws = wb.worksheet("LOCATIONS")
        col_a = ws.col_values(1)
    except Exception:
        return set()

    locations = set()
    header_names = {"location", "lokasi", "precise location", "locations", "kode lokasi"}
    for i, val in enumerate(col_a):
        loc = str(val).strip()
        if not loc:
            continue
        if i == 0 and loc.lower() in header_names:
            continue
        locations.add(loc)
    return locations

def get_valid_counters(wb):
    """Return a set of allowed counter names from COUNTERS column A."""
    try:
        ws = wb.worksheet("COUNTERS")
        col_a = ws.col_values(1)
    except Exception:
        return set()

    counters = set()
    header_names = {
        "counter", "counter name", "nama", "nama petugas", "petugas",
        "name", "counters", "id badge", "badge",
    }
    for i, val in enumerate(col_a):
        name = str(val).strip()
        if not name:
            continue
        if i == 0 and name.lower() in header_names:
            continue
        counters.add(name)
    return counters

def get_row_counter_name(row):
    """Read counter from new or legacy column header."""
    return str(row.get("Counter Name") or row.get("Counter Team") or "").strip()

# --- HTML INTERFACE WITH DUPLICATE PROTECTION & AUTO-RESET ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Aeris Opname 2026</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <script src="https://unpkg.com/html5-qrcode"></script>
</head>
<body class="bg-zinc-50 font-sans text-zinc-900 antialiased min-h-screen flex flex-col">

    <!-- Toast -->
    <div id="toast" class="fixed top-4 left-4 right-4 z-[60] mx-auto max-w-md translate-y-[-120%] opacity-0 transition-all duration-300 pointer-events-none">
        <div id="toastInner" class="rounded-xl px-4 py-3 text-sm font-medium shadow-lg border"></div>
    </div>

    <!-- Sticky header: brand + counter + location -->
    <header class="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-zinc-200 shadow-sm">
        <div class="max-w-md lg:max-w-5xl mx-auto px-4 pt-3 pb-3">
            <h1 class="text-lg font-bold text-zinc-900 tracking-tight">Aeris Beaute</h1>
            <p class="text-xs text-zinc-500 mb-3">Stock Opname 2026</p>
            <div class="space-y-3">
                <div>
                    <label for="counterName" class="block text-xs font-medium text-zinc-600 mb-1">Petugas</label>
                    <div class="flex gap-2">
                        <input type="text" id="counterName" autocomplete="off" onblur="onCounterNameInput()" placeholder="Scan ID badge atau ketik nama" class="flex-1 min-w-0 border border-amber-200 p-2.5 rounded-lg bg-amber-50 text-sm font-semibold text-amber-800 placeholder-amber-400/80 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
                        <button type="button" onclick="openScanModal('counter')" class="shrink-0 px-3 py-2.5 rounded-lg bg-violet-600 hover:bg-violet-700 text-white text-xs font-semibold transition">Scan</button>
                    </div>
                </div>
                <div>
                    <label for="location" class="block text-xs font-medium text-zinc-600 mb-1">Lokasi</label>
                    <div class="flex gap-2">
                        <input type="text" id="location" readonly placeholder="Scan ID badge dulu" class="flex-1 min-w-0 border border-amber-200 p-2.5 rounded-lg bg-amber-50 font-mono text-sm font-semibold text-amber-800 placeholder-amber-400/80">
                        <button type="button" id="scanLocationBtn" onclick="openScanModal('location')" disabled class="shrink-0 px-3 py-2.5 rounded-lg bg-violet-600 hover:bg-violet-700 text-white text-xs font-semibold transition disabled:opacity-40 disabled:cursor-not-allowed">Scan</button>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <main class="flex-1 max-w-md lg:max-w-5xl w-full mx-auto px-4 py-4 pb-28 lg:pb-6">

        <!-- Stepper -->
        <nav class="flex items-center gap-2 mb-5" aria-label="Count progress">
            <div id="step1Indicator" class="flex-1 flex items-center gap-2 min-w-0">
                <span id="step1Dot" class="shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-violet-600 bg-violet-600 text-white">1</span>
                <span class="text-xs font-medium text-zinc-700 truncate hidden sm:inline">Location</span>
            </div>
            <div class="h-px w-4 bg-zinc-200 shrink-0" aria-hidden="true"></div>
            <div id="step2Indicator" class="flex-1 flex items-center gap-2 min-w-0">
                <span id="step2Dot" class="shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400">2</span>
                <span class="text-xs font-medium text-zinc-400 truncate hidden sm:inline">Product</span>
            </div>
            <div class="h-px w-4 bg-zinc-200 shrink-0" aria-hidden="true"></div>
            <div id="step3Indicator" class="flex-1 flex items-center gap-2 min-w-0">
                <span id="step3Dot" class="shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400">3</span>
                <span class="text-xs font-medium text-zinc-400 truncate hidden sm:inline">Count</span>
            </div>
        </nav>

        <!-- Mobile tabs -->
        <div class="lg:hidden flex rounded-lg bg-zinc-100 p-1 mb-4" role="tablist">
            <button id="tabCount" type="button" onclick="switchTab('count')" role="tab" aria-selected="true" class="flex-1 py-2 text-sm font-semibold rounded-md bg-white text-zinc-900 shadow-sm transition">Count</button>
            <button id="tabHistory" type="button" onclick="switchTab('history')" role="tab" aria-selected="false" class="flex-1 py-2 text-sm font-semibold rounded-md text-zinc-500 transition">History</button>
        </div>

        <div class="lg:grid lg:grid-cols-2 lg:gap-6 lg:items-start">

            <!-- Count panel -->
            <div id="panelCount" class="space-y-4">

                <!-- Step 2: Product -->
                <section id="step2Card" class="bg-white rounded-xl border border-zinc-200 shadow-sm overflow-hidden opacity-60 pointer-events-none transition" aria-disabled="true">
                    <div class="flex items-center gap-3 px-4 py-3 border-b border-zinc-100 bg-zinc-50/80">
                        <span class="text-xs font-semibold text-zinc-400 bg-zinc-100 px-2 py-0.5 rounded">Step 2</span>
                        <h2 class="text-sm font-semibold text-zinc-500">Product</h2>
                    </div>
                    <div class="p-4 space-y-3">
                        <button type="button" id="scanSkuBtn" onclick="openScanModal('sku')" disabled class="w-full py-3 px-4 rounded-lg border-2 border-zinc-200 text-zinc-400 font-semibold text-sm transition disabled:opacity-50 disabled:cursor-not-allowed enabled:border-violet-600 enabled:text-violet-700 enabled:hover:bg-violet-50">
                            Scan SKU QR
                        </button>
                        <div>
                            <label for="skuType" class="block text-sm font-medium text-zinc-600 mb-1">Goods type</label>
                            <select id="skuType" onchange="updateCategories()" disabled class="w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 font-medium transition">
                                <option value="">Locked — scan location first</option>
                            </select>
                        </div>
                        <div>
                            <label for="skuCategory" class="block text-sm font-medium text-zinc-600 mb-1">Category</label>
                            <select id="skuCategory" onchange="updateSkus()" disabled class="w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 font-medium transition">
                                <option value="">Locked</option>
                            </select>
                        </div>
                        <div>
                            <label for="skuSelector" class="block text-sm font-medium text-zinc-600 mb-1">SKU code</label>
                            <select id="skuSelector" onchange="updateStepperUI(); updateSubmitState();" disabled class="w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 font-medium transition">
                                <option value="">Locked</option>
                            </select>
                        </div>
                    </div>
                </section>

                <!-- Step 3: Count -->
                <section id="step3Card" class="bg-white rounded-xl border border-zinc-200 shadow-sm overflow-hidden opacity-60 pointer-events-none transition" aria-disabled="true">
                    <div class="flex items-center gap-3 px-4 py-3 border-b border-zinc-100 bg-zinc-50/80">
                        <span class="text-xs font-semibold text-zinc-400 bg-zinc-100 px-2 py-0.5 rounded">Step 3</span>
                        <h2 class="text-sm font-semibold text-zinc-500">Quantity</h2>
                    </div>
                    <div class="p-4 space-y-4">
                        <div class="flex items-center justify-center gap-3">
                            <button type="button" onclick="adjustCount(-10)" class="text-xs font-semibold text-zinc-500 bg-zinc-100 hover:bg-zinc-200 px-3 py-2 rounded-lg transition">−10</button>
                            <button type="button" onclick="adjustCount(-1)" class="flex h-12 w-12 items-center justify-center rounded-full bg-zinc-100 hover:bg-zinc-200 text-xl font-bold text-zinc-700 transition">−</button>
                            <input type="number" id="count" oninput="updateSubmitState()" onfocus="this.select()" onclick="this.select()" class="w-28 border border-zinc-200 rounded-xl text-center text-4xl font-bold tabular-nums text-zinc-900 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none" value="0" min="0">
                            <button type="button" onclick="adjustCount(1)" class="flex h-12 w-12 items-center justify-center rounded-full bg-violet-600 hover:bg-violet-700 text-xl font-bold text-white transition">+</button>
                            <button type="button" onclick="adjustCount(10)" class="text-xs font-semibold text-zinc-500 bg-zinc-100 hover:bg-zinc-200 px-3 py-2 rounded-lg transition">+10</button>
                        </div>
                        <div>
                            <label for="notes" class="block text-sm font-medium text-zinc-600 mb-1">Notes <span class="text-zinc-400 font-normal">(optional)</span></label>
                            <input type="text" id="notes" class="w-full border border-zinc-200 p-3 rounded-lg text-sm focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none transition" placeholder="e.g. damaged box">
                        </div>
                    </div>
                </section>

                <div id="footerSubmit" class="fixed bottom-0 left-0 right-0 z-30 bg-white/95 backdrop-blur-md border-t border-zinc-200 p-4 pb-[max(1rem,env(safe-area-inset-bottom))] lg:relative lg:z-auto lg:border-0 lg:bg-transparent lg:p-0 lg:mt-2 lg:pb-0">
                    <button id="submitBtn" type="button" onclick="submitData()" disabled class="w-full bg-emerald-600 hover:bg-emerald-700 disabled:bg-zinc-300 disabled:cursor-not-allowed text-white font-semibold text-base py-3.5 rounded-xl shadow-sm active:scale-[0.98] transition flex items-center justify-center gap-2">
                        <span id="submitBtnLabel">Submit to master sheet</span>
                        <svg id="submitSpinner" class="hidden h-5 w-5 animate-spin" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"></path></svg>
                    </button>
                </div>
            </div>

            <!-- History panel -->
            <div id="panelHistory" class="hidden lg:block">
                <div class="bg-white rounded-xl border border-zinc-200 shadow-sm p-4 space-y-3 lg:sticky lg:top-52">
                    <div class="flex justify-between items-center">
                        <h2 class="text-sm font-semibold text-zinc-800">Recent activity</h2>
                        <button type="button" onclick="fetchHistory()" class="text-sm font-medium text-violet-600 hover:text-violet-800">Refresh</button>
                    </div>
                    <div id="historyContainer" class="space-y-2 max-h-[calc(100vh-10rem)] overflow-y-auto text-sm">
                        <p class="text-zinc-400 text-center py-8">Loading history…</p>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <!-- Scanner modal -->
    <div id="scanModal" class="fixed inset-0 z-50 hidden" role="dialog" aria-modal="true" aria-labelledby="scanModalTitle">
        <div class="absolute inset-0 bg-black/70" onclick="closeScanModal()"></div>
        <div class="absolute inset-x-0 bottom-0 max-h-[92vh] flex flex-col bg-white rounded-t-2xl shadow-2xl lg:inset-auto lg:top-1/2 lg:left-1/2 lg:-translate-x-1/2 lg:-translate-y-1/2 lg:w-full lg:max-w-lg lg:rounded-2xl">
            <div class="flex items-center justify-between px-4 py-3 border-b border-zinc-100 shrink-0">
                <h3 id="scanModalTitle" class="text-base font-semibold text-zinc-900">Scan QR code</h3>
                <button type="button" onclick="closeScanModal()" class="text-sm font-medium text-zinc-500 hover:text-zinc-800 px-2 py-1">Cancel</button>
            </div>
            <div class="p-4 overflow-hidden flex-1 min-h-0">
                <div id="reader" class="w-full aspect-square max-h-[55vh] mx-auto rounded-xl overflow-hidden bg-black"></div>
                <p class="text-center text-xs text-zinc-500 mt-3">Point your camera at the QR code</p>
            </div>
        </div>
    </div>

    <script>
        const skuTree = {{ sku_tree|tojson|safe }};
        const validLocations = new Set({{ valid_locations|tojson|safe }});
        const validCounters = new Set({{ valid_counters|tojson|safe }});
        let currentTarget = '';
        let scannerRunning = false;
        const html5QrcodeScanner = new Html5Qrcode("reader");

        const CLS = {
            counterLocked: "flex-1 min-w-0 border border-amber-200 p-2.5 rounded-lg bg-amber-50 text-sm font-semibold text-amber-800 placeholder-amber-400/80 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none",
            counterUnlocked: "flex-1 min-w-0 border border-emerald-200 p-2.5 rounded-lg bg-emerald-50 text-sm font-semibold text-emerald-800 focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none",
            locLocked: "flex-1 min-w-0 border border-amber-200 p-2.5 rounded-lg bg-amber-50 font-mono text-sm font-semibold text-amber-800 placeholder-amber-400/80",
            locUnlocked: "flex-1 min-w-0 border border-emerald-200 p-2.5 rounded-lg bg-emerald-50 font-mono text-sm font-semibold text-emerald-800",
            selLocked: "w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 font-medium transition",
            selUnlocked: "w-full border border-zinc-200 p-3 rounded-lg bg-white text-zinc-900 font-medium focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none transition",
        };

        const STEP_DOT = {
            pending: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400",
            active: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-violet-600 bg-violet-600 text-white",
            done: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-emerald-500 bg-emerald-500 text-white",
        };

        window.onload = () => {
            lockAfterCounter();
            updateStepperUI();
            updateSubmitState();
            switchTab('count');
        };

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

        function setStepCardEnabled(cardId, enabled) {
            const card = document.getElementById(cardId);
            card.classList.toggle('opacity-60', !enabled);
            card.classList.toggle('pointer-events-none', !enabled);
            card.setAttribute('aria-disabled', !enabled);
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

            if (!counterOk) {
                setStepDot(1, 'pending');
                setStepDot(2, 'pending');
                setStepDot(3, 'pending');
                setStepCardEnabled('step2Card', false);
                setStepCardEnabled('step3Card', false);
                return;
            }

            if (!loc) {
                setStepDot(1, 'active');
                setStepDot(2, 'pending');
                setStepDot(3, 'pending');
                setStepCardEnabled('step2Card', false);
                setStepCardEnabled('step3Card', false);
            } else if (!sku) {
                setStepDot(1, 'done');
                setStepDot(2, 'active');
                setStepDot(3, 'pending');
                setStepCardEnabled('step2Card', true);
                setStepCardEnabled('step3Card', false);
            } else {
                setStepDot(1, 'done');
                setStepDot(2, 'done');
                setStepDot(3, 'active');
                setStepCardEnabled('step2Card', true);
                setStepCardEnabled('step3Card', true);
            }
        }

        function updateSubmitState() {
            const ready = isValidCounter(document.getElementById('counterName').value)
                && document.getElementById('location').value.trim()
                && document.getElementById('skuSelector').value
                && document.getElementById('count').value !== '';
            const btn = document.getElementById('submitBtn');
            btn.disabled = !ready || btn.dataset.loading === '1';
        }

        function isValidCounter(name) {
            const trimmed = String(name).trim();
            return trimmed && validCounters.has(trimmed);
        }

        function applyCounterScan(text) {
            const trimmed = text.trim();
            if (!validCounters.size) {
                showToast('Daftar petugas belum dimuat. Hubungi admin.', 'error');
                return false;
            }
            if (!isValidCounter(trimmed)) {
                showToast('ID badge tidak dikenali. Scan badge yang benar atau ketik nama sesuai daftar.', 'warning');
                return false;
            }
            document.getElementById('counterName').value = trimmed;
            unlockAfterCounter();
            return true;
        }

        function onCounterNameInput() {
            if (isValidCounter(document.getElementById('counterName').value)) {
                unlockAfterCounter();
            } else {
                lockAfterCounter();
            }
        }

        function unlockAfterCounter() {
            const counterInput = document.getElementById('counterName');
            counterInput.className = CLS.counterUnlocked;
            document.getElementById('scanLocationBtn').disabled = false;
            const locInput = document.getElementById('location');
            if (!locInput.value.trim()) {
                locInput.placeholder = 'Scan lokasi';
            }
            updateStepperUI();
            updateSubmitState();
            fetchHistory();
        }

        function lockAfterCounter() {
            const counterInput = document.getElementById('counterName');
            if (!isValidCounter(counterInput.value)) {
                counterInput.className = CLS.counterLocked;
            }
            document.getElementById('scanLocationBtn').disabled = true;
            const locInput = document.getElementById('location');
            locInput.placeholder = 'Scan ID badge dulu';
            lockFormPostSubmit();
            updateStepperUI();
            updateSubmitState();
            const container = document.getElementById('historyContainer');
            container.innerHTML = '<p class="text-zinc-400 text-center py-8">Scan ID badge atau ketik nama petugas.</p>';
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
            toast.classList.remove('translate-y-[-120%]', 'opacity-0', 'pointer-events-none');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => {
                toast.classList.add('translate-y-[-120%]', 'opacity-0', 'pointer-events-none');
            }, type === 'warning' ? 6000 : 3500);
        }

        function isValidLocation(code) {
            const trimmed = code.trim();
            return trimmed && validLocations.has(trimmed);
        }

        function applyLocationScan(text) {
            if (!isValidCounter(document.getElementById('counterName').value)) {
                showToast('Scan ID badge petugas dulu.', 'warning');
                return false;
            }
            const trimmed = text.trim();
            if (!validLocations.size) {
                showToast('Daftar lokasi belum dimuat. Hubungi admin.', 'error');
                return false;
            }
            if (!isValidLocation(trimmed)) {
                showToast('Bukan kode lokasi valid. Scan QR lokasi, bukan SKU.', 'warning');
                return false;
            }
            document.getElementById('location').value = trimmed;
            unlockFormForLocation();
            return true;
        }

        function unlockFormForLocation() {
            const locInput = document.getElementById('location');
            locInput.className = CLS.locUnlocked;

            document.getElementById('scanSkuBtn').disabled = false;

            const typeSelect = document.getElementById('skuType');
            typeSelect.disabled = false;
            typeSelect.className = CLS.selUnlocked;
            typeSelect.innerHTML = '<option value="">Choose goods type</option>';
            Object.keys(skuTree).sort().forEach(type => {
                typeSelect.options[typeSelect.options.length] = new Option(type, type);
            });

            updateStepperUI();
            updateSubmitState();
        }

        function lockFormPostSubmit() {
            const locInput = document.getElementById('location');
            locInput.value = '';
            locInput.className = CLS.locLocked;
            locInput.placeholder = 'Scan lokasi';

            document.getElementById('scanSkuBtn').disabled = true;

            ['skuType', 'skuCategory', 'skuSelector'].forEach(id => {
                const el = document.getElementById(id);
                el.disabled = true;
                el.innerHTML = id === 'skuType'
                    ? '<option value="">Locked — scan location first</option>'
                    : '<option value="">Locked</option>';
                el.className = CLS.selLocked;
            });

            updateStepperUI();
            updateSubmitState();
        }

        function updateCategories() {
            const typeVal = document.getElementById('skuType').value;
            const catSelect = document.getElementById('skuCategory');
            const skuSelect = document.getElementById('skuSelector');

            catSelect.innerHTML = '<option value="">Choose category</option>';
            skuSelect.innerHTML = '<option value="">Choose SKU</option>';
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

            skuSelect.innerHTML = '<option value="">Choose SKU</option>';

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

        async function openScanModal(target) {
            if (target === 'location' && document.getElementById('scanLocationBtn').disabled) return;
            if (target === 'sku' && document.getElementById('scanSkuBtn').disabled) return;

            currentTarget = target;
            const scanTitles = {
                counter: 'Scan ID badge',
                location: 'Scan location QR',
                sku: 'Scan SKU QR',
            };
            document.getElementById('scanModalTitle').textContent = scanTitles[target] || 'Scan QR';
            document.getElementById('scanModal').classList.remove('hidden');
            document.body.classList.add('overflow-hidden');

            await new Promise(r => setTimeout(r, 150));

            if (scannerRunning) return;
            try {
                scannerRunning = true;
                await html5QrcodeScanner.start(
                    { facingMode: "environment" },
                    { fps: 15, qrbox: { width: 250, height: 250 } },
                    async (decodedText) => {
                        const text = decodedText.trim();
                        if (currentTarget === 'counter') {
                            applyCounterScan(text);
                        } else if (currentTarget === 'location') {
                            applyLocationScan(text);
                        } else if (currentTarget === 'sku') {
                            let found = false;
                            for (const type in skuTree) {
                                for (const cat in skuTree[type]) {
                                    if (skuTree[type][cat].includes(text)) {
                                        document.getElementById('skuType').value = type;
                                        updateCategories();
                                        document.getElementById('skuCategory').value = cat;
                                        updateSkus();
                                        document.getElementById('skuSelector').value = text;
                                        found = true;
                                        break;
                                    }
                                }
                                if (found) break;
                            }
                            if (!found) {
                                showToast('SKU not found: ' + text, 'warning');
                            } else {
                                updateStepperUI();
                                updateSubmitState();
                            }
                        }
                        await closeScanModal();
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
            try {
                if (scannerRunning && html5QrcodeScanner.isScanning) {
                    await html5QrcodeScanner.stop();
                }
            } catch (e) {}
            scannerRunning = false;
            document.getElementById('scanModal').classList.add('hidden');
            document.body.classList.remove('overflow-hidden');
        }

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
            const counterName = document.getElementById('counterName').value.trim();
            const container = document.getElementById('historyContainer');

            if (!isValidCounter(counterName)) {
                container.innerHTML = '<p class="text-zinc-400 text-center py-8">Scan ID badge atau ketik nama petugas.</p>';
                return;
            }

            container.innerHTML = '<p class="text-zinc-400 text-center py-8 animate-pulse">Loading…</p>';

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
                                    <span class="font-mono text-sm font-semibold text-violet-700 truncate">${escapeHtml(item.location)}</span>
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

            const counterName = document.getElementById('counterName').value.trim();

            if (!isValidCounter(counterName)) {
                showToast('Scan ID badge petugas dulu.', 'warning');
                return;
            }

            if (!locInput || !skuInput || countInput === '') {
                showToast('Complete location, SKU, and count before submitting.', 'warning');
                return;
            }

            if (!isValidLocation(locInput)) {
                showToast('Lokasi tidak valid. Scan QR lokasi yang benar.', 'warning');
                return;
            }

            btn.disabled = true;
            btn.dataset.loading = '1';
            label.textContent = 'Saving…';
            spinner.classList.remove('hidden');

            const payload = {
                counter_name: counterName,
                location: locInput,
                sku: skuInput,
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
                    document.getElementById('count').value = '0';
                    document.getElementById('notes').value = '';
                    lockFormPostSubmit();
                    fetchHistory();
                } else {
                    showToast('Sync failed. Try again.', 'error');
                }
            } catch (err) {
                showToast('Network error. Try again.', 'error');
            } finally {
                btn.dataset.loading = '0';
                label.textContent = 'Submit to master sheet';
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

# --- ROUTES ---

@app.route('/')
def home():
    sku_tree = {}
    valid_locations = []
    valid_counters = []
    try:
        wb = get_spreadsheet()
        valid_locations = sorted(get_valid_locations(wb))
        valid_counters = sorted(get_valid_counters(wb))
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
        sku_tree = {
            "Finished": {
                "Brushes": ["AR-BRSH-01", "AR-BRSH-02"]
            }
        }
        
    return render_template_string(
        HTML_TEMPLATE,
        sku_tree=sku_tree,
        valid_locations=valid_locations,
        valid_counters=valid_counters,
    )

@app.route('/history', methods=['GET'])
def history():
    try:
        target_name = request.args.get('name', '').strip()
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        all_records = sheet.get_all_records()
        
        counter_data = [
            {
                "id": row.get("Log ID"),
                "location": row.get("Precise Location"),
                "sku": row.get("SKU Code"),
                "count": row.get("Physical Count"),
                "timestamp": row.get("Timestamp"),
                "notes": row.get("Notes") or ""
            }
            for row in all_records if get_row_counter_name(row) == target_name
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
        
        counter_name = str(data.get('counter_name', '')).strip()
        valid_counters = get_valid_counters(wb)

        if not valid_counters:
            return jsonify({
                "status": "error",
                "message": "Daftar petugas tidak tersedia. Periksa tab COUNTERS di sheet.",
            }), 400

        if not counter_name or counter_name not in valid_counters:
            return jsonify({
                "status": "invalid_counter",
                "message": "Nama petugas tidak valid. Scan ID badge atau ketik nama yang benar.",
            }), 400

        # --- FEATURE 1: REAL-TIME DUPLICATE DETECTOR INTERCEPTOR ---
        all_records = sheet.get_all_records()
        for row in all_records:
            if (str(row.get("Precise Location")).strip() == str(data['location']).strip() and 
                str(row.get("SKU Code")).strip() == str(data['sku']).strip() and 
                get_row_counter_name(row) == counter_name):
                
                return jsonify({
                    "status": "duplicate",
                    "message": (
                        f"Sudah tercatat: {data['sku']} di {data['location']} "
                        f"(jumlah {row.get('Physical Count')}). "
                        f"Ubah lewat Edit di tab Riwayat."
                    ),
                }), 409
        
        valid_locations = get_valid_locations(wb)
        loc_string = str(data['location']).strip()

        if not valid_locations:
            return jsonify({
                "status": "error",
                "message": "Daftar lokasi tidak tersedia. Periksa tab LOCATIONS di sheet.",
            }), 400

        if loc_string not in valid_locations:
            return jsonify({
                "status": "invalid_location",
                "message": f"Lokasi tidak valid: {loc_string}. Scan QR lokasi yang benar.",
            }), 400

        # Process standard row generation if duplicate test passes
        log_id = str(uuid.uuid4())[:8] 
        parts = loc_string.split('-')
        
        zone = parts[0] if len(parts) > 0 else loc_string[:1]
        shelf = parts[1] if len(parts) > 1 else ""
        bin_code = parts[2] if len(parts) > 2 else ""
        
        jakarta_tz = pytz.timezone('Asia/Jakarta')
        now_wib = datetime.now(jakarta_tz)
        timestamp = now_wib.strftime("%d/%m/%Y %H:%M:%S")
        notes = data['notes'].strip()
        
        row_to_append = [
            log_id,              # Column A: Log ID
            counter_name,        # Column B: Counter Name
            zone,                # Column C: Zone
            shelf,               # Column D: Shelf
            bin_code,            # Column E: Bin
            loc_string,          # Column F: Precise Location
            data['sku'],         # Column G: SKU Code
            "",                  # Column H: Item Name
            int(data['count']),  # Column I: Physical Count
            timestamp,           # Column J: Timestamp
            notes                # Column K: Notes
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
            sheet.update_cell(cell.row, 9, new_count)
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