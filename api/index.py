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

    <!-- Sticky header -->
    <header class="sticky top-0 z-40 bg-white/90 backdrop-blur-md border-b border-zinc-200">
        <div class="max-w-md lg:max-w-5xl mx-auto px-4 py-3 flex items-center justify-between gap-3">
            <div class="min-w-0">
                <h1 class="text-lg font-bold text-zinc-900 tracking-tight truncate">Aeris Beaute</h1>
                <p class="text-xs text-zinc-500">Stock Opname 2026</p>
            </div>
            <div class="shrink-0 w-36">
                <label for="counterTeam" class="sr-only">Counter team</label>
                <select id="counterTeam" onchange="fetchHistory()" class="w-full border border-zinc-200 text-sm font-medium rounded-lg px-3 py-2 bg-white focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none">
                    <option value="Team 1">Team 1</option>
                    <option value="Team 2">Team 2</option>
                    <option value="Team 3">Team 3</option>
                    <option value="Team 4">Team 4</option>
                </select>
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

                <!-- Step 1: Location -->
                <section id="step1Card" class="bg-white rounded-xl border border-zinc-200 shadow-sm overflow-hidden">
                    <div class="flex items-center gap-3 px-4 py-3 border-b border-zinc-100 bg-zinc-50/80">
                        <span class="text-xs font-semibold text-violet-600 bg-violet-50 px-2 py-0.5 rounded">Step 1</span>
                        <h2 class="text-sm font-semibold text-zinc-800">Location</h2>
                    </div>
                    <div class="p-4 space-y-3">
                        <button type="button" onclick="openScanModal('location')" class="w-full py-3 px-4 rounded-lg bg-violet-600 hover:bg-violet-700 text-white font-semibold text-sm transition shadow-sm">
                            Scan location QR
                        </button>
                        <div>
                            <label for="location" class="block text-sm font-medium text-zinc-600 mb-1">Precise location</label>
                            <input type="text" id="location" readonly placeholder="Scan location QR to unlock" class="w-full border border-amber-200 p-3 rounded-lg bg-amber-50 font-mono text-sm font-semibold text-amber-800 placeholder-amber-400/80">
                        </div>
                    </div>
                </section>

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
                <div class="bg-white rounded-xl border border-zinc-200 shadow-sm p-4 space-y-3 lg:sticky lg:top-20">
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
        let currentTarget = '';
        let scannerRunning = false;
        const html5QrcodeScanner = new Html5Qrcode("reader");

        const CLS = {
            locLocked: "w-full border border-amber-200 p-3 rounded-lg bg-amber-50 font-mono text-sm font-semibold text-amber-800 placeholder-amber-400/80",
            locUnlocked: "w-full border border-emerald-200 p-3 rounded-lg bg-emerald-50 font-mono text-sm font-semibold text-emerald-800",
            selLocked: "w-full border border-zinc-200 p-3 rounded-lg bg-zinc-50 text-zinc-400 font-medium transition",
            selUnlocked: "w-full border border-zinc-200 p-3 rounded-lg bg-white text-zinc-900 font-medium focus:border-violet-500 focus:ring-2 focus:ring-violet-500/20 focus:outline-none transition",
        };

        const STEP_DOT = {
            pending: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-zinc-200 bg-white text-zinc-400",
            active: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-violet-600 bg-violet-600 text-white",
            done: "shrink-0 flex h-8 w-8 items-center justify-center rounded-full text-xs font-bold border-2 border-emerald-500 bg-emerald-500 text-white",
        };

        window.onload = () => {
            fetchHistory();
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
            const loc = document.getElementById('location').value.trim();
            const sku = document.getElementById('skuSelector').value;

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
            const ready = document.getElementById('location').value.trim()
                && document.getElementById('skuSelector').value
                && document.getElementById('count').value !== '';
            const btn = document.getElementById('submitBtn');
            btn.disabled = !ready || btn.dataset.loading === '1';
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
            locInput.placeholder = 'Scan location QR to unlock';

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
            if (target === 'sku' && document.getElementById('scanSkuBtn').disabled) return;

            currentTarget = target;
            document.getElementById('scanModalTitle').textContent =
                target === 'location' ? 'Scan location QR' : 'Scan SKU QR';
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
                        if (currentTarget === 'location') {
                            document.getElementById('location').value = text;
                            unlockFormForLocation();
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
            const team = document.getElementById('counterTeam').value;
            const container = document.getElementById('historyContainer');
            container.innerHTML = '<p class="text-zinc-400 text-center py-8 animate-pulse">Loading…</p>';

            try {
                const response = await fetch(`/history?team=${encodeURIComponent(team)}`);
                const data = await response.json();

                if (data.length === 0) {
                    container.innerHTML = '<p class="text-zinc-400 text-center py-8">No records for this team yet.</p>';
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
                                <p class="text-xs text-zinc-400 mt-0.5 truncate">${escapeHtml(item.notes || 'No notes')}</p>
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

            if (!locInput || !skuInput || countInput === '') {
                showToast('Complete location, SKU, and count before submitting.', 'warning');
                return;
            }

            btn.disabled = true;
            btn.dataset.loading = '1';
            label.textContent = 'Saving…';
            spinner.classList.remove('hidden');

            const payload = {
                team: document.getElementById('counterTeam').value,
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
                    showToast(result.message || 'Duplicate entry blocked.', 'warning');
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
    try:
        wb = get_spreadsheet()
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
        
    return render_template_string(HTML_TEMPLATE, sku_tree=sku_tree)

@app.route('/history', methods=['GET'])
def history():
    try:
        target_team = request.args.get('team')
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        all_records = sheet.get_all_records()
        
        team_data = [
            {
                "id": row.get("Log ID"),
                "location": row.get("Precise Location"),
                "sku": row.get("SKU Code"),
                "count": row.get("Physical Count"),
                "notes": row.get("Notes")
            }
            for row in all_records if str(row.get("Counter Team")).strip() == target_team
        ]
        return jsonify(list(reversed(team_data))), 200
    except Exception:
        return jsonify([]), 500

@app.route('/submit', methods=['POST'])
def submit():
    try:
        data = request.json
        wb = get_spreadsheet()
        sheet = wb.worksheet("Raw Counts")
        
        # --- FEATURE 1: REAL-TIME DUPLICATE DETECTOR INTERCEPTOR ---
        all_records = sheet.get_all_records()
        for row in all_records:
            if (str(row.get("Precise Location")).strip() == str(data['location']).strip() and 
                str(row.get("SKU Code")).strip() == str(data['sku']).strip() and 
                str(row.get("Counter Team")).strip() == str(data['team']).strip()):
                
                return jsonify({
                    "status": "duplicate",
                    "message": f"Your team already recorded SKU [{data['sku']}] at location [{data['location']}] with an entry of {row.get('Physical Count')} units.\\n\\nTo modify this quantity, scroll down to your 'Recent Activity' log below, click 'Edit' on that record, and update the value instead of creating a duplicate line item."
                }), 409
        
        # Process standard row generation if duplicate test passes
        log_id = str(uuid.uuid4())[:8] 
        loc_string = str(data['location']).strip()
        parts = loc_string.split('-')
        
        zone = parts[0] if len(parts) > 0 else loc_string[:1]
        shelf = parts[1] if len(parts) > 1 else ""
        bin_code = parts[2] if len(parts) > 2 else ""
        
        jakarta_tz = pytz.timezone('Asia/Jakarta')
        now_wib = datetime.now(jakarta_tz)
        time_string = now_wib.strftime("[%H:%M:%S]")
        combined_notes = f"{time_string} {data['notes']}".strip()
        
        row_to_append = [
            log_id,              # Column A: Log ID
            data['team'],        # Column B: Counter Team[cite: 1]
            zone,                # Column C: Zone[cite: 1]
            shelf,               # Column D: Shelf[cite: 1]
            bin_code,            # Column E: Bin[cite: 1]
            loc_string,          # Column F: Precise Location[cite: 1]
            data['sku'],         # Column G: SKU Code[cite: 1]
            "",                  # Column H: Item Name[cite: 1]
            int(data['count']),  # Column I: Physical Count[cite: 1]
            combined_notes       # Column J: Notes[cite: 1]
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