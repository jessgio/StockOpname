
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
        // #region agent log
        dbg('boot', 'parsed', {counterLookupKeys: Object.keys(counterLookup).length, html5Defined: typeof Html5Qrcode !== 'undefined'}, 'A');
        // #endregion
        let currentTarget = '';
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

        // #region agent log
        function dbg(location, message, data, hypothesisId) {
            const payload = {sessionId: '8114e2', runId: 'pre-fix', hypothesisId, location, message, data: data || {}, timestamp: Date.now()};
            fetch('http://127.0.0.1:7715/ingest/d5497b62-266c-4d71-9f16-7243fc1f0e15', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-Debug-Session-Id': '8114e2'}, body: JSON.stringify(payload)}).catch(() => {});
            fetch('/api/debug-log', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)}).catch(() => {});
        }
        window.addEventListener('error', (e) => {
            dbg('window.onerror', e.message || 'error', {filename: e.filename, lineno: e.lineno}, 'A');
        });
        window.addEventListener('unhandledrejection', (e) => {
            dbg('unhandledrejection', String(e.reason), {}, 'A');
        });
        // #endregion

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
            let s = String(code || '').trim().replace(/\ufeff/g, '').replace(/\u200b/g, '').replace(/\r/g, '').replace(/\n/g, '');
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

        function updateLocationUI() {
            const locInput = document.getElementById('location');
            const hint = document.getElementById('locationHint');
            const counterOk = isValidCounter(document.getElementById('counterName').value);
            if (!counterOk) {
                if (hint) hint.textContent = 'Isi nama petugas dulu (scan badge atau ketik).';
                locInput.placeholder = 'Belum di-scan';
            } else if (locInput.value.trim()) {
                if (hint) hint.textContent = 'Lokasi terisi. Tekan Scan untuk mengganti.';
                locInput.placeholder = locInput.value;
            } else {
                if (hint) hint.textContent = 'Tekan Scan untuk QR lokasi.';
                locInput.placeholder = 'Belum di-scan';
            }
        }

        function lockSkuFields() {
            document.getElementById('scanSkuBtn').disabled = true;
            ['skuType', 'skuCategory', 'skuSelector'].forEach(id => {
                const el = document.getElementById(id);
                el.disabled = true;
                el.innerHTML = id === 'skuType'
                    ? '<option value="">Scan lokasi terlebih dahulu</option>'
                    : '<option value="">Terkunci</option>';
                el.className = CLS.selLocked;
            });
        }

        function resetLocationAndSku() {
            const locInput = document.getElementById('location');
            locInput.value = '';
            locInput.className = CLS.locLocked;
            updateLocationUI();
            lockSkuFields();
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
                document.getElementById('scanLocationBtn').disabled = false;
                setStepCardEnabled('step2Card', true);
                updateLocationUI();
                if (locInput.value.trim() && isValidLocation(locInput.value)) {
                    unlockFormForLocation();
                } else {
                    if (locInput.value.trim() && !isValidLocation(locInput.value)) {
                        locInput.value = '';
                    }
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
                updateLocationUI();
                lockSkuFields();
                historyContainer.innerHTML = '<p class="text-zinc-400 text-center py-8">Scan ID badge atau ketik nama petugas.</p>';
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
            let hitTestId = null;
            if (counterBtn) {
                const r = counterBtn.getBoundingClientRect();
                const topEl = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
                hitTestId = topEl ? (topEl.id || topEl.tagName) : null;
            }
            // #region agent log
            dbg('bindScanButtons', 'elements', {
                counterFound: !!counterBtn,
                counterDisabled: counterBtn ? counterBtn.disabled : null,
                openScanModalType: typeof openScanModal,
                hitTestId,
                modalHidden: document.getElementById('scanModal')?.hidden,
                modalDisplay: document.getElementById('scanModal')?.style?.display,
            }, 'C');
            // #endregion
            if (counterBtn) {
                counterBtn.addEventListener('pointerdown', () => {
                    // #region agent log
                    dbg('scanCounterBtn', 'pointerdown', {}, 'C');
                    // #endregion
                });
                counterBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    // #region agent log
                    dbg('scanCounterBtn', 'click', {openScanModalType: typeof openScanModal}, 'B');
                    // #endregion
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
            // #region agent log
            dbg('initApp', 'start', {readyState: document.readyState}, 'A');
            // #endregion
            try {
                resetPageInteractionState();
                bindScanButtons();
                showLookupWarnings(lookupWarnings);
                await loadLookups();
                await syncUIState();
                switchTab('count');
                // #region agent log
                dbg('initApp', 'done', {}, 'A');
                // #endregion
            } catch (err) {
                // #region agent log
                dbg('initApp', 'error', {err: String(err)}, 'A');
                // #endregion
                throw err;
            }
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

        function setStepCardEnabled(cardId, enabled) {
            if (cardId === 'step1Card') return;
            const card = document.getElementById(cardId);
            if (!card) return;
            card.classList.toggle('opacity-60', !enabled);
            card.classList.toggle('step-locked', !enabled);
            card.setAttribute('aria-disabled', String(!enabled));
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
            if (validCounters.has(trimmed)) return trimmed;
            const key = normalizeScanText(trimmed, 'counter').toLowerCase();
            return counterLookup[key] || null;
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
                showToast(`Petugas tidak dikenali: "${scanned.slice(0, 40)}". Periksa tab COUNTERS.`, 'warning');
                return false;
            }
            document.getElementById('counterName').value = resolved;
            unlockAfterCounter();
            await syncUIState();
            return true;
        }

        function onCounterNameInput() {
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
            updateLocationUI();

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
            // #region agent log
            dbg('openScanModal', 'entry', {
                target,
                html5Defined: typeof Html5Qrcode !== 'undefined',
                locBtnDisabled: document.getElementById('scanLocationBtn')?.disabled,
            }, 'B');
            // #endregion
            if (target === 'location' && document.getElementById('scanLocationBtn').disabled) {
                // #region agent log
                dbg('openScanModal', 'earlyReturn', {reason: 'locationDisabled'}, 'B');
                // #endregion
                return;
            }
            if (target === 'sku' && document.getElementById('scanSkuBtn').disabled) {
                // #region agent log
                dbg('openScanModal', 'earlyReturn', {reason: 'skuDisabled'}, 'B');
                // #endregion
                return;
            }
            if (typeof Html5Qrcode === 'undefined') {
                // #region agent log
                dbg('openScanModal', 'earlyReturn', {reason: 'html5Missing'}, 'B');
                // #endregion
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
            // #region agent log
            dbg('openScanModal', 'modalShown', {
                hidden: modal.hidden,
                display: modal.style.display,
                isOpen: modal.classList.contains('is-open'),
            }, 'D');
            // #endregion
            resetScanDebounce();

            await ensureScannerStopped();
            await new Promise(r => setTimeout(r, 200));

            const scanner = getScanner();
            if (!scanner) {
                // #region agent log
                dbg('openScanModal', 'noScanner', {}, 'B');
                // #endregion
                showToast('Pemindai QR tidak siap. Muat ulang halaman.', 'error');
                closeScanModal();
                return;
            }

            try {
                scannerRunning = true;
                // #region agent log
                dbg('openScanModal', 'scannerStart', {target}, 'D');
                // #endregion
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
                // #region agent log
                dbg('openScanModal', 'scannerStartError', {err: String(err)}, 'D');
                // #endregion
                showToast('Could not start camera. Check permissions.', 'error');
                closeScanModal();
            }
        }

        async function closeScanModal() {
            // #region agent log
            dbg('closeScanModal', 'called', {stack: (new Error()).stack?.slice(0, 200)}, 'D');
            // #endregion
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
            label.textContent = 'Saving…';
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
                    document.getElementById('count').value = '0';
                    document.getElementById('notes').value = '';
                    resetLocationAndSku();
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
    