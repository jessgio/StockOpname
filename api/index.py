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
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Aeris Opname 2026</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <script src="https://unpkg.com/html5-qrcode"></script>
</head>
<body class="bg-gray-50 p-4 font-sans text-gray-800 antialiased">
    <div class="max-w-md mx-auto space-y-6">
        
        <!-- MAIN FORM PANEL -->
        <div class="bg-white rounded-2xl shadow-xl p-6 space-y-5 border border-gray-100">
            <div class="text-center space-y-1">
                <h2 class="text-2xl font-extrabold text-indigo-600 tracking-tight">Aeris Beaute</h2>
                <p class="text-xs font-semibold text-gray-400 uppercase tracking-wider">Stock Opname System v2026</p>
            </div>
            
            <hr class="border-gray-100">

            <!-- Team Choice -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Counter Team</label>
                <select id="counterTeam" onchange="fetchHistory()" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-medium bg-white transition">
                    <option value="Team 1">Team 1</option>
                    <option value="Team 2">Team 2</option>
                    <option value="Team 3">Team 3</option>
                    <option value="Team 4">Team 4</option>
                </select>
            </div>

            <!-- Scanner Frame -->
            <div class="border-2 border-dashed border-indigo-200 p-2 rounded-2xl bg-indigo-50/30 overflow-hidden">
                <div id="reader" class="w-full rounded-xl overflow-hidden bg-black"></div>
                <div class="flex justify-around space-x-3 mt-2">
                    <button onclick="startScan('location')" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-4 rounded-xl text-sm transition shadow-sm w-1/2">
                        📷 Scan Location QR
                    </button>
                    <button onclick="startScan('sku')" id="scanSkuBtn" class="bg-gray-400 text-white font-bold py-2.5 px-4 rounded-xl text-sm transition shadow-sm w-1/2" disabled>
                        🏷️ Scan SKU Barcode
                    </button>
                </div>
            </div>

            <!-- Location Output (Triggers SKU Unlock) -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Precise Location</label>
                <input type="text" id="location" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 bg-orange-50 font-mono font-bold text-orange-700 border-orange-300 placeholder-orange-400" readonly placeholder="⚠️ SCAN LOCATION QR TO UNLOCK FORM">
            </div>

            <!-- Cascading Level 1 -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">1. Goods Type</label>
                <select id="skuType" onchange="updateCategories()" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold bg-gray-50 text-gray-400 transition" disabled>
                    <option value="">-- Locked --</option>
                </select>
            </div>

            <!-- Cascading Level 2 -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">2. Product Category</label>
                <select id="skuCategory" onchange="updateSkus()" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold bg-gray-50 text-gray-400 transition" disabled>
                    <option value="">-- Locked --</option>
                </select>
            </div>

            <!-- Cascading Level 3 -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">3. Target SKU Code</label>
                <select id="skuSelector" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold bg-gray-50 text-gray-400 transition" disabled>
                    <option value="">-- Locked --</option>
                </select>
            </div>

            <!-- Counter Control -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Physical Count</label>
                <div class="flex items-center space-x-2 mt-1">
                    <button type="button" onclick="adjustCount(-10)" class="bg-gray-200 hover:bg-gray-300 font-extrabold text-xl px-4 py-2 rounded-xl transition">-10</button>
                    <button type="button" onclick="adjustCount(-1)" class="bg-gray-200 hover:bg-gray-300 font-extrabold text-xl px-4 py-2 rounded-xl transition">-1</button>
                    <input type="number" id="count" class="w-full border-2 border-gray-200 p-3 rounded-xl focus:border-indigo-500 focus:outline-none text-center text-2xl font-black text-gray-900" value="0" min="0">
                    <button type="button" onclick="adjustCount(1)" class="bg-gray-200 hover:bg-gray-300 font-extrabold text-xl px-4 py-2 rounded-xl transition">+1</button>
                    <button type="button" onclick="adjustCount(10)" class="bg-gray-200 hover:bg-gray-300 font-extrabold text-xl px-4 py-2 rounded-xl transition">+10</button>
                </div>
            </div>

            <!-- Notes -->
            <div>
                <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Notes</label>
                <input type="text" id="notes" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none text-sm transition" placeholder="e.g., damaged box">
            </div>

            <button id="submitBtn" onclick="submitData()" class="w-full bg-emerald-500 hover:bg-emerald-600 active:scale-[0.98] text-white font-extrabold text-lg p-4 rounded-xl shadow-md hover:shadow-lg transition transform duration-150">
                📤 SUBMIT TO MASTER SHEET
            </button>
        </div>

        <!-- ACTIVITY LOG CONTEXT CONTAINER -->
        <div class="bg-white rounded-2xl shadow-lg p-6 border border-gray-100 space-y-4">
            <div class="flex justify-between items-center">
                <h3 class="text-md font-bold text-gray-700 uppercase tracking-wide">📋 Your Recent Activity</h3>
                <button onclick="fetchHistory()" class="text-indigo-600 hover:text-indigo-800 text-xs font-bold flex items-center">🔄 Refresh</button>
            </div>
            <div id="historyContainer" class="space-y-3 max-h-64 overflow-y-auto pr-1 text-sm">
                <p class="text-gray-400 text-center py-4">Select a team to stream history data...</p>
            </div>
        </div>
    </div>

    <script>
        const skuTree = {{ sku_tree|tojson|safe }};
        let currentTarget = '';
        const html5QrcodeScanner = new Html5Qrcode("reader");

        window.onload = () => { 
            fetchHistory(); 
        };

        function unlockFormForLocation() {
            // Style shift for location field
            const locInput = document.getElementById('location');
            locInput.classList.replace('bg-orange-50', 'bg-emerald-50');
            locInput.classList.replace('text-orange-700', 'text-emerald-700');
            locInput.classList.replace('border-orange-300', 'border-emerald-300');

            // Unlock SKU barcode scan button
            const scanSkuBtn = document.getElementById('scanSkuBtn');
            scanSkuBtn.disabled = false;
            scanSkuBtn.classList.replace('bg-gray-400', 'bg-blue-600');
            scanSkuBtn.classList.add('hover:bg-blue-700');

            // Unlock Type Dropdown
            const typeSelect = document.getElementById('skuType');
            typeSelect.disabled = false;
            typeSelect.classList.replace('bg-gray-50', 'bg-white');
            typeSelect.classList.replace('text-gray-400', 'text-gray-800');
            
            typeSelect.innerHTML = '<option value="">-- Choose Finished / Unfinished --</option>';
            Object.keys(skuTree).sort().forEach(type => {
                typeSelect.options[typeSelect.options.length] = new Option(type, type);
            });
        }

        function lockFormPostSubmit() {
            // Re-lock location field styles
            const locInput = document.getElementById('location');
            locInput.value = '';
            locInput.className = "w-full border-2 border-orange-300 p-3 rounded-xl mt-1 bg-orange-50 font-mono font-bold text-orange-700 placeholder-orange-400";
            
            // Re-lock barcode scanner button
            const scanSkuBtn = document.getElementById('scanSkuBtn');
            scanSkuBtn.disabled = true;
            scanSkuBtn.className = "bg-gray-400 text-white font-bold py-2.5 px-4 rounded-xl text-sm transition shadow-sm w-1/2";

            // Re-lock cascade levels completely
            ['skuType', 'skuCategory', 'skuSelector'].forEach(id => {
                const el = document.getElementById(id);
                el.disabled = true;
                el.innerHTML = '<option value="">-- Locked --</option>';
                el.className = "w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold bg-gray-50 text-gray-400 transition";
            });
        }

        function updateCategories() {
            const typeVal = document.getElementById('skuType').value;
            const catSelect = document.getElementById('skuCategory');
            const skuSelect = document.getElementById('skuSelector');
            
            catSelect.innerHTML = '<option value="">-- Choose Category --</option>';
            skuSelect.innerHTML = '<option value="">-- Choose SKU --</option>';
            skuSelect.disabled = true;
            skuSelect.className = "w-full border-2 border-gray-200 p-3 rounded-xl mt-1 font-semibold bg-gray-50 text-gray-400 transition";

            if (!typeVal || !skuTree[typeVal]) {
                catSelect.disabled = true;
                catSelect.className = "w-full border-2 border-gray-200 p-3 rounded-xl mt-1 font-semibold bg-gray-50 text-gray-400 transition";
                return;
            }

            catSelect.disabled = false;
            catSelect.className = "w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold bg-white text-gray-800 transition";
            Object.keys(skuTree[typeVal]).sort().forEach(cat => {
                catSelect.options[catSelect.options.length] = new Option(cat, cat);
            });
        }

        function updateSkus() {
            const typeVal = document.getElementById('skuType').value;
            const catVal = document.getElementById('skuCategory').value;
            const skuSelect = document.getElementById('skuSelector');
            
            skuSelect.innerHTML = '<option value="">-- Choose SKU --</option>';

            if (!catVal || !skuTree[typeVal] || !skuTree[typeVal][catVal]) {
                skuSelect.disabled = true;
                skuSelect.className = "w-full border-2 border-gray-200 p-3 rounded-xl mt-1 font-semibold bg-gray-50 text-gray-400 transition";
                return;
            }

            skuSelect.disabled = false;
            skuSelect.className = "w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold bg-white text-gray-800 transition";
            skuTree[typeVal][catVal].sort().forEach(sku => {
                skuSelect.options[skuSelect.options.length] = new Option(sku, sku);
            });
        }

        function startScan(target) {
            currentTarget = target;
            html5QrcodeScanner.start(
                { facingMode: "environment" },
                { fps: 15, qrbox: { width: 250, height: 250 } },
                (decodedText) => {
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
                            alert('⚠️ Barcode Matrix Match Failure: ' + text);
                        }
                    }
                    html5QrcodeScanner.stop();
                },
                (errorMessage) => {}
            );
        }

        function adjustCount(amount) {
            const countInput = document.getElementById('count');
            let currentVal = parseInt(countInput.value) || 0;
            currentVal += amount;
            if (currentVal < 0) currentVal = 0;
            countInput.value = currentVal;
        }

        async function fetchHistory() {
            const team = document.getElementById('counterTeam').value;
            const container = document.getElementById('historyContainer');
            container.innerHTML = `<p class="text-gray-400 text-center py-4 animate-pulse">Loading sheet records...</p>`;
            
            try {
                const response = await fetch(`/history?team=${encodeURIComponent(team)}`);
                const data = await response.json();
                
                if (data.length === 0) {
                    container.innerHTML = `<p class="text-gray-400 text-center py-4">No records found for this team yet.</p>`;
                    return;
                }
                
                container.innerHTML = data.map(item => `
                    <div class="bg-gray-50 border border-gray-100 p-3 rounded-xl flex justify-between items-center shadow-sm">
                        <div class="space-y-0.5">
                            <div class="font-mono font-bold text-indigo-600">${item.location}</div>
                            <div class="font-semibold text-gray-700">${item.sku}</div>
                            <div class="text-xs text-gray-400 font-medium">${item.notes || 'No remarks'}</div>
                        </div>
                        <div class="flex items-center space-x-3">
                            <div class="text-right font-black text-xl text-gray-900 px-2">${item.count}</div>
                            <div class="flex flex-col space-y-1">
                                <button onclick="editItem('${item.id}', ${item.count})" class="bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold px-2 py-1 rounded-md transition">Edit</button>
                                <button onclick="deleteItem('${item.id}')" class="bg-rose-50 hover:bg-rose-100 text-rose-600 text-xs font-bold px-2 py-1 rounded-md transition">Delete</button>
                            </div>
                        </div>
                    </div>
                `).join('');
            } catch (err) {
                container.innerHTML = `<p class="text-rose-500 text-center py-4">Failed to fetch history feed.</p>`;
            }
        }

        async function submitData() {
            const locInput = document.getElementById('location').value;
            const skuInput = document.getElementById('skuSelector').value;
            const countInput = document.getElementById('count').value;
            const btn = document.getElementById('submitBtn');

            if (!locInput || !skuInput || countInput === '') {
                alert('Please fill out Location, SKU selection, and Count parameters.');
                return;
            }

            btn.disabled = true;
            btn.innerText = "TRANSMITTING DATA...";
            btn.classList.replace('bg-emerald-500', 'bg-gray-400');

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
                    // Triggered by Duplicate Warning interceptor
                    alert(`⚠️ DUPLICATE ENTRY BLOCKER:\\n\\n${result.message}`);
                } else if (response.ok) {
                    alert('✅ Data logged safely!');
                    document.getElementById('count').value = '0';
                    document.getElementById('notes').value = '';
                    lockFormPostSubmit(); // Hard clear and re-lock until next location scan
                    fetchHistory();
                } else {
                    alert('❌ Connection Error: Sync failed.');
                }
            } catch (err) {
                alert('❌ Transmission Failed.');
            } finally {
                btn.disabled = false;
                btn.innerText = "SUBMIT TO MASTER SHEET";
                if (document.getElementById('location').value !== '') {
                    btn.classList.replace('bg-gray-400', 'bg-emerald-500');
                }
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