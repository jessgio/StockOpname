import os
import json
from flask import Flask, render_template_string, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
import uuid

app = Flask(__name__)

# Define scopes globally
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_sheet_client():
    """Establishes connection to the spreadsheet on-demand instead of blocking global app startup."""
    creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_raw:
        raise RuntimeError("GOOGLE_CREDENTIALS environment variable is missing!")
    
    creds_dict = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    
    # Connect directly to your master file setup
    return client.open("Aeris Beaute - Stock Opname Master Template").worksheet("Raw Counts")

# --- EMBEDDED HTML INTERFACE ---
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
    <div class="max-w-md mx-auto bg-white rounded-2xl shadow-xl p-6 space-y-5 border border-gray-100">
        
        <div class="text-center space-y-1">
            <h2 class="text-2xl font-extrabold text-indigo-600 tracking-tight">Aeris Beaute</h2>
            <p class="text-xs font-semibold text-gray-400 uppercase tracking-wider">Stock Opname System v2026</p>
        </div>
        
        <hr class="border-gray-100">

        <!-- Team Input -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Counter Team</label>
            <select id="counterTeam" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-medium bg-white transition">
                <option value="Team 1">Team 1</option>
                <option value="Team 2">Team 2</option>
                <option value="Team 3">Team 3</option>
                <option value="Team 4">Team 4</option>
            </select>
        </div>

        <!-- Viewfinder Panel -->
        <div class="border-2 border-dashed border-indigo-200 p-2 rounded-2xl bg-indigo-50/30 overflow-hidden">
            <div id="reader" class="w-full rounded-xl overflow-hidden bg-black"></div>
            <div class="flex justify-around space-x-3 mt-2">
                <button onclick="startScan('location')" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-4 rounded-xl text-sm transition shadow-sm w-1/2">
                    📷 Scan Location QR
                </button>
                <button onclick="startScan('sku')" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-4 rounded-xl text-sm transition shadow-sm w-1/2">
                    🏷️ Scan SKU Barcode
                </button>
            </div>
        </div>

        <!-- Location Output -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Precise Location</label>
            <input type="text" id="location" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 bg-gray-100 font-mono font-bold text-indigo-700" readonly placeholder="Scan Location QR Code First">
        </div>

        <!-- SKU Code -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">SKU Code</label>
            <input type="text" id="sku" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold transition" placeholder="Scan item or type manually">
        </div>

        <!-- Metric Counter Control Box -->
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

        <!-- User Notes -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Notes</label>
            <input type="text" id="notes" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none text-sm transition" placeholder="e.g., damaged box, promo pack">
        </div>

        <!-- Submitting Trigger Button -->
        <button id="submitBtn" onclick="submitData()" class="w-full bg-emerald-500 hover:bg-emerald-600 active:scale-[0.98] text-white font-extrabold text-lg p-4 rounded-xl shadow-md hover:shadow-lg transition transform duration-150">
            📤 SUBMIT TO MASTER SHEET
        </button>
    </div>

    <script>
        let currentTarget = '';
        const html5QrcodeScanner = new Html5Qrcode("reader");

        function startScan(target) {
            currentTarget = target;
            html5QrcodeScanner.start(
                { facingMode: "environment" },
                { fps: 15, qrbox: { width: 250, height: 250 } },
                (decodedText) => {
                    document.getElementById(currentTarget).value = decodedText;
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

        async function submitData() {
            const locInput = document.getElementById('location').value;
            const skuInput = document.getElementById('sku').value.trim();
            const countInput = document.getElementById('count').value;
            const btn = document.getElementById('submitBtn');

            if (!locInput || !skuInput || countInput === '') {
                alert('Please fill out Location, SKU, and Count before submitting.');
                return;
            }

            // Simple UI lock to prevent aggressive accidental multi-clicks
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

                if (response.ok) {
                    alert('✅ Data logged safely to the Control Desk!');
                    document.getElementById('sku').value = '';
                    document.getElementById('count').value = '0';
                    document.getElementById('notes').value = '';
                } else {
                    alert('❌ Connection Error: Verification failed at Control Desk.');
                }
            } catch (err) {
                alert('❌ Transmission Failed: Check network connectivity.');
            } finally {
                btn.disabled = false;
                btn.innerText = "SUBMIT TO MASTER SHEET";
                btn.classList.replace('bg-gray-400', 'bg-emerald-500');
            }
        }
    </script>
</body>
</html>
"""

# --- ROUTES ---

@app.route('/')
def home():
    """Serves the frontend layout instantly with zero network lookup latency."""
    return render_template_string(HTML_TEMPLATE)

@app.route('/submit', methods=['POST'])
def submit():
    """Handles the heavy data synchronization step entirely within an active lifecycle request."""
    try:
        data = request.json
        log_id = str(uuid.uuid4())[:8] 
        
        # Pull out Zone prefix cleanly (e.g. "A" from "A-01-A")[cite: 1]
        zone_prefix = data['location'].split('-')[0] if '-' in data['location'] else data['location'][:1]
        
        row_to_append = [
            log_id,              # Column A: Log ID
            data['team'],        # Column B: Counter Team[cite: 1]
            zone_prefix,         # Column C: Zone[cite: 1]
            data['location'],    # Column D: Precise Location[cite: 1]
            data['sku'],         # Column E: SKU Code[cite: 1]
            "",                  # Column F: Item Name (Handled by Sheet Formula)[cite: 1]
            int(data['count']),  # Column G: Physical Count[cite: 1]
            data['notes']        # Column H: Notes[cite: 1]
        ]
        
        # Connect dynamically and push data instantly
        sheet = get_sheet_client()
        sheet.append_row(row_to_append)
        
        return jsonify({"status": "success"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    
# [ADD THIS AT THE VERY BOTTOM OF YOUR api/index.py FILE]

@app.route('/favicon.ico')
@app.route('/favicon.png')
def favicon():
    """Safely intercepts automated browser icon requests to prevent 500 routing crashes."""
    return '', 204

# Essential fallback binding for Vercel's WSGI routing engine
app = app