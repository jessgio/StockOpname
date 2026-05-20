import os
import json
from flask import Flask, render_template_string, request, jsonify
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import uuid

app = Flask(__name__)

# 1. Access Google Credentials safely from Vercel's environment variables
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_raw = os.environ.get("GOOGLE_CREDENTIALS")

if creds_raw:
    creds_dict = json.loads(creds_raw)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    # Connects to your exact master spreadsheet template
    sheet = client.open("Aeris Beaute - Stock Opname Master Template").worksheet("Raw Counts")
else:
    raise RuntimeError("GOOGLE_CREDENTIALS environment variable is missing!")

# 2. Embedded HTML Mobile UI with Tailwind CSS and Html5-Qrcode
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Aeris Opname 2026</title>
    <!-- Tailwind CSS for modern look and feel -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <!-- Open-source high-speed QR/Barcode scanning framework -->
    <script src="https://unpkg.com/html5-qrcode"></script>
</head>
<body class="bg-gray-50 p-4 font-sans text-gray-800 antialiased">
    <div class="max-w-md mx-auto bg-white rounded-2xl shadow-xl p-6 space-y-5 border border-gray-100">
        
        <div class="text-center space-y-1">
            <h2 class="text-2xl font-extrabold text-indigo-600 tracking-tight">Aeris Beaute</h2>
            <p class="text-xs font-semibold text-gray-400 uppercase tracking-wider">Stock Opname System v2026</p>
        </div>
        
        <hr class="border-gray-100">

        <!-- Step 1: Team Selection -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Counter Team</label>
            <select id="counterTeam" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-medium bg-white transition">
                <option value="Team 1">Team 1</option>
                <option value="Team 2">Team 2</option>
                <option value="Team 3">Team 3</option>
                <option value="Team 4">Team 4</option>
            </select>
        </div>

        <!-- Camera Scanner Viewfinder -->
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

        <!-- Step 2: Location Data -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Precise Location</label>
            <input type="text" id="location" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 bg-gray-100 font-mono font-bold text-indigo-700" readonly placeholder="Scan Location QR Code First">
        </div>

        <!-- Step 3: SKU Entry -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">SKU Code</label>
            <input type="text" id="sku" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none font-semibold transition" placeholder="Scan item or type manually">
        </div>

        <!-- Step 4: Physical Count Input -->
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

        <!-- Step 5: Optional Notes -->
        <div>
            <label class="block text-xs font-bold text-gray-500 uppercase tracking-wide">Notes</label>
            <input type="text" id="notes" class="w-full border-2 border-gray-200 p-3 rounded-xl mt-1 focus:border-indigo-500 focus:outline-none text-sm transition" placeholder="e.g., damaged box, promo sticker bundle">
        </div>

        <!-- Submit Button -->
        <button onclick="submitData()" class="w-full bg-emerald-500 hover:bg-emerald-600 active:scale-[0.98] text-white font-extrabold text-lg p-4 rounded-xl shadow-md hover:shadow-lg transition transform duration-150">
            📤 SUBMIT TO MASTER SHEET
        </button>
    </div>

    <script>
        let currentTarget = '';
        const html5QrcodeScanner = new Html5Qrcode("reader");

        function startScan(target) {
            currentTarget = target;
            // Configured to favor the high-resolution back camera on modern mobile phones
            html5QrcodeScanner.start(
                { facingMode: "environment" },
                { fps: 15, qrbox: { width: 250, height: 250 } },
                (decodedText) => {
                    document.getElementById(currentTarget).value = decodedText;
                    html5QrcodeScanner.stop();
                },
                (errorMessage) => { /* Silently suppress continuous scanner logging noise */ }
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

            if (!locInput || !skuInput || countInput === '') {
                alert('🚨 Operational Error: Please fill out Location, SKU, and Count before submitting.');
                return;
            }

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
                    // Clear product metrics so they don't double count, but KEEP team and location active 
                    // so they can rapidly scan multiple distinct SKUs sitting on the exact same shelf.
                    document.getElementById('sku').value = '';
                    document.getElementById('count').value = '0';
                    document.getElementById('notes').value = '';
                } else {
                    alert('❌ Connection Error: Data could not be saved. Contact Control Desk immediately.');
                }
            } catch (err) {
                alert('❌ Transmission Failed: Check cellular connection.');
            }
        }
    </script>
</body>
</html>
"""

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/submit', methods=['POST'])
def submit():
    data = request.json
    log_id = str(uuid.uuid4())[:8] 
    
    # Automatically extracts the top-level Zone letter (e.g., extracts "A" out of "A-01-A")[cite: 1]
    zone_prefix = data['location'].split('-')[0] if '-' in data['location'] else data['location'][:1]
    
    # Maps exactly to your 8 actual spreadsheet headers[cite: 1]
    row_to_append = [
        log_id,              # Column A: Log ID[cite: 1]
        data['team'],        # Column B: Counter Team[cite: 1]
        zone_prefix,         # Column C: Zone[cite: 1]
        data['location'],    # Column D: Precise Location[cite: 1]
        data['sku'],         # Column E: SKU Code[cite: 1]
        "",                  # Column F: Item Name (Left blank for your sheet's automatic XLOOKUP/VLOOKUP formula)[cite: 1]
        int(data['count']),  # Column G: Physical Count[cite: 1]
        data['notes']        # Column H: Notes[cite: 1]
    ]
    
    sheet.append_row(row_to_append)
    return jsonify({"status": "success"}), 200

# Entry point wrapper required by Vercel for serverless Python executions
def handler(request, client):
    return app(request, client)