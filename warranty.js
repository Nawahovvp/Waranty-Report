const SCRAP_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C1MeLvv2PrOEQsqmXQplZt1xLuD5Z4KUGOrnGjYrh_A/gviz/tq?tqx=out:csv&sheet=Scrap';
const WORKFILTER_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C-_TRQcz2w0ia9bF9BfFcln5X_dX8T_cNfVUI-oU8ho/gviz/tq?tqx=out:csv&sheet=WorkFilter';
const EMPLOYEE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1eqVoLsZxGguEbRCC5rdI4iMVtQ7CK4T3uXRdx8zE3uw/gviz/tq?tqx=out:csv&sheet=EmployeeWeb';
const WARRANTY_SHEET_URL = 'https://docs.google.com/spreadsheets/d/19WWSESBcerEUlEDzM4lu3cglPoTHOWUl4P0prexqYpc/gviz/tq?tqx=out:csv&sheet=WarrantyData';
const TECHNICIAN_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1EcP7N1Yr6fAzl17ItoqPstaxrIIf1Aly/gviz/tq?tqx=out:csv&sheet=Data';
const PARTS_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C-_TRQcz2w0ia9bF9BfFcln5X_dX8T_cNfVUI-oU8ho/gviz/tq?tqx=out:csv&sheet=Parts';

// PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL HERE
const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwq_ocJV50pVeUhuYJ5NJ114PPLDzCXuelrkYP5Mz33krtkbFp7VP5pzQEuzXnrGlDyXA/exec';

const COLUMNS = [
    { header: '<input type="checkbox" id="selectAll" onclick="toggleAllCheckboxes(this)">', source: 'C', key: 'checkbox' },
    { header: 'Status', source: 'C', key: 'status' },
    { header: 'Work Order', source: 'S', key: 'work order' },
    { header: 'Spare Part Code', source: 'S', key: 'Spare Part Code' },
    { header: 'Spare Part Name', source: 'S', key: 'Spare Part Name' },
    { header: 'Old Material Code', source: 'S', key: 'old material code' },
    { header: 'Person', source: 'P', key: 'Person' }, // Added Person Column
    { header: 'Qty', source: 'S', key: 'qty' },
    { header: 'Serial Number', source: 'W', key: 'Serial Number' },
    { header: 'Store Code', source: 'W', key: 'Store Code' },
    { header: 'Store Name', source: 'W', key: 'Store Name' },
    { header: 'Date Received', source: 'S', key: 'วันที่รับซาก' },
    { header: 'รหัสช่าง', source: 'S', key: 'รหัสช่าง' },
    { header: 'ชื่อช่าง', source: 'S', key: 'ชื่อช่าง' },
    { header: 'Mobile', source: 'T', key: 'Phone' },
    { header: 'Receiver', source: 'S', key: 'ผู้รับซาก' },
    { header: 'Plant', source: 'S', key: 'plant' },
    { header: 'Keep', source: 'S', key: 'Keep' },
    { header: 'CI Name', source: 'W', key: 'CI Name' },
    { header: 'Problem', source: 'W', key: 'Problem' },
    { header: 'Product Type', source: 'W', key: 'Product Type' },
    { header: 'Product', source: 'W', key: 'Product' }
];

const PLANT_MAPPING = {
    '301': 'นวนคร',
    '303': 'SA-สงขลา',
    '304': 'พระราม 3',
    '305': 'ราชบุรี',
    '307': 'สุราษฎร์',
    '309': 'นครราชสีมา',
    '310': 'SA-อุดร 1',
    '311': 'ศรีราชา',
    '312': 'พิษณุโลก',
    '313': 'ภูเก็ต',
    '315': 'SA-อยุธยา',
    '319': 'ขอนแก่น',
    '318': 'คลังวัตถุดิบ',
    '320': 'ลำปาง',
    '322': 'SA-อุดร 2',
    '323': 'SA-ลำลูกกา',
    '324': 'SA-ปัตตานี',
    '326': 'วิภาวดี 62',
    '330': 'ประเวศ',
    '362': 'SA-หนองแขม',
    '363': 'SA-ปากเกร็ด',
    '364': 'SA-บางบัวทอง',
    '365': 'SA-สมุทรปราการ',
    '366': 'เชียงใหม่',
    '367': 'SA-ฉะเชิงเทรา',
    '368': 'SA-ร้อยเอ็ด',
    '369': 'ระยอง'
};

const BOOKING_COLUMNS = [
    { header: '<input type="checkbox" id="selectAllBooking" onclick="toggleAllBookingCheckboxes(this)">', key: 'checkbox' },
    { header: 'Status', key: 'CustomStatus' }, // Added Status
    { header: 'ใบจองรถ', key: 'Booking Slip' }, // Moved to 2nd position
    { header: 'Work Order', key: 'Work Order' },
    { header: 'Spare Part Code', key: 'Spare Part Code' },
    { header: 'Spare Part Name', key: 'Spare Part Name' },
    { header: 'Old Material Code', key: 'Old Material Code' },
    { header: 'Qty', key: 'Qty' },
    { header: 'Serial Number', key: 'Serial Number' },
    { header: 'Store Code', key: 'Store Code' },
    { header: 'Store Name', key: 'Store Name' },
    { header: 'Claim Receiver', key: 'Claim Receiver' }, // Added Claim Receiver
    { header: 'รหัสช่าง', key: 'รหัสช่าง' },
    { header: 'ชื่อช่าง', key: 'ชื่อช่าง' },
    { header: 'Mobile', key: 'Mobile' },
    { header: 'Plantcenter', key: 'Plantcenter' },
    { header: 'Plant', key: 'Plant' },
    { header: 'Product', key: 'Product' },
    { header: 'Warranty Action', key: 'Warranty Action' },
    { header: 'Recorder', key: 'Recorder' },
    { header: 'Recripte', key: 'Recripte' },
    { header: 'Recripte Date', key: 'RecripteDate' },
    { header: 'Timestamp', key: 'Timestamp' },
    { header: 'วันที่จองรถ', key: 'Booking Date' },
];


const SUPPLIER_COLUMNS = [
    { header: '<input type="checkbox" id="selectAllSupplier" onclick="toggleAllSupplierCheckboxes(this)">', key: 'checkbox' }, // Enabled header
    { header: 'Status', key: 'CustomStatus' }, // Added Status
    { header: 'ใบจองรถ', key: 'Booking Slip' },
    { header: 'Work Order', key: 'Work Order' },
    { header: 'Spare Part Code', key: 'Spare Part Code' },
    { header: 'Spare Part Name', key: 'Spare Part Name' },
    { header: 'Old Material Code', key: 'Old Material Code' },
    { header: 'Qty', key: 'Qty' },
    { header: 'Serial Number', key: 'Serial Number' },
    { header: 'Store Code', key: 'Store Code' },
    { header: 'Store Name', key: 'Store Name' },
    { header: 'Claim Receiver', key: 'Claim Receiver' },
    { header: 'รหัสช่าง', key: 'รหัสช่าง' },
    { header: 'ชื่อช่าง', key: 'ชื่อช่าง' },
    { header: 'Mobile', key: 'Mobile' },
    { header: 'Plantcenter', key: 'Plantcenter' },
    { header: 'Plant', key: 'Plant' },
    { header: 'Product', key: 'Product' },
    { header: 'Warranty Action', key: 'Warranty Action' },
    { header: 'Recorder', key: 'Recorder' },
    { header: 'Recripte', key: 'Recripte' },
    { header: 'Recripte Date', key: 'RecripteDate' },
    { header: 'Timestamp', key: 'Timestamp' },
    { header: 'วันที่จองรถ', key: 'Booking Date' },
    { header: 'Claim Date', key: 'Claim Date' },
    { header: 'ClaimSup', key: 'ClaimSup' }
];

function getComputedStatus(item) {
    if (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '') {
        return '<span class="status-badge status-claim-sent">ส่งเคลมแล้ว</span>';
    }
    if (item['Recripte'] && String(item['Recripte']).trim() !== '') {
        return '<span class="status-badge status-waiting-claim">รอส่งเคลม</span>';
    }
    if (item['Booking Slip'] && String(item['Booking Slip']).trim() !== '') {
        return '<span class="status-badge status-transit">ระหว่างขนส่ง</span>';
    }
    return '<span class="status-badge status-local">คลังพื้นที่</span>';
}



let currentBookingPage = 1;
const ITEMS_PER_PAGE_BOOKING = 20;

let allEmployees = [];
let fullData = []; // Master data (filtered by plant)
let displayedData = []; // Data currently shown (filtered by UI)
let globalBookingData = []; // Store raw warranty data for second tab
let currentPage = 1;
let editingItem = null; // Item currently being edited
let currentSupplierPage = 1;
let currentClaimSentPage = 1;
let currentHistoryPage = 1;
const ITEMS_PER_PAGE = 20;
let supplierProductOptions = new Set(); // Store selected products for Supplier Tab filter



async function fetchData(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Failed to fetch ${url}`);
        const csvText = await response.text();
        return new Promise((resolve, reject) => {
            Papa.parse(csvText, {
                header: true,
                skipEmptyLines: true,
                complete: (results) => resolve(results.data),
                error: (err) => reject(err)
            });
        });
    } catch (e) {
        console.error("Fetch Error:", e);
        return [];
    }
}

async function initAuth() {
    allEmployees = await fetchData(EMPLOYEE_SHEET_URL);
    allEmployees = await fetchData(EMPLOYEE_SHEET_URL);

    const savedUser = localStorage.getItem('currentUser');
    if (savedUser) {
        const user = JSON.parse(savedUser);
        showApp(user);
    } else {
        document.getElementById('loginOverlay').style.display = 'flex';
    }
}

function handleLogin() {
    const idInput = document.getElementById('idInput').value.trim();
    const passInput = document.getElementById('passInput').value.trim();
    const errorMsg = document.getElementById('loginError');

    if (!idInput || !passInput) {
        errorMsg.textContent = "Please enter ID and Password";
        errorMsg.style.display = 'block';
        return;
    }

    const user = allEmployees.find(emp => emp['IDRec'] === idInput);

    if (user) {
        const expectedPass = user['IDRec'].slice(-4);
        if (passInput === expectedPass) {
            localStorage.setItem('currentUser', JSON.stringify(user));
            errorMsg.style.display = 'none';
            showApp(user);
        } else {
            errorMsg.textContent = "Incorrect Password";
            errorMsg.style.display = 'block';
        }
    } else {
        errorMsg.textContent = "User ID not found";
        errorMsg.style.display = 'block';
    }
}

function handleLogout() {
    localStorage.removeItem('currentUser');
    location.reload();
}

function showApp(user) {
    document.getElementById('loginOverlay').style.display = 'none';
    document.getElementById('mainApp').style.display = 'flex';

    document.getElementById('displayUserName').textContent = user['Name'] || user['IDRec'];
    const details = [
        user['ตำแหน่ง'],
        user['หน่วยงาน'],
        user['Team'],
        user['Plant']
    ].filter(Boolean).join(' | ');
    document.getElementById('displayUserDetail').textContent = details;

    loadTableData();
}

function toggleMenu() {
    const menu = document.getElementById('userDropdown');
    menu.classList.toggle('show');
}

window.onclick = function (event) {
    if (!event.target.closest('.menu-container')) {
        const dropdowns = document.getElementsByClassName("dropdown-menu");
        for (let i = 0; i < dropdowns.length; i++) {
            const openDropdown = dropdowns[i];
            if (openDropdown.classList.contains('show')) {
                openDropdown.classList.remove('show');
            }
        }
    }
}



function switchTab(tabId) {
    // Update Header
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    // Match clicking button text or index? Simplest to assume 2 tabs.
    // Actually, querying by onclick attribute is messy.
    // Let's rely on event target or simple class toggle.
    // Refactor: We need to highlight the correct nav button.
    // Let's assume order: 0 = main, 1 = booking
    const tabs = document.querySelectorAll('.tab-btn');
    // Hide all contents first
    document.querySelectorAll('.tab-content').forEach(el => el.style.display = 'none');
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));

    if (tabId === 'main') {
        tabs[0].classList.add('active');
        document.getElementById('tab-content-main').style.display = 'flex';
        document.getElementById('tab-content-main').classList.add('active');
    } else if (tabId === 'booking') {
        tabs[1].classList.add('active');
        document.getElementById('tab-content-booking').style.display = 'flex';
        document.getElementById('tab-content-booking').classList.add('active');
        populateBookingFilter();
        renderBookingTable();
    } else if (tabId === 'navanakorn') {
        tabs[2].classList.add('active');
        document.getElementById('tab-content-navanakorn').style.display = 'block';
        document.getElementById('tab-content-navanakorn').classList.add('active');
        document.getElementById('tab-content-navanakorn').classList.add('active');
        backToDeck('navanakorn'); // Reset View
        renderDeckView('0301', 'navaNakornDeck', 'navanakorn'); // Pass tab key
    } else if (tabId === 'vibhavadi') {
        tabs[3].classList.add('active');
        document.getElementById('tab-content-vibhavadi').style.display = 'block';
        document.getElementById('tab-content-vibhavadi').classList.add('active');
        backToDeck('vibhavadi'); // Reset View
        renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi'); // Pass tab key
    } else if (tabId === 'supplier') {
        tabs[4].classList.add('active');
        document.getElementById('tab-content-supplier').style.display = 'flex';
        document.getElementById('tab-content-supplier').classList.add('active');
        populateSupplierFilter(); // NEW: Filter for supplier
        renderSupplierTable(); // Render Table
    } else if (tabId === 'claimSent') {
        tabs[5].classList.add('active');
        document.getElementById('tab-content-claimSent').style.display = 'flex';
        document.getElementById('tab-content-claimSent').classList.add('active');
        populateClaimSentFilter();
        renderClaimSentTable();
    } else if (tabId === 'history') {
        tabs[6].classList.add('active');
        document.getElementById('tab-content-history').style.display = 'flex';
        document.getElementById('tab-content-history').classList.add('active');
        renderHistoryTable();
    }
}

async function loadTableData() {
    try {
        const [scrapData, workFilterData, warrantyData, technicianData, partsData] = await Promise.all([
            fetchData(SCRAP_SHEET_URL),
            fetchData(WORKFILTER_SHEET_URL),
            fetchData(WARRANTY_SHEET_URL),
            fetchData(TECHNICIAN_SHEET_URL),
            fetchData(PARTS_SHEET_URL)
        ]);

        // Create Parts Map (Spare Part Code -> Person) AND (Spare Part Code -> Name)
        const partsMap = new Map();
        const partsNameMap = new Map();

        if (partsData && partsData.length > 0) {
            const headers = Object.keys(partsData[0]);
            // Find valid keys dynamically (Case Insensitive, Trimmed)
            const codeKey = headers.find(h => h.toLowerCase().includes('spare') && h.toLowerCase().includes('code')) || 'Spare Part Code';
            const personKey = headers.find(h => h.toLowerCase().includes('person')) || 'Person';
            const nameKey = headers.find(h => h.toLowerCase().includes('spare') && h.toLowerCase().includes('name')) || 'Spare Part Name';

            partsData.forEach(row => {
                const rawCode = row[codeKey];
                const rawPerson = row[personKey];
                const rawName = row[nameKey];

                const code = rawCode ? String(rawCode).trim() : '';
                const person = rawPerson ? String(rawPerson).trim() : '';
                const name = rawName ? String(rawName).trim() : '';

                if (code) {
                    partsMap.set(code, person);
                    if (name) partsNameMap.set(code, name);

                    // DEBUG: Check specific failing code
                    if (code.includes('30053786')) { // Fuzzy check for debug
                    }
                }
            });
        }

        // Create Set of Saved Keys (KEY column in WarrantyData)
        const workFilterMap = new Map();
        const workFilterDigitMap = new Map(); // Fuzzy Match Helper
        let wfCodeKey = 'Work Order Number 1';
        let wfProductKey = 'Product';

        if (workFilterData.length > 0) {
            const headers = Object.keys(workFilterData[0]);
            wfCodeKey = headers.find(h => h.toLowerCase().includes('work') && (h.toLowerCase().includes('order') || h.toLowerCase().includes('no'))) || headers.find(h => h.toLowerCase().includes('key')) || 'Work Order Number 1';
            wfProductKey = headers.find(h => h.toLowerCase().includes('product')) || 'Product';
            console.log(`[DEBUG] Detected WorkFilter Keys - Code: ${wfCodeKey}, Product: ${wfProductKey}`);
        }

        workFilterData.forEach(row => {
            const rawKey = row[wfCodeKey];
            const key = rawKey?.trim().toLowerCase();
            if (key) workFilterMap.set(key, row);



            // Digit Map
            if (rawKey) {
                const digits = String(rawKey).replace(/\D/g, '');
                if (digits.length > 3) { // Only map if sufficient length to avoid false positives with small numbers
                    workFilterDigitMap.set(digits, row);
                }
            }
        });
        console.log(`[DEBUG] workFilterMap size: ${workFilterMap.size}`);

        // Get Current User Plant & Normalize
        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
        const userPlant = currentUser['Plant'] ? currentUser['Plant'].toString().trim().padStart(4, '0') : null;

        // Create Set of Saved Keys AND Filter Global Booking Data
        let warrantyWOKey = 'Work Order';
        if (warrantyData.length > 0) {
            const wHeaders = Object.keys(warrantyData[0]);
            warrantyWOKey = wHeaders.find(h => h.toLowerCase().includes('work') && h.toLowerCase().includes('order')) || 'Work Order';
            console.log('[DEBUG] Warranty Row Keys:', wHeaders);
            console.log('[DEBUG] Detected Warranty WO Key:', warrantyWOKey);
        }

        globalBookingData = warrantyData.filter(row => {
            const action = row['ActionStatus'] || row['Warranty Action'] || '';
            return action === 'เคลมประกัน';
        }).map(item => {
            // Enrich with Product from WorkFilter if missing
            if (!item['Product']) {
                const rawWO = String(item[warrantyWOKey] || '');
                const wo = rawWO.trim().toLowerCase();

                // 1. Exact Match
                if (wo && workFilterMap.has(wo)) {
                    item['Product'] = workFilterMap.get(wo)[wfProductKey];
                }
                // 2. Fuzzy Digit Match
                else {
                    const digits = rawWO.replace(/\D/g, '');
                    if (digits.length > 3 && workFilterDigitMap.has(digits)) {
                        item['Product'] = workFilterDigitMap.get(digits)[wfProductKey];
                        // console.log(`[DEBUG] Fuzzy Match for ${wo} -> Product: ${item['Product']}`);
                    }
                }
            }
            return item;
        });

        const itemsWithProduct = globalBookingData.filter(i => i['Product']).length;
        console.log(`[DEBUG] GlobalBookingData items with Product: ${itemsWithProduct} / ${globalBookingData.length}`);

        // Render Booking Table immediately
        populateBookingFilter();
        renderBookingTable();

        // Create Map of keys -> Status (Warranty Action)
        const savedStatusMap = new Map();
        warrantyData.forEach(row => {
            const key = row['KEY'] ? row['KEY'].toString().trim() : '';
            if (key) {
                savedStatusMap.set(key, row['ActionStatus'] || row['Warranty Action'] || 'บันทึกแล้ว');
            }
        });

        // Create Map of Technician ID -> Phone
        const technicianMap = new Map();
        if (technicianData) {
            technicianData.forEach(row => {
                // Ensure ID is treated as a string and trimmed
                let id = row['IDEmployee'];
                if (id) {
                    id = String(id).trim();
                    // If duplicate, keep the first one (or robustly check if phone exists)
                    if (!technicianMap.has(id)) {
                        technicianMap.set(id, row['Phone'] || '');
                    }
                }
            });
        }

        // Detect Keys in Scrap Data (Robustness)
        let scrapCodeKey = 'Spare Part Code';
        let scrapNameKey = 'Spare Part Name';
        let scrapWOKey = 'work order';

        if (scrapData && scrapData.length > 0) {
            const scrapHeaders = Object.keys(scrapData[0]);
            scrapCodeKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'spare part code') || scrapCodeKey;
            scrapNameKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'spare part name') || scrapNameKey;
            scrapWOKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'work order') || scrapWOKey;
        }

        // Join Data & Normalize Data Plant
        fullData = scrapData.map(scrapRow => {
            // Normalize Plant in Data (e.g. "326" -> "0326")
            if (scrapRow['plant']) {
                scrapRow['plant'] = scrapRow['plant'].toString().trim().padStart(4, '0');
            }

            const workOrderKey = scrapRow[scrapWOKey]?.trim();

            // Generate KEY for status check
            // Use detected keys for access
            const codeVal = scrapRow[scrapCodeKey] || '';
            const uniqueKey = (workOrderKey || '') + codeVal;
            const statusValue = savedStatusMap.get(uniqueKey) || '';

            // Get Person and Name from Parts Map
            const partCode = codeVal.trim();
            const personValue = partsMap.get(partCode) || '';

            // Override Spare Part Name if found in Parts DB
            if (partCode && partsNameMap.has(partCode)) {
                scrapRow[scrapNameKey] = partsNameMap.get(partCode);
            }

            // Ensure we rely on standard keys for the rest of the app ??
            // OR normalize the row object itself?
            // Safer to normalize the row object to have standard keys!
            if (scrapCodeKey !== 'Spare Part Code') scrapRow['Spare Part Code'] = scrapRow[scrapCodeKey];
            if (scrapNameKey !== 'Spare Part Name') scrapRow['Spare Part Name'] = scrapRow[scrapNameKey];
            if (scrapWOKey !== 'work order') scrapRow['work order'] = scrapRow[scrapWOKey];


            return {
                scrap: scrapRow,
                fullRow: workFilterMap.get(workOrderKey) || {},
                status: statusValue,
                technicianPhone: technicianMap.get(String(scrapRow['รหัสช่าง'] || '').trim()) || '', // Safe Lookup
                person: personValue // Added Person
            };
        });

        // Filter by User Plant if exists
        if (userPlant) {
            fullData = fullData.filter(item => item.scrap['plant'] === userPlant);
        }

        // Initialize Displayed Data
        // Map person to globalBookingData for fallback logic
        globalBookingData = globalBookingData.map(item => {
            const partCode = item['Spare Part Code']?.trim();
            const personValue = partCode ? (partsMap.get(partCode) || '') : '';
            return {
                ...item,
                person: personValue
            };
        });

        displayedData = [...fullData];
        populateFilters();

        sortDisplayedData(); // Ensure sorted initially


        populateFilters();

        // Init Other Tabs - Split try/catch for robustness
        try {
            console.log('Calling populateBookingFilter...');
            if (typeof populateBookingFilter === 'function') populateBookingFilter();
            renderBookingTable();
        } catch (e) { console.error("Error init Booking:", e); }

        try {
            console.log('Calling populateSupplierFilter...');
            if (typeof populateSupplierFilter === 'function') populateSupplierFilter();
            else console.error('populateSupplierFilter is NOT a function');
            renderSupplierTable();
        } catch (e) { console.error("Error init Supplier:", e); }

        try {
            populateClaimSentFilter();
            renderClaimSentTable();
            renderHistoryTable();
            // Deck Views
            renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
            renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
        } catch (e) { console.error("Error init rest:", e); }

        sortDisplayedData(); // Ensure sorted initially
        updateTotalQty();

        const tableHeader = document.getElementById('tableHeader');
        tableHeader.innerHTML = '';
        COLUMNS.forEach(col => {
            const th = document.createElement('th');
            // Check if header contains HTML (like input checkbox)
            if (col.header.includes('<')) {
                th.innerHTML = col.header;
            } else {
                th.textContent = col.header;
            }
            tableHeader.appendChild(th);
        });

        document.getElementById('loadingIndicator').style.display = 'none';

        // Initial Render
        currentPage = 1; // Reset to page 1 on new data load
        renderTable();

    } catch (err) {
        console.error(err);
        document.getElementById('loadingIndicator').innerHTML = `<div class="error">Error loading data: ${err.message}</div>`;
    }
}

function populateFilters() {
    const filterSelect = document.getElementById('sparePartFilter');

    // Get unique Spare Part Names from fullData
    const uniqueParts = [...new Set(fullData.map(item => item.scrap['Spare Part Name']).filter(Boolean))].sort();

    // Keep current selection if possible, otherwise reset
    const currentSelection = filterSelect.value;

    filterSelect.innerHTML = '<option value="">All Spare Parts</option>';

    uniqueParts.forEach(part => {
        const option = document.createElement('option');
        option.value = part;
        option.textContent = part;
        filterSelect.appendChild(option);
    });

    if (uniqueParts.includes(currentSelection)) {
        filterSelect.value = currentSelection;
    }
}

function handleFilterChange() {
    currentPage = 1;
    applyFilters();
}

function handleSearch() {
    currentPage = 1;
    applyFilters();
}

function sortDisplayedData() {
    displayedData.sort((a, b) => {
        const nameA = a.scrap['Spare Part Name'] || '';
        const nameB = b.scrap['Spare Part Name'] || '';
        const nameCompare = nameA.localeCompare(nameB, 'th');
        if (nameCompare !== 0) return nameCompare;

        const codeA = a.scrap['Spare Part Code'] || '';
        const codeB = b.scrap['Spare Part Code'] || '';
        return codeA.localeCompare(codeB);
    });
}

function applyFilters() {
    const selectedPart = document.getElementById('sparePartFilter').value;
    const selectedStatus = document.getElementById('statusFilter').value;
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();

    displayedData = fullData.filter(item => {
        // 1. Filter by Spare Part
        const matchPart = selectedPart ? item.scrap['Spare Part Name'] === selectedPart : true;

        // 2. Filter by Status
        let matchStatus = true;
        if (selectedStatus === 'Pending') {
            matchStatus = !item.status; // Empty status means pending
        } else if (selectedStatus) {
            matchStatus = item.status === selectedStatus;
        }

        // 2. Filter by Search Term (All Columns)
        let matchSearch = true;
        if (searchTerm) {
            // Combine all values from scrap and fullRow
            const allValues = [
                ...Object.values(item.scrap),
                ...Object.values(item.fullRow)
            ].join(' ').toLowerCase();
            matchSearch = allValues.includes(searchTerm);
        }

        return matchPart && matchStatus && matchSearch;
    });

    sortDisplayedData();
    updateTotalQty();
    renderTable();
}

function updateTotalQty() {
    const total = displayedData.reduce((sum, item) => {
        const qty = parseFloat(item.scrap['qty']) || 0;
        return sum + qty;
    }, 0);
    document.getElementById('totalQtyValue').textContent = total.toLocaleString() + ' PC';
}

function renderTable() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';

    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = displayedData.slice(startIndex, endIndex);

    // Calculate Group Stats (Total & Completed)
    const groupStats = displayedData.reduce((acc, item) => {
        const code = item.scrap['Spare Part Code'] || 'Unknown';
        const qty = parseFloat(item.scrap['qty']) || 0;

        if (!acc[code]) acc[code] = { total: 0, completed: 0 };

        acc[code].total += qty;
        // Count as completed if Status is not empty
        if (item.status && String(item.status).trim() !== '') {
            acc[code].completed += qty;
        }
        return acc;
    }, {});

    let previousPartCode = null;
    // If we're on page > 1, check the last item of previous page to see if we are continuing a group
    if (currentPage > 1 && displayedData[startIndex - 1]) {
        previousPartCode = displayedData[startIndex - 1].scrap['Spare Part Code'];
    }

    pageData.forEach((item, index) => {
        const currentPartName = item.scrap['Spare Part Name'] || 'Unknown';
        const currentPartCode = item.scrap['Spare Part Code'] || 'Unknown';

        // Insert Group Header if Code Changed
        if (currentPartCode !== previousPartCode) {
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#f8fafc';
            headerRow.style.fontWeight = 'bold';

            const headerCell = document.createElement('td');
            headerCell.colSpan = COLUMNS.length;
            headerCell.style.padding = '12px';
            headerCell.style.borderTop = '2px solid #e2e8f0';
            headerCell.style.color = '#334155';

            const stats = groupStats[currentPartCode] || { total: 0, completed: 0 };

            headerCell.innerHTML = `
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div style="display:flex; align-items:center; gap:0.5rem;">
                        <span>📦 ${currentPartName}</span>
                        <span style="font-size:0.85em; color:#64748b; font-weight:normal;">(${currentPartCode})</span>
                        <span style="font-size:0.85em; color:#0369a1; font-weight:bold;">(${stats.completed.toLocaleString()}/${stats.total.toLocaleString()} Pc)</span>
                    </div>
                </div>
            `;

            headerRow.appendChild(headerCell);
            tableBody.appendChild(headerRow);
            previousPartCode = currentPartCode;
        }

        const tr = document.createElement('tr');

        COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = '';

            if (col.source === 'S') value = item.scrap[col.key] || '';
            else if (col.source === 'C') value = item.status;
            else if (col.source === 'T') value = item.technicianPhone; // Tech Source
            else if (col.source === 'P') value = item.person; // Person Source
            else value = item.fullRow[col.key] || '';

            // Checkbox Column
            if (col.key === 'checkbox') {
                const isSaved = !!item.status;
                const isLE = item.fullRow['Product'] === 'L&E';
                // Check if Store Code is missing
                const storeCode = item.fullRow['Store Code'] || '';
                const isMissingStore = !storeCode || String(storeCode).trim() === '';

                // User requested to allow selecting Poom/L&E, so we remove isLE from disabled check
                // BUT if Store Code is missing, we MUST disable it.
                const disabledAttr = (isSaved || isMissingStore) ? 'disabled' : '';
                td.innerHTML = `<input type="checkbox" class="row-checkbox" value="${startIndex + index}" onchange="handleCheckboxChange()" ${disabledAttr}>`;
                td.style.textAlign = 'center';
                tr.appendChild(td);
                return; // Skip other rendering for this col
            }

            // Apply Green Tab Style for specific columns if Saved
            // Apply Green Tab Style for specific columns if Saved (Status is not empty)
            const greenColumns = ['status', 'work order', 'Spare Part Code', 'Spare Part Name', 'old material code', 'qty', 'Serial Number', 'Store Code', 'Store Name'];

            // Check Store Code for Status Display
            const storeCode = item.fullRow['Store Code'] || '';
            const isMissingStore = !storeCode || String(storeCode).trim() === '';

            if (item.status && greenColumns.includes(col.key)) {
                const span = document.createElement('span');
                span.textContent = value;

                // Default Green (Claim Warranty / บันทึกแล้ว)
                let bgColor = '#dcfce7'; // Green-100
                let textColor = '#166534'; // Green-700

                if (item.status === 'Sworp') {
                    bgColor = '#e0e7ff'; // Indigo-100
                    textColor = '#3730a3'; // Indigo-800
                } else if (item.status === 'หมดประกัน') {
                    bgColor = '#ffedd5'; // Orange-100
                    textColor = '#9a3412'; // Orange-800
                } else if (item.status === 'ชำรุด') {
                    bgColor = '#fce7f3'; // Pink-100
                    textColor = '#9d174d'; // Pink-800
                }

                span.style.backgroundColor = bgColor;
                span.style.color = textColor;
                span.style.fontSize = '0.875rem'; // Standardize
                span.style.fontWeight = '600';
                span.style.padding = '0.25rem 0.5rem';
                span.style.borderRadius = '4px';
                span.style.display = 'inline-block'; // Ensure padding works
                td.appendChild(span);
            }
            // NEW: If Status is empty AND Store Code is missing (for Status Column or others?)
            else if (!item.status && isMissingStore && col.key === 'status') {
                const span = document.createElement('span');
                span.textContent = 'Wait Data';
                span.style.backgroundColor = '#f1f5f9'; // Slate-100
                span.style.color = '#ef4444'; // Red-500
                span.style.fontSize = '0.875rem'; // Standardize
                span.style.fontWeight = '600';
                span.style.padding = '0.25rem 0.5rem';
                span.style.borderRadius = '4px';
                span.style.display = 'inline-block';
                td.appendChild(span);
            }
            else {
                td.textContent = value;
                td.style.fontSize = '0.875rem'; // Standardize base text
            }

            if (col.key === 'Serial Number') {
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.fontWeight = '500';
                td.title = 'Click to edit Serial Number';
                td.onclick = () => openEditModal(item);
            }



            if (col.key === 'Store Code') {
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.fontWeight = '500';
                td.title = 'Click to view details';
                td.onclick = () => openStoreModal(item);
            }

            // Person (Claim Receiver) Click
            if (col.key === 'Person' || col.source === 'P') {
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.fontWeight = '500';
                td.title = 'Click to edit Person';
                td.onclick = () => openMainPersonModal(item);
            }

            // Mobile (Phone) Click
            if (col.key === 'Mobile' || col.source === 'T' || col.key === 'Phone') {
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.fontWeight = '500';
                td.title = 'Click to edit Mobile';
                td.onclick = () => openMainMobileModal(item);
                // If empty, show [Add]
                if (!value) {
                    td.innerText = '[Add]';
                    td.style.fontSize = '0.875rem'; // Standardize
                }
            }


            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    renderPagination();
}

function renderPagination() {
    const container = document.getElementById('paginationControls');
    container.innerHTML = '';

    const totalPages = Math.ceil(displayedData.length / ITEMS_PER_PAGE);
    if (totalPages === 0) return;

    const createButton = (text, onClick, disabled = false, isActive = false) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.onclick = onClick;
        btn.disabled = disabled;
        if (isActive) btn.classList.add('active');
        return btn;
    };

    // << First
    container.appendChild(createButton('<<', () => changePage(1), currentPage === 1));

    // < Prev
    container.appendChild(createButton('<', () => changePage(currentPage - 1), currentPage === 1));

    // Current Page Number
    container.appendChild(createButton(currentPage, () => { }, false, true));

    // > Next
    container.appendChild(createButton('>', () => changePage(currentPage + 1), currentPage === totalPages));

    // >> Last
    container.appendChild(createButton('>>', () => changePage(totalPages), currentPage === totalPages));
}

function changePage(newPage) {
    currentPage = newPage;
    renderTable();
    // Maybe scroll to top of table?
    document.querySelector('.table-container').scrollTop = 0;
}

function openEditModal(item) {
    editingItem = item;
    const currentSerial = item.fullRow['Serial Number'] || '';
    document.getElementById('editSerialInput').value = currentSerial;
    document.getElementById('editModal').style.display = 'flex';
}

function closeEditModal() {
    document.getElementById('editModal').style.display = 'none';
    editingItem = null;
}

function saveSerialNumber() {
    if (!editingItem) return;

    const newSerial = document.getElementById('editSerialInput').value.trim();

    // Update the item in memory
    // Note: Serial Number comes from WorkFilter ('W'), so we update fullRow
    editingItem.fullRow['Serial Number'] = newSerial;

    // Re-render to show changes
    renderTable();

    // Close modal
    closeEditModal();
}

function selectReceiver(btn, name) {
    // Reset buttons
    document.querySelectorAll('.receiver-btn').forEach(b => {
        b.style.backgroundColor = '#f1f5f9';
        b.style.color = '#64748b';
        b.style.border = '1px solid #e2e8f0';
    });

    // Highlight selected
    btn.style.backgroundColor = '#3b82f6';
    btn.style.color = 'white';
    btn.style.border = 'none';

    document.getElementById('store_receiver').value = name;
}

function openStoreModal(item) {
    editingItem = item;

    // Show/Hide Receiver Section based on Product
    const isLE = item.fullRow['Product'] === 'L&E';
    // Reset Receiver Selection
    document.getElementById('store_receiver').value = '';
    document.querySelectorAll('.receiver-btn').forEach(b => {
        b.style.backgroundColor = '#f1f5f9';
        b.style.color = '#64748b';
        b.style.border = '1px solid #e2e8f0';
    });

    if (isLE) {
        receiverSection.style.display = 'block';

        // Priority: Saved Receiver > Person Column > Default 'Poom'
        const savedReceiver = item.fullRow['Claim Receiver'] || item['Claim Receiver'];
        const personReceiver = item.person;
        const defaultReceiver = savedReceiver || personReceiver || 'Poom';

        document.getElementById('store_receiver').value = defaultReceiver;
        document.querySelectorAll('.receiver-btn').forEach(b => {
            // Reset
            b.style.backgroundColor = '#f1f5f9';
            b.style.color = '#64748b';
            b.style.border = '1px solid #e2e8f0';

            if (b.textContent === defaultReceiver) {
                b.style.backgroundColor = '#3b82f6';
                b.style.color = 'white';
                b.style.border = 'none';
            }
        });
    } else {
        // Non-L&E: Use Person if available
        receiverSection.style.display = 'block'; // Always show for clarity? Or keep hidden?
        // User asked "Warranty Claim Processing window... if no selection... take Person".
        // This implies the window is OPEN. So we should probably show it?
        // But original code hid it for non-L&E. 
        // Let's safe-play: If hidden, user can't select. 
        // But the REQUEST implies we can Save WITHOUT selection.
        // So we must ensure 'store_receiver' is populated even if hidden.

        // Let's keep original hiding logic BUT populate the hidden value.
        // Or maybe show it now? User didn't ask to show it.
        // Let's stick to: Populate the value.
        receiverSection.style.display = 'none';

        // Use Person as default for non-L&E too
        const defaultReceiver = item.person || '';
        document.getElementById('store_receiver').value = defaultReceiver;
    }

    // Populate fields
    // S Source
    document.getElementById('store_workOrder').value = item.scrap['work order'] || '';
    document.getElementById('store_partCode').value = item.scrap['Spare Part Code'] || '';
    document.getElementById('store_partName').value = item.scrap['Spare Part Name'] || '';
    document.getElementById('store_qty').value = item.scrap['qty'] || '';

    // W Source
    document.getElementById('store_serial').value = item.fullRow['Serial Number'] || '';
    document.getElementById('store_code').value = item.fullRow['Store Code'] || '';
    document.getElementById('store_name').value = item.fullRow['Store Name'] || '';

    // Load saved status or default
    const currentStatus = item.status || item.fullRow['ActionStatus'] || 'เคลมประกัน';
    setButtonActive(currentStatus);

    // Toggle Buttons based on saved status
    const btnSave = document.getElementById('btnSaveStore');
    const btnUpdate = document.getElementById('btnUpdateStore');
    const btnDelete = document.getElementById('btnDeleteStore');

    if (item.status) {
        btnSave.style.display = 'none';
        btnUpdate.style.display = 'inline-block';
        btnDelete.style.display = 'inline-block';
    } else {
        btnSave.style.display = 'inline-block';
        btnUpdate.style.display = 'none';
        btnDelete.style.display = 'none';
    }

    document.getElementById('storeModal').style.display = 'flex';
}

function selectAction(btn, value) {
    // Remove active class from all
    document.querySelectorAll('.status-btn').forEach(b => b.classList.remove('active'));
    // Add to clicked
    btn.classList.add('active');
    // Set hidden input
    document.getElementById('store_actionStatus').value = value;
}

function setButtonActive(value) {
    document.getElementById('store_actionStatus').value = value;
    document.querySelectorAll('.status-btn').forEach(btn => {
        if (btn.textContent.trim() === value) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

function closeStoreModal() {
    document.getElementById('storeModal').style.display = 'none';
    editingItem = null;
}

function saveStoreDetail() {
    if (!editingItem) return;

    // REMOVED 'Already Saved' Blocker to allow edits

    const serialInput = document.getElementById('store_serial').value.trim();
    const actionStatus = document.getElementById('store_actionStatus').value;

    // Validation: Check if empty or less than 8 characters (ONLY IF Claim Warranty)
    if (actionStatus === 'เคลมประกัน') {
        const isLE = editingItem.fullRow['Product'] === 'L&E';
        // If L&E, must be > 5 digits (so length >= 6). Else must be >= 8.
        const minLength = isLE ? 6 : 8;

        if (!serialInput || serialInput.length < minLength) {
            Swal.fire({
                icon: 'warning',
                title: 'ตรวจสอบข้อมูล',
                text: `กรุณาตรวจสอบ Serial Number ให้ถูกต้อง (L&E > 5 หลัก, อื่นๆ >= 8 หลัก)`,
            });
            return;
        }
    }

    // Only Serial Number is allowed to be edited
    editingItem.fullRow['Serial Number'] = serialInput;
    // Save Action Status
    editingItem.fullRow['ActionStatus'] = document.getElementById('store_actionStatus').value;

    renderTable();

    // Sync to Google Sheet (must be before closing modal which clears editingItem)
    sendDatatoGAS(editingItem);

    closeStoreModal();
}

async function deleteStoreDetail() {
    if (!editingItem) return;

    // Confirm Delete
    const result = await Swal.fire({
        title: 'ยืนยันการลบ?',
        text: "คุณต้องการลบข้อมูลนี้ออกจากระบบใช่หรือไม่?",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#ef4444',
        confirmButtonText: 'ลบข้อมูล (Delete)'
    });

    if (!result.isConfirmed) return;

    // Show Loading
    Swal.fire({ title: 'Deleting...', text: 'Updating Google Sheet...', didOpen: () => Swal.showLoading() });

    const payload = {
        'work order': editingItem.scrap['work order'],
        'Spare Part Code': editingItem.scrap['Spare Part Code'],
        'operation': 'delete' // Special flag for GAS
    };

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        Swal.fire('Deleted!', 'Record has been deleted.', 'success');

        // Clear local status
        editingItem.status = '';
        // editingItem.fullRow = {}; // REMOVED: Do not clear fullRow as it contains WorkFilter data

        if (editingItem.fullRow) {
            editingItem.fullRow['ActionStatus'] = '';
        }

        renderTable();
        closeStoreModal();

    } catch (error) {
        console.error(error);
        Swal.fire('Error', 'Failed to delete.', 'error');
    }
}

function sendDatatoGAS(item) {
    if (!GAS_API_URL) {
        alert("Please configure the Google Apps Script URL in the code to save to Sheets.");
        return;
    }

    // Get Current User (Fix ReferenceError)
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

    // Combine data sources
    const payload = {
        ...item.scrap,
        ...item.fullRow,
        'user': currentUser.IDRec || 'Unknown',
        'ActionStatus': item.fullRow['ActionStatus'] || 'เคลมประกัน',
        'Qty': document.getElementById('store_qty').value || item.scrap['qty'] || '',
        'Serial Number': document.getElementById('store_serial').value || item.fullRow['Serial Number'] || '',
        'Phone': item.technicianPhone || '', // Add Phone (Mobile)
        'รหัสช่าง': item.scrap['รหัสช่าง'] || '',
        'ชื่อช่าง': item.scrap['ชื่อช่าง'] || '',
        'Claim Receiver': (item.fullRow['ActionStatus'] === 'เคลมประกัน') ?
            (document.getElementById('store_receiver').value || item.person || '') : ''
    };

    // Remove 'Values' key if present to avoid confusion
    delete payload['Values'];

    // UI Feedback (SweetAlert2)
    const saveBtn = document.querySelector('#storeModal .btn-primary');
    const originalText = saveBtn.textContent;
    saveBtn.disabled = true;

    // Show Loading
    Swal.fire({
        title: 'Saving...',
        text: 'Please wait while we save your data.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    fetch(GAS_API_URL, {
        method: 'POST',
        mode: 'no-cors', // Important for GAS Web App
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload)
    })
        .then(() => {
            // Success
            Swal.fire({
                icon: 'success',
                title: 'Saved!',
                text: 'Data successfully saved to Google Sheet.',
                timer: 2000,
                showConfirmButton: false
            });

            // Optimistic UI Update
            item.status = payload['ActionStatus'];
            // Also update the 'user' field in fullRow so it shows up if we look at it
            item.fullRow['user'] = payload['user'];

            // -------------------------------------------------------------
            // START: Sync Recording Booking Data (Tab 2)
            // -------------------------------------------------------------
            // 1. Remove existing entry if exists (to avoid duplicates or update status)
            // Key is Work Order + Part Code, effectively unique.
            const targetKey = (payload['work order'] || '') + (payload['Spare Part Code'] || '');

            globalBookingData = globalBookingData.filter(row => {
                const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || '');
                return rowKey !== targetKey;
            });

            // 2. If status is "เคลมประกัน", add to globalBookingData
            if (payload['ActionStatus'] === 'เคลมประกัน') {
                // Construct Booking Row from Payload (Map S/W/C keys -> Booking Keys)
                // NOTE: payload has keys from scrap (S) and fullRow (W) directly mixed in
                const newBookingRow = {
                    'Work Order': payload['work order'],
                    'Spare Part Code': payload['Spare Part Code'],
                    'Spare Part Name': payload['Spare Part Name'],
                    'Old Material Code': payload['old material code'],
                    'Qty': payload['Qty'], // Use updated Qty
                    'Serial Number': payload['Serial Number'], // Use updated Serial
                    'Store Code': payload['Store Code'],
                    'Store Name': payload['Store Name'],
                    'รหัสช่าง': payload['รหัสช่าง'],
                    'ชื่อช่าง': payload['ชื่อช่าง'],
                    'Mobile': payload['Phone'], // Mapped Phone -> Mobile
                    'Plant': payload['plant'],
                    'Claim Receiver': payload['Claim Receiver'],
                    'Product': payload['Product'],
                    'Warranty Action': payload['ActionStatus'],
                    'Recorder': payload['user'],
                    'Timestamp': new Date().toISOString(), // Current Time
                    'Booking Slip': '', // New entry has no slip yet
                    'Booking Date': ''
                };
                globalBookingData.push(newBookingRow);
            }

            // 3. Re-render Tab 2
            populateBookingFilter(); // Refresh filters as new part might be added
            renderBookingTable();
            // -------------------------------------------------------------
            // END: Sync Recording Booking Data
            // -------------------------------------------------------------

            renderTable(); // Re-render DOM
        })
        .catch(error => {
            console.error('Error saving to GAS:', error);
            // Error
            Swal.fire({
                icon: 'error',
                title: 'Oops...',
                text: 'Failed to save to Google Sheet. Please try again.',
            });
        })
        .finally(() => {
            saveBtn.textContent = originalText;
            saveBtn.disabled = false;
        });
}

function toggleAllCheckboxes(source) {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = source.checked;
    });
    handleCheckboxChange();
}

function handleCheckboxChange() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    const bulkActions = document.getElementById('bulkActions');
    if (checkboxes.length > 0) {
        bulkActions.style.display = 'flex';
    } else {
        bulkActions.style.display = 'none';
    }
}

async function processBulkAction(actionName) {
    const savedKeys = new Set();
    fullData.forEach(item => {
        if (item.status === 'บันทึกแล้ว') {
            // Reconstruct key
            const uniqueKey = (item.scrap['work order'] || '') + (item.scrap['Spare Part Code'] || '');
            savedKeys.add(uniqueKey);
        }
    });

    const checkboxes = document.querySelectorAll('.row-checkbox:checked');

    // 1. Gather Items
    const selectedItems = [];
    for (const cb of checkboxes) {
        const globalIndex = parseInt(cb.value);
        const item = displayedData[globalIndex];

        if (item) {
            // Check if already saved
            if (item.status === 'บันทึกแล้ว') {
                Swal.fire({
                    icon: 'warning',
                    title: 'บางรายการถูกบันทึกแล้ว',
                    text: `รายการ ${item.scrap['work order']} ถูกบันทึกไปแล้ว ไม่สามารถแก้ไขได้`,
                });
                return; // Abort whole batch? Or continue? User usually wants safe abort.
            }

            // Validate Serial Number for EACH item? 
            const serial = item.fullRow['Serial Number'] || '';
            const isLE = item.fullRow['Product'] === 'L&E';
            const minLength = isLE ? 6 : 8;

            if (actionName === 'เคลมประกัน' && (!serial || serial.length < minLength)) {
                Swal.fire({
                    icon: 'warning',
                    title: 'ข้อมูลไม่ถูกต้อง',
                    text: `รายการ ${item.scrap['work order']}: Serial Number ไม่ถูกต้อง (L&E > 5 หลัก, อื่นๆ >= 8 หลัก)`,
                });
                return;
            }

            // NEW: Validation for Mobile (Phone) if Claim Warranty
            if (actionName === 'เคลมประกัน') {
                // Check if Mobile is empty
                const mobile = item.technicianPhone || '';
                if (!mobile || String(mobile).trim() === '') {
                    Swal.fire({
                        icon: 'warning',
                        title: 'ข้อมูลไม่ครบถ้วน',
                        text: `รายการ ${item.scrap['work order']}: กรุณาระบุเบอร์โทรศัพท์ (Mobile) ก่อนบันทึก`,
                    });
                    return;
                }
            }

            selectedItems.push(item);
        }
    }

    // 2. Confirm
    const result = await Swal.fire({
        title: 'ยืนยันการบันทึก?',
        text: `คุณต้องการบันทึก ${selectedItems.length} รายการ ด้วยสถานะ "${actionName}" ใช่หรือไม่?`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'ยืนยัน',
        cancelButtonText: 'ยกเลิก'
    });

    if (!result.isConfirmed) return;

    // 3. Process Batch (One by One)
    let successCount = 0;
    let failCount = 0;

    Swal.fire({
        title: 'กำลังบันทึก...',
        html: `กำลังประมวลผล 0 / ${selectedItems.length}`,
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];
        Swal.getHtmlContainer().textContent = `กำลังประมวลผล ${i + 1} / ${selectedItems.length}`;

        // Get Current User
        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

        const payload = {
            ...item.scrap,
            ...item.fullRow,
            'user': currentUser.IDRec || 'Unknown',
            'ActionStatus': actionName,
            'Qty': item.scrap['qty'] || '',
            'Serial Number': item.fullRow['Serial Number'] || '',
            'Phone': item.technicianPhone || '', // Add Phone (Mobile)
            'รหัสช่าง': item.scrap['รหัสช่าง'] || '',
            'ชื่อช่าง': item.scrap['ชื่อช่าง'] || '',
            'Claim Receiver': (actionName === 'เคลมประกัน') ? (item.person || '') : ''
        };
        delete payload['Values'];

        try {
            await fetch(GAS_API_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            successCount++;
            // Optimistic Update
            item.status = actionName;
            item.fullRow['ActionStatus'] = actionName;
        } catch (e) {
            console.error(e);
            failCount++;
        }
    }

    // 4. Finish

    // -------------------------------------------------------------
    // START: Sync Recording Booking Data (Tab 2) for Bulk Action
    // -------------------------------------------------------------
    // Since we processed items in a loop, we can just rebuild the data 
    // or (better) update incrementally. Let's do it by re-scanning `selectedItems`.
    // NOTE: `item` in selectedItems is the REFERENCE to `displayedData`.
    // It has been updated optimistically: item.status = actionName.

    selectedItems.forEach(item => {
        // Construct Payload-like object for mapping
        const itemPayload = {
            ...item.scrap,
            ...item.fullRow, // Contains updated Serial and ActionStatus
            'Qty': item.scrap['qty'], // Bulk doesn't change Qty
            'Phone': item.technicianPhone
        };

        const targetKey = (itemPayload['work order'] || '') + (itemPayload['Spare Part Code'] || '');

        // 1. Remove existing
        globalBookingData = globalBookingData.filter(row => {
            const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || '');
            return rowKey !== targetKey;
        });

        // 2. Add if claim
        if (actionName === 'เคลมประกัน') {
            const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
            const newBookingRow = {
                'Work Order': itemPayload['work order'],
                'Spare Part Code': itemPayload['Spare Part Code'],
                'Spare Part Name': itemPayload['Spare Part Name'],
                'Old Material Code': itemPayload['old material code'],
                'Qty': itemPayload['Qty'],
                'Serial Number': itemPayload['Serial Number'],
                'Store Code': itemPayload['Store Code'],
                'Store Name': itemPayload['Store Name'],
                'รหัสช่าง': itemPayload['รหัสช่าง'],
                'ชื่อช่าง': itemPayload['ชื่อช่าง'],
                'Mobile': itemPayload['Phone'],
                'Plant': itemPayload['plant'],
                'Claim Receiver': (actionName === 'เคลมประกัน') ? (item.person || '') : '',
                'Product': itemPayload['Product'],
                'Warranty Action': actionName,
                'Recorder': currentUser.IDRec || 'Unknown',
                'Timestamp': new Date().toISOString(),
                'Booking Slip': '',
                'Booking Date': ''
            };
            globalBookingData.push(newBookingRow);
        }
    });

    populateBookingFilter();
    renderBookingTable();

    // Sync Other Tabs (User Request: Rerender All Tabs)
    try {
        renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
        renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');

        populateSupplierFilter();
        renderSupplierTable();

        populateClaimSentFilter();
        renderClaimSentTable();

        renderHistoryTable();
    } catch (err) {
        console.warn('Error refreshing other tabs:', err);
    }

    // -------------------------------------------------------------
    // END: Sync Recording Booking Data
    // -------------------------------------------------------------

    renderTable();
    toggleAllCheckboxes({ checked: false }); // Deselect all

    Swal.fire({
        icon: successCount > 0 ? 'success' : 'error',
        title: 'เสร็จสิ้น',
        text: `บันทึกสำเร็จ: ${successCount} รายการ\nล้มเหลว: ${failCount} รายการ`,
    });
}

// --- Global Selection State for Booking Tab ---
let currentBookingDisplayedData = []; // Populated in renderBookingTable
let selectedBookingKeys = new Set(); // Stores unique keys (e.g., Work Order + Spare Part Code)

// Helper to generate unique key
function getBookingKey(row) {
    return (row['Work Order'] || '') + '_' + (row['Spare Part Code'] || '') + '_' + (row['Old Material Code'] || '');
}

function handleBookingPartFilterChange() {
    selectedBookingKeys.clear();
    renderBookingTable();
}

function toggleAllBookingCheckboxes(source) {
    const isChecked = source.checked;

    // Use currentDisplayedData to Select/Deselect All filtered items
    currentBookingDisplayedData.forEach(row => {
        // Skip Disabled Items (Already processed)
        const hasBookingSlip = row['Booking Slip'] && String(row['Booking Slip']).trim() !== '';
        const hasRecripte = row['Recripte'] && String(row['Recripte']).trim() !== '';
        // Note: Missing Store Code or other disable logic should be consistent? 
        // For Booking tab, logic is mainly Slip/Recripte.

        if (!hasBookingSlip && !hasRecripte) {
            const key = getBookingKey(row);
            if (isChecked) {
                selectedBookingKeys.add(key);
            } else {
                selectedBookingKeys.delete(key);
            }
        }
    });

    // Re-render to update checkboxes state
    renderBookingTable();
}

function handleBookingCheckboxChange(checkbox) {
    // If checkbox is passed (from onchange="handleBookingCheckboxChange(this)")
    if (checkbox) {
        const key = checkbox.value; // Value will be the unique key
        if (checkbox.checked) {
            selectedBookingKeys.add(key);
        } else {
            selectedBookingKeys.delete(key);
        }
    }

    // Update Button States (Active/Inactive)
    updateBookingActionButtons();
}

function updateBookingActionButtons() {
    const actionDiv = document.getElementById('bookingBulkActions');
    const sendNavaBtn = document.querySelector('button[onclick="sendToNavaNakorn()"]');
    const sendVibhavadiBtn = document.querySelector('button[onclick="sendToVibhavadi()"]');

    if (selectedBookingKeys.size > 0) {
        actionDiv.style.display = 'flex';
    } else {
        actionDiv.style.display = 'none';
        return;
    }

    // Validation Flags
    let hasPoom = false;
    let hasNonPoom = false;

    // Iterate Selected Items to check Poom/Not status
    // Need to lookup item from keys. Since keys are in Set, we iterate global or displayed?
    // Safer to iterate selectedBookingKeys and find items in user's view context or valid data.
    // Optimization check: iterating selectedBookingKeys against globalBookingData map? 
    // Let's just create a Lookup Map for performance if slow.
    // For now, iterate selected keys and find data (or store extra metadata in Set? No)
    // Actually, we can check `currentBookingDisplayedData` or `globalBookingData`.

    // PERF: Create Map for fast lookup
    const lookupMap = new Map();
    // Assuming globalBookingData covers everything
    globalBookingData.forEach(row => lookupMap.set(getBookingKey(row), row));

    selectedBookingKeys.forEach(key => {
        const row = lookupMap.get(key);
        if (row) {
            const receiver = (row['Claim Receiver'] || row.person || '').toLowerCase();
            const isPoom = receiver.includes('poom');
            if (isPoom) hasPoom = true;
            else hasNonPoom = true;
        }
    });

    // 2. Nava Nakorn Button Logic: Block if "Poom" is selected
    if (sendNavaBtn) {
        if (hasPoom) {
            sendNavaBtn.disabled = true;
            sendNavaBtn.style.opacity = '0.5';
            sendNavaBtn.style.cursor = 'not-allowed';
            sendNavaBtn.title = 'Items assigned to "Poom" cannot be sent to Nava Nakorn.';
        } else {
            sendNavaBtn.disabled = false;
            sendNavaBtn.style.opacity = '1';
            sendNavaBtn.style.cursor = 'pointer';
            sendNavaBtn.title = '';
        }
    }

    // 3. Vibhavadi Button Logic: Block if "Non-Poom" is selected (Strict Poom Only)
    if (sendVibhavadiBtn) {
        if (hasNonPoom) {
            sendVibhavadiBtn.disabled = true;
            sendVibhavadiBtn.style.opacity = '0.5';
            sendVibhavadiBtn.style.cursor = 'not-allowed';
            sendVibhavadiBtn.title = 'Only items assigned to "Poom" can be sent to Vibhavadi.';
        } else {
            sendVibhavadiBtn.disabled = false;
            sendVibhavadiBtn.style.opacity = '1';
            sendVibhavadiBtn.style.cursor = 'pointer';
            sendVibhavadiBtn.title = '';
        }
    }
}

// REMOVED DUPLICATE handleBookingCheckboxChange HERE


async function processBookingAction(destination, targetPlantCode) {
    if (selectedBookingKeys.size === 0) return;

    // 1. Gather Selected Items from Global Data
    const selectedItems = [];

    // Create Lookup Map (Work Order + Code -> Item)
    const lookupMap = new Map();
    // Assuming globalBookingData covers everything needed
    globalBookingData.forEach(row => lookupMap.set(getBookingKey(row), row));

    selectedBookingKeys.forEach(key => {
        const item = lookupMap.get(key);
        if (item) {
            selectedItems.push(item);
        }
    });

    if (selectedItems.length === 0) return;

    if (selectedItems.length === 0) return; // Should be covered by checkbox check, but safety.

    // --- VALIDATION: Block Poom -> Nava Nakorn ---
    if (destination.includes('นวนคร')) {
        const hasPoom = selectedItems.some(item => {
            const receiver = (item['Claim Receiver'] || item.person || '').toLowerCase();
            return receiver.includes('poom');
        });

        if (hasPoom) {
            Swal.fire({
                icon: 'warning',
                title: 'ไม่อนุญาต',
                text: 'รายการที่มี Receiver เป็น "Poom" ไม่สามารถส่งไป "นวนคร" ได้',
                confirmButtonText: 'ตกลง'
            });
            return;
        }
    }

    // 3. Prompt for Booking Slip
    let plantCenterCode = targetPlantCode || '';
    // Fallback if not passed (legacy)
    if (!plantCenterCode) {
        if (destination.includes('นวนคร')) plantCenterCode = '0301';
        if (destination.includes('วิภาวดี')) plantCenterCode = '0326';
    }

    const { value: bookingSlip, isConfirmed } = await Swal.fire({
        title: 'กรอกเลขใบจองรถ',
        html: `คุณกำลังจะส่ง ${selectedItems.length} รายการ ไปยัง <b>${destination}</b><br>
               (รหัส: <span style="color: var(--primary-color); font-weight: bold;">${plantCenterCode}</span>)`,
        input: 'text',
        inputPlaceholder: 'เลขใบจองรถ',
        showCancelButton: true,
        confirmButtonText: 'บันทึก',
        cancelButtonText: 'ยกเลิก',
        preConfirm: (value) => {
            if (!value) {
                Swal.showValidationMessage('กรุณากรอกเลขใบจองรถ!');
            }
            return value;
        }
    });

    if (!isConfirmed) return;

    // 4. Process Updates
    let successCount = 0;
    let failCount = 0;
    const bookingDate = new Date(); // Timestamp

    const totalToProcess = selectedItems.length;

    Swal.fire({
        title: 'กำลังบันทึก...',
        html: `กำลังประมวลผล 0 / ${totalToProcess}`,
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    // 4.1 Process Selected Items (Update)
    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];
        Swal.getHtmlContainer().textContent = `กำลังส่งข้อมูล ${i + 1} / ${totalToProcess}`;

        // Manually format date to DD/MM/YYYY
        const day = String(bookingDate.getDate()).padStart(2, '0');
        const month = String(bookingDate.getMonth() + 1).padStart(2, '0');
        const year = bookingDate.getFullYear();
        const formattedDate = `${day}/${month}/${year}`;

        const payload = {
            ...item,
            'Key': (item['Work Order'] || '') + (item['Spare Part Code'] || ''),
            'Booking Slip': bookingSlip,
            'Booking Date': formattedDate,
            'Plantcenter': plantCenterCode,
        };

        try {
            await fetch(GAS_API_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            // Optimistic Update
            item['Booking Slip'] = bookingSlip;
            item['Booking Date'] = payload['Booking Date'];
            item['Plantcenter'] = plantCenterCode;
            successCount++;

        } catch (e) {
            console.error(e);
            failCount++;
        }
    }



    renderBookingTable();
    selectedBookingKeys.clear();
    renderBookingTable();

    Swal.fire({
        icon: successCount > 0 ? 'success' : 'error',
        title: 'เสร็จสิ้น',
        text: `อัปเดตเรียบร้อย: ${successCount} รายการ (ล้มเหลว: ${failCount})`,
    });
}

async function sendToNavaNakorn() {
    await processBookingAction('ส่งคลังนวนคร', '0301');
}

async function sendToVibhavadi() {
    await processBookingAction('ส่งคลังวิภาวดี', '0326');
}

// --- Booking Receiver Filter Logic ---
let bookingReceiverOptions = new Set(); // Stores selected receiver names

function toggleReceiverDropdown() {
    const dropdown = document.getElementById('bookingReceiverDropdown');
    const isHidden = dropdown.style.display === 'none';
    dropdown.style.display = isHidden ? 'block' : 'none';
}

function handleReceiverCheckboxChange() {
    selectedBookingKeys.clear(); // Clear selection on filter change to avoid ghost selections

    bookingReceiverOptions.clear();
    const checkboxes = document.querySelectorAll('.receiver-checkbox:checked');
    checkboxes.forEach(cb => {
        bookingReceiverOptions.add(cb.value);
    });

    // Update Label
    const label = document.getElementById('receiverFilterLabel');
    if (bookingReceiverOptions.size === 0) {
        label.textContent = 'All Receivers';
    } else {
        label.textContent = `${bookingReceiverOptions.size} Selected`;
    }

    renderBookingTable();
}

function populateBookingReceiverFilter() {
    const dropdown = document.getElementById('bookingReceiverDropdown');
    dropdown.innerHTML = ''; // Clear existing

    if (!globalBookingData || globalBookingData.length === 0) return;

    // Extract Unique Receivers (Claim Receiver or person)
    // Note: globalBookingData has 'person' mapped from partsMap in loadTableData
    const receivers = new Set();
    globalBookingData.forEach(row => {
        const p = row['person'] || row['Claim Receiver'];
        if (p) receivers.add(String(p).trim());
    });

    const sortedReceivers = Array.from(receivers).sort();

    if (sortedReceivers.length === 0) {
        dropdown.innerHTML = '<div style="padding:0.5rem; color:#64748b; font-size:0.875rem;">No receivers found</div>';
        return;
    }

    sortedReceivers.forEach(receiver => {
        const div = document.createElement('div');
        div.style.padding = '0.25rem 0';
        div.style.display = 'flex';
        div.style.alignItems = 'center';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = receiver;
        checkbox.className = 'receiver-checkbox';
        checkbox.style.marginRight = '0.5rem';
        // Maintain selection if re-populating (though re-populating wipes selections usually)
        if (bookingReceiverOptions.has(receiver)) {
            checkbox.checked = true;
        }
        checkbox.onchange = handleReceiverCheckboxChange;

        const label = document.createElement('label');
        label.textContent = receiver;
        label.style.fontSize = '0.875rem';
        label.style.cursor = 'pointer';
        label.onclick = () => checkbox.click(); // Label clicks checkbox

        div.appendChild(checkbox);
        div.appendChild(label);
        dropdown.appendChild(div);
    });
}

// Close Dropdown when clicking outside
window.addEventListener('click', function (e) {
    const dropdown = document.getElementById('bookingReceiverDropdown');
    const button = document.querySelector('button[onclick="toggleReceiverDropdown()"]');
    if (dropdown && button && !dropdown.contains(e.target) && !button.contains(e.target)) {
        dropdown.style.display = 'none';
    }
});

function populateBookingFilter() {
    const filterSelect = document.getElementById('bookingPartFilter');
    // Keep "All" option
    filterSelect.innerHTML = '<option value="">All Spare Parts</option>';

    if (!globalBookingData || globalBookingData.length === 0) return;

    // Extract Unique Spare Part Names
    const parts = new Set();
    globalBookingData.forEach(row => {
        const name = row['Spare Part Name'];
        if (name) parts.add(name);
    });

    // API sort
    const sortedParts = Array.from(parts).sort();

    sortedParts.forEach(part => {
        const option = document.createElement('option');
        option.value = part;
        option.textContent = part;
        filterSelect.appendChild(option);
    });

    populateBookingReceiverFilter(); // Call Receiver Filter Population
}

function renderBookingTable() {
    const tableHeader = document.getElementById('bookingTableHeader');
    const tableBody = document.getElementById('bookingTableBody');
    const filterValue = document.getElementById('bookingPartFilter').value;

    if (!globalBookingData || globalBookingData.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="' + BOOKING_COLUMNS.length + '" style="text-align:center; padding: 2rem;">No data found in Warranty Sheet.</td></tr>';
        return;
    }

    // FILTERING
    let filteredData = globalBookingData;

    // 1. Plant Filter (Show ONLY current user's plant for outgoing bookings)
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const userPlant = currentUser['Plant'] ? currentUser['Plant'].toString().trim().padStart(4, '0') : null;

    if (userPlant) {
        filteredData = filteredData.filter(row => {
            const rowPlant = row['Plant'] ? row['Plant'].toString().trim().padStart(4, '0') : '';
            return rowPlant === userPlant;
        });
    }

    // 2. Status Filter: SHOW 'Local Stock' OR 'In Transit'
    // Exclude: 'Claim Sent' OR 'Received'
    filteredData = filteredData.filter(item => {
        const hasClaimSup = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
        // Received at Central (Recripte exists) -> Exclude (Moves to Supplier Tab)
        const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';

        if (hasClaimSup) return false;
        if (hasRecripte) return false;

        // Remaining:
        // 1. No Booking Slip = Local Stock
        // 2. Has Booking Slip + No Recripte = In Transit
        return true;
    });

    // 2. Part Name Filter
    if (filterValue) {
        filteredData = filteredData.filter(row => row['Spare Part Name'] === filterValue);
    }

    // 3. Multi-Select Receiver Filter
    if (bookingReceiverOptions.size > 0) {
        filteredData = filteredData.filter(row => {
            const val = row['Claim Receiver'] || row['person'] || '';
            return bookingReceiverOptions.has(String(val).trim());
        });
    }

    // Sort Data (Strict Match with processBookingAction)
    const sortedData = [...filteredData].reverse().sort((a, b) => {
        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });

    // Populate Global Displayed Data (Filtered & Sorted)
    currentBookingDisplayedData = sortedData;

    // Calculate Group Totals
    const groupTotals = sortedData.reduce((acc, item) => {
        const name = item['Spare Part Name'] || 'Unknown';
        const qty = parseFloat(item['Qty']) || 0;
        acc[name] = (acc[name] || 0) + qty;
        return acc;
    }, {});

    // Headers
    tableHeader.innerHTML = '';
    BOOKING_COLUMNS.forEach(col => {
        const th = document.createElement('th');
        if (col.header.includes('<')) {
            th.innerHTML = col.header;
        } else {
            th.textContent = col.header;
        }
        tableHeader.appendChild(th);
    });

    // Check Select All State (Data Driven)
    const selectAllCheckbox = document.getElementById('selectAllBooking');
    if (selectAllCheckbox) {
        let allEnabledSelected = true;
        let anySelected = false;
        let hasOneEnabled = false;

        currentBookingDisplayedData.forEach(row => {
            const hasBookingSlip = row['Booking Slip'] && String(row['Booking Slip']).trim() !== '';
            const hasRecripte = row['Recripte'] && String(row['Recripte']).trim() !== '';
            if (!hasBookingSlip && !hasRecripte) {
                hasOneEnabled = true;
                const key = getBookingKey(row);
                if (!selectedBookingKeys.has(key)) {
                    allEnabledSelected = false;
                } else {
                    anySelected = true;
                }
            }
        });

        if (hasOneEnabled && allEnabledSelected) {
            selectAllCheckbox.checked = true;
            selectAllCheckbox.indeterminate = false;
        } else if (anySelected) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = true;
        } else {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        }
    }

    // Pagination Logic
    const startIndex = (currentBookingPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = sortedData.slice(startIndex, endIndex);

    // Rows
    tableBody.innerHTML = '';

    // Track previous item for Grouping
    let previousPartName = null;
    if (currentBookingPage > 1 && sortedData[startIndex - 1]) {
        previousPartName = sortedData[startIndex - 1]['Spare Part Name'];
    }

    pageData.forEach((row, index) => {
        const currentPartName = row['Spare Part Name'] || 'Unknown';
        const currentPartCode = row['Spare Part Code'] || '';

        // Insert Group Header if Name Changed
        if (currentPartName !== previousPartName) {
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#f8fafc';
            headerRow.style.fontWeight = 'bold';

            const headerCell = document.createElement('td');
            headerCell.colSpan = BOOKING_COLUMNS.length;
            headerCell.style.padding = '12px';
            headerCell.style.borderTop = '2px solid #e2e8f0';
            headerCell.style.color = '#334155';

            const total = groupTotals[currentPartName] || 0;
            headerCell.innerHTML = `
                 <div style="display: flex; justify-content: space-between; align-items: center;">
                     <div style="display:flex; align-items:center; gap:0.5rem;">
                         <span>📦 ${currentPartName}</span>
                         <span style="font-size:0.85em; color:#64748b; font-weight:normal;">(${currentPartCode})</span>
                         <span style="font-size:0.85em; color:#0369a1; font-weight:bold;">(${total.toLocaleString()} Pc)</span>
                     </div>
                 </div>
             `;
            headerRow.appendChild(headerCell);
            tableBody.appendChild(headerRow);
            previousPartName = currentPartName;
        }

        const tr = document.createElement('tr');
        BOOKING_COLUMNS.forEach(col => {
            const td = document.createElement('td');

            if (col.key === 'checkbox') {
                // Adjust index for value to match global index in sortedData
                const globalIndex = startIndex + index;
                // Disable if Booking Slip exists OR Recripte exists
                const hasBookingSlip = row['Booking Slip'] && String(row['Booking Slip']).trim() !== '';
                const hasRecripte = row['Recripte'] && String(row['Recripte']).trim() !== '';
                const disabledAttr = (hasBookingSlip || hasRecripte) ? 'disabled' : '';

                // Key for selection
                const key = getBookingKey(row);
                const isChecked = selectedBookingKeys.has(key) ? 'checked' : '';

                // We encode Key in the value for easy lookup, or use handleBookingCheckboxChange(this) to access value
                td.innerHTML = `<input type="checkbox" class="booking-checkbox" value="${key}" onchange="handleBookingCheckboxChange(this)" ${disabledAttr} ${isChecked}>`;
                td.style.textAlign = 'center';
                tr.appendChild(td);
                return;
            }

            let value = row[col.key] || '';

            if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(row);
                tr.appendChild(td);
                return;
            }

            // Format Timestamp
            if (col.key === 'Timestamp' && value) {
                try {
                    const date = new Date(value);
                    if (!isNaN(date)) {
                        value = date.toLocaleString('th-TH');
                    }
                } catch (e) { }
            }

            // Format Booking Date (Always Strip Time)
            if (col.key === 'Booking Date' && value) {
                // If value looks like ISO or DateTime string
                let s = String(value);
                if (s.indexOf('T') > -1) s = s.split('T')[0];
                if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                value = s;
            }

            // Format Mobile Number
            if (col.key === 'Mobile') {
                if (value) {
                    value = formatPhoneNumber(value);
                } else {
                    // Make clickable if empty
                    td.style.cursor = 'pointer';
                    td.style.color = '#3b82f6'; // Blue-500
                    td.innerText = '[Add]';
                    td.style.fontSize = '0.8em';
                    td.onclick = () => openMobileModal(row);
                }
            }

            // Booking Slip Click Action
            if (col.key === 'Booking Slip' && value && String(value).trim() !== '') {
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.fontWeight = '500';
                td.title = 'Click to Edit/Delete';
                td.onclick = () => openBookingSlipModal(row);
            }

            // Color Badge for Claim Receiver
            if (col.key === 'Claim Receiver') {
                td.style.cursor = 'pointer';
                td.title = 'Click to Edit Receiver';
                td.onclick = () => openClaimReceiverModal(row);

                if (value) {
                    const badge = document.createElement('span');
                    badge.textContent = value;
                    badge.style.padding = '0.25rem 0.5rem';
                    badge.style.borderRadius = '4px';
                    badge.style.fontSize = '0.75rem';
                    badge.style.fontWeight = '500';

                    // Assign colors based on value (Case Insensitive)
                    const valLower = String(value).toLowerCase().trim();
                    if (valLower === 'mai') {
                        badge.style.backgroundColor = '#dbeafe'; // Blue 100
                        badge.style.color = '#1e40af'; // Blue 800
                    } else if (valLower === 'mon') {
                        badge.style.backgroundColor = '#dcfce7'; // Green 100
                        badge.style.color = '#166534'; // Green 800
                    } else if (valLower === 'poom') {
                        badge.style.backgroundColor = '#fae8ff'; // Purple 100
                        badge.style.color = '#6b21a8'; // Purple 800
                    } else {
                        badge.style.backgroundColor = '#f1f5f9'; // Slate 100
                        badge.style.color = '#475569'; // Slate 600
                    }

                    td.appendChild(badge);
                } else {
                    td.textContent = '-';
                    td.style.color = '#ccc';
                }
                tr.appendChild(td);
                return;
            }

            td.textContent = value;
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Render Pagination Controls
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentBookingPage > totalPages) currentBookingPage = 1;
    renderBookingPagination(totalPages);

    // Call update action buttons to ensure correct state after render
    updateBookingActionButtons();
}

function renderBookingPagination(totalPages) {
    const container = document.getElementById('bookingPaginationControls');
    container.innerHTML = '';

    if (totalPages <= 1) return;

    const createButton = (text, onClick, disabled = false, isActive = false) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.onclick = onClick;
        btn.disabled = disabled;
        if (isActive) btn.classList.add('active');
        return btn;
    };

    // << First
    container.appendChild(createButton('<<', () => changeBookingPage(1), currentBookingPage === 1));

    // < Prev
    container.appendChild(createButton('<', () => changeBookingPage(currentBookingPage - 1), currentBookingPage === 1));

    // Current Page Number
    container.appendChild(createButton(currentBookingPage, () => { }, false, true));

    // > Next
    container.appendChild(createButton('>', () => changeBookingPage(currentBookingPage + 1), currentBookingPage === totalPages));

    // >> Last
    container.appendChild(createButton('>>', () => changeBookingPage(totalPages), currentBookingPage === totalPages));
}

function changeBookingPage(newPage) {
    currentBookingPage = newPage;
    renderBookingTable();
    // Scroll to top of table container
    const tableContainer = document.querySelector('#tab-content-booking .table-container');
    if (tableContainer) tableContainer.scrollTop = 0;
}
async function openBookingSlipModal(item) {
    const currentSlip = item['Booking Slip'] || '';

    // SweetAlert with 3 Buttons: Edit, Delete, Cancel
    const result = await Swal.fire({
        title: 'จัดการใบจองรถ',
        html: `
    <p>Current Slip: <b>${currentSlip || '-'}</b></p>
    <p>เลือกการทำงาน:</p>
`,
        showDenyButton: true,
        showCancelButton: true,
        confirmButtonText: 'แก้ไข (Edit)',
        denyButtonText: 'ลบข้อมูล (Delete)',
        cancelButtonText: 'ยกเลิก',
        denyButtonColor: '#d33'
    });

    if (result.isConfirmed) {
        // EDIT
        const { value: newSlip } = await Swal.fire({
            title: 'แก้ไขเลขใบจองรถ',
            input: 'text',
            inputValue: currentSlip,
            showCancelButton: true,
            confirmButtonText: 'บันทึก',
            inputValidator: (value) => {
                if (!value) return 'กรุณากรอกข้อมูล!';
            }
        });

        if (newSlip) {
            const d = new Date();
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear();
            const formattedDate = `${day}/${month}/${year}`;

            updateBookingSlip(item, newSlip, formattedDate);
        }

    } else if (result.isDenied) {
        // DELETE
        const confirmDelete = await Swal.fire({
            title: 'ยืนยันการลบ?',
            text: "ข้อมูลใบจองรถและวันที่จองจะถูกลบออก",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'ลบข้อมูล',
            confirmButtonColor: '#d33'
        });

        if (confirmDelete.isConfirmed) {
            updateBookingSlip(item, '', '', ''); // Clear Slip + Date + Plantcenter
        }
    }
}

async function updateBookingSlip(item, newSlip, newDate, newPlantCenter = null) {
    Swal.fire({
        title: 'Updating...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    const payload = {
        ...item,
        'Booking Slip': newSlip,
        'Booking Date': newDate,
        'Plantcenter': newPlantCenter !== null ? newPlantCenter : item['Plantcenter']
    };

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        // Optimistic UI Update
        item['Booking Slip'] = newSlip;
        item['Booking Date'] = newDate;
        if (newPlantCenter !== null) {
            item['Plantcenter'] = newPlantCenter;
        }

        renderBookingTable(); // Re-render DOM

        Swal.fire({
            icon: 'success',
            title: 'Updated',
            timer: 1500,
            showConfirmButton: false
        });

    } catch (error) {
        console.error('Error updating receiver:', error);
        Swal.fire('Error', 'Failed to update', 'error');
    }
}


function formatPhoneNumber(phone) {
    if (!phone) return '';

    // Remove all non-numeric characters
    let cleaned = ('' + phone).replace(/\D/g, '');

    // Case 1: 10 digits starting with 0 (e.g., 0812345678)
    if (cleaned.length === 10 && cleaned.startsWith('0')) {
        return cleaned.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }

    // Case 2: 9 digits (missing leading 0) (e.g., 812345678)
    if (cleaned.length === 9) {
        cleaned = '0' + cleaned;
        return cleaned.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }

    // Default: return original
    return phone;
}

async function openClaimReceiverModal(item) {
    const currentReceiver = item['Claim Receiver'] || '';

    // Create HTML for buttons
    const receiverOptions = ['Mai', 'Mon', 'Poom', 'Not'];
    let buttonsHtml = '<div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">';

    receiverOptions.forEach(opt => {
        let btnColor = '#f1f5f9';
        let txtColor = '#64748b';
        if (opt === 'Mai') { btnColor = '#dbeafe'; txtColor = '#1e40af'; }
        if (opt === 'Mon') { btnColor = '#dcfce7'; txtColor = '#166534'; }
        if (opt === 'Poom') { btnColor = '#fae8ff'; txtColor = '#6b21a8'; }
        if (opt === 'Not') { btnColor = '#ffedd5'; txtColor = '#c2410c'; } // Orange/Amber
        if (opt === 'Clear') { btnColor = '#fee2e2'; txtColor = '#991b1b'; } // Redish for clear

        // Check if currently selected (highlight border?)
        const isSelected = (currentReceiver.toLowerCase() === opt.toLowerCase()) || (opt === 'Clear' && !currentReceiver);
        const borderStyle = isSelected ? '2px solid var(--primary-color)' : '1px solid transparent';

        buttonsHtml += `
            <button type="button" class="swal2-confirm swal2-styled"
                style="background-color: ${btnColor}; color: ${txtColor}; border: ${borderStyle}; margin: 5px; flex: 1 0 40%;"
                onclick="window.selectedClaimReceiver = '${opt === 'Clear' ? '' : opt}'; 
                         document.querySelectorAll('.receiver-opt-btn').forEach(b => b.style.border='1px solid transparent'); 
                         this.style.border='2px solid var(--primary-color)';"
                class="receiver-opt-btn">
                ${opt === 'Clear' ? 'ล้างค่า (Clear)' : opt}
            </button>
        `;
    });
    buttonsHtml += '</div>';

    // Reset global selection
    window.selectedClaimReceiver = currentReceiver;

    const result = await Swal.fire({
        title: 'เลือกผู้รับงาน (Claim Receiver)',
        html: buttonsHtml,
        showCancelButton: true,
        confirmButtonText: 'บันทึก (Save)',
        cancelButtonText: 'ยกเลิก',
        preConfirm: () => {
            return window.selectedClaimReceiver;
        }
    });

    if (result.isConfirmed) {
        const newReceiver = result.value;
        if (newReceiver !== currentReceiver) {
            updateClaimReceiver(item, newReceiver);
        }
    }
}

async function updateClaimReceiver(item, newReceiver) {
    Swal.fire({
        title: 'Updating...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    const payload = {
        ...item,
        'Claim Receiver': newReceiver,
    };

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        // Optimistic UI Update
        item['Claim Receiver'] = newReceiver;
        renderBookingTable(); // Re-render DOM

        Swal.fire({
            icon: 'success',
            title: 'Updated',
            timer: 1500,
            showConfirmButton: false
        });

    } catch (error) {
        console.error('Error updating receiver:', error);
        Swal.fire('Error', 'Failed to update receiver', 'error');
    }
}

async function openMobileModal(item) {
    const { value: newMobile } = await Swal.fire({
        title: 'เพิ่มเบอร์โทรศัพท์ (Add Mobile)',
        input: 'tel',
        inputLabel: 'Phone Number',
        inputPlaceholder: 'Enter phone number...',
        showCancelButton: true,
        confirmButtonText: 'บันทึก (Save)',
        cancelButtonText: 'ยกเลิก',
        inputValidator: (value) => {
            if (!value) {
                return 'You need to write something!';
            }
        }
    });

    if (newMobile) {
        updateMobile(item, newMobile);
    }
}

async function updateMobile(item, newMobile) {
    Swal.fire({
        title: 'Updating...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    const payload = {
        ...item,
        'Mobile': newMobile,
    };

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        // Optimistic UI Update
        item['Mobile'] = newMobile;
        renderBookingTable(); // Re-render DOM to show new number

        Swal.fire({
            icon: 'success',
            title: 'Updated',
            timer: 1500,
            showConfirmButton: false
        });

    } catch (error) {
        console.error('Error updating mobile:', error);
        Swal.fire('Error', 'Failed to update mobile', 'error');
    }
}

// ==========================================
// NEW: Main Table Person/Mobile Modal Logic
// ==========================================

async function openMainPersonModal(item) {
    const currentReceiver = item.person || '';

    // Create HTML for buttons (reuse style from openClaimReceiverModal)
    const receiverOptions = ['Mai', 'Mon', 'Poom', 'Not'];
    let buttonsHtml = '<div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">';

    receiverOptions.forEach(opt => {
        let btnColor = '#f1f5f9';
        let txtColor = '#64748b';
        if (opt === 'Mai') { btnColor = '#dbeafe'; txtColor = '#1e40af'; }
        if (opt === 'Mon') { btnColor = '#dcfce7'; txtColor = '#166534'; }
        if (opt === 'Poom') { btnColor = '#fae8ff'; txtColor = '#6b21a8'; }
        if (opt === 'Not') { btnColor = '#ffedd5'; txtColor = '#c2410c'; }

        const isSelected = (currentReceiver.toLowerCase() === opt.toLowerCase());
        const borderStyle = isSelected ? '2px solid var(--primary-color)' : '1px solid transparent';

        buttonsHtml += `
            <button type="button" class="swal2-confirm swal2-styled"
                style="background-color: ${btnColor}; color: ${txtColor}; border: ${borderStyle}; margin: 5px; flex: 1 0 40%;"
                onclick="window.selectedMainPerson = '${opt}'; 
                         document.querySelectorAll('.receiver-opt-btn').forEach(b => b.style.border='1px solid transparent'); 
                         this.style.border='2px solid var(--primary-color)';"
                class="receiver-opt-btn">
                ${opt}
            </button>
        `;
    });
    buttonsHtml += '</div>';

    window.selectedMainPerson = currentReceiver; // Default

    const result = await Swal.fire({
        title: 'Select Person (Claim Receiver)',
        html: buttonsHtml,
        showCancelButton: true,
        confirmButtonText: 'Save',
        cancelButtonText: 'Cancel',
        preConfirm: () => {
            return window.selectedMainPerson;
        }
    });

    if (result.isConfirmed) {
        const newPerson = result.value;
        if (newPerson && newPerson !== currentReceiver) {
            saveMainPerson(item, newPerson);
        }
    }
}

async function saveMainPerson(item, newPerson) {
    Swal.fire({
        title: 'Saving...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

    // Construct Payload similar to sendDatatoGAS but minimal/targeted
    // Note: We need to send enough data for GAS to identify the row and update it.
    // GAS script usually needs 'work order' and 'Spare Part Code' (or unique key) + 'operation' or 'user' etc.
    // Re-using the Standard Payload structure ensures best compatibility.

    const payload = {
        ...item.scrap,
        ...item.fullRow,
        'user': currentUser.IDRec || 'Unknown',
        'ActionStatus': item.status || item.fullRow['ActionStatus'] || 'เคลมประกัน',
        'Claim Receiver': newPerson, // The value to update
        'Person': newPerson // Ensure mapping
    };
    delete payload['Values'];

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        // Optimistic Update
        item.person = newPerson;

        // Also update fullRow if it maps there
        if (item.fullRow) item.fullRow['Claim Receiver'] = newPerson;

        renderTable();

        Swal.fire({
            icon: 'success',
            title: 'Saved',
            timer: 1500,
            showConfirmButton: false
        });

    } catch (e) {
        console.error(e);
        Swal.fire('Error', 'Failed to save person', 'error');
    }
}

async function openMainMobileModal(item) {
    const currentMobile = item.technicianPhone || '';

    const { value: newMobile } = await Swal.fire({
        title: 'Edit Mobile Number',
        input: 'tel',
        inputValue: currentMobile,
        inputLabel: 'Phone Number',
        inputPlaceholder: 'Enter phone number...',
        showCancelButton: true,
        confirmButtonText: 'Save',
        cancelButtonText: 'Cancel'
    });

    if (newMobile !== undefined && newMobile !== currentMobile) {
        saveMainMobile(item, newMobile);
    }
}

async function saveMainMobile(item, newMobile) {
    Swal.fire({
        title: 'Saving...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

    const payload = {
        ...item.scrap,
        ...item.fullRow,
        'user': currentUser.IDRec || 'Unknown',
        'ActionStatus': item.status || item.fullRow['ActionStatus'] || 'เคลมประกัน',
        'Phone': newMobile
    };
    delete payload['Values'];

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        // Optimistic Update
        item.technicianPhone = newMobile;
        renderTable();

        Swal.fire({
            icon: 'success',
            title: 'Saved',
            timer: 1500,
            showConfirmButton: false
        });

    } catch (e) {
        console.error(e);
        Swal.fire('Error', 'Failed to save mobile', 'error');
    }
}

initAuth();

document.getElementById('passInput').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        handleLogin();
    }
});

function renderDeckView(targetPlantCode, containerId, tabKey) {
    const deckContainer = document.getElementById(containerId);
    deckContainer.innerHTML = '';

    const uniqueMap = new Map();

    const sortedData = [...globalBookingData].sort((a, b) => {
        const parseDate = (dateStr) => {
            if (!dateStr) return 0;
            const [d, m, y] = dateStr.split('/');
            return new Date(y, m - 1, d).getTime();
        };
        return parseDate(b['Booking Date']) - parseDate(a['Booking Date']);
    });

    sortedData.forEach(item => {
        const slip = item['Booking Slip'];
        const pcCode = String(item['Plantcenter'] || '').trim();

        // Filter by Plant Code (Normalize: remove leading zeros)
        const normalize = (s) => s.replace(/^0+/, '');
        if (!slip) return;
        // DEBUG: Check filtering
        // if (targetPlantCode === '0326' && normalize(pcCode) === '326') console.log('Found Vibhavadi Item:', slip);

        if (normalize(pcCode) !== normalize(targetPlantCode)) return;

        const date = item['Booking Date'] || '-';
        const plant = item['Plant'] || '-';
        const receiver = item['Claim Receiver'] || '-';

        // Key: Slip + Receiver (with fallback to Person)
        const receiverVal = item['Claim Receiver'] || item.person || 'Unknown';
        const key = slip + '|' + receiverVal;
        const hasRecripte = item['Recripte'] && item['Recripte'].trim() !== '';
        const recripteInc = hasRecripte ? 1 : 0;

        if (!uniqueMap.has(key)) {
            // Store the effective receiver in the item for display purposes if needed?
            // Actually, better not mutate 'item' here unpredictably. 
            // We can just use the logic when rendering.
            uniqueMap.set(key, {
                ...item,
                _effectiveReceiver: receiverVal,
                _count: 1,
                _recripteCount: recripteInc
            });
        } else {
            // Increment count
            const existing = uniqueMap.get(key);
            existing._count = (existing._count || 1) + 1;
            existing._recripteCount = (existing._recripteCount || 0) + recripteInc;
        }
    });

    if (uniqueMap.size === 0) {
        deckContainer.innerHTML = '<p style="text-align: center; color: var(--text-secondary); grid-column: 1/-1;">No booking records found for this plant.</p>';
        return;
    }

    uniqueMap.forEach((item, key) => {
        const slip = item['Booking Slip'];
        const date = item['Booking Date'];
        const plantCode = item['Plant'] || '';
        const plantName = PLANT_MAPPING[plantCode] || plantCode; // Resolve Name
        // Use the effective receiver we stored, or re-derive
        const receiver = item._effectiveReceiver || item['Claim Receiver'] || item.person || 'Unknown';
        const plantCenter = item['Plantcenter'] || '-';
        const count = item._count || 1;
        const recCount = item._recripteCount || 0;

        // Status Logic
        let statusText = 'ระหว่างขนส่ง';
        let statusBg = '#fef3c7'; // yellow-100
        let statusColor = '#92400e'; // yellow-800

        if (recCount === 0) {
            statusText = 'ระหว่างขนส่ง';
            statusBg = '#fef3c7';
            statusColor = '#92400e';
        } else if (recCount < count) {
            statusText = 'ระหว่างตรวจรับ';
            statusBg = '#e0f2fe'; // blue-100
            statusColor = '#0369a1'; // blue-700
        } else if (recCount === count) {
            statusText = 'เสร็จสิ้น';
            statusBg = '#dcfce7'; // green-100
            statusColor = '#166534'; // green-800
        }

        const card = document.createElement('div');
        card.className = 'deck-card';
        card.style.cursor = 'pointer';
        card.setAttribute('data-key', key);
        // Pass 'this' to access the element
        card.onclick = function () { toggleDetailView(this, tabKey, slip, receiver); };
        card.innerHTML = `
            <div class="card-header" style="display: block;">
                <!-- Row 1: Slip and Date -->
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
                    <div class="card-title">📄 ${slip}</div>
                    <div class="card-badge" style="background-color: #f3f4f6; color: #374151;">📅 ${date}</div>
                </div>
                <!-- Row 2: Status and Count -->
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <span class="card-badge" style="background-color: ${statusBg}; color: ${statusColor}; font-weight:bold; font-size: 0.75rem;">${statusText}</span>
                    <div class="card-badge" style="background-color: #dbeafe; color: #1e40af;">รับแล้ว ${recCount}/${count} item</div>
                </div>
            </div>
            <!-- Date moved to header
            <div class="card-detail-row">
                <span class="card-label">Booking Date:</span>
                <span class="card-value">${date}</span>
            </div>
            -->
            <!-- Receiver Removed -->
            <!-- 
            <div class="card-detail-row">
                <span class="card-label">Receiver:</span>
                <span class="card-value">${receiver}</span>
            </div>
            -->
            <div style="display: flex; justify-content: space-between; align-items: center; font-weight: 700; font-size: 1rem; color: #1e293b;">
                <div>${plantName}</div>
                <div style="color: #475569;">${receiver}</div>
            </div>
        `;
        deckContainer.appendChild(card);
    });
}

function backToDeck(tabKey) {
    // Legacy function, might not be needed if toggle logic works, but keep for safety reset
    const deckContainer = document.getElementById(tabKey === 'navanakorn' ? 'navaNakornDeck' : 'vibhavadiDeck');
    const cards = deckContainer.getElementsByClassName('deck-card');

    // Show all cards
    Array.from(cards).forEach(c => c.style.display = 'flex');

    document.getElementById(tabKey + 'TableWrapper').style.display = 'none';
}

function toggleDetailView(cardElement, tabKey, slip, targetReceiver) {
    const deckWrapper = document.getElementById(tabKey + 'DeckWrapper');
    const tableWrapper = document.getElementById(tabKey + 'TableWrapper');
    const deckContainer = cardElement.parentElement; // Need this to find siblings

    // Check if currently active
    const isActive = cardElement.classList.contains('active-card');

    if (isActive) {
        // COLLAPSE
        const cards = deckContainer.getElementsByClassName('deck-card');
        Array.from(cards).forEach(c => c.style.display = 'flex');

        cardElement.classList.remove('active-card');
        cardElement.style.border = '1px solid var(--border-color)';

        deckWrapper.style.height = '100%';
        deckWrapper.style.flex = '';

        tableWrapper.style.display = 'none';

        // Clear Context
        currentDetailContext = null;
    } else {
        // EXPAND
        const cards = deckContainer.getElementsByClassName('deck-card');
        Array.from(cards).forEach(c => {
            if (c !== cardElement) {
                c.style.display = 'none';
            }
        });

        cardElement.classList.add('active-card');

        deckWrapper.style.height = 'auto';
        deckWrapper.style.flex = '0 0 auto';

        tableWrapper.style.display = 'flex';
        tableWrapper.style.height = 'auto';
        tableWrapper.style.flex = '1';

        // Set Context and Render
        currentDetailContext = { tabKey, slip, targetReceiver };
        renderTopLevelDetailTable(tabKey, slip, targetReceiver);
    }
}
// Global context for detail view
let currentDetailContext = null;

function renderTopLevelDetailTable(tabKey, slip, targetReceiver) {
    const tableWrapper = document.getElementById(tabKey + 'TableWrapper');
    const header = document.getElementById(tabKey + 'ReviewHeader');
    const thead = document.getElementById(tabKey + 'TableHeader');
    const tbody = document.getElementById(tabKey + 'TableBody');

    // Filter Data
    const detailData = globalBookingData.filter(item => {
        const itemReceiver = item['Claim Receiver'] || item.person || 'Unknown';
        return item['Booking Slip'] === slip && itemReceiver === targetReceiver;
    });

    // Update Header with Save Button
    // Use unique class or relative finding. Let's just use class 'bulk-save-btn' and hide it.
    // We will control it via traversing up from checkbox.
    header.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
            <span>Booking Details: ${slip} (${targetReceiver})</span>
            <button class="btn btn-primary bulk-save-btn" style="display: none;" onclick="saveBulkReviewItems(this)">
                Save Selected
            </button>
        </div>
    `;

    // Render Table Header
    thead.innerHTML = '';
    // thead is actually the TR element (id=...TableHeader)
    BOOKING_COLUMNS.forEach(col => {
        const th = document.createElement('th');
        if (col.key === 'checkbox') {
            th.innerHTML = '<input type="checkbox" class="select-all-review" onclick="toggleAllReviewCheckboxes(this)">';
        } else {
            th.innerHTML = col.header;
        }
        thead.appendChild(th);
    });

    // Render Table Body
    tbody.innerHTML = '';

    if (detailData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="' + BOOKING_COLUMNS.length + '" style="text-align:center;">No data found.</td></tr>';
        return;
    }

    detailData.forEach((item, index) => {
        const tr = document.createElement('tr');
        BOOKING_COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = item[col.key] || '';

            if (col.key === 'checkbox') {
                // VALUE should be index in detailData to identify it
                // We add a data attribute to the checkbox to find the row or data later? 
                // Or just keep using value as index. But we need to know WHICH detailData it refers to?
                // Ideally, we should store the context ON the DOM element (the table or wrapper).
                // For now, let's assume one active context. But the ID conflict was the main issue.

                // Disable if Booking Slip exists OR Recripte exists
                // Note: item['Booking Slip'] should match 'slip' from context, so implicitly it exists if we are here?
                // Wait, 'slip' from context is the grouping key. 
                // BUT the requirement is: "Tab Booking Slip Navanakhon ... table that shows when card is pressed -> Checkbox not work when Recripte has data"
                // AND previously "Booking Slip" (Car Reservation) has data.
                // In this view, they are grouped BY Booking Slip, so they ALL have Booking Slip data.
                // So technically, ALL checkboxes should be disabled if we follow "Booking Slip" rule strictly?
                // The user said: "Tab ... Table that shows when card is pressed ... give Checkbox NOT work when column 'Booking Slip' has data".
                // Since this IS the table for a specific Booking Slip, yes, they all have it.
                // However, maybe the user wants to disable it ONLY if it's already "Processed" or something further?
                // But let's follow the user's literal request.
                // AND now "Recripte" has data.

                const hasBookingSlip = item['Booking Slip'] && String(item['Booking Slip']).trim() !== '';
                const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
                const disabledAttr = (hasBookingSlip || hasRecripte) ? 'disabled' : '';

                td.innerHTML = `<input type="checkbox" class="review-checkbox" value="${index}" onchange="handleReviewCheckboxChange(this)" ${disabledAttr}>`;
                td.style.textAlign = 'center';
            } else if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
            } else if (col.key === 'Work Order' || col.key === 'Serial Number') {
                // Make Work Order and Serial Number clickable
                td.textContent = value;
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.textDecoration = 'underline';
                td.onclick = function (e) {
                    e.stopPropagation();
                    openWorkOrderModal(item);
                };
            } else {
                if (col.key === 'Timestamp' && value) {
                    const date = new Date(value);
                    if (!isNaN(date.getTime())) {
                        value = date.toLocaleString();
                    }
                }
                if (col.key === 'Booking Date' && value) {
                    let s = String(value);
                    if (s.indexOf('T') > -1) s = s.split('T')[0];
                    if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                    value = s;
                }
                td.textContent = value;
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

function toggleAllReviewCheckboxes(source) {
    // Find the table that contains this checkbox
    const table = source.closest('table');
    if (!table) return;
    const checkboxes = table.querySelectorAll('.review-checkbox');
    checkboxes.forEach(cb => cb.checked = source.checked);

    // Trigger update on one of them to update button
    if (checkboxes.length > 0) {
        handleReviewCheckboxChange(checkboxes[0]);
    }
}

function handleReviewCheckboxChange(source) {
    // Traverse up to find the wrapper, then the header, then the button?
    // Structure: Wrapper -> [Header(contains button), Table]
    // source is in Table -> tbody -> tr -> td
    const wrapper = source.closest('.detail-content-wrapper') || source.closest('[id$="TableWrapper"]');
    if (!wrapper) return;

    const table = wrapper.querySelector('table');
    const checkboxes = table.querySelectorAll('.review-checkbox:checked');
    const saveBtn = wrapper.querySelector('.bulk-save-btn'); // Find by class in this wrapper

    if (saveBtn) {
        saveBtn.style.display = checkboxes.length > 0 ? 'block' : 'none';
        saveBtn.textContent = `Save Selected (${checkboxes.length})`;
    }
}

async function saveBulkReviewItems(btnElement) {
    if (!currentDetailContext) return;
    // WARNING: currentDetailContext might be stale if user switched tabs and opened another detail WITHOUT closing the first correctly?
    // toggleDetailView cleans up, so it should be fine. 
    // But to be safe, we can re-derive context or trust 'currentDetailContext' belongs to the visible view.

    const { slip, targetReceiver, tabKey } = currentDetailContext;

    // Re-derive data to match indices
    const detailData = globalBookingData.filter(item => {
        const itemReceiver = item['Claim Receiver'] || item.person || 'Unknown';
        return item['Booking Slip'] === slip && itemReceiver === targetReceiver;
    });

    // Find checkboxes in the same container as the button
    const wrapper = btnElement.closest('.detail-content-wrapper') || btnElement.closest('[id$="TableWrapper"]');
    if (!wrapper) return;

    const checkboxes = wrapper.querySelectorAll('.review-checkbox:checked');
    const selectedItems = [];
    checkboxes.forEach(cb => {
        const idx = parseInt(cb.value);
        if (detailData[idx]) selectedItems.push(detailData[idx]);
    });

    if (selectedItems.length === 0) return;

    // Get Current User
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const recripteName = currentUser.name || currentUser.IDRec || 'Unknown';
    const recripteDate = new Date();
    const recripteDateStr = recripteDate.toLocaleString('en-GB');

    Swal.fire({
        title: 'Saving...',
        html: `Updating <b>1</b> / ${selectedItems.length} items...`,
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];

        // Update Progress
        if (Swal.getHtmlContainer()) {
            Swal.getHtmlContainer().innerHTML = `Updating <b>${i + 1}</b> / ${selectedItems.length} items...`;
        }

        // Payload
        const payload = {
            ...item,
            'Recripte': recripteName,
            'RecripteDate': recripteDateStr,
            'user': recripteName // Optional implicit update
        };

        try {
            await fetch(GAS_API_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            // Update Local
            item['Recripte'] = recripteName;
            item['RecripteDate'] = recripteDateStr;
            successCount++;
        } catch (e) {
            console.error(e);
            failCount++;
        }
    }

    Swal.fire({
        icon: 'success',
        title: 'Updated',
        text: `Success: ${successCount}, Failed: ${failCount}`,
        timer: 1500,
        showConfirmButton: false
    });

    // Re-render table and DECK to update counts
    renderTopLevelDetailTable(tabKey, slip, targetReceiver);

    // Refresh Deck Card Badge
    // We need to re-render the deck or update valid badges.
    // Easiest is to re-render deck view for that plant.
    if (tabKey === 'navanakorn') renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
    if (tabKey === 'vibhavadi') renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
}

function openWorkOrderModal(item) {
    editingItem = item;

    document.getElementById('woDetail_WorkOrder').value = item['Work Order'] || '';
    document.getElementById('woDetail_Qty').value = item['Qty'] || '';
    document.getElementById('woDetail_PartCode').value = item['Spare Part Code'] || '';
    document.getElementById('woDetail_PartName').value = item['Spare Part Name'] || '';
    document.getElementById('woDetail_StoreCode').value = item['Store Code'] || '';
    document.getElementById('woDetail_StoreName').value = item['Store Name'] || '';
    document.getElementById('woDetail_SerialNumber').value = item['Serial Number'] || '';
    const currentAction = item['Warranty Action'] || '';
    document.getElementById('woDetail_ActionValue').value = currentAction;

    // Set active button
    const buttons = document.querySelectorAll('#workOrderModal .status-btn');
    buttons.forEach(btn => {
        btn.classList.remove('active');
        // Simple check: if button's onclick contains the action string
        if (btn.getAttribute('onclick').includes(`'${currentAction}'`)) {
            btn.classList.add('active');
        }
    });
    document.getElementById('woDetail_Note').value = item['Note'] || ''; // Assuming 'Note' exists

    document.getElementById('workOrderModal').style.display = 'flex';
}

function closeWorkOrderModal() {
    document.getElementById('workOrderModal').style.display = 'none';
    editingItem = null;
}

async function saveWorkOrderDetail() {
    if (!editingItem) return;

    const newSerial = document.getElementById('woDetail_SerialNumber').value;
    const newNote = document.getElementById('woDetail_Note').value;
    const newAction = document.getElementById('woDetail_ActionValue').value;
    const newQty = document.getElementById('woDetail_Qty').value;

    // Get Current User
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const recripteName = currentUser.name || currentUser.IDRec || 'Unknown';
    const recripteDate = new Date();

    // Prepare Payload
    const payload = {
        ...editingItem, // Include original data
        'Serial Number': newSerial,
        'Qty': newQty,
        'Note': newNote,
        'Warranty Action': newAction,
        'Recripte': recripteName,
        'RecripteDate': recripteDate.toLocaleString('en-GB'), // DD/MM/YYYY, HH:mm:ss
        'user': recripteName,
        'ActionStatus': newAction
    };

    // UI Feedback: Saving...
    Swal.fire({
        title: 'Saving...',
        text: 'Updating Work Order details...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    try {
        await fetch(GAS_API_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        // Success: Update Local Data
        editingItem['Serial Number'] = newSerial;
        editingItem['Note'] = newNote;
        editingItem['Warranty Action'] = newAction;
        editingItem['Qty'] = newQty;
        // Also update the new fields locally
        editingItem['Recripte'] = recripteName;
        editingItem['RecripteDate'] = payload['RecripteDate'];

        // Refresh Table
        if (currentDetailContext) {
            renderTopLevelDetailTable(currentDetailContext.tabKey, currentDetailContext.slip, currentDetailContext.targetReceiver);

            // Refresh Deck View as well to update badge counts
            if (currentDetailContext.tabKey === 'navanakorn') {
                renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
            } else if (currentDetailContext.tabKey === 'vibhavadi') {
                renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
            }
        }

        closeWorkOrderModal();

        Swal.fire({
            icon: 'success',
            title: 'Saved!',
            text: 'Work Order updated successfully.',
            timer: 1500,
            showConfirmButton: false
        });

    } catch (error) {
        console.error('Error saving Work Order:', error);
        Swal.fire({
            icon: 'error',
            title: 'Error',
            text: 'Failed to save changes to Google Sheet.',
        });
    }
}

function selectWorkOrderAction(element, action) {
    const container = element.parentElement;
    const buttons = container.querySelectorAll('.status-btn');
    buttons.forEach(btn => btn.classList.remove('active'));
    element.classList.add('active');
    document.getElementById('woDetail_ActionValue').value = action;
}









function populateSupplierProductFilter(data) {
    const dropdown = document.getElementById('supplierProductDropdown');
    if (!dropdown) return;

    // Unique Products from provided data (or globalBookingData if not provided)
    // We want to show ALL products available in the Supplier view (filtered by Recripte)
    let sourceData = data;
    if (!sourceData) {
        sourceData = globalBookingData.filter(item => item['Recripte'] && item['Recripte'].trim() !== '');
    }

    const uniqueProducts = new Set();
    sourceData.forEach(item => {
        if (item['Product']) uniqueProducts.add(item['Product']);
    });

    // Also include any currently selected options to avoid them disappearing if filtered out?
    // Actually, usually filters show what's available. 
    // But if we use 'data' which is ALREADY filtered, options disappear.
    // So we should use the BASE supplier data (globalBookingData filtered by Recripte).

    // Sorted Array
    const sortedProducts = Array.from(uniqueProducts).sort();

    // Preserve Current Selection State
    // If options change, we might need to cleanup selection? Or keep it?
    // Let's keep it.

    dropdown.innerHTML = '';

    sortedProducts.forEach(prod => {
        const div = document.createElement('div');
        div.style.padding = '0.25rem 0.5rem';
        div.style.display = 'flex';
        div.style.alignItems = 'center';
        div.style.gap = '0.5rem';
        div.style.cursor = 'pointer';
        div.className = 'hover:bg-gray-100'; // Tailwind syntax if available, else style
        div.onmouseover = function () { this.style.backgroundColor = '#f3f4f6'; };
        div.onmouseout = function () { this.style.backgroundColor = 'transparent'; };

        const isChecked = supplierProductOptions.has(prod) ? 'checked' : '';

        div.innerHTML = `
            <input type="checkbox" class="supplier-product-checkbox" value="${prod}" ${isChecked} onchange="handleSupplierProductCheckboxChange()">
            <span onclick="this.previousElementSibling.click();" style="flex:1;">${prod}</span>
        `;
        dropdown.appendChild(div);
    });
}

function toggleSupplierProductDropdown() {
    const dropdown = document.getElementById('supplierProductDropdown');
    if (dropdown) {
        dropdown.style.display = dropdown.style.display === 'none' ? 'block' : 'none';
    }
}

function handleSupplierProductCheckboxChange() {
    supplierProductOptions.clear();
    const checkboxes = document.querySelectorAll('.supplier-product-checkbox:checked');
    checkboxes.forEach(cb => {
        supplierProductOptions.add(cb.value);
    });

    // Update Label
    const label = document.getElementById('supplierProductFilterLabel');
    if (label) {
        if (supplierProductOptions.size === 0) {
            label.textContent = 'All Products';
        } else {
            label.textContent = `Selected (${supplierProductOptions.size})`;
        }
    }

    renderSupplierTable();
}

// Close Dropdown on Click Outside
window.addEventListener('click', function (e) {
    const dropdown = document.getElementById('supplierProductDropdown');
    const button = document.querySelector('button[onclick="toggleSupplierProductDropdown()"]');
    if (dropdown && button && !dropdown.contains(e.target) && !button.contains(e.target)) {
        dropdown.style.display = 'none';
    }
});



function populateSupplierFilter() {
    console.log("--- populateSupplierFilter called ---");
    const filterSelect = document.getElementById('supplierPartFilter');
    if (!filterSelect) return;

    // Filter Data: Items relevant for Supplier Tab (Has Recripte)
    let supplierItems = globalBookingData.filter(item => {
        return item['Recripte'] && String(item['Recripte']).trim() !== '';
    });
    console.log(`[DEBUG] populateSupplierFilter: ${supplierItems.length} supplier items found.`);
    if (supplierItems.length > 0) console.log('[DEBUG] First supplier item Product:', supplierItems[0]['Product']);

    // 1. Populate Part Filter
    const parts = new Set();
    supplierItems.forEach(item => {
        const name = item['Spare Part Name'];
        if (name) parts.add(name);
    });

    const sortedParts = Array.from(parts).sort();

    // Save current selection
    const currentPart = filterSelect.value;

    filterSelect.innerHTML = '<option value="">All Spare Parts</option>';
    sortedParts.forEach(part => {
        const option = document.createElement('option');
        option.value = part;
        option.textContent = part;
        if (part === currentPart) option.selected = true;
        filterSelect.appendChild(option);
    });



    // 2. Populate Product Filter
    populateSupplierProductFilter(supplierItems);
}

function renderSupplierTable() {
    console.log("--- renderSupplierTable called ---");
    const tableHeader = document.getElementById('supplierTableHeader');
    const tableBody = document.getElementById('supplierTableBody');

    // Filter Data: Items with 'Recripte'
    const filterSelect = document.getElementById('supplierPartFilter');
    const filterValue = filterSelect ? filterSelect.value : '';

    let supplierData = globalBookingData.filter(item => {
        return item['Recripte'] && item['Recripte'].trim() !== '';
    });

    // Populate Product Filter Options (using base supplier data)
    // Only populate if not already populated? To prevent re-rendering dropdown while open.
    // Ideally we check if dropdown is empty.
    const dropdown = document.getElementById('supplierProductDropdown');
    if (dropdown && dropdown.children.length === 0) {
        populateSupplierProductFilter(supplierData);
    }
    // Alternatively, update counts? For now just ensure it's populated.
    // If we want it to update dynamically based on other filters, we would pass filtered data.
    // User asked for "add filter", usually implies standard filter behavior.

    // Apply Product Filter
    if (supplierProductOptions.size > 0) {
        supplierData = supplierData.filter(item => supplierProductOptions.has(item['Product']));
    }



    if (filterValue) {
        supplierData = supplierData.filter(item => item['Spare Part Name'] === filterValue);
    }


    // Sort Data: Primary = Spare Part Name (Asc), Secondary = Booking Date (Desc)
    const sortedData = [...supplierData].reverse().sort((a, b) => {
        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });

    // Calculate Group Totals (across ALL sorted data to show correct totals despite pagination)
    const groupTotals = sortedData.reduce((acc, item) => {
        const name = item['Spare Part Name'] || 'Unknown';
        const qty = parseFloat(item['Qty']) || 0;
        acc[name] = (acc[name] || 0) + qty;
        return acc;
    }, {});


    // Headers
    tableHeader.innerHTML = '';
    SUPPLIER_COLUMNS.forEach(col => {
        const th = document.createElement('th');
        if (col.header.includes('<')) {
            th.innerHTML = col.header;
        } else {
            th.textContent = col.header;
        }
        tableHeader.appendChild(th);
    });

    // Pagination Logic
    const startIndex = (currentSupplierPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = sortedData.slice(startIndex, endIndex);

    tableBody.innerHTML = '';

    // Track previous item for Grouping (handle pagination boundary)
    let previousPartName = null;
    if (currentSupplierPage > 1 && sortedData[startIndex - 1]) {
        previousPartName = sortedData[startIndex - 1]['Spare Part Name'];
    }

    pageData.forEach((item, index) => {
        const currentPartName = item['Spare Part Name'] || 'Unknown';
        const currentPartCode = item['Spare Part Code'] || '';

        // Insert Group Header if Name Changed
        if (currentPartName !== previousPartName) {
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#f8fafc';
            headerRow.style.fontWeight = 'bold';

            const headerCell = document.createElement('td');
            headerCell.colSpan = SUPPLIER_COLUMNS.length;
            headerCell.style.padding = '12px';
            headerCell.style.borderTop = '2px solid #e2e8f0';
            headerCell.style.color = '#334155';

            const total = groupTotals[currentPartName] || 0;
            headerCell.innerHTML = `
                 <div style="display: flex; justify-content: space-between; align-items: center;">
                     <div style="display:flex; align-items:center; gap:0.5rem;">
                         <span>📦 ${currentPartName}</span>
                         <span style="font-size:0.85em; color:#64748b; font-weight:normal;">(${currentPartCode})</span>
                         <span style="font-size:0.85em; color:#0369a1; font-weight:bold;">(${total.toLocaleString()} Pc)</span>
                     </div>
                 </div>
             `;
            headerRow.appendChild(headerCell);
            tableBody.appendChild(headerRow);
            previousPartName = currentPartName;
        }

        const tr = document.createElement('tr');
        SUPPLIER_COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = item[col.key] || '';

            if (col.key === 'checkbox') {
                const globalIndex = startIndex + index;
                // Disable if ClaimSup has data
                const isClaimed = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
                const disabledAttr = isClaimed ? 'disabled' : '';

                td.innerHTML = `<input type="checkbox" class="supplier-checkbox" value="${globalIndex}" ${disabledAttr} onchange="handleSupplierCheckboxChange()">`;
                td.style.textAlign = 'center';
                if (isClaimed) {
                    td.style.opacity = '0.5'; // Visual cue
                }
            } else if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
            } else if (col.key === 'Work Order' || col.key === 'Serial Number') {
                td.textContent = value;
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.textDecoration = 'underline';
                td.onclick = function (e) {
                    e.stopPropagation();
                    openWorkOrderModal(item);
                };
            } else {
                // Date formatting
                if ((col.key === 'Timestamp' || col.key === 'RecripteDate' || col.key === 'Claim Date') && value) {
                    const date = new Date(value);
                    // Simple check if it looks like a date object string or ISO
                    if (!isNaN(date.getTime()) && String(value).length > 10) {
                        // Check length to avoid "DD/MM/YYYY" being parsed as valid date in some browsers if not careful, 
                        // but usually "DD/MM/YYYY" parsing is quirky. 
                        // If we saved it as DD/MM/YYYY string, we should just show it?
                        // Our previous code saved 'RecripteDate' as LocaleString.
                        // Let's just trust textContent mostly unless it's ISO.
                        if (String(value).includes('T')) {
                            value = date.toLocaleString('en-GB');
                        }
                    }
                }
                // If value is manually formatted DD/MM/YYYY, leave it.
                if (col.key === 'Booking Date' && value) {
                    let s = String(value);
                    if (s.indexOf('T') > -1) s = s.split('T')[0];
                    if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                    value = s;
                }
                td.textContent = value;
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Render Pagination Controls
    console.log(`Rendering Supplier Pagination. Total Data: ${sortedData.length}, Items/Page: ${ITEMS_PER_PAGE}`);
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    console.log(`Calculated Total Pages: ${totalPages}, Current Page: ${currentSupplierPage}`);

    if (currentSupplierPage > totalPages) currentSupplierPage = 1;
    renderSupplierPagination(totalPages);
}

function handleSupplierCheckboxChange() {
    const checkboxes = document.querySelectorAll('.supplier-checkbox:checked');
    const bulkActions = document.getElementById('supplierBulkActions');
    const selectAll = document.getElementById('selectAllSupplier');

    // Update Select All Header
    const allCheckboxes = document.querySelectorAll('.supplier-checkbox');
    if (allCheckboxes.length > 0) {
        if (checkboxes.length === allCheckboxes.length) {
            selectAll.checked = true;
            selectAll.indeterminate = false;
        } else if (checkboxes.length > 0) {
            selectAll.checked = false;
            selectAll.indeterminate = true;
        } else {
            selectAll.checked = false;
            selectAll.indeterminate = false;
        }
    }

    if (checkboxes.length > 0) {
        bulkActions.style.display = 'flex';
    } else {
        bulkActions.style.display = 'none';
    }
}

function toggleAllSupplierCheckboxes(source) {
    const checkboxes = document.querySelectorAll('.supplier-checkbox');
    checkboxes.forEach(cb => cb.checked = source.checked);
    handleSupplierCheckboxChange();
}

function populateSupplierFilter() {
    const filterSelect = document.getElementById('supplierPartFilter');
    filterSelect.innerHTML = '<option value="">All Spare Parts</option>';

    // Filter Data: Items with 'Recripte'
    const supplierData = globalBookingData.filter(item => {
        return item['Recripte'] && item['Recripte'].trim() !== '';
    });

    populateSupplierProductFilter(supplierData); // Populate Product Filter when Part Filter is populated

    const uniqueParts = [...new Set(supplierData.map(item => item['Spare Part Name']).filter(Boolean))].sort();

    uniqueParts.forEach(part => {
        const option = document.createElement('option');
        option.value = part;
        option.textContent = part;
        filterSelect.appendChild(option);
    });
}


async function sendClaim() {
    const checkboxes = document.querySelectorAll('.supplier-checkbox:checked');
    if (checkboxes.length === 0) return;

    // Filter logic to get data based on search/sort
    // We must mirror logic in renderSupplierTable to map indices correctly
    const filterSelect = document.getElementById('supplierPartFilter');
    const filterValue = filterSelect ? filterSelect.value : '';

    // Filter Data: Items with 'Recripte'
    let allowedData = globalBookingData.filter(item => {
        return item['Recripte'] && item['Recripte'].trim() !== '';
    });

    // Product Filter (Must match renderSupplierTable)
    if (supplierProductOptions.size > 0) {
        allowedData = allowedData.filter(item => supplierProductOptions.has(item['Product']));
    }

    // UI Filter (Spare Part)
    if (filterValue) {
        allowedData = allowedData.filter(item => item['Spare Part Name'] === filterValue);
    }

    // Sort Data: Primary = Spare Part Name (Asc), Secondary = Booking Date (Desc) - matches render
    const sortedData = [...allowedData].reverse().sort((a, b) => {
        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });

    // Gather Items
    const selectedItems = [];
    checkboxes.forEach(cb => {
        const idx = parseInt(cb.value);
        if (sortedData[idx]) {
            selectedItems.push(sortedData[idx]);
        }
    });

    if (selectedItems.length === 0) return;

    // Auto-fill logic (No Prompt)
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const claimSup = currentUser.IDRec || 'Unknown';

    // Date = Today (DD/MM/YYYY)
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const formattedDate = `${day}/${month}/${year}`;

    // Confirm Action
    const result = await Swal.fire({
        title: 'Confirm Send Claim?',
        html: `Update <b>${selectedItems.length}</b> items?<br><br>Date: <b>${formattedDate}</b><br>Supplier: <b>${claimSup}</b>`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'Yes, Send',
        cancelButtonText: 'Cancel'
    });

    if (!result.isConfirmed) return;

    Swal.fire({
        title: 'Saving...',
        html: `Updating ${selectedItems.length} items...`,
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];

        const payload = {
            ...item,
            'Claim Date': formattedDate,
            'ClaimSup': claimSup,
            'user': JSON.parse(localStorage.getItem('currentUser') || '{}').IDRec || 'Unknown'
        };

        try {
            await fetch(GAS_API_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            // Update Local
            item['Claim Date'] = formattedDate;
            item['ClaimSup'] = claimSup;
            successCount++;
        } catch (e) {
            console.error(e);
            failCount++;
        }
    }

    renderSupplierTable();
    handleSupplierCheckboxChange(); // Updated UI state

    Swal.fire({
        icon: 'success',
        title: 'Completed',
        text: `Success: ${successCount}, Failed: ${failCount}`,
    });

}

function renderSupplierPagination(totalPages) {
    console.log(`renderSupplierPagination called with totalPages: ${totalPages}`);
    const container = document.getElementById('supplierPaginationControls');
    console.log("Container found:", container);

    if (!container) {
        console.error("Pagination container #supplierPaginationControls NOT found!");
        return;
    }
    container.style.display = 'flex'; // Force display
    container.innerHTML = '';

    // if (totalPages <= 1) return; // Force show for user feedback

    const createButton = (text, onClick, disabled = false, isActive = false) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.onclick = onClick;
        btn.disabled = disabled;
        if (isActive) btn.classList.add('active');
        return btn;
    };

    // << First
    container.appendChild(createButton('<<', () => changeSupplierPage(1), currentSupplierPage === 1));

    // < Prev
    container.appendChild(createButton('<', () => changeSupplierPage(currentSupplierPage - 1), currentSupplierPage === 1));

    // Current Page Number
    container.appendChild(createButton(currentSupplierPage, () => { }, false, true));

    // > Next
    container.appendChild(createButton('>', () => changeSupplierPage(currentSupplierPage + 1), currentSupplierPage >= totalPages));

    // >> Last
    container.appendChild(createButton('>>', () => changeSupplierPage(totalPages), currentSupplierPage >= totalPages));
}

function changeSupplierPage(newPage) {
    if (newPage < 1) return;
    currentSupplierPage = newPage;
    renderSupplierTable(); // Re-render table with new page
}

// ==========================================================
// CLAIM SENT TAB LOGIC
// ==========================================================

function populateClaimSentFilter() {
    const filterSelect = document.getElementById('claimSentPartFilter');
    if (!filterSelect) return;

    // Filter Items: Must have 'Recripte' AND 'ClaimSup'
    const claimSentItems = globalBookingData.filter(item => {
        const hasRecripte = item['Recripte'] && item['Recripte'].trim() !== '';
        const hasClaimSup = item['ClaimSup'] && item['ClaimSup'].trim() !== '';
        return hasRecripte && hasClaimSup;
    });

    const parts = [...new Set(claimSentItems.map(item => item['Spare Part Name']).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));

    filterSelect.innerHTML = '<option value="">All Spare Parts</option>';
    parts.forEach(part => {
        const option = document.createElement('option');
        option.value = part;
        option.textContent = part;
        filterSelect.appendChild(option);
    });
}

function renderClaimSentTable() {
    console.log("--- renderClaimSentTable called ---");
    const tableHeader = document.getElementById('claimSentTableHeader');
    const tableBody = document.getElementById('claimSentTableBody');

    // Filter Data: 'Recripte' AND 'ClaimSup'
    const filterSelect = document.getElementById('claimSentPartFilter');
    const filterValue = filterSelect ? filterSelect.value : '';

    let allowedData = globalBookingData.filter(item => {
        const hasRecripte = item['Recripte'] && item['Recripte'].trim() !== '';
        const hasClaimSup = item['ClaimSup'] && item['ClaimSup'].trim() !== '';
        return hasRecripte && hasClaimSup;
    });

    if (filterValue) {
        allowedData = allowedData.filter(item => item['Spare Part Name'] === filterValue);
    }

    // Sort Data: Spare Part Name (Asc)
    const sortedData = [...allowedData].reverse().sort((a, b) => {
        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });

    // Calculate Group Totals
    const groupTotals = sortedData.reduce((acc, item) => {
        const name = item['Spare Part Name'] || 'Unknown';
        const qty = parseFloat(item['Qty']) || 0;
        acc[name] = (acc[name] || 0) + qty;
        return acc;
    }, {});

    // Headers (Same as Supplier)
    tableHeader.innerHTML = '';
    SUPPLIER_COLUMNS.forEach(col => {
        // Skip Checkbox for Sent Claim view (read-only for now)
        if (col.key === 'checkbox') return;

        const th = document.createElement('th');
        if (col.header.includes('<')) {
            th.innerHTML = col.header;
        } else {
            th.textContent = col.header;
        }
        tableHeader.appendChild(th);
    });

    // Pagination Logic
    console.log(`Rendering Claim Sent Pagination. Total Data: ${sortedData.length}`);
    const startIndex = (currentClaimSentPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = sortedData.slice(startIndex, endIndex);

    tableBody.innerHTML = '';

    // Track previous for grouping
    let previousPartName = null;
    if (currentClaimSentPage > 1 && sortedData[startIndex - 1]) {
        previousPartName = sortedData[startIndex - 1]['Spare Part Name'];
    }

    pageData.forEach((item, index) => {
        const currentPartName = item['Spare Part Name'] || 'Unknown';
        const currentPartCode = item['Spare Part Code'] || '';

        // Group Header
        if (currentPartName !== previousPartName) {
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#f8fafc';
            headerRow.style.fontWeight = 'bold';

            const headerCell = document.createElement('td');
            headerCell.colSpan = SUPPLIER_COLUMNS.length - 1; // -1 for checkbox
            headerCell.style.padding = '12px';
            headerCell.style.borderTop = '2px solid #e2e8f0';
            headerCell.style.color = '#334155';

            const total = groupTotals[currentPartName] || 0;
            headerCell.innerHTML = `
                 <div style="display: flex; justify-content: space-between; align-items: center;">
                     <div style="display:flex; align-items:center; gap:0.5rem;">
                         <span>📦 ${currentPartName}</span>
                         <span style="font-size:0.85em; color:#64748b; font-weight:normal;">(${currentPartCode})</span>
                         <span style="font-size:0.85em; color:#0369a1; font-weight:bold;">(${total.toLocaleString()} Pc)</span>
                     </div>
                 </div>
             `;
            headerRow.appendChild(headerCell);
            tableBody.appendChild(headerRow);
            previousPartName = currentPartName;
        }

        const tr = document.createElement('tr');
        SUPPLIER_COLUMNS.forEach(col => {
            if (col.key === 'checkbox') return; // Skip checkbox

            const td = document.createElement('td');
            let value = item[col.key] || '';

            if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
                tr.appendChild(td);
                return;
            }

            if (col.key === 'Work Order' || col.key === 'Serial Number') {
                td.textContent = value;
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.textDecoration = 'underline';
                td.onclick = function (e) {
                    e.stopPropagation();
                    openWorkOrderModal(item);
                };
            } else {
                if ((col.key === 'Timestamp' || col.key === 'RecripteDate' || col.key === 'Claim Date') && value) {
                    const date = new Date(value);
                    if (!isNaN(date.getTime()) && String(value).length > 10) {
                        if (String(value).includes('T')) {
                            value = date.toLocaleString('en-GB');
                        }
                    }
                }
                if (col.key === 'Booking Date' && value) {
                    let s = String(value);
                    if (s.indexOf('T') > -1) s = s.split('T')[0];
                    if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                    value = s;
                }
                td.textContent = value;
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Render Pagination Controls
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentClaimSentPage > totalPages) currentClaimSentPage = 1;
    renderClaimSentPagination(totalPages);
}

function renderClaimSentPagination(totalPages) {
    const container = document.getElementById('claimSentPaginationControls');
    if (!container) return;

    container.style.display = 'flex';
    container.innerHTML = '';

    const createButton = (text, onClick, disabled = false, isActive = false) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.onclick = onClick;
        btn.disabled = disabled;
        if (isActive) btn.classList.add('active');
        return btn;
    };

    container.appendChild(createButton('<<', () => changeClaimSentPage(1), currentClaimSentPage === 1));
    container.appendChild(createButton('<', () => changeClaimSentPage(currentClaimSentPage - 1), currentClaimSentPage === 1));
    container.appendChild(createButton(currentClaimSentPage, () => { }, false, true));
    container.appendChild(createButton('>', () => changeClaimSentPage(currentClaimSentPage + 1), currentClaimSentPage >= totalPages));
    container.appendChild(createButton('>>', () => changeClaimSentPage(totalPages), currentClaimSentPage >= totalPages));
}

function changeClaimSentPage(newPage) {
    if (newPage < 1) return;
    currentClaimSentPage = newPage;
    renderClaimSentTable();
}

// ==========================================================
// CLAIM HISTORY TAB LOGIC
// ==========================================================

function renderHistoryTable() {
    console.log("--- renderHistoryTable called ---");
    const tableHeader = document.getElementById('historyTableHeader');
    const tableBody = document.getElementById('historyTableBody');
    const searchInput = document.getElementById('historySearchInput');
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';

    // Data Source: Global Booking Data (ALL)
    let allowedData = globalBookingData;

    // Search Filter
    if (searchTerm) {
        allowedData = allowedData.filter(item => {
            return Object.values(item).some(val =>
                String(val).toLowerCase().includes(searchTerm)
            );
        });
    }

    // Sort Data: Timestamp Descending (Newest First)
    // Or Booking Date? Let's use Timestamp if available, or fall back to Booking Date
    const sortedData = [...allowedData].sort((a, b) => {
        // Safe Date Parsing helper
        const parse = (d) => {
            if (!d) return 0;
            const t = new Date(d).getTime();
            return isNaN(t) ? 0 : t;
        };
        const timeA = parse(a['Timestamp']);
        const timeB = parse(b['Timestamp']);
        return timeB - timeA;
    });

    // Headers
    tableHeader.innerHTML = '';
    SUPPLIER_COLUMNS.forEach(col => {
        // Skip Checkbox for History (ReadOnly)
        if (col.key === 'checkbox') return;

        const th = document.createElement('th');
        if (col.header.includes('<')) {
            th.innerHTML = col.header;
        } else {
            th.textContent = col.header;
        }
        tableHeader.appendChild(th);
    });

    // Pagination Logic
    console.log(`Rendering History Pagination. Total Data: ${sortedData.length}`);
    const startIndex = (currentHistoryPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = sortedData.slice(startIndex, endIndex);

    tableBody.innerHTML = '';

    pageData.forEach((item, index) => {
        const tr = document.createElement('tr');
        SUPPLIER_COLUMNS.forEach(col => {
            if (col.key === 'checkbox') return;

            const td = document.createElement('td');
            let value = item[col.key] || '';

            if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
                tr.appendChild(td);
                return;
            }

            if (col.key === 'Work Order' || col.key === 'Serial Number') {
                td.textContent = value;
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.textDecoration = 'underline';
                td.onclick = function (e) {
                    e.stopPropagation();
                    openWorkOrderModal(item);
                };
            } else {
                if ((col.key === 'Timestamp' || col.key === 'RecripteDate' || col.key === 'Claim Date') && value) {
                    const date = new Date(value);
                    if (!isNaN(date.getTime()) && String(value).length > 10) {
                        if (String(value).includes('T')) {
                            value = date.toLocaleString('en-GB');
                        }
                    }
                }
                if (col.key === 'Booking Date' && value) {
                    let s = String(value);
                    if (s.indexOf('T') > -1) s = s.split('T')[0];
                    if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                    value = s;
                }
                td.textContent = value;
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Render Pagination Controls
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentHistoryPage > totalPages) currentHistoryPage = 1;
    renderHistoryPagination(totalPages);
}

function renderHistoryPagination(totalPages) {
    const container = document.getElementById('historyPaginationControls');
    if (!container) return;

    container.style.display = 'flex';
    container.innerHTML = '';

    const createButton = (text, onClick, disabled = false, isActive = false) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.onclick = onClick;
        btn.disabled = disabled;
        if (isActive) btn.classList.add('active');
        return btn;
    };

    container.appendChild(createButton('<<', () => changeHistoryPage(1), currentHistoryPage === 1));
    container.appendChild(createButton('<', () => changeHistoryPage(currentHistoryPage - 1), currentHistoryPage === 1));
    container.appendChild(createButton(currentHistoryPage, () => { }, false, true));
    container.appendChild(createButton('>', () => changeHistoryPage(currentHistoryPage + 1), currentHistoryPage >= totalPages));
    container.appendChild(createButton('>>', () => changeHistoryPage(totalPages), currentHistoryPage >= totalPages));
}

function changeHistoryPage(newPage) {
    if (newPage < 1) return;
    currentHistoryPage = newPage;
    renderHistoryTable();
}
