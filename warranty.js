const SCRAP_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C1MeLvv2PrOEQsqmXQplZt1xLuD5Z4KUGOrnGjYrh_A/gviz/tq?tqx=out:csv&sheet=Scrap';
const WORKFILTER_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C-_TRQcz2w0ia9bF9BfFcln5X_dX8T_cNfVUI-oU8ho/gviz/tq?tqx=out:csv&sheet=WorkFilter';
const EMPLOYEE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1eqVoLsZxGguEbRCC5rdI4iMVtQ7CK4T3uXRdx8zE3uw/gviz/tq?tqx=out:csv&sheet=EmployeeWeb';
const WARRANTY_SHEET_URL = 'https://docs.google.com/spreadsheets/d/19WWSESBcerEUlEDzM4lu3cglPoTHOWUl4P0prexqYpc/gviz/tq?tqx=out:csv&sheet=WarrantyData';
const TECHNICIAN_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1EcP7N1Yr6fAzl17ItoqPstaxrIIf1Aly/gviz/tq?tqx=out:csv&sheet=Data';

// PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL HERE
const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwq_ocJV50pVeUhuYJ5NJ114PPLDzCXuelrkYP5Mz33krtkbFp7VP5pzQEuzXnrGlDyXA/exec';

const COLUMNS = [
    { header: '<input type="checkbox" id="selectAll" onclick="toggleAllCheckboxes(this)">', source: 'C', key: 'checkbox' },
    { header: 'Status', source: 'C', key: 'status' },
    { header: 'Work Order', source: 'S', key: 'work order' },
    { header: 'Spare Part Code', source: 'S', key: 'Spare Part Code' },
    { header: 'Spare Part Name', source: 'S', key: 'Spare Part Name' },
    { header: 'Old Material Code', source: 'S', key: 'old material code' },
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

const BOOKING_COLUMNS = [
    { header: '<input type="checkbox" id="selectAllBooking" onclick="toggleAllBookingCheckboxes(this)">', key: 'checkbox' },
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
    { header: 'Timestamp', key: 'Timestamp' },
    { header: 'วันที่จองรถ', key: 'Booking Date' },
];

let allEmployees = [];
let fullData = []; // Master data (filtered by plant)
let displayedData = []; // Data currently shown (filtered by UI)
let globalBookingData = []; // Store raw warranty data for second tab
let currentPage = 1;
let editingItem = null; // Item currently being edited
const ITEMS_PER_PAGE = 20;

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
    console.log("Employees Loaded:", allEmployees.length);

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
    if (tabId === 'main') {
        tabs[0].classList.add('active');
        document.getElementById('tab-content-main').classList.add('active');
        document.getElementById('tab-content-booking').classList.remove('active');
    } else {
        tabs[1].classList.add('active');
        document.getElementById('tab-content-main').classList.remove('active');
        document.getElementById('tab-content-booking').classList.add('active');
    }
}

async function loadTableData() {
    try {
        const [scrapData, workFilterData, warrantyData, technicianData] = await Promise.all([
            fetchData(SCRAP_SHEET_URL),
            fetchData(WORKFILTER_SHEET_URL),
            fetchData(WARRANTY_SHEET_URL),
            fetchData(TECHNICIAN_SHEET_URL)
        ]);

        // Create Set of Saved Keys (KEY column in WarrantyData)
        // Create Set of Saved Keys (KEY column in WarrantyData)
        const workFilterMap = new Map();
        workFilterData.forEach(row => {
            const key = row['Work Order Number 1']?.trim();
            if (key) workFilterMap.set(key, row);
        });

        // Get Current User Plant & Normalize
        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
        const userPlant = currentUser['Plant'] ? currentUser['Plant'].toString().trim().padStart(4, '0') : null;

        // Create Set of Saved Keys AND Filter Global Booking Data
        // Filter warrantyData by userPlant
        if (userPlant) {
            globalBookingData = warrantyData.filter(row => {
                const rowPlant = row['Plant'] ? row['Plant'].toString().trim().padStart(4, '0') : '';
                const action = row['ActionStatus'] || row['Warranty Action'] || '';
                return rowPlant === userPlant && action === 'เคลมประกัน';
            });
        } else {
            globalBookingData = warrantyData.filter(row => {
                const action = row['ActionStatus'] || row['Warranty Action'] || '';
                return action === 'เคลมประกัน';
            });
        }

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

        // Join Data & Normalize Data Plant
        fullData = scrapData.map(scrapRow => {
            // Normalize Plant in Data (e.g. "326" -> "0326")
            if (scrapRow['plant']) {
                scrapRow['plant'] = scrapRow['plant'].toString().trim().padStart(4, '0');
            }

            const workOrderKey = scrapRow['work order']?.trim();

            // Generate KEY for status check (Work Order + Spare Part Code)
            const uniqueKey = (scrapRow['work order'] || '') + (scrapRow['Spare Part Code'] || '');
            const statusValue = savedStatusMap.get(uniqueKey) || '';

            return {
                scrap: scrapRow,
                fullRow: workFilterMap.get(workOrderKey) || {},
                status: statusValue,
                technicianPhone: technicianMap.get(String(scrapRow['รหัสช่าง'] || '').trim()) || '' // Safe Lookup
            };
        });

        // Filter by User Plant if exists
        if (userPlant) {
            fullData = fullData.filter(item => item.scrap['plant'] === userPlant);
        }

        // Initialize Displayed Data
        displayedData = [...fullData];
        populateFilters();
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

    pageData.forEach((item, index) => {
        const tr = document.createElement('tr');

        COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = '';

            if (col.source === 'S') value = item.scrap[col.key] || '';
            else if (col.source === 'C') value = item.status;
            else if (col.source === 'T') value = item.technicianPhone; // Tech Source
            else value = item.fullRow[col.key] || '';

            // Checkbox Column
            if (col.key === 'checkbox') {
                const isSaved = !!item.status;
                const isLE = item.fullRow['Product'] === 'L&E';
                const disabledAttr = (isSaved || isLE) ? 'disabled' : '';
                td.innerHTML = `<input type="checkbox" class="row-checkbox" value="${startIndex + index}" onchange="handleCheckboxChange()" ${disabledAttr}>`;
                td.style.textAlign = 'center';
                tr.appendChild(td);
                return; // Skip other rendering for this col
            }

            // Apply Green Tab Style for specific columns if Saved
            // Apply Green Tab Style for specific columns if Saved (Status is not empty)
            const greenColumns = ['status', 'work order', 'Spare Part Code', 'Spare Part Name', 'old material code', 'qty', 'Serial Number', 'Store Code', 'Store Name'];
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
                span.style.fontSize = '0.75rem'; // Small size
                span.style.fontWeight = '600';
                span.style.padding = '0.25rem 0.5rem';
                span.style.borderRadius = '4px';
                span.style.display = 'inline-block'; // Ensure padding works
                td.appendChild(span);
            } else {
                td.textContent = value;
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
        // Default to Poom for L&E
        const defaultReceiver = 'Poom';
        document.getElementById('store_receiver').value = defaultReceiver;
        document.querySelectorAll('.receiver-btn').forEach(b => {
            if (b.textContent === defaultReceiver) {
                b.style.backgroundColor = '#3b82f6';
                b.style.color = 'white';
                b.style.border = 'none';
            }
        });
    } else {
        receiverSection.style.display = 'none';
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
            ((String(item.scrap['Spare Part Code']) === '30027106') ? 'Mai' :
                (item.fullRow['Product'] === 'L&E') ? (document.getElementById('store_receiver').value || '') : 'Mon')
            : ''
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
            'Claim Receiver': (actionName === 'เคลมประกัน') ?
                ((String(item.scrap['Spare Part Code']) === '30027106') ? 'Mai' :
                    (item.fullRow['Product'] !== 'L&E') ? 'Mon' : '')
                : ''
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

function toggleAllBookingCheckboxes(source) {
    const checkboxes = document.querySelectorAll('.booking-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = source.checked;
    });
    handleBookingCheckboxChange();
}

function handleBookingCheckboxChange() {
    const checkboxes = document.querySelectorAll('.booking-checkbox:checked');
    const actionDiv = document.getElementById('bookingBulkActions');
    if (checkboxes.length > 0) {
        actionDiv.style.display = 'flex';
    } else {
        actionDiv.style.display = 'none';
    }
}

async function processBookingAction(destination) {
    const checkboxes = document.querySelectorAll('.booking-checkbox:checked');
    if (checkboxes.length === 0) return;

    // 1. Re-derive filtered & sorted data to match indices
    const filterValue = document.getElementById('bookingPartFilter').value;
    let filteredData = globalBookingData;
    if (filterValue) {
        filteredData = globalBookingData.filter(row => row['Spare Part Name'] === filterValue);
    }
    const sortedData = [...filteredData].reverse();

    // 2. Identify selected items
    const selectedItems = [];
    checkboxes.forEach(cb => {
        const index = parseInt(cb.value);
        if (sortedData[index]) selectedItems.push(sortedData[index]);
    });

    if (selectedItems.length === 0) return;

    // 3. Prompt for Booking Slip
    let plantCenterCode = '';
    if (destination.includes('นวนคร')) plantCenterCode = '0301';
    if (destination.includes('วิภาวดี')) plantCenterCode = '0326';

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

    Swal.fire({
        title: 'กำลังบันทึก...',
        html: `กำลังประมวลผล 0 / ${selectedItems.length}`,
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];
        Swal.getHtmlContainer().textContent = `กำลังประมวลผล ${i + 1} / ${selectedItems.length}`;

        // Construct Payload
        // Debug: Check if destination is received
        // console.log("Sending to:", destination); 

        // Manually format date to DD/MM/YYYY to ensure no time component
        const day = String(bookingDate.getDate()).padStart(2, '0');
        const month = String(bookingDate.getMonth() + 1).padStart(2, '0');
        const year = bookingDate.getFullYear();
        const formattedDate = `${day}/${month}/${year}`;

        const payload = {
            ...item,
            'Booking Slip': bookingSlip,
            'Booking Date': formattedDate,
            'Plantcenter': plantCenterCode, // Use Code (0301/0326)
            // GAS needs 'KEY' or Work Order + Part Code. 
            // 'item' from globalBookingData already has all columns.
        };

        // DEBUG: Alert for the first item to verify payload
        if (i === 0) {
            console.log("PAYLOAD DEBUG:", JSON.stringify(payload, null, 2));
        }

        // Cleanup: GAS might fail if unknown keys are abundant? 
        // Code.gs ignores unknown keys.
        // IMPORTANT: ActionStatus must be preserved or set. 
        // item already has it.

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
            item['Plantcenter'] = plantCenterCode; // Update UI with Code
            successCount++;

        } catch (e) {
            console.error(e);
            failCount++;
        }
    }

    renderBookingTable();
    toggleAllBookingCheckboxes({ checked: false });

    Swal.fire({
        icon: successCount > 0 ? 'success' : 'error',
        title: 'เสร็จสิ้น',
        text: `อัปเดต ${successCount} รายการเรียบร้อย (ล้มเหลว: ${failCount})`,
    });
}

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
    if (filterValue) {
        filteredData = globalBookingData.filter(row => row['Spare Part Name'] === filterValue);
    }

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

    // Rows
    tableBody.innerHTML = '';
    // Display newest first (reverse order) if preferred, or as is. 
    // Usually Sheets appends to bottom, so reverse is better for "Latest".
    const sortedData = [...filteredData].reverse();

    sortedData.forEach((row, index) => {
        const tr = document.createElement('tr');
        BOOKING_COLUMNS.forEach(col => {
            const td = document.createElement('td');

            if (col.key === 'checkbox') {
                td.innerHTML = `<input type="checkbox" class="booking-checkbox" value="${index}" onchange="handleBookingCheckboxChange()">`;
                td.style.textAlign = 'center';
                tr.appendChild(td);
                return;
            }

            let value = row[col.key] || '';

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
            if (col.key === 'Mobile' && value) {
                value = formatPhoneNumber(value);
            }

            // Booking Slip Click Action
            if (col.key === 'Booking Slip') {
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
        'Plantcenter': newPlantCenter !== null ? newPlantCenter : item['Plantcenter'] // Update if provided, else keep same
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
        console.error(error);
        Swal.fire('Error', 'Failed to update.', 'error');
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
    const receiverOptions = ['Mai', 'Mon', 'Poom'];
    let buttonsHtml = '<div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">';

    receiverOptions.forEach(opt => {
        let btnColor = '#f1f5f9';
        let txtColor = '#64748b';
        if (opt === 'Mai') { btnColor = '#dbeafe'; txtColor = '#1e40af'; }
        if (opt === 'Mon') { btnColor = '#dcfce7'; txtColor = '#166534'; }
        if (opt === 'Poom') { btnColor = '#fae8ff'; txtColor = '#6b21a8'; }
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

initAuth();

document.getElementById('passInput').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        handleLogin();
    }
});
