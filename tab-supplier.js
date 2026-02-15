let supplierProductOptions = new Set();
let currentSupplierPage = 1;
let currentClaimSentPage = 1;
let currentHistoryPage = 1;
let currentSupplierDisplayedData = [];
// Global Selection Helper
let selectedSupplierKeys = new Set();
let activeClaimDateFilter = null;
function getSupplierKey(item) {
    return (item['Work Order'] || '') + '|' + (item['Spare Part Code'] || '') + '|' + (item['Serial Number'] || '');
}

function isDateToday(dateStr) {
    if (!dateStr) return false;
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const todayStr = `${day}/${month}/${year}`;
    if (dateStr === todayStr) return true;

    // Handle DD/MM/YYYY explicitly to prevent Invalid Date in some browsers
    if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        if (parts.length === 3) {
            return `${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[2]}` === todayStr;
        }
    }

    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
        const dDay = String(d.getDate()).padStart(2, '0');
        const dMonth = String(d.getMonth() + 1).padStart(2, '0');
        const dYear = d.getFullYear();
        return `${dDay}/${dMonth}/${dYear}` === todayStr;
    }
    return false;
}

function handleSupplierSearch() { currentSupplierPage = 1; renderSupplierTable(); }

function populateSupplierProductFilter(data) {
    const dropdown = document.getElementById('supplierProductDropdown');
    if (!dropdown) return;
    let sourceData = data;
    if (!sourceData) {
        sourceData = globalBookingData.filter(item => {
            const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
            const hasClaimSup = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
            const claimDate = item['Claim Date'] ? String(item['Claim Date']).trim() : '';
            if (!hasRecripte) return false;
            if (!hasClaimSup || !claimDate) return true;
            return isDateToday(claimDate);
        });
    }
    const uniqueProducts = new Set();
    sourceData.forEach(item => {
        if (item['Product']) uniqueProducts.add(item['Product']);
    });
    const sortedProducts = Array.from(uniqueProducts).sort();
    dropdown.innerHTML = '';
    sortedProducts.forEach(prod => {
        const div = document.createElement('div');
        div.style.padding = '0.25rem 0.5rem';
        div.style.display = 'flex';
        div.style.alignItems = 'center';
        div.style.gap = '0.5rem';
        div.style.cursor = 'pointer';
        div.className = 'hover:bg-gray-100';
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
    const label = document.getElementById('supplierProductFilterLabel');
    if (label) {
        if (supplierProductOptions.size === 0) label.textContent = 'All Products';
        else label.textContent = `Selected (${supplierProductOptions.size})`;
    }
    renderSupplierTable();
}

window.addEventListener('click', function (e) {
    const dropdown = document.getElementById('supplierProductDropdown');
    const button = document.querySelector('button[onclick="toggleSupplierProductDropdown()"]');
    if (dropdown && button && !dropdown.contains(e.target) && !button.contains(e.target)) {
        dropdown.style.display = 'none';
    }
});

function populateSupplierFilter() {
    const filterSelect = document.getElementById('supplierPartFilter');
    let supplierItems = globalBookingData.filter(item => {
        const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
        const hasClaimSup = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
        const claimDate = item['Claim Date'] ? String(item['Claim Date']).trim() : '';
        if (!hasRecripte) return false;
        if (!hasClaimSup || !claimDate) return true;
        return isDateToday(claimDate);
    });
    if (filterSelect) {
        const parts = new Set();
        supplierItems.forEach(item => {
            const name = item['Spare Part Name'];
            if (name) parts.add(name);
        });
        const sortedParts = Array.from(parts).sort();
        const currentPart = filterSelect.value;
        filterSelect.innerHTML = '<option value="">All Spare Parts</option>';
        sortedParts.forEach(part => {
            const option = document.createElement('option');
            option.value = part;
            option.textContent = part;
            if (part === currentPart) option.selected = true;
            filterSelect.appendChild(option);
        });
    }
    // 2. Populate Product Filter
    populateSupplierProductFilter(supplierItems);

    // 3. Populate Receiver Filter (New)
    const receiverFilter = document.getElementById('supplierReceiverFilter');
    if (receiverFilter) {
        // Collect unique receivers
        const receivers = new Set();
        supplierItems.forEach(item => {
            const receiver = item['Claim Receiver'];
            if (receiver) receivers.add(receiver);
        });
        const sortedReceivers = Array.from(receivers).sort();

        // Save current selection
        const currentReceiver = receiverFilter.value;

        receiverFilter.innerHTML = '<option value="">All Receivers</option>';
        sortedReceivers.forEach(recv => {
            const option = document.createElement('option');
            option.value = recv;
            option.textContent = recv;
            if (recv === currentReceiver) option.selected = true;
            receiverFilter.appendChild(option);
        });
    }
}

function renderSupplierTable() {
    const tableHeader = document.getElementById('supplierTableHeader');
    const tableBody = document.getElementById('supplierTableBody');
    const filterSelect = document.getElementById('supplierPartFilter');
    const searchInput = document.getElementById('supplierSearchInput');
    const statusFilter = document.getElementById('supplierStatusFilter');
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const filterValue = filterSelect ? filterSelect.value : '';
    const statusValue = statusFilter ? statusFilter.value : '';
    let supplierData = globalBookingData.filter(item => {
        const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
        const hasClaimSup = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
        const claimDate = item['Claim Date'] ? String(item['Claim Date']).trim() : '';
        if (!hasRecripte) return false;
        if (!hasClaimSup || !claimDate) return true;
        return isDateToday(claimDate);
    });

    // 1. FILTER BY USER PLANT (Added)
    const userPlant = getEffectiveUserPlant();
    const userStatus = getCurrentUserStatus();
    if (userPlant) {
        supplierData = supplierData.filter(item => {
            // Allow if User Status matches Claim Receiver
            if (userStatus && item['Claim Receiver'] === userStatus) return true;

            const itemPlant = item['Plant'] || item['plant'] || item['PLANT'];
            return String(itemPlant).trim().padStart(4, '0') === userPlant;
        });
    }
    const dropdown = document.getElementById('supplierProductDropdown');
    if (dropdown && dropdown.children.length === 0) {
        populateSupplierProductFilter(supplierData);
    }
    if (supplierProductOptions.size > 0) {
        supplierData = supplierData.filter(item => supplierProductOptions.has(item['Product']));
    }

    // Apply Receiver Filter (New)
    const receiverFilter = document.getElementById('supplierReceiverFilter');
    const receiverValue = receiverFilter ? receiverFilter.value : '';
    if (receiverValue) {
        supplierData = supplierData.filter(item => item['Claim Receiver'] === receiverValue);
    }

    if (filterValue) {
        supplierData = supplierData.filter(item => item['Spare Part Name'] === filterValue);
    }
    if (statusValue) {
        supplierData = supplierData.filter(item => {
            const isClaimed = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
            const status = isClaimed ? 'ส่งเคลมแล้ว' : 'รอส่งเคลม';
            return status === statusValue;
        });
    }
    if (searchTerm) {
        supplierData = supplierData.filter(item => {
            return Object.values(item).some(val => String(val || '').toLowerCase().includes(searchTerm));
        });
    }
    const sortedData = [...supplierData].reverse().sort((a, b) => {
        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });

    // Store globally for bulk actions
    currentSupplierDisplayedData = sortedData;

    const groupTotals = sortedData.reduce((acc, item) => {
        const name = item['Spare Part Name'] || 'Unknown';
        const qty = parseFloat(item['Qty']) || 0;
        acc[name] = (acc[name] || 0) + qty;
        return acc;
    }, {});

    tableHeader.innerHTML = '';
    SUPPLIER_COLUMNS.forEach(col => {
        const th = document.createElement('th');
        if (col.header.includes('<')) th.innerHTML = col.header;
        else th.textContent = col.header;
        tableHeader.appendChild(th);
    });

    const startIndex = (currentSupplierPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = sortedData.slice(startIndex, endIndex);
    tableBody.innerHTML = '';

    let previousPartName = null;
    if (currentSupplierPage > 1 && sortedData[startIndex - 1]) {
        previousPartName = sortedData[startIndex - 1]['Spare Part Name'];
    }

    pageData.forEach((item, index) => {
        const currentPartName = item['Spare Part Name'] || 'Unknown';
        const currentPartCode = item['Spare Part Code'] || '';
        if (currentPartName !== previousPartName) {
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#f8fafc';
            headerRow.style.fontWeight = 'bold';
            const total = groupTotals[currentPartName] || 0;
            headerRow.innerHTML = `<td colspan="${SUPPLIER_COLUMNS.length}" style="padding: 12px; border-top: 2px solid #e2e8f0; color: #334155;"><div style="display: flex; justify-content: space-between; align-items: center;"><div style="display:flex; align-items:center; gap:0.5rem;"><span>📦 ${currentPartName}</span><span style="font-size:0.85em; color:#64748b; font-weight:normal;">(${currentPartCode})</span><span style="font-size:0.85em; color:#0369a1; font-weight:bold;">(${total.toLocaleString()} Pc)</span></div></div></td>`;
            tableBody.appendChild(headerRow);
            previousPartName = currentPartName;
        }
        const tr = document.createElement('tr');
        SUPPLIER_COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = item[col.key] || '';
            if (col.key === 'checkbox') {
                const globalIndex = startIndex + index;
                const isClaimed = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
                const disabledAttr = isClaimed ? 'disabled' : '';

                // New Logic: Check Set
                const key = getSupplierKey(item);
                const isChecked = selectedSupplierKeys.has(key) ? 'checked' : '';

                td.innerHTML = `<input type="checkbox" class="supplier-checkbox" value="${globalIndex}" ${disabledAttr} ${isChecked} onchange="handleSupplierCheckboxChange(this)">`;
                td.style.textAlign = 'center';
                if (isClaimed) td.style.opacity = '0.5';
            } else if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
                const isClaimed = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
                if (!isClaimed) {
                    td.style.cursor = 'pointer';
                    td.title = 'Click to Send Claim';
                    td.onclick = function (e) {
                        e.stopPropagation();
                        sendSingleClaim(item);
                    };
                } else {
                    td.style.cursor = 'pointer';
                    td.title = 'Click to Cancel Claim';
                    td.onclick = function (e) {
                        e.stopPropagation();
                        cancelSingleClaim(item);
                    };
                }
            } else if (col.key === 'Work Order' || col.key === 'Serial Number') {
                td.textContent = value;
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.textDecoration = 'underline';
                td.onclick = function (e) { e.stopPropagation(); openWorkOrderModal(item); };
            } else {
                if ((col.key === 'Timestamp' || col.key === 'RecripteDate' || col.key === 'Claim Date') && value) {
                    const date = new Date(value);
                    if (!isNaN(date.getTime()) && String(value).length > 10 && String(value).includes('T')) value = date.toLocaleString('en-GB');
                }
                if (col.key === 'Booking Date' && value) {
                    let s = String(value);
                    if (s.indexOf('T') > -1) s = s.split('T')[0];
                    if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                    value = s;
                }
                if (col.key === 'Plantcenter' && value) {
                    const v = String(value).replace(/^0+/, '');
                    if (v === '301') value = 'คลังนวนคร';
                    else if (v === '326') value = 'คลังวิภาวดี';
                }
                td.textContent = value;
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentSupplierPage > totalPages) currentSupplierPage = 1;
    renderGenericPagination('supplierPaginationControls', currentSupplierPage, totalPages, changeSupplierPage);

    // Update Select All State
    const selectAll = document.getElementById('selectAllSupplier');
    const bulkActions = document.getElementById('supplierBulkActions');

    // Only items that can be selected (not already claimed)
    const validItems = sortedData.filter(item => !item['ClaimSup']);

    if (validItems.length > 0) {
        const allSelected = validItems.every(item => selectedSupplierKeys.has(getSupplierKey(item)));
        const someSelected = validItems.some(item => selectedSupplierKeys.has(getSupplierKey(item)));

        if (allSelected) { selectAll.checked = true; selectAll.indeterminate = false; }
        else if (someSelected) { selectAll.checked = false; selectAll.indeterminate = true; }
        else { selectAll.checked = false; selectAll.indeterminate = false; }
    } else {
        selectAll.checked = false; selectAll.indeterminate = false;
    }

    if (selectedSupplierKeys.size > 0) {
        bulkActions.style.display = 'flex';
        // Optional: Update button text to show count
        // const btn = bulkActions.querySelector('button');
        // if(btn) btn.innerHTML = `Send Claim (${selectedSupplierKeys.size})`;
    } else {
        bulkActions.style.display = 'none';
    }
}

function toggleAllSupplierCheckboxes(source) {
    if (!currentSupplierDisplayedData || currentSupplierDisplayedData.length === 0) return;

    const isChecked = source.checked;

    currentSupplierDisplayedData.forEach(item => {
        if (!item['ClaimSup']) {
            const key = getSupplierKey(item);
            if (isChecked) selectedSupplierKeys.add(key);
            else selectedSupplierKeys.delete(key);
        }
    });

    renderSupplierTable();
}

function handleSupplierCheckboxChange(checkbox) {
    // If called from individual checkbox change
    if (checkbox) {
        const globalIndex = parseInt(checkbox.value);
        // We use currentSupplierDisplayedData because sortedData is local to render. 
        // globalIndex passed to value in render is accurate to the array index.
        const item = currentSupplierDisplayedData[globalIndex];

        if (item) {
            const key = getSupplierKey(item);
            if (checkbox.checked) selectedSupplierKeys.add(key);
            else selectedSupplierKeys.delete(key);
        }
    }
    // Just refresh UI state
    renderSupplierTable();
}

async function sendSingleClaim(item) {
    if (!item) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const claimSup = currentUser.IDRec || 'Unknown';
    const now = new Date();
    const formattedDate = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

    const result = await Swal.fire({
        title: 'Confirm Send Claim?',
        html: `Update <b>1</b> item?<br>Item: <b>${item['Spare Part Name']}</b><br><br>Date: <b>${formattedDate}</b><br>Supplier: <b>${claimSup}</b>`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'Yes, Send',
        cancelButtonText: 'Cancel'
    });

    if (!result.isConfirmed) return;

    // Robust Data Lookup
    let dateReceived = item['วันที่รับซาก'];
    let receiver = item['ผู้รับซาก'];
    let keep = item['Keep'];
    let ciName = item['CI Name'];
    let problem = item['Problem'];
    let productType = item['Product Type'];

    if ((!dateReceived || !receiver || !keep || !ciName || !problem || !productType) && typeof fullData !== 'undefined') {
        const targetKey = ((item['Work Order'] || '') + (item['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
        const match = fullData.find(d => {
            const dKey = ((d.scrap['work order'] || '') + (d.scrap['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
            return dKey === targetKey;
        });
        if (match) {
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = { ...item, 'Claim Date': formattedDate, 'ClaimSup': claimSup, 'user': currentUser.IDRec || 'Unknown', 'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || '' };

    // Optimistic
    item['Claim Date'] = formattedDate;
    item['ClaimSup'] = claimSup;

    SaveQueue.add(payload);

    renderSupplierTable();
    try {
        populateClaimSentFilter();
        renderClaimSentTable();
        renderHistoryTable();
    } catch (e) { console.error(e); }
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Item sent to claim' });
}

async function cancelSingleClaim(item) {
    if (!item) return;

    const result = await Swal.fire({
        title: 'ยกเลิกการส่งเคลม?',
        text: `คุณต้องการยกเลิกสถานะ "ส่งเคลมแล้ว" ของรายการ: ${item['Spare Part Name']}?`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ใช่, ยกเลิก',
        cancelButtonText: 'ไม่',
        confirmButtonColor: '#d33'
    });

    if (!result.isConfirmed) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

    // Robust Data Lookup
    let dateReceived = item['วันที่รับซาก'];
    let receiver = item['ผู้รับซาก'];
    let keep = item['Keep'];
    let ciName = item['CI Name'];
    let problem = item['Problem'];
    let productType = item['Product Type'];

    if ((!dateReceived || !receiver || !keep || !ciName || !problem || !productType) && typeof fullData !== 'undefined') {
        const targetKey = ((item['Work Order'] || '') + (item['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
        const match = fullData.find(d => {
            const dKey = ((d.scrap['work order'] || '') + (d.scrap['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
            return dKey === targetKey;
        });
        if (match) {
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = { ...item, 'Claim Date': '', 'ClaimSup': '', 'user': currentUser.IDRec || 'Unknown', 'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || '' };

    // Optimistic
    item['Claim Date'] = '';
    item['ClaimSup'] = '';

    SaveQueue.add(payload);

    renderSupplierTable();
    try {
        populateClaimSentFilter();
        renderClaimSentTable();
        renderHistoryTable();
    } catch (e) { console.error(e); }
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Claim cancelled' });
}

async function sendClaim() {
    // Collect Items from globalBookingData (or currentSupplierDisplayedData) that match the keys
    // BUT currentSupplierDisplayedData is filtered. If user selects all, changes filter, currentSupplierDisplayedData changes.
    // However, key is unique. We should iterate globalBookingData (or at least all allowedData re-calculated, but Set has keys).
    // Safest is to find items in globalBookingData that match keys AND are valid for 'Send Claim'.

    // Efficiency: iterating globalBookingData is fine.

    if (selectedSupplierKeys.size === 0) return;

    // We need to find the actual objects.
    // We can use the logic from 'processBookingAction' with a lookup map.
    const lookupMap = new Map();
    globalBookingData.forEach(row => {
        // Only consider items that are NOT claimed yet? logic below will overwrite anyway but it's cleaner.
        if (!row['ClaimSup']) {
            lookupMap.set(getSupplierKey(row), row);
        }
    });

    const selectedItems = [];
    selectedSupplierKeys.forEach(key => {
        const item = lookupMap.get(key);
        if (item) selectedItems.push(item);
    });

    if (selectedItems.length === 0) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const claimSup = currentUser.IDRec || 'Unknown';
    const now = new Date();
    const formattedDate = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

    const result = await Swal.fire({ title: 'Confirm Send Claim?', html: `Update <b>${selectedItems.length}</b> items?<br><br>Date: <b>${formattedDate}</b><br>Supplier: <b>${claimSup}</b>`, icon: 'question', showCancelButton: true, confirmButtonText: 'Yes, Send', cancelButtonText: 'Cancel' });
    if (!result.isConfirmed) return;

    selectedItems.forEach(item => {
        // Robust Data Lookup: Ensure critical fields exist before sending
        let dateReceived = item['วันที่รับซาก'];
        let receiver = item['ผู้รับซาก'];
        let keep = item['Keep'];
        let ciName = item['CI Name'];
        let problem = item['Problem'];
        let productType = item['Product Type'];

        if ((!dateReceived || !receiver || !keep || !ciName || !problem || !productType) && typeof fullData !== 'undefined') {
            const targetKey = ((item['Work Order'] || '') + (item['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
            const match = fullData.find(d => {
                const dKey = ((d.scrap['work order'] || '') + (d.scrap['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
                return dKey === targetKey;
            });
            if (match) {
                const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
                if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
                if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
                if (!keep) keep = getVal(match.scrap, 'Keep');
                if (!ciName) ciName = match.fullRow['CI Name'] || '';
                if (!problem) problem = match.fullRow['Problem'] || '';
                if (!productType) productType = match.fullRow['Product Type'] || '';
            }
        }

        const payload = { ...item, 'Claim Date': formattedDate, 'ClaimSup': claimSup, 'user': JSON.parse(localStorage.getItem('currentUser') || '{}').IDRec || 'Unknown', 'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || '' };

        // Optimistic
        item['Claim Date'] = formattedDate;
        item['ClaimSup'] = claimSup;

        SaveQueue.add(payload);
    });

    // CLEAR SELECTION AFTER SEND
    selectedSupplierKeys.clear();

    renderSupplierTable();
    handleSupplierCheckboxChange(); // Updates UI
    try {
        populateClaimSentFilter();
        renderClaimSentTable();
        renderHistoryTable();
    } catch (e) { console.error(e); }
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: `Sent ${selectedItems.length} items to claim` });
}

function changeSupplierPage(newPage) { if (newPage < 1) return; currentSupplierPage = newPage; renderSupplierTable(); }

// ==========================================================
// CLAIM SENT TAB LOGIC
// ==========================================================

function populateClaimSentFilter() {
    const receiverSelect = document.getElementById('claimSentReceiverFilter');

    // Filter Items: Must have 'Recripte' AND 'ClaimSup'
    const claimSentItems = globalBookingData.filter(item => {
        const hasRecripte = item['Recripte'] && item['Recripte'].trim() !== '';
        const hasClaimSup = item['ClaimSup'] && item['ClaimSup'].trim() !== '';
        return hasRecripte && hasClaimSup;
    });

    if (receiverSelect) {
        const receivers = [...new Set(claimSentItems.map(item => item['Claim Receiver']).filter(Boolean))].sort();
        receiverSelect.innerHTML = '<option value="">All Receivers</option>';
        receivers.forEach(receiver => {
            const option = document.createElement('option');
            option.value = receiver;
            option.textContent = receiver;
            receiverSelect.appendChild(option);
        });
    }
}

function renderClaimSentTable() {
    console.log("--- renderClaimSentTable called ---");
    const tableHeader = document.getElementById('claimSentTableHeader');
    const tableBody = document.getElementById('claimSentTableBody');

    // Filter Data: 'Recripte' AND 'ClaimSup'
    // Filter Data: 'Recripte' AND 'ClaimSup'
    const receiverSelect = document.getElementById('claimSentReceiverFilter');
    const searchInput = document.getElementById('claimSentSearchInput');

    const receiverValue = receiverSelect ? receiverSelect.value : '';
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';

    let allowedData = globalBookingData.filter(item => {
        const hasRecripte = item['Recripte'] && item['Recripte'].trim() !== '';
        const hasClaimSup = item['ClaimSup'] && item['ClaimSup'].trim() !== '';
        return hasRecripte && hasClaimSup;
    });

    // 1. FILTER BY USER PLANT (Added)
    const userPlant = getEffectiveUserPlant();
    const userStatus = getCurrentUserStatus();
    if (userPlant) {
        allowedData = allowedData.filter(item => {
            // Allow if User Status matches Claim Receiver
            if (userStatus && item['Claim Receiver'] === userStatus) return true;

            const itemPlant = item['Plant'] || item['plant'] || item['PLANT'];
            return String(itemPlant).trim().padStart(4, '0') === userPlant;
        });
    }

    if (receiverValue) {
        allowedData = allowedData.filter(item => item['Claim Receiver'] === receiverValue);
    }

    if (activeClaimDateFilter) {
        allowedData = allowedData.filter(item => (item['Claim Date'] || 'Unknown') === activeClaimDateFilter);
    }

    if (searchTerm) {
        allowedData = allowedData.filter(item => {
            return Object.values(item).some(val => String(val || '').toLowerCase().includes(searchTerm));
        });
    }

    // Sort Data: Claim Date (Desc), then Spare Part Name (Asc) - Requested by User
    const sortedData = [...allowedData].sort((a, b) => {
        const getDateTimestamp = (d) => {
            if (!d) return 0;
            const parts = String(d).split('/');
            if (parts.length === 3) {
                return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
            }
            return 0;
        };

        const dateA = getDateTimestamp(a['Claim Date']);
        const dateB = getDateTimestamp(b['Claim Date']);

        if (dateB !== dateA) return dateB - dateA; // Descending Date

        const prodA = a['Product'] || '';
        const prodB = b['Product'] || '';
        const prodCompare = prodA.localeCompare(prodB, 'th');
        if (prodCompare !== 0) return prodCompare;

        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });

    // Calculate Group Totals for Dates and Products
    const groupTotals = sortedData.reduce((acc, item) => {
        const date = item['Claim Date'] || 'Unknown';
        const prod = item['Product'] || 'Unknown';
        const qty = parseFloat(item['Qty']) || 0;

        if (!acc[date]) acc[date] = { total: 0, products: {} };
        acc[date].total += qty;

        if (!acc[date].products[prod]) acc[date].products[prod] = 0;
        acc[date].products[prod] += qty;

        return acc;
    }, {});

    // Headers (Same as Supplier)
    tableHeader.innerHTML = '';
    SUPPLIER_COLUMNS.forEach(col => {
        const th = document.createElement('th');
        if (col.key === 'checkbox') {
            th.innerHTML = '<input type="checkbox" id="selectAllClaimSent" onclick="toggleAllClaimSentCheckboxes(this)">';
        } else if (col.header.includes('<')) {
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

    // Track previous for grouping (Claim Date)
    let previousClaimDate = null;
    let previousProduct = null;

    if (currentClaimSentPage > 1 && sortedData[startIndex - 1]) {
        previousClaimDate = sortedData[startIndex - 1]['Claim Date'];
        previousProduct = sortedData[startIndex - 1]['Product'];
    }

    // Store for "Select All" logic
    currentClaimSentDisplayedData = sortedData;

    pageData.forEach((item, index) => {
        const currentClaimDate = item['Claim Date'] || 'Unknown';
        const currentProduct = item['Product'] || 'Unknown';

        // Group Header (Claim Date)
        if (currentClaimDate !== previousClaimDate) {
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#e0f2fe'; // Light Blue for Date
            headerRow.style.fontWeight = 'bold';

            const headerCell = document.createElement('td');
            headerCell.colSpan = SUPPLIER_COLUMNS.length;
            headerCell.style.padding = '12px';
            headerCell.style.borderTop = '2px solid #cbd5e1';
            headerCell.style.color = '#0f172a';

            const total = groupTotals[currentClaimDate]?.total || 0;
            const safeDate = currentClaimDate.replace(/'/g, "\\'");
            const isFilterActive = activeClaimDateFilter === currentClaimDate;
            const filterLabel = isFilterActive ? '❌ Show All' : '🔍 Filter Only';
            const filterStyle = isFilterActive ? 'color: #ef4444; font-weight: bold;' : 'color: #64748b;';

            headerCell.innerHTML = `
                 <div style="display: flex; justify-content: space-between; align-items: center; cursor: pointer;" onclick="filterClaimSentByDate('${safeDate}')" title="${isFilterActive ? 'Click to show all dates' : 'Click to show only this date'}">
                     <div style="display:flex; align-items:center; gap:0.5rem;">
                         <span style="font-size: 1.1em;">📅 ${currentClaimDate}</span>
                         <span style="font-size:0.9em; color:#0369a1; font-weight:bold;">(Total: ${total.toLocaleString()} Pc)</span>
                     </div>
                     <div style="font-size: 0.8em; ${filterStyle}">${filterLabel}</div>
                 </div>
             `;
            headerRow.appendChild(headerCell);
            tableBody.appendChild(headerRow);
            
            previousClaimDate = currentClaimDate;
            previousProduct = null; // Reset product grouping when date changes
        }

        // Sub-Group Header (Product)
        if (currentProduct !== previousProduct) {
            const prodHeaderRow = document.createElement('tr');
            prodHeaderRow.className = 'sub-group-header-row';
            prodHeaderRow.style.backgroundColor = '#f8fafc'; // Very light gray for Product
            prodHeaderRow.style.fontWeight = 'bold';

            const prodHeaderCell = document.createElement('td');
            prodHeaderCell.colSpan = SUPPLIER_COLUMNS.length;
            prodHeaderCell.style.padding = '8px 12px 8px 32px'; // Indent
            prodHeaderCell.style.borderTop = '1px solid #e2e8f0';
            prodHeaderCell.style.color = '#334155';

            // Check if all eligible items in this group are selected
            const groupItems = sortedData.filter(d => (d['Claim Date'] || 'Unknown') === currentClaimDate && (d['Product'] || 'Unknown') === currentProduct);
            const eligibleGroupItems = groupItems.filter(item => {
                 const action = item['Warranty Action'] || item['ActionStatus'];
                 const hasAction = action && action !== 'เคลมประกัน' && action !== 'Pending' && action !== 'บันทึกแล้ว';
                 return !hasAction && (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '');
            });
            const allSelected = eligibleGroupItems.length > 0 && eligibleGroupItems.every(d => selectedClaimSentKeys.has(getSupplierKey(d)));
            const checkedStr = allSelected ? 'checked' : '';
            const safeDate = currentClaimDate.replace(/'/g, "\\'");
            const safeProduct = currentProduct.replace(/'/g, "\\'");

            const prodTotal = groupTotals[currentClaimDate]?.products[currentProduct] || 0;
            prodHeaderCell.innerHTML = `
                 <div style="display: flex; justify-content: space-between; align-items: center;">
                     <div style="display:flex; align-items:center; gap:0.5rem;">
                         <input type="checkbox" onclick="toggleClaimSentProductGroup(this, '${safeDate}', '${safeProduct}')" ${checkedStr} style="cursor: pointer;">
                         <span>📦 ${currentProduct}</span>
                         <span style="font-size:0.85em; color:#475569; font-weight:normal;">(${prodTotal.toLocaleString()} Pc)</span>
                     </div>
                 </div>
             `;
            prodHeaderRow.appendChild(prodHeaderCell);
            tableBody.appendChild(prodHeaderRow);

            previousProduct = currentProduct;
        }

        const tr = document.createElement('tr');
        SUPPLIER_COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = item[col.key] || '';

            if (col.key === 'checkbox') {
                const globalIndex = startIndex + index;
                const key = getSupplierKey(item);

                // Logic to determine if status is 'ส่งเคลมแล้ว'
                const action = item['Warranty Action'] || item['ActionStatus'];
                const hasAction = action && action !== 'เคลมประกัน' && action !== 'Pending' && action !== 'บันทึกแล้ว';
                const isClaimSent = !hasAction && (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '');

                if (isClaimSent) {
                    const isChecked = selectedClaimSentKeys.has(key) ? 'checked' : '';
                    td.innerHTML = `<input type="checkbox" class="claim-sent-checkbox" value="${globalIndex}" ${isChecked} onchange="handleClaimSentCheckboxChange(this)">`;
                } else {
                    td.innerHTML = '';
                }
                td.style.textAlign = 'center';
            } else if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
                td.style.cursor = 'pointer';
                td.title = 'Click to update warranty status';
                td.onclick = function (e) {
                    e.stopPropagation();
                    openWorkOrderModal(item, 'warrantyDecision');
                };
                tr.appendChild(td);
                return;
            } else if (col.key === 'Work Order' || col.key === 'Serial Number') {
                td.textContent = value;
                // Action Removed as requested
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
                if (col.key === 'Plantcenter' && value) {
                    const v = String(value).replace(/^0+/, '');
                    if (v === '301') value = 'คลังนวนคร';
                    else if (v === '326') value = 'คลังวิภาวดี';
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
    renderGenericPagination('claimSentPaginationControls', currentClaimSentPage, totalPages, changeClaimSentPage);

    // Update Select All & Buttons
    const selectAll = document.getElementById('selectAllClaimSent');
    const bulkActions = document.getElementById('claimSentBulkActions');

    // Filter only eligible items for "Select All" status checks
    const eligibleItems = sortedData.filter(item => {
        const action = item['Warranty Action'] || item['ActionStatus'];
        const hasAction = action && action !== 'เคลมประกัน' && action !== 'Pending' && action !== 'บันทึกแล้ว';
        return !hasAction && (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '');
    });

    if (eligibleItems.length > 0) {
        const allSelected = eligibleItems.every(item => selectedClaimSentKeys.has(getSupplierKey(item)));
        const someSelected = eligibleItems.some(item => selectedClaimSentKeys.has(getSupplierKey(item)));

        if (selectAll) {
            selectAll.disabled = false;
            if (allSelected) {
                selectAll.checked = true;
                selectAll.indeterminate = false;
            } else if (someSelected) {
                selectAll.checked = false;
                selectAll.indeterminate = true;
            } else {
                selectAll.checked = false;
                selectAll.indeterminate = false;
            }
        }
    } else {
        if (selectAll) {
            selectAll.checked = false;
            selectAll.indeterminate = false;
            selectAll.disabled = true; // Disable if no items can be selected
        }
    }

    if (bulkActions) {
        bulkActions.style.display = selectedClaimSentKeys.size > 0 ? 'flex' : 'none';
        // Optional: Update button text with count
    }
}

// Global Set for Selected Keys
let selectedClaimSentKeys = new Set();
let currentClaimSentDisplayedData = [];

function toggleAllClaimSentCheckboxes(source) {
    if (!currentClaimSentDisplayedData || currentClaimSentDisplayedData.length === 0) return;
    const isChecked = source.checked;
    currentClaimSentDisplayedData.forEach(item => {
        // Same logic: ONLY select if status is 'ส่งเคลมแล้ว'
        const action = item['Warranty Action'] || item['ActionStatus'];
        const hasAction = action && action !== 'เคลมประกัน' && action !== 'Pending' && action !== 'บันทึกแล้ว';
        const isClaimSent = !hasAction && (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '');

        if (isClaimSent) {
            const key = getSupplierKey(item);
            if (isChecked) selectedClaimSentKeys.add(key);
            else selectedClaimSentKeys.delete(key);
        }
    });
    renderClaimSentTable();
}

function handleClaimSentCheckboxChange(checkbox) {
    if (checkbox) {
        // We need to resolve item from index relative to current page or global?
        // render renders from pageData, but value is globalIndex into sortedData (currentClaimSentDisplayedData)
        // Check line: value="${globalIndex}" where globalIndex = startIndex + index
        const globalIndex = parseInt(checkbox.value);
        const item = currentClaimSentDisplayedData[globalIndex];
        if (item) {
            const key = getSupplierKey(item);
            if (checkbox.checked) selectedClaimSentKeys.add(key);
            else selectedClaimSentKeys.delete(key);
        }
    }
    renderClaimSentTable();
}

function toggleClaimSentProductGroup(source, date, product) {
    const isChecked = source.checked;
    if (!currentClaimSentDisplayedData) return;

    currentClaimSentDisplayedData.forEach(item => {
        const itemDate = item['Claim Date'] || 'Unknown';
        const itemProduct = item['Product'] || 'Unknown';
        
        if (itemDate === date && itemProduct === product) {
             const action = item['Warranty Action'] || item['ActionStatus'];
             const hasAction = action && action !== 'เคลมประกัน' && action !== 'Pending' && action !== 'บันทึกแล้ว';
             const isClaimSent = !hasAction && (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '');
             
             if (isClaimSent) {
                 const key = getSupplierKey(item);
                 if (isChecked) selectedClaimSentKeys.add(key);
                 else selectedClaimSentKeys.delete(key);
             }
        }
    });
    renderClaimSentTable();
}

function filterClaimSentByDate(dateStr) {
    if (activeClaimDateFilter === dateStr) {
        activeClaimDateFilter = null; // Toggle off
    } else {
        activeClaimDateFilter = dateStr;
    }
    currentClaimSentPage = 1;
    renderClaimSentTable();
}

async function handleWarrantyAction(status) {
    if (selectedClaimSentKeys.size === 0) return;

    // Resolve Items
    const lookupMap = new Map();
    globalBookingData.forEach(row => lookupMap.set(getSupplierKey(row), row));

    const selectedItems = [];
    selectedClaimSentKeys.forEach(key => {
        const item = lookupMap.get(key);
        if (item) selectedItems.push(item);
    });

    if (selectedItems.length === 0) return;

    const result = await Swal.fire({
        title: `ยืนยันผลการเคลม: ${status}?`,
        html: `คุณกำลังจะบันทึกผล <b>${status}</b> สำหรับ <b>${selectedItems.length}</b> รายการ`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'บันทึก (Save)',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: status === 'รับประกัน' ? '#10b981' : '#ef4444'
    });

    if (!result.isConfirmed) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const userName = currentUser.name || currentUser.IDRec || 'Unknown';

    selectedItems.forEach(item => {
        // Robust Data Lookup (simplified here as item should be robust)
        // ... (We rely on item having necessary fields or globalBookingData syncing being enough)

        const payload = {
            ...item,
            'Warranty Action': status,
            'ActionStatus': status, // Sync standard status
            'user': userName
        };

        // Logic for Finish and Datefinish
        if (status === 'เคลมสำเร็จ' || status === 'ไม่รับประกัน') {
            // payload['Finish'] = status; // REMOVED
            payload['Datefinish'] = new Date().toLocaleString('en-GB');
        }

        // Optimistic Update
        item['Warranty Action'] = status;
        item['ActionStatus'] = status;
        // if (payload['Finish']) item['Finish'] = payload['Finish']; // REMOVED
        if (payload['Datefinish']) item['Datefinish'] = payload['Datefinish'];

        SaveQueue.add(payload);
    });

    selectedClaimSentKeys.clear();
    renderClaimSentTable();

    // Refresh others
    if (typeof renderSupplierTable === 'function') renderSupplierTable();
    if (typeof renderHistoryTable === 'function') renderHistoryTable();

    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: `Updated ${selectedItems.length} items to ${status}` });
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

    // 1. FILTER BY USER PLANT (Added)
    const userPlant = getEffectiveUserPlant();
    const userStatus = getCurrentUserStatus();
    if (userPlant) {
        allowedData = allowedData.filter(item => {
            // Allow if User Status matches Claim Receiver
            if (userStatus && item['Claim Receiver'] === userStatus) return true;

            const itemPlant = item['Plant'] || item['plant'] || item['PLANT'];
            return String(itemPlant).trim().padStart(4, '0') === userPlant;
        });

        // Debug Alert Removed as per user request
    } else {
        // Warn if no plant found (unless Admin)
        if (!isUserAdmin()) {
            Swal.fire({
                toast: true,
                position: 'top-end',
                icon: 'warning',
                title: 'No Plant Found',
                text: 'User profile has no Plant data. Showing all.'
            });
        }
    }

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
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentHistoryPage > totalPages) currentHistoryPage = 1;
    renderGenericPagination('historyPaginationControls', currentHistoryPage, totalPages, changeHistoryPage);
}

function changeHistoryPage(newPage) {
    if (newPage < 1) return;
    currentHistoryPage = newPage;
    renderHistoryTable();
}
