let currentBookingPage = 1;
let currentBookingDisplayedData = [];
let selectedBookingKeys = new Set();
let bookingReceiverOptions = new Set();

function handleBookingPartFilterChange() { selectedBookingKeys.clear(); renderBookingTable(); }

function handleBookingSearch() { currentBookingPage = 1; renderBookingTable(); }

function handleBookingStatusFilterChange() { currentBookingPage = 1; renderBookingTable(); }

function toggleAllBookingCheckboxes(source) {
    const isChecked = source.checked;
    currentBookingDisplayedData.forEach(row => {
        const hasBookingSlip = row['Booking Slip'] && String(row['Booking Slip']).trim() !== '';
        const hasRecripte = row['Recripte'] && String(row['Recripte']).trim() !== '';
        if (!hasBookingSlip && !hasRecripte) {
            const key = getBookingKey(row);
            if (isChecked) selectedBookingKeys.add(key); else selectedBookingKeys.delete(key);
        }
    });
    renderBookingTable();
}

function handleBookingCheckboxChange(checkbox) {
    if (checkbox) {
        const key = checkbox.value;
        if (checkbox.checked) selectedBookingKeys.add(key); else selectedBookingKeys.delete(key);
    }
    updateBookingActionButtons();
}

function updateBookingActionButtons() {
    const actionDiv = document.getElementById('bookingBulkActions');
    const sendNavaBtn = document.querySelector('button[onclick="sendToNavaNakorn()"]');
    const sendVibhavadiBtn = document.querySelector('button[onclick="sendToVibhavadi()"]');
    if (selectedBookingKeys.size > 0) actionDiv.style.display = 'flex'; else { actionDiv.style.display = 'none'; return; }

    let hasPoom = false; let hasNonPoom = false;
    const lookupMap = new Map();
    globalBookingData.forEach(row => lookupMap.set(getBookingKey(row), row));
    selectedBookingKeys.forEach(key => {
        const row = lookupMap.get(key);
        if (row) {
            const receiver = (row['Claim Receiver'] || row.person || '').toLowerCase();
            if (receiver.includes('poom')) hasPoom = true; else hasNonPoom = true;
        }
    });

    if (sendNavaBtn) {
        if (hasPoom) { sendNavaBtn.disabled = true; sendNavaBtn.style.opacity = '0.5'; sendNavaBtn.style.cursor = 'not-allowed'; sendNavaBtn.title = 'Items assigned to "Poom" cannot be sent to Nava Nakorn.'; }
        else { sendNavaBtn.disabled = false; sendNavaBtn.style.opacity = '1'; sendNavaBtn.style.cursor = 'pointer'; sendNavaBtn.title = ''; }
    }
    if (sendVibhavadiBtn) {
        if (hasNonPoom) { sendVibhavadiBtn.disabled = true; sendVibhavadiBtn.style.opacity = '0.5'; sendVibhavadiBtn.style.cursor = 'not-allowed'; sendVibhavadiBtn.title = 'Only items assigned to "Poom" can be sent to Vibhavadi.'; }
        else { sendVibhavadiBtn.disabled = false; sendVibhavadiBtn.style.opacity = '1'; sendVibhavadiBtn.style.cursor = 'pointer'; sendVibhavadiBtn.title = ''; }
    }
}

async function processBookingAction(destination, targetPlantCode) {
    if (selectedBookingKeys.size === 0) return;
    const selectedItems = [];
    const lookupMap = new Map();
    globalBookingData.forEach(row => lookupMap.set(getBookingKey(row), row));
    selectedBookingKeys.forEach(key => { const item = lookupMap.get(key); if (item) selectedItems.push(item); });
    if (selectedItems.length === 0) return;

    if (destination.includes('นวนคร')) {
        const hasPoom = selectedItems.some(item => { const receiver = (item['Claim Receiver'] || item.person || '').toLowerCase(); return receiver.includes('poom'); });
        if (hasPoom) { Swal.fire({ icon: 'warning', title: 'ไม่อนุญาต', text: 'รายการที่มี Receiver เป็น "Poom" ไม่สามารถส่งไป "นวนคร" ได้', confirmButtonText: 'ตกลง' }); return; }
    }

    let plantCenterCode = targetPlantCode || '';
    if (!plantCenterCode) { if (destination.includes('นวนคร')) plantCenterCode = '0301'; if (destination.includes('วิภาวดี')) plantCenterCode = '0326'; }

    const { value: bookingSlip, isConfirmed } = await Swal.fire({
        title: 'กรอกเลขใบจองรถ', html: `คุณกำลังจะส่ง ${selectedItems.length} รายการ ไปยัง <b>${destination}</b><br>(รหัส: <span style="color: var(--primary-color); font-weight: bold;">${plantCenterCode}</span>)`,
        input: 'text', inputPlaceholder: 'เลขใบจองรถ', showCancelButton: true, confirmButtonText: 'บันทึก', cancelButtonText: 'ยกเลิก', preConfirm: (value) => { if (!value) Swal.showValidationMessage('กรุณากรอกเลขใบจองรถ!'); return value; }
    });

    if (!isConfirmed) return;

    const bookingDate = new Date();
    const day = String(bookingDate.getDate()).padStart(2, '0');
    const month = String(bookingDate.getMonth() + 1).padStart(2, '0');
    const year = bookingDate.getFullYear();
    const formattedDate = `${day}/${month}/${year}`;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

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

        const payload = {
            ...item,
            'Booking Slip': bookingSlip,
            'Booking Date': formattedDate,
            'Plantcenter': plantCenterCode,
            'user': currentUser.IDRec || 'Unknown',
            'วันที่รับซาก': dateReceived || '',
            'ผู้รับซาก': receiver || '',
            'Keep': keep || '',
            'CI Name': ciName || '',
            'Problem': problem || '',
            'Product Type': productType || ''
        };
        payload['Key'] = (item['Work Order'] || '') + (item['Spare Part Code'] || '');

        // Optimistic Update
        item['Booking Slip'] = bookingSlip;
        item['Booking Date'] = formattedDate;
        item['Plantcenter'] = plantCenterCode;
        SaveQueue.add(payload);
    });

    selectedBookingKeys.clear();
    renderBookingTable();
    setTimeout(() => {
        const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true, didOpen: (toast) => { toast.style.zIndex = '20001'; } });
        Toast.fire({ icon: 'success', title: `Saved ${selectedItems.length} items successfully` });
    }, 500);
}

async function sendToNavaNakorn() { await processBookingAction('ส่งคลังนวนคร', '0301'); }
async function sendToVibhavadi() { await processBookingAction('ส่งคลังวิภาวดี', '0326'); }

function toggleReceiverDropdown() {
    const dropdown = document.getElementById('bookingReceiverDropdown');
    dropdown.style.display = dropdown.style.display === 'none' ? 'block' : 'none';
}

function handleReceiverCheckboxChange() {
    selectedBookingKeys.clear();
    bookingReceiverOptions.clear();
    const checkboxes = document.querySelectorAll('.receiver-checkbox:checked');
    checkboxes.forEach(cb => bookingReceiverOptions.add(cb.value));
    const label = document.getElementById('receiverFilterLabel');
    label.textContent = bookingReceiverOptions.size === 0 ? 'All Receivers' : `${bookingReceiverOptions.size} Selected`;
    renderBookingTable();
}

function populateBookingReceiverFilter() {
    const dropdown = document.getElementById('bookingReceiverDropdown');
    dropdown.innerHTML = '';
    if (!globalBookingData || globalBookingData.length === 0) return;
    const receivers = new Set();
    globalBookingData.forEach(row => { const p = row['person'] || row['Claim Receiver']; if (p) receivers.add(String(p).trim()); });
    const sortedReceivers = Array.from(receivers).sort();
    if (sortedReceivers.length === 0) { dropdown.innerHTML = '<div style="padding:0.5rem; color:#64748b; font-size:0.875rem;">No receivers found</div>'; return; }
    sortedReceivers.forEach(receiver => {
        const div = document.createElement('div');
        div.style.padding = '0.25rem 0'; div.style.display = 'flex'; div.style.alignItems = 'center';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox'; checkbox.value = receiver; checkbox.className = 'receiver-checkbox'; checkbox.style.marginRight = '0.5rem';
        if (bookingReceiverOptions.has(receiver)) checkbox.checked = true;
        checkbox.onchange = handleReceiverCheckboxChange;
        const label = document.createElement('label');
        label.textContent = receiver; label.style.fontSize = '0.875rem'; label.style.cursor = 'pointer'; label.onclick = () => checkbox.click();
        div.appendChild(checkbox); div.appendChild(label); dropdown.appendChild(div);
    });
}

window.addEventListener('click', function (e) {
    const dropdown = document.getElementById('bookingReceiverDropdown');
    const button = document.querySelector('button[onclick="toggleReceiverDropdown()"]');
    if (dropdown && button && !dropdown.contains(e.target) && !button.contains(e.target)) dropdown.style.display = 'none';
});

function populateBookingFilter() {
    const filterSelect = document.getElementById('bookingPartFilter');
    if (filterSelect) {
        filterSelect.innerHTML = '<option value="">All Spare Parts</option>';
        if (!globalBookingData || globalBookingData.length === 0) return;
        const parts = new Set();
        globalBookingData.forEach(row => { const name = row['Spare Part Name']; if (name) parts.add(name); });
        Array.from(parts).sort().forEach(part => {
            const option = document.createElement('option'); option.value = part; option.textContent = part; filterSelect.appendChild(option);
        });
    }
    populateBookingReceiverFilter();
}

function renderBookingTable() {
    const tableHeader = document.getElementById('bookingTableHeader');
    const tableBody = document.getElementById('bookingTableBody');
    const filterSelect = document.getElementById('bookingPartFilter');
    const filterValue = filterSelect ? filterSelect.value : '';
    const searchInput = document.getElementById('bookingSearchInput');
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const statusFilter = document.getElementById('bookingStatusFilter');
    const statusFilterValue = statusFilter ? statusFilter.value : '';

    // Render Headers FIRST (Always show headers)
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

    if (!globalBookingData || globalBookingData.length === 0) {
        currentBookingDisplayedData = [];
        tableBody.innerHTML = '<tr><td colspan="' + BOOKING_COLUMNS.length + '" style="text-align:center; padding: 2rem;">No data found in Warranty Sheet.</td></tr>';
        return;
    }

    let filteredData = globalBookingData;
    const userPlant = (typeof getEffectiveUserPlant === 'function') ? getEffectiveUserPlant() : null;
    const userStatus = getCurrentUserStatus(); // Get user status

    if (userPlant) {
        filteredData = filteredData.filter(row => {
            // Allow if User Status matches Claim Receiver
            if (userStatus && row['Claim Receiver'] === userStatus) return true;

            const rowPlant = normalizePlantCode(row['Plant']);
            return rowPlant === userPlant;
        });
    }

    filteredData = filteredData.filter(item => {
        const hasClaimSup = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';
        const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
        if (hasClaimSup) return false;
        if (hasRecripte) return false;
        
        // Filter only 'เคลมประกัน' for Booking Tab
        const action = item['ActionStatus'] || item['Warranty Action'] || '';
        if (action !== 'เคลมประกัน') return false;
        
        return true;
    });

    if (filterValue) filteredData = filteredData.filter(row => row['Spare Part Name'] === filterValue);
    if (bookingReceiverOptions.size > 0) filteredData = filteredData.filter(row => { const val = row['Claim Receiver'] || row['person'] || ''; return bookingReceiverOptions.has(String(val).trim()); });
    if (searchTerm) {
        filteredData = filteredData.filter(row => {
            return Object.values(row).some(val => String(val || '').toLowerCase().includes(searchTerm));
        });
    }
    if (statusFilterValue) {
        filteredData = filteredData.filter(item => {
            const hasBookingSlip = item['Booking Slip'] && String(item['Booking Slip']).trim() !== '';
            // Note: ClaimSup and Recripte are already filtered out above
            const status = hasBookingSlip ? 'ระหว่างขนส่ง' : 'คลังพื้นที่';

            if (statusFilterValue === 'TodayOrLocal') {
                if (status === 'คลังพื้นที่') return true;
                const bDate = String(item['Booking Date'] || '').trim();
                if (!bDate) return false;
                const now = new Date();
                const day = String(now.getDate()).padStart(2, '0');
                const month = String(now.getMonth() + 1).padStart(2, '0');
                const year = now.getFullYear();
                return bDate === `${day}/${month}/${year}`;
            }

            return status === statusFilterValue;
        });
    }

    const sortedData = [...filteredData].reverse().sort((a, b) => { const nameA = a['Spare Part Name'] || ''; const nameB = b['Spare Part Name'] || ''; return nameA.localeCompare(nameB, 'th'); });
    currentBookingDisplayedData = sortedData;

    const groupTotals = sortedData.reduce((acc, item) => { const name = item['Spare Part Name'] || 'Unknown'; const qty = parseFloat(item['Qty']) || 0; acc[name] = (acc[name] || 0) + qty; return acc; }, {});

    const selectAllCheckbox = document.getElementById('selectAllBooking');
    if (selectAllCheckbox) {
        let allEnabledSelected = true; let anySelected = false; let hasOneEnabled = false;
        currentBookingDisplayedData.forEach(row => {
            const hasBookingSlip = row['Booking Slip'] && String(row['Booking Slip']).trim() !== '';
            const hasRecripte = row['Recripte'] && String(row['Recripte']).trim() !== '';
            if (!hasBookingSlip && !hasRecripte) {
                hasOneEnabled = true;
                const key = getBookingKey(row);
                if (!selectedBookingKeys.has(key)) allEnabledSelected = false; else anySelected = true;
            }
        });
        if (hasOneEnabled && allEnabledSelected) { selectAllCheckbox.checked = true; selectAllCheckbox.indeterminate = false; }
        else if (anySelected) { selectAllCheckbox.checked = false; selectAllCheckbox.indeterminate = true; }
        else { selectAllCheckbox.checked = false; selectAllCheckbox.indeterminate = false; }
    }

    // Handle No Matching Records
    if (sortedData.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="' + BOOKING_COLUMNS.length + '" style="text-align:center; padding: 2rem;">No matching records found.</td></tr>';
        renderGenericPagination('bookingPaginationControls', 1, 0, changeBookingPage);
        updateBookingActionButtons();
        if (selectAllCheckbox) selectAllCheckbox.disabled = true;
        return;
    } else if (selectAllCheckbox) {
        selectAllCheckbox.disabled = false;
    }

    const startIndex = (currentBookingPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = sortedData.slice(startIndex, endIndex);
    tableBody.innerHTML = '';

    let previousPartName = null;
    if (currentBookingPage > 1 && sortedData[startIndex - 1]) {
        previousPartName = sortedData[startIndex - 1]['Spare Part Name'];
    }

    pageData.forEach((row, index) => {
        const currentPartName = row['Spare Part Name'] || 'Unknown';
        const currentPartCode = row['Spare Part Code'] || '';

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
                const globalIndex = startIndex + index;
                const hasBookingSlip = row['Booking Slip'] && String(row['Booking Slip']).trim() !== '';
                const hasRecripte = row['Recripte'] && String(row['Recripte']).trim() !== '';
                const disabledAttr = (hasBookingSlip || hasRecripte) ? 'disabled' : '';
                const key = getBookingKey(row);
                const isChecked = selectedBookingKeys.has(key) ? 'checked' : '';

                td.innerHTML = `<input type="checkbox" class="booking-checkbox" value="${key}" onchange="handleBookingCheckboxChange(this)" ${disabledAttr} ${isChecked}>`;
                td.style.textAlign = 'center';
                tr.appendChild(td);
                return;
            }

            let value = row[col.key] || '';

            if (col.key === 'CustomStatus') {
                const statusHtml = getComputedStatus(row);
                // Create a clickable wrapper
                const wrapper = document.createElement('div');
                wrapper.innerHTML = statusHtml;
                wrapper.style.cursor = 'pointer';
                wrapper.onclick = (e) => {
                    e.stopPropagation(); // Prevent row click if any
                    handleBookingStatusClick(row);
                };
                td.appendChild(wrapper);
                tr.appendChild(td);
                return;
            }

            if (col.key === 'Timestamp' && value) {
                try {
                    const date = new Date(value);
                    if (!isNaN(date)) {
                        value = date.toLocaleString('th-TH');
                    }
                } catch (e) { }
            }

            if (col.key === 'Booking Date' && value) {
                let s = String(value);
                if (s.indexOf('T') > -1) s = s.split('T')[0];
                if (s.indexOf(' ') > -1) s = s.split(' ')[0];
                value = s;
            }

            if (col.key === 'Plant' && value) {
                value = String(value).replace(/^0+/, '');
            }

            if (col.key === 'Plantcenter' && value) {
                const v = String(value).replace(/^0+/, '');
                if (v === '301') value = 'คลังนวนคร';
                else if (v === '326') value = 'คลังวิภาวดี';
            }

            if (col.key === 'Mobile') {
                if (value) {
                    value = formatPhoneNumber(value);
                } else {
                    td.style.cursor = 'pointer';
                    td.style.color = '#3b82f6';
                    td.innerText = '[Add]';
                    td.style.fontSize = '0.8em';
                    td.onclick = () => openMobileModal(row);
                }
            }

            if (col.key === 'Booking Slip' && value && String(value).trim() !== '') {
                td.style.cursor = 'pointer';
                td.style.color = 'var(--primary-color)';
                td.style.fontWeight = '500';
                td.title = 'Click to Edit/Delete';
                td.onclick = () => openBookingSlipModal(row);
            }

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

                    const valLower = String(value).toLowerCase().trim();
                    if (valLower === 'mai') {
                        badge.style.backgroundColor = '#dbeafe';
                        badge.style.color = '#1e40af';
                    } else if (valLower === 'mon') {
                        badge.style.backgroundColor = '#dcfce7';
                        badge.style.color = '#166534';
                    } else if (valLower === 'poom') {
                        badge.style.backgroundColor = '#fae8ff';
                        badge.style.color = '#6b21a8';
                    } else {
                        badge.style.backgroundColor = '#f1f5f9';
                        badge.style.color = '#475569';
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

    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentBookingPage > totalPages) currentBookingPage = 1;
    renderGenericPagination('bookingPaginationControls', currentBookingPage, totalPages, changeBookingPage);
    updateBookingActionButtons();
}

function changeBookingPage(newPage) {
    currentBookingPage = newPage;
    renderBookingTable();
    const tableContainer = document.querySelector('#tab-content-booking .table-container');
    if (tableContainer) tableContainer.scrollTop = 0;
}

function handleBookingStatusClick(row) {
    if (!row) return;

    // Use work order and spare part code to find the full data object
    const workOrder = row['Work Order'] || '';
    const partCode = row['Spare Part Code'] || '';
    const key = (workOrder + partCode).replace(/\s/g, '').toLowerCase();

    // Try to find the matching item in fullData (which has the correct structure for openStoreModal)
    let match = null;
    if (typeof fullData !== 'undefined') {
        match = fullData.find(item => {
            const itemWO = item.scrap['work order'] || '';
            const itemCode = item.scrap['Spare Part Code'] || '';
            const itemKey = (itemWO + itemCode).replace(/\s/g, '').toLowerCase();
            return itemKey === key;
        });
    }

    if (match) {
        // Attach booking info to the matched item for the modal to use
        match.bookingSlip = row['Booking Slip'];
        match.bookingDate = row['Booking Date'];
        openStoreModal(match);
    } else {
        // Fallback: Create a temporary object structure compatible with openStoreModal
        // This handles cases where data might not be fully synced or in fullData yet
        const tempItem = {
            scrap: {
                'work order': row['Work Order'],
                'Spare Part Code': row['Spare Part Code'],
                'Spare Part Name': row['Spare Part Name'],
                'qty': row['Qty'],
                'รหัสช่าง': '', // Might be missing in booking view
                'ชื่อช่าง': '',
                'Keep': row['Keep'],
                'วันที่รับซาก': row['วันที่รับซาก'],
                'ผู้รับซาก': row['ผู้รับซาก']
            },
            fullRow: {
                'Serial Number': row['Serial Number'],
                'Store Code': row['Store Code'],
                'Store Name': row['Store Name'],
                'CI Name': row['CI Name'],
                'Product Type': row['Product Type'],
                'Problem': row['Problem'],
                'Product': row['Product'],
                'ActionStatus': row['ActionStatus'] || row['Warranty Action'] || 'เคลมประกัน',
                'Claim Receiver': row['Claim Receiver']
            },
            status: row['ActionStatus'] || row['Warranty Action'] || 'เคลมประกัน',
            technicianPhone: row['Mobile'] || '',
            person: row['Claim Receiver'] || row['person'],
            bookingSlip: row['Booking Slip'],
            bookingDate: row['Booking Date']
        };
        openStoreModal(tempItem);
    }
}