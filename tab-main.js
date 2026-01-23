function populateFilters() {
    const filterSelect = document.getElementById('sparePartFilter');
    if (!filterSelect) return;
    const uniqueParts = [...new Set(fullData.map(item => item.scrap['Spare Part Name']).filter(Boolean))].sort();
    const currentSelection = filterSelect.value;
    filterSelect.innerHTML = '<option value="">All Spare Parts</option>';
    uniqueParts.forEach(part => {
        const option = document.createElement('option');
        option.value = part; option.textContent = part;
        filterSelect.appendChild(option);
    });
    if (uniqueParts.includes(currentSelection)) filterSelect.value = currentSelection;
}

function handleFilterChange() { currentPage = 1; applyFilters(); }
function handleSearch() { currentPage = 1; applyFilters(); }

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
    const filterSelect = document.getElementById('sparePartFilter');
    const selectedPart = filterSelect ? filterSelect.value : '';
    const selectedStatus = document.getElementById('statusFilter').value;
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();

    displayedData = fullData.filter(item => {
        const matchPart = selectedPart ? item.scrap['Spare Part Name'] === selectedPart : true;
        let matchStatus = true;
        if (selectedStatus === 'Pending') matchStatus = !item.status;
        else if (selectedStatus) matchStatus = item.status === selectedStatus;
        let matchSearch = true;
        if (searchTerm) {
            const allValues = [...Object.values(item.scrap || {}), ...Object.values(item.fullRow || {}), item.status || '', item.technicianPhone || '', item.person || ''].join(' ').toLowerCase();
            matchSearch = allValues.includes(searchTerm);
        }
        return matchPart && matchStatus && matchSearch;
    });
    sortDisplayedData();
    updateTotalQty();
    renderTable();
}

function updateTotalQty() {
    const total = displayedData.reduce((sum, item) => sum + (parseFloat(item.scrap['qty']) || 0), 0);
    document.getElementById('totalQtyValue').textContent = total.toLocaleString() + ' PC';
}

function renderTable() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';
    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE;
    const endIndex = startIndex + ITEMS_PER_PAGE;
    const pageData = displayedData.slice(startIndex, endIndex);

    const groupStats = displayedData.reduce((acc, item) => {
        const code = item.scrap['Spare Part Code'] || 'Unknown';
        const qty = parseFloat(item.scrap['qty']) || 0;
        if (!acc[code]) acc[code] = { total: 0, completed: 0 };
        acc[code].total += qty;
        if (item.status && String(item.status).trim() !== '') acc[code].completed += qty;
        return acc;
    }, {});

    let previousPartCode = null;
    if (currentPage > 1 && displayedData[startIndex - 1]) previousPartCode = displayedData[startIndex - 1].scrap['Spare Part Code'];

    pageData.forEach((item, index) => {
        const currentPartName = item.scrap['Spare Part Name'] || 'Unknown';
        const currentPartCode = item.scrap['Spare Part Code'] || 'Unknown';

        if (currentPartCode !== previousPartCode) {
            const stats = groupStats[currentPartCode] || { total: 0, completed: 0 };
            const headerRow = document.createElement('tr');
            headerRow.className = 'group-header-row';
            headerRow.style.backgroundColor = '#f8fafc'; headerRow.style.fontWeight = 'bold';
            headerRow.innerHTML = `<td colspan="${COLUMNS.length}" style="padding: 12px; border-top: 2px solid #e2e8f0; color: #334155;"><div style="display: flex; justify-content: space-between; align-items: center;"><div style="display:flex; align-items:center; gap:0.5rem;"><span>📦 ${currentPartName}</span><span style="font-size:0.85em; color:#64748b; font-weight:normal;">(${currentPartCode})</span><span style="font-size:0.85em; color:#0369a1; font-weight:bold;">(${stats.completed.toLocaleString()}/${stats.total.toLocaleString()} Pc)</span></div></div></td>`;
            tableBody.appendChild(headerRow);
            previousPartCode = currentPartCode;
        }

        const tr = document.createElement('tr');
        COLUMNS.forEach(col => {
            const td = document.createElement('td');
            let value = '';
            if (col.source === 'S') value = item.scrap[col.key] || '';
            else if (col.source === 'C') value = item.status;
            else if (col.source === 'T') value = item.technicianPhone;
            else if (col.source === 'P') value = item.person;
            else value = item.fullRow[col.key] || '';

            if (col.key === 'checkbox') {
                const isSaved = !!item.status;
                const storeCode = item.fullRow['Store Code'] || '';
                const isMissingStore = !storeCode || String(storeCode).trim() === '';
                const disabledAttr = (isSaved || isMissingStore) ? 'disabled' : '';
                td.innerHTML = `<input type="checkbox" class="row-checkbox" value="${startIndex + index}" onchange="handleCheckboxChange()" ${disabledAttr}>`;
                td.style.textAlign = 'center';
                tr.appendChild(td);
                return;
            }

            const greenColumns = ['status', 'work order', 'Spare Part Code', 'Spare Part Name', 'old material code', 'qty', 'Serial Number', 'Store Code', 'Store Name'];
            const storeCode = item.fullRow['Store Code'] || '';
            const isMissingStore = !storeCode || String(storeCode).trim() === '';
            const isLE = item.fullRow['Product'] === 'L&E';
            const hasMobile = item.technicianPhone && String(item.technicianPhone).trim() !== '';
            const isReadyForClaim = !item.status && !isSerialInvalidForClaim(item) && (isLE || hasMobile);
            const readyColumns = ['work order', 'Spare Part Code', 'Spare Part Name'];

            if (item.status && greenColumns.includes(col.key)) {
                const span = document.createElement('span');
                span.textContent = value;
                let bgColor = '#dcfce7'; let textColor = '#166534';
                if (item.status === 'Sworp') { bgColor = '#e0e7ff'; textColor = '#3730a3'; }
                else if (item.status === 'หมดประกัน') { bgColor = '#ffedd5'; textColor = '#9a3412'; }
                else if (item.status === 'ชำรุด') { bgColor = '#fce7f3'; textColor = '#9d174d'; }
                span.style.backgroundColor = bgColor; span.style.color = textColor; span.style.fontSize = '0.875rem'; span.style.fontWeight = '600'; span.style.padding = '0.25rem 0.5rem'; span.style.borderRadius = '4px'; span.style.display = 'inline-block';
                td.appendChild(span);
            } else if (isReadyForClaim && readyColumns.includes(col.key)) {
                const span = document.createElement('span');
                span.textContent = value;
                span.style.backgroundColor = '#dcfce7'; span.style.color = '#166534'; span.style.fontSize = '0.875rem'; span.style.fontWeight = '600'; span.style.padding = '0.25rem 0.5rem'; span.style.borderRadius = '4px'; span.style.display = 'inline-block';
                td.appendChild(span);
            } else if (!item.status && isMissingStore && col.key === 'status') {
                const span = document.createElement('span');
                span.textContent = 'Wait Data';
                span.style.backgroundColor = '#f1f5f9'; span.style.color = '#ef4444'; span.style.fontSize = '0.875rem'; span.style.fontWeight = '600'; span.style.padding = '0.25rem 0.5rem'; span.style.borderRadius = '4px'; span.style.display = 'inline-block';
                td.appendChild(span);
            } else if (!item.status && col.key === 'status') {
                const span = document.createElement('span');
                span.textContent = 'Pending';
                span.style.backgroundColor = '#fff7ed'; span.style.color = '#ea580c'; span.style.fontSize = '0.875rem'; span.style.fontWeight = '600'; span.style.padding = '0.25rem 0.5rem'; span.style.borderRadius = '4px'; span.style.display = 'inline-block';
                td.appendChild(span);
            } else {
                td.textContent = value;
                td.style.fontSize = '0.875rem';
            }

            if (col.key === 'Serial Number') { 
                td.style.cursor = 'pointer'; 
                td.style.color = 'var(--primary-color)'; 
                td.style.fontWeight = '500'; 
                td.title = 'Click to edit Serial Number'; 
                td.onclick = () => openEditModal(item); 

                // Validation for Claim Insurance
                if (isSerialInvalidForClaim(item)) {
                    td.classList.add('invalid-serial');
                    td.title = 'Invalid Serial Number for Claim Insurance! Click to fix.';
                }
            }
            if (col.key === 'status') { td.style.cursor = 'pointer'; td.title = 'Click to view details'; td.onclick = () => openStoreModal(item); }
            if (col.key === 'Person' || col.source === 'P') { td.style.cursor = 'pointer'; td.style.color = 'var(--primary-color)'; td.style.fontWeight = '500'; td.title = 'Click to edit Person'; td.onclick = () => openMainPersonModal(item); }
            if (col.key === 'Mobile' || col.source === 'T' || col.key === 'Phone') {
                td.style.cursor = 'pointer'; td.style.color = 'var(--primary-color)'; td.style.fontWeight = '500'; td.title = 'Click to edit Mobile'; td.onclick = () => openMainMobileModal(item);
                if (!value) { 
                    td.innerText = '[Add]'; 
                    td.style.fontSize = '0.875rem'; 
                    if (!item.status) {
                        if (!isLE) {
                            td.classList.add('invalid-serial');
                            td.title = 'Mobile Number is required! Click to add.';
                        }
                    }
                }
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    renderPagination();
}

function isSerialInvalidForClaim(item) {
    // Skip validation if already saved/processed
    if (item.status) return false;

    const serial = item.fullRow['Serial Number'] || '';
    const isLE = item.fullRow['Product'] === 'L&E';
    const minLength = isLE ? 6 : 8;

    if (!serial || serial.length < minLength) return true;
    if (/[#$%\u0E00-\u0E7F\s]/.test(serial)) return true;
    const keep = item.scrap['Keep'] || '';
    if (keep === 'SCOTMAN' && !String(serial).startsWith('IDP020') && !String(serial).startsWith('NW508')) return true;
    const product = item.fullRow['Product'] || '';
    if (product === 'SCOTMAN' && !String(serial).startsWith('IDP020') && !String(serial).startsWith('NW508')) return true;

    return false;
}

function renderPagination() {
    const totalPages = Math.ceil(displayedData.length / ITEMS_PER_PAGE);
    renderGenericPagination('paginationControls', currentPage, totalPages, changePage);
}

function changePage(newPage) {
    currentPage = newPage;
    renderTable();
    document.querySelector('.table-container').scrollTop = 0;
}

function toggleAllCheckboxes(source) {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach(checkbox => { checkbox.checked = source.checked; });
    handleCheckboxChange();
}

function handleCheckboxChange() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    const bulkActions = document.getElementById('bulkActions');
    if (checkboxes.length > 0) bulkActions.style.display = 'flex';
    else bulkActions.style.display = 'none';
}

async function processBulkAction(actionName) {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    const selectedItems = [];
    for (const cb of checkboxes) {
        const globalIndex = parseInt(cb.value);
        const item = displayedData[globalIndex];
        if (item) {
            if (item.status === 'บันทึกแล้ว') { Swal.fire({ icon: 'warning', title: 'บางรายการถูกบันทึกแล้ว', text: `รายการ ${item.scrap['work order']} ถูกบันทึกไปแล้ว ไม่สามารถแก้ไขได้` }); return; }
            const serial = item.fullRow['Serial Number'] || '';
            const isLE = item.fullRow['Product'] === 'L&E';
            const minLength = isLE ? 6 : 8;
            if (actionName === 'เคลมประกัน' && (!serial || serial.length < minLength)) { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `ข้อมูล Serial Number ไม่ถูกต้อง โปรดดำเนินการ` }); return; }
            if (actionName === 'เคลมประกัน' && /[#$%\u0E00-\u0E7F\s]/.test(serial)) { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `รายการ ${item.scrap['work order']}: Serial Number ห้ามมีอักขระพิเศษ, ภาษาไทย หรือช่องว่าง` }); return; }
            const keep = item.scrap['Keep'] || '';
            if (actionName === 'เคลมประกัน' && keep === 'SCOTMAN' && !String(serial).startsWith('IDP020') && !String(serial).startsWith('NW508')) { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `รายการ ${item.scrap['work order']}: Keep เป็น SCOTMAN Serial Number ต้องขึ้นต้นด้วย IDP020 หรือ NW508` }); return; }
            const product = item.fullRow['Product'] || '';
            if (actionName === 'เคลมประกัน' && product === 'SCOTMAN' && !String(serial).startsWith('IDP020') && !String(serial).startsWith('NW508')) { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `รายการ ${item.scrap['work order']}: Product เป็น SCOTMAN Serial Number ต้องขึ้นต้นด้วย IDP020 หรือ NW508` }); return; }
            if (actionName === 'เคลมประกัน') {
                const mobile = item.technicianPhone || '';
                const isLE = item.fullRow['Product'] === 'L&E';
                if (!isLE && (!mobile || String(mobile).trim() === '')) { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ครบถ้วน', text: `รายการ ${item.scrap['work order']}: กรุณาระบุเบอร์โทรศัพท์ (Mobile) ก่อนบันทึก` }); return; }
            }
            selectedItems.push(item);
        }
    }
    // ... (rest of bulk action logic similar to original)
    // Note: Full implementation would be lengthy, assuming user has it or I should provide full.
    // I will provide the rest of the function in next block if needed, but for brevity I'll stop here as the pattern is clear.
    // Wait, user asked for "remaining files". I should provide full content.
    
    const result = await Swal.fire({ title: 'ยืนยันการบันทึก?', text: `คุณต้องการบันทึก ${selectedItems.length} รายการ ด้วยสถานะ "${actionName}" ใช่หรือไม่?`, icon: 'question', showCancelButton: true, confirmButtonText: 'ยืนยัน', cancelButtonText: 'ยกเลิก' });
    if (!result.isConfirmed) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

    // --- OPTIMISTIC UPDATE & QUEUE ---
    selectedItems.forEach(item => {
        const getVal = (obj, keyName) => {
             if (!obj) return '';
             if (obj[keyName]) return obj[keyName];
             const cleanKey = keyName.replace(/\s+/g, '').toLowerCase();
             const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey);
             return found ? obj[found] : '';
        };

        const getRobustVal = (itm, keyName) => {
            let val = getVal(itm.scrap, keyName);
            if (val) return val;
            if (typeof fullData !== 'undefined') {
                const wo = getVal(itm.scrap, 'work order') || getVal(itm.fullRow, 'Work Order');
                const code = getVal(itm.scrap, 'Spare Part Code') || getVal(itm.fullRow, 'Spare Part Code');
                const targetKey = (wo + code).replace(/\s/g, '').toLowerCase();
                const match = fullData.find(d => {
                    const dWO = d.scrap['work order'] || '';
                    const dCode = d.scrap['Spare Part Code'] || '';
                    return (dWO + dCode).replace(/\s/g, '').toLowerCase() === targetKey;
                });
                if (match) return getVal(match.scrap, keyName);
            }
            return '';
        };

        const payload = { ...item.scrap, ...item.fullRow, 'user': currentUser.IDRec || 'Unknown', 'ActionStatus': actionName, 'Qty': item.scrap['qty'] || '', 'Serial Number': item.fullRow['Serial Number'] || '', 'Phone': item.technicianPhone || '', 'รหัสช่าง': item.scrap['รหัสช่าง'] || '', 'ชื่อช่าง': item.scrap['ชื่อช่าง'] || '', 'Claim Receiver': (actionName === 'เคลมประกัน') ? (item.person || '') : '',
            'วันที่รับซาก': getRobustVal(item, 'วันที่รับซาก'),
            'ผู้รับซาก': getRobustVal(item, 'ผู้รับซาก'),
            'Keep': getRobustVal(item, 'Keep'),
            'CI Name': item.fullRow['CI Name'] || getVal(item.fullRow, 'CI Name'),
            'Problem': item.fullRow['Problem'] || getVal(item.fullRow, 'Problem'),
            'Product Type': item.fullRow['Product Type'] || getVal(item.fullRow, 'Product Type')
        };
        
        // Explicitly set Key to avoid Unknown in logs/backend
        payload['Key'] = (getVal(item.scrap, 'work order') || '') + (getVal(item.scrap, 'Spare Part Code') || '');
        
        delete payload['Values'];
        
        // Update Local State Immediately
        item.status = actionName;
        item.fullRow['ActionStatus'] = actionName;
        item.fullRow['user'] = payload['user'];

        // Add to Background Queue
        SaveQueue.add(payload);
    });

    selectedItems.forEach(item => {
        const itemPayload = { ...item.scrap, ...item.fullRow, 'Qty': item.scrap['qty'], 'Phone': item.technicianPhone };
        const targetKey = (itemPayload['work order'] || '') + (itemPayload['Spare Part Code'] || '');
        
        // Preserve existing data if available
        const existingRow = globalBookingData.find(row => { const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || ''); return rowKey === targetKey; });
        const preservedSlip = existingRow ? (existingRow['Booking Slip'] || '') : '';
        const preservedDate = existingRow ? (existingRow['Booking Date'] || '') : '';
        const preservedPlantCenter = existingRow ? (existingRow['Plantcenter'] || '') : '';
        const preservedReceiver = existingRow ? (existingRow['Claim Receiver'] || '') : '';

        globalBookingData = globalBookingData.filter(row => { const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || ''); return rowKey !== targetKey; });
        if (actionName === 'เคลมประกัน') {
            const getVal = (obj, keyName) => {
                 if (obj[keyName]) return obj[keyName];
                 const found = Object.keys(obj).find(k => k.toLowerCase().trim() === keyName.toLowerCase());
                 return found ? obj[found] : '';
            };

            const newBookingRow = { 'Work Order': itemPayload['work order'], 'Spare Part Code': itemPayload['Spare Part Code'], 'Spare Part Name': itemPayload['Spare Part Name'], 'Old Material Code': itemPayload['old material code'], 'Qty': itemPayload['Qty'], 'Serial Number': itemPayload['Serial Number'], 'Store Code': itemPayload['Store Code'], 'Store Name': itemPayload['Store Name'], 'รหัสช่าง': itemPayload['รหัสช่าง'], 'ชื่อช่าง': itemPayload['ชื่อช่าง'], 'Mobile': itemPayload['Phone'], 'Plant': itemPayload['plant'], 'Claim Receiver': preservedReceiver || (item.person || ''), 'person': item.person, 'Product': itemPayload['Product'], 'Warranty Action': actionName, 'Recorder': currentUser.IDRec || 'Unknown', 'Timestamp': new Date().toISOString(), 'Booking Slip': preservedSlip, 'Booking Date': preservedDate, 'Plantcenter': preservedPlantCenter,
                'วันที่รับซาก': getVal(item.scrap, 'วันที่รับซาก'),
                'ผู้รับซาก': getVal(item.scrap, 'ผู้รับซาก'),
                'Keep': getVal(item.scrap, 'Keep'),
                'CI Name': item.fullRow['CI Name'] || '',
                'Problem': item.fullRow['Problem'] || '',
                'Product Type': item.fullRow['Product Type'] || ''
            };
            globalBookingData.push(newBookingRow);
        }
    });

    populateBookingFilter(); renderBookingTable();
    try { renderDeckView('0301', 'navaNakornDeck', 'navanakorn'); renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi'); populateSupplierFilter(); renderSupplierTable(); populateClaimSentFilter(); renderClaimSentTable(); renderHistoryTable(); } catch (err) { console.warn('Error refreshing other tabs:', err); }
    renderTable();
    toggleAllCheckboxes({ checked: false });
    
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: `Added ${selectedItems.length} items to save queue` });
}

function sendDatatoGAS(item) {
    if (!GAS_API_URL) { alert("Please configure the Google Apps Script URL."); return; }
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    
    const getVal = (obj, keyName) => {
         if (!obj) return '';
         if (obj[keyName]) return obj[keyName];
         const cleanKey = keyName.replace(/\s+/g, '').toLowerCase();
         const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey);
         return found ? obj[found] : '';
    };

    const getRobustVal = (itm, keyName) => {
        let val = getVal(itm.scrap, keyName);
        if (val) return val;
        if (typeof fullData !== 'undefined') {
            const wo = getVal(itm.scrap, 'work order') || getVal(itm.fullRow, 'Work Order');
            const code = getVal(itm.scrap, 'Spare Part Code') || getVal(itm.fullRow, 'Spare Part Code');
            const targetKey = (wo + code).replace(/\s/g, '').toLowerCase();
            const match = fullData.find(d => {
                const dWO = d.scrap['work order'] || '';
                const dCode = d.scrap['Spare Part Code'] || '';
                return (dWO + dCode).replace(/\s/g, '').toLowerCase() === targetKey;
            });
            if (match) return getVal(match.scrap, keyName);
        }
        return '';
    };

    const payload = { ...item.scrap, ...item.fullRow, 'user': currentUser.IDRec || 'Unknown', 'ActionStatus': item.fullRow['ActionStatus'] || 'เคลมประกัน', 'Qty': document.getElementById('store_qty').value || item.scrap['qty'] || '', 'Serial Number': document.getElementById('store_serial').value || item.fullRow['Serial Number'] || '', 'Phone': item.technicianPhone || '', 'รหัสช่าง': item.scrap['รหัสช่าง'] || '', 'ชื่อช่าง': item.scrap['ชื่อช่าง'] || '', 'Claim Receiver': (item.fullRow['ActionStatus'] === 'เคลมประกัน') ? (document.getElementById('store_receiver').value || item.person || '') : '',
        'วันที่รับซาก': getRobustVal(item, 'วันที่รับซาก'),
        'ผู้รับซาก': getRobustVal(item, 'ผู้รับซาก'),
        'Keep': getRobustVal(item, 'Keep'),
        'CI Name': item.fullRow['CI Name'] || getVal(item.fullRow, 'CI Name'),
        'Problem': item.fullRow['Problem'] || getVal(item.fullRow, 'Problem'),
        'Product Type': item.fullRow['Product Type'] || getVal(item.fullRow, 'Product Type')
    };
    
    // Explicitly set Key
    payload['Key'] = (getVal(item.scrap, 'work order') || '') + (getVal(item.scrap, 'Spare Part Code') || '');

    delete payload['Values'];

    // --- OPTIMISTIC UPDATE (Update Local Data Immediately) ---
    item.status = payload['ActionStatus'];
    item.fullRow['user'] = payload['user'];
    
    // Update Global Booking Data locally
    const targetKey = (payload['work order'] || '') + (payload['Spare Part Code'] || '');
    
    // Preserve existing data if available
    const existingRow = globalBookingData.find(row => { const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || ''); return rowKey === targetKey; });
    const preservedSlip = existingRow ? (existingRow['Booking Slip'] || '') : '';
    const preservedDate = existingRow ? (existingRow['Booking Date'] || '') : '';
    const preservedPlantCenter = existingRow ? (existingRow['Plantcenter'] || '') : '';

    globalBookingData = globalBookingData.filter(row => { const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || ''); return rowKey !== targetKey; });
    
    if (payload['ActionStatus'] === 'เคลมประกัน') {
        const getVal = (obj, keyName) => {
             if (obj[keyName]) return obj[keyName];
             const found = Object.keys(obj).find(k => k.toLowerCase().trim() === keyName.toLowerCase());
             return found ? obj[found] : '';
        };

        const newBookingRow = { 'Work Order': payload['work order'], 'Spare Part Code': payload['Spare Part Code'], 'Spare Part Name': payload['Spare Part Name'], 'Old Material Code': payload['old material code'], 'Qty': payload['Qty'], 'Serial Number': payload['Serial Number'], 'Store Code': payload['Store Code'], 'Store Name': payload['Store Name'], 'รหัสช่าง': payload['รหัสช่าง'], 'ชื่อช่าง': payload['ชื่อช่าง'], 'Mobile': payload['Phone'], 'Plant': payload['plant'], 'Claim Receiver': payload['Claim Receiver'], 'person': item.person, 'Product': payload['Product'], 'Warranty Action': payload['ActionStatus'], 'Recorder': payload['user'], 'Timestamp': new Date().toISOString(), 'Booking Slip': preservedSlip, 'Booking Date': preservedDate, 'Plantcenter': preservedPlantCenter,
            'วันที่รับซาก': getVal(item.scrap, 'วันที่รับซาก'),
            'ผู้รับซาก': getVal(item.scrap, 'ผู้รับซาก'),
            'Keep': getVal(item.scrap, 'Keep'),
            'CI Name': item.fullRow['CI Name'] || '',
            'Problem': item.fullRow['Problem'] || '',
            'Product Type': item.fullRow['Product Type'] || ''
        };
        globalBookingData.push(newBookingRow);
    }

    // Refresh UI immediately
    populateBookingFilter(); 
    renderBookingTable(); 
    renderTable();

    // --- BACKGROUND SAVE ---
    SaveQueue.add(payload);

    // Show simple toast instead of blocking alert
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Saved to queue' });
}