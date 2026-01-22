function populateFilters() {
    const filterSelect = document.getElementById('sparePartFilter');
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
    const selectedPart = document.getElementById('sparePartFilter').value;
    const selectedStatus = document.getElementById('statusFilter').value;
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();

    displayedData = fullData.filter(item => {
        const matchPart = selectedPart ? item.scrap['Spare Part Name'] === selectedPart : true;
        let matchStatus = true;
        if (selectedStatus === 'Pending') matchStatus = !item.status;
        else if (selectedStatus) matchStatus = item.status === selectedStatus;
        let matchSearch = true;
        if (searchTerm) {
            const allValues = [...Object.values(item.scrap), ...Object.values(item.fullRow)].join(' ').toLowerCase();
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

            if (item.status && greenColumns.includes(col.key)) {
                const span = document.createElement('span');
                span.textContent = value;
                let bgColor = '#dcfce7'; let textColor = '#166534';
                if (item.status === 'Sworp') { bgColor = '#e0e7ff'; textColor = '#3730a3'; }
                else if (item.status === 'หมดประกัน') { bgColor = '#ffedd5'; textColor = '#9a3412'; }
                else if (item.status === 'ชำรุด') { bgColor = '#fce7f3'; textColor = '#9d174d'; }
                span.style.backgroundColor = bgColor; span.style.color = textColor; span.style.fontSize = '0.875rem'; span.style.fontWeight = '600'; span.style.padding = '0.25rem 0.5rem'; span.style.borderRadius = '4px'; span.style.display = 'inline-block';
                td.appendChild(span);
            } else if (!item.status && isMissingStore && col.key === 'status') {
                const span = document.createElement('span');
                span.textContent = 'Wait Data';
                span.style.backgroundColor = '#f1f5f9'; span.style.color = '#ef4444'; span.style.fontSize = '0.875rem'; span.style.fontWeight = '600'; span.style.padding = '0.25rem 0.5rem'; span.style.borderRadius = '4px'; span.style.display = 'inline-block';
                td.appendChild(span);
            } else {
                td.textContent = value;
                td.style.fontSize = '0.875rem';
            }

            if (col.key === 'Serial Number') { td.style.cursor = 'pointer'; td.style.color = 'var(--primary-color)'; td.style.fontWeight = '500'; td.title = 'Click to edit Serial Number'; td.onclick = () => openEditModal(item); }
            if (col.key === 'Store Code') { td.style.cursor = 'pointer'; td.style.color = 'var(--primary-color)'; td.style.fontWeight = '500'; td.title = 'Click to view details'; td.onclick = () => openStoreModal(item); }
            if (col.key === 'Person' || col.source === 'P') { td.style.cursor = 'pointer'; td.style.color = 'var(--primary-color)'; td.style.fontWeight = '500'; td.title = 'Click to edit Person'; td.onclick = () => openMainPersonModal(item); }
            if (col.key === 'Mobile' || col.source === 'T' || col.key === 'Phone') {
                td.style.cursor = 'pointer'; td.style.color = 'var(--primary-color)'; td.style.fontWeight = '500'; td.title = 'Click to edit Mobile'; td.onclick = () => openMainMobileModal(item);
                if (!value) { td.innerText = '[Add]'; td.style.fontSize = '0.875rem'; }
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    renderPagination();
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
            if (actionName === 'เคลมประกัน' && (!serial || serial.length < minLength)) { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `รายการ ${item.scrap['work order']}: Serial Number ไม่ถูกต้อง (L&E > 5 หลัก, อื่นๆ >= 8 หลัก)` }); return; }
            if (actionName === 'เคลมประกัน') {
                const mobile = item.technicianPhone || '';
                if (!mobile || String(mobile).trim() === '') { Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ครบถ้วน', text: `รายการ ${item.scrap['work order']}: กรุณาระบุเบอร์โทรศัพท์ (Mobile) ก่อนบันทึก` }); return; }
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

    let successCount = 0; let failCount = 0;
    Swal.fire({ title: 'กำลังบันทึก...', html: `กำลังประมวลผล 0 / ${selectedItems.length}`, allowOutsideClick: false, didOpen: () => Swal.showLoading() });

    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];
        Swal.getHtmlContainer().textContent = `กำลังประมวลผล ${i + 1} / ${selectedItems.length}`;
        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
        const payload = { ...item.scrap, ...item.fullRow, 'user': currentUser.IDRec || 'Unknown', 'ActionStatus': actionName, 'Qty': item.scrap['qty'] || '', 'Serial Number': item.fullRow['Serial Number'] || '', 'Phone': item.technicianPhone || '', 'รหัสช่าง': item.scrap['รหัสช่าง'] || '', 'ชื่อช่าง': item.scrap['ชื่อช่าง'] || '', 'Claim Receiver': (actionName === 'เคลมประกัน') ? (item.person || '') : '' };
        delete payload['Values'];
        try {
            await postToGAS(payload);
            successCount++;
            item.status = actionName;
            item.fullRow['ActionStatus'] = actionName;
        } catch (e) { console.error(e); failCount++; }
    }

    selectedItems.forEach(item => {
        const itemPayload = { ...item.scrap, ...item.fullRow, 'Qty': item.scrap['qty'], 'Phone': item.technicianPhone };
        const targetKey = (itemPayload['work order'] || '') + (itemPayload['Spare Part Code'] || '');
        globalBookingData = globalBookingData.filter(row => { const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || ''); return rowKey !== targetKey; });
        if (actionName === 'เคลมประกัน') {
            const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
            const newBookingRow = { 'Work Order': itemPayload['work order'], 'Spare Part Code': itemPayload['Spare Part Code'], 'Spare Part Name': itemPayload['Spare Part Name'], 'Old Material Code': itemPayload['old material code'], 'Qty': itemPayload['Qty'], 'Serial Number': itemPayload['Serial Number'], 'Store Code': itemPayload['Store Code'], 'Store Name': itemPayload['Store Name'], 'รหัสช่าง': itemPayload['รหัสช่าง'], 'ชื่อช่าง': itemPayload['ชื่อช่าง'], 'Mobile': itemPayload['Phone'], 'Plant': itemPayload['plant'], 'Claim Receiver': (actionName === 'เคลมประกัน') ? (item.person || '') : '', 'Product': itemPayload['Product'], 'Warranty Action': actionName, 'Recorder': currentUser.IDRec || 'Unknown', 'Timestamp': new Date().toISOString(), 'Booking Slip': '', 'Booking Date': '' };
            globalBookingData.push(newBookingRow);
        }
    });

    populateBookingFilter(); renderBookingTable();
    try { renderDeckView('0301', 'navaNakornDeck', 'navanakorn'); renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi'); populateSupplierFilter(); renderSupplierTable(); populateClaimSentFilter(); renderClaimSentTable(); renderHistoryTable(); } catch (err) { console.warn('Error refreshing other tabs:', err); }
    renderTable();
    toggleAllCheckboxes({ checked: false });
    Swal.fire({ icon: successCount > 0 ? 'success' : 'error', title: 'เสร็จสิ้น', text: `บันทึกสำเร็จ: ${successCount} รายการ\nล้มเหลว: ${failCount} รายการ` });
}

function sendDatatoGAS(item) {
    if (!GAS_API_URL) { alert("Please configure the Google Apps Script URL."); return; }
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const payload = { ...item.scrap, ...item.fullRow, 'user': currentUser.IDRec || 'Unknown', 'ActionStatus': item.fullRow['ActionStatus'] || 'เคลมประกัน', 'Qty': document.getElementById('store_qty').value || item.scrap['qty'] || '', 'Serial Number': document.getElementById('store_serial').value || item.fullRow['Serial Number'] || '', 'Phone': item.technicianPhone || '', 'รหัสช่าง': item.scrap['รหัสช่าง'] || '', 'ชื่อช่าง': item.scrap['ชื่อช่าง'] || '', 'Claim Receiver': (item.fullRow['ActionStatus'] === 'เคลมประกัน') ? (document.getElementById('store_receiver').value || item.person || '') : '' };
    delete payload['Values'];
    const saveBtn = document.querySelector('#storeModal .btn-primary');
    const originalText = saveBtn.textContent;
    saveBtn.disabled = true;
    Swal.fire({ title: 'Saving...', text: 'Please wait while we save your data.', allowOutsideClick: false, didOpen: () => { Swal.showLoading(); } });

    postToGAS(payload).then(() => {
        Swal.fire({ icon: 'success', title: 'Saved!', text: 'Data successfully saved to Google Sheet.', timer: 2000, showConfirmButton: false });
        item.status = payload['ActionStatus'];
        item.fullRow['user'] = payload['user'];
        const targetKey = (payload['work order'] || '') + (payload['Spare Part Code'] || '');
        globalBookingData = globalBookingData.filter(row => { const rowKey = (row['Work Order'] || '') + (row['Spare Part Code'] || ''); return rowKey !== targetKey; });
        if (payload['ActionStatus'] === 'เคลมประกัน') {
            const newBookingRow = { 'Work Order': payload['work order'], 'Spare Part Code': payload['Spare Part Code'], 'Spare Part Name': payload['Spare Part Name'], 'Old Material Code': payload['old material code'], 'Qty': payload['Qty'], 'Serial Number': payload['Serial Number'], 'Store Code': payload['Store Code'], 'Store Name': payload['Store Name'], 'รหัสช่าง': payload['รหัสช่าง'], 'ชื่อช่าง': payload['ชื่อช่าง'], 'Mobile': payload['Phone'], 'Plant': payload['plant'], 'Claim Receiver': payload['Claim Receiver'], 'Product': payload['Product'], 'Warranty Action': payload['ActionStatus'], 'Recorder': payload['user'], 'Timestamp': new Date().toISOString(), 'Booking Slip': '', 'Booking Date': '' };
            globalBookingData.push(newBookingRow);
        }
        populateBookingFilter(); renderBookingTable(); renderTable();
    }).catch(error => { console.error('Error saving to GAS:', error); Swal.fire({ icon: 'error', title: 'Oops...', text: 'Failed to save to Google Sheet. Please try again.' }); })
    .finally(() => { saveBtn.textContent = originalText; saveBtn.disabled = false; });
}