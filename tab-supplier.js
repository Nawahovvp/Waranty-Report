let supplierProductOptions = new Set();
let currentSupplierPage = 1;
let currentClaimSentPage = 1;
let currentHistoryPage = 1;

function populateSupplierProductFilter(data) {
    const dropdown = document.getElementById('supplierProductDropdown');
    if (!dropdown) return;
    let sourceData = data;
    if (!sourceData) {
        sourceData = globalBookingData.filter(item => item['Recripte'] && item['Recripte'].trim() !== '');
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
    if (!filterSelect) return;
    let supplierItems = globalBookingData.filter(item => {
        return item['Recripte'] && String(item['Recripte']).trim() !== '';
    });
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
    populateSupplierProductFilter(supplierItems);
}

function renderSupplierTable() {
    const tableHeader = document.getElementById('supplierTableHeader');
    const tableBody = document.getElementById('supplierTableBody');
    const filterSelect = document.getElementById('supplierPartFilter');
    const filterValue = filterSelect ? filterSelect.value : '';
    let supplierData = globalBookingData.filter(item => {
        return item['Recripte'] && item['Recripte'].trim() !== '';
    });
    const dropdown = document.getElementById('supplierProductDropdown');
    if (dropdown && dropdown.children.length === 0) {
        populateSupplierProductFilter(supplierData);
    }
    if (supplierProductOptions.size > 0) {
        supplierData = supplierData.filter(item => supplierProductOptions.has(item['Product']));
    }
    if (filterValue) {
        supplierData = supplierData.filter(item => item['Spare Part Name'] === filterValue);
    }
    const sortedData = [...supplierData].reverse().sort((a, b) => {
        const nameA = a['Spare Part Name'] || '';
        const nameB = b['Spare Part Name'] || '';
        return nameA.localeCompare(nameB, 'th');
    });
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
                td.innerHTML = `<input type="checkbox" class="supplier-checkbox" value="${globalIndex}" ${disabledAttr} onchange="handleSupplierCheckboxChange()">`;
                td.style.textAlign = 'center';
                if (isClaimed) td.style.opacity = '0.5';
            } else if (col.key === 'CustomStatus') {
                td.innerHTML = getComputedStatus(item);
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
                td.textContent = value;
            }
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    const totalPages = Math.ceil(sortedData.length / ITEMS_PER_PAGE);
    if (currentSupplierPage > totalPages) currentSupplierPage = 1;
    renderGenericPagination('supplierPaginationControls', currentSupplierPage, totalPages, changeSupplierPage);
}

function toggleAllSupplierCheckboxes(source) {
    const checkboxes = document.querySelectorAll('.supplier-checkbox:not(:disabled)');
    checkboxes.forEach(cb => cb.checked = source.checked);
    handleSupplierCheckboxChange();
}

function handleSupplierCheckboxChange() {
    const checkboxes = document.querySelectorAll('.supplier-checkbox:checked');
    const bulkActions = document.getElementById('supplierBulkActions');
    const selectAll = document.getElementById('selectAllSupplier');
    const allCheckboxes = document.querySelectorAll('.supplier-checkbox:not(:disabled)');
    if (allCheckboxes.length > 0) {
        if (checkboxes.length === allCheckboxes.length) { selectAll.checked = true; selectAll.indeterminate = false; }
        else if (checkboxes.length > 0) { selectAll.checked = false; selectAll.indeterminate = true; }
        else { selectAll.checked = false; selectAll.indeterminate = false; }
    }
    if (checkboxes.length > 0) bulkActions.style.display = 'flex'; else bulkActions.style.display = 'none';
}

async function sendClaim() {
    const checkboxes = document.querySelectorAll('.supplier-checkbox:checked');
    if (checkboxes.length === 0) return;
    const filterSelect = document.getElementById('supplierPartFilter');
    const filterValue = filterSelect ? filterSelect.value : '';
    let allowedData = globalBookingData.filter(item => item['Recripte'] && item['Recripte'].trim() !== '');
    if (supplierProductOptions.size > 0) allowedData = allowedData.filter(item => supplierProductOptions.has(item['Product']));
    if (filterValue) allowedData = allowedData.filter(item => item['Spare Part Name'] === filterValue);
    const sortedData = [...allowedData].reverse().sort((a, b) => { const nameA = a['Spare Part Name'] || ''; const nameB = b['Spare Part Name'] || ''; return nameA.localeCompare(nameB, 'th'); });
    const selectedItems = [];
    checkboxes.forEach(cb => { const idx = parseInt(cb.value); if (sortedData[idx]) selectedItems.push(sortedData[idx]); });
    if (selectedItems.length === 0) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const claimSup = currentUser.IDRec || 'Unknown';
    const now = new Date();
    const formattedDate = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

    const result = await Swal.fire({ title: 'Confirm Send Claim?', html: `Update <b>${selectedItems.length}</b> items?<br><br>Date: <b>${formattedDate}</b><br>Supplier: <b>${claimSup}</b>`, icon: 'question', showCancelButton: true, confirmButtonText: 'Yes, Send', cancelButtonText: 'Cancel' });
    if (!result.isConfirmed) return;

    Swal.fire({ title: 'Saving...', html: `Updating ${selectedItems.length} items...`, allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    let successCount = 0; let failCount = 0;
    for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];
        const payload = { ...item, 'Claim Date': formattedDate, 'ClaimSup': claimSup, 'user': JSON.parse(localStorage.getItem('currentUser') || '{}').IDRec || 'Unknown' };
        try { await postToGAS(payload); item['Claim Date'] = formattedDate; item['ClaimSup'] = claimSup; successCount++; } catch (e) { console.error(e); failCount++; }
    }
    renderSupplierTable(); handleSupplierCheckboxChange();
    Swal.fire({ icon: 'success', title: 'Completed', text: `Success: ${successCount}, Failed: ${failCount}` });
}

function changeSupplierPage(newPage) { if (newPage < 1) return; currentSupplierPage = newPage; renderSupplierTable(); }

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
    renderGenericPagination('claimSentPaginationControls', currentClaimSentPage, totalPages, changeClaimSentPage);
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
    renderGenericPagination('historyPaginationControls', currentHistoryPage, totalPages, changeHistoryPage);
}

function changeHistoryPage(newPage) {
    if (newPage < 1) return;
    currentHistoryPage = newPage;
    renderHistoryTable();
}