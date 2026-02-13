const cards = deckContainer.getElementsByClassName('deck-card');

// Show all cards
Array.from(cards).forEach(c => c.style.display = 'flex');

document.getElementById(tabKey + 'TableWrapper').style.display = 'none';


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
            <span style="font-weight: bold; font-size: 1.1rem; color: #1e293b;">รายการอะไหล่ส่งเคลม: ${slip} (${targetReceiver})</span>
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
        toast: true,
        position: 'top-end',
        showConfirmButton: false,
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
            await postToGAS(payload);

            // Update Local
            item['Recripte'] = recripteName;
            item['RecripteDate'] = recripteDateStr;
            successCount++;
        } catch (e) {
            console.error(e);
            failCount++;
        }
    }

    const Toast = Swal.mixin({
        toast: true,
        position: 'top-end',
        showConfirmButton: false,
        timer: 3000,
        timerProgressBar: true
    });
    Toast.fire({
        icon: 'success',
        title: 'Updated Successfully',
        text: `Success: ${successCount}, Failed: ${failCount}`
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

    // UI Feedback: Saving... (Toast)
    Swal.fire({
        title: 'Saving...',
        text: 'Updating Work Order details...',
        toast: true,
        position: 'top-end',
        showConfirmButton: false,
        didOpen: () => Swal.showLoading()
    });

    try {
        await postToGAS(payload);

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

        const Toast = Swal.mixin({
            toast: true,
            position: 'top-end',
            showConfirmButton: false,
            timer: 3000,
            timerProgressBar: true
        });
        Toast.fire({
            icon: 'success',
            title: 'Saved!',
            text: 'Work Order updated successfully.'
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

    // 1. FILTER BY USER PLANT (Added)
    const userPlant = getEffectiveUserPlant();
    if (userPlant) {
        supplierData = supplierData.filter(item => {
            const itemPlant = item['Plant'] || item['plant'] || item['PLANT'];
            return String(itemPlant).trim().padStart(4, '0') === userPlant;
        });
        // Console log for debug if needed, but keeping it clean
    }

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
    renderGenericPagination('supplierPaginationControls', currentSupplierPage, totalPages, changeSupplierPage);
}

function toggleAllSupplierCheckboxes(source) {
    const checkboxes = document.querySelectorAll('.supplier-checkbox');
    checkboxes.forEach(cb => cb.checked = source.checked);
    handleSupplierCheckboxChange();
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
            await postToGAS(payload);

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

    // 1. FILTER BY USER PLANT (Added)
    const userPlant = getEffectiveUserPlant();
    if (userPlant) {
        allowedData = allowedData.filter(item => {
            const itemPlant = item['Plant'] || item['plant'] || item['PLANT'];
            return String(itemPlant).trim().padStart(4, '0') === userPlant;
        });
    }

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

    // Filter by User Plant (Refactored)
    const userPlant = getEffectiveUserPlant();

    console.log('[History Filter] Effective User Plant:', userPlant);

    if (userPlant) {
        allowedData = allowedData.filter(item => {
            const itemPlant = item['Plant'] || item['plant'] || item['PLANT'];
            return String(itemPlant).trim().padStart(4, '0') === userPlant;
        });

        console.log(`[History Filter] Filtered Result: ${allowedData.length} items returned.`);

        // Debug Alert Removed
    } else {
        console.warn('[History Filter] No Plant found for user. Showing ALL data.');
        Swal.fire({
            toast: true,
            position: 'top-end',
            icon: 'warning',
            title: 'No Plant Found',
            text: 'User profile has no Plant data. Showing all.'
        });
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
