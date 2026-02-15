// Global context for detail view
let currentDetailContext = null;
let deckFilterState = {
    navanakorn: 'Pending',
    vibhavadi: 'Pending'
};

function setDeckFilter(tabKey, status) {
    deckFilterState[tabKey] = status;
    
    // Update UI buttons
    const btnPending = document.getElementById(`btn-${tabKey}-pending`);
    const btnAll = document.getElementById(`btn-${tabKey}-all`);
    
    if (btnPending && btnAll) {
        if (status === 'Pending') {
            btnPending.style.backgroundColor = '#f59e0b'; // Orange/Amber
            btnPending.style.color = 'white';
            btnPending.style.borderColor = '#d97706';
            btnAll.style.backgroundColor = '#e5e7eb';
            btnAll.style.color = '#374151';
            btnAll.style.borderColor = '#d1d5db';
        } else {
            btnAll.style.backgroundColor = '#3b82f6'; // Blue
            btnAll.style.color = 'white';
            btnAll.style.borderColor = '#2563eb';
            btnPending.style.backgroundColor = '#e5e7eb';
            btnPending.style.color = '#374151';
            btnPending.style.borderColor = '#d1d5db';
        }
    }

    // Re-render based on tabKey
    if (tabKey === 'navanakorn') renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
    if (tabKey === 'vibhavadi') renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
}

function renderDeckView(targetPlantCode, containerId, tabKey) {
    const deckContainer = document.getElementById(containerId);
    if (!deckContainer) return;
    deckContainer.innerHTML = '';

    const filterId = tabKey + 'PlantFilter';
    populateDeckPlantFilter(targetPlantCode, filterId);
    const filterElement = document.getElementById(filterId);
    const selectedPlant = filterElement ? filterElement.value : '';

    const receiverFilterId = tabKey + 'ReceiverFilter';
    populateDeckReceiverFilter(targetPlantCode, receiverFilterId);
    const receiverFilterElement = document.getElementById(receiverFilterId);
    const selectedReceiver = receiverFilterElement ? receiverFilterElement.value : '';

    const uniqueMap = new Map();

    // Get User Plant for security filtering
    const userPlant = (typeof getEffectiveUserPlant === 'function') ? getEffectiveUserPlant() : null;

    globalBookingData.forEach(item => {
        const slip = item['Booking Slip'];
        const pcCode = String(item['Plantcenter'] || '').trim();
        const normalize = (s) => s.replace(/^0+/, '');

        if (!slip) return;
        if (normalize(pcCode) !== normalize(targetPlantCode)) return;

        // SECURITY FILTER: Matches User Plant?
        if (userPlant) {
            const userStatus = getCurrentUserStatus();
            // Allow if User Status matches Claim Receiver
            if (userStatus && item['Claim Receiver'] === userStatus) {
                // Allow
            } else {
                const itemPlant = String(item['Plant'] || '').trim().padStart(4, '0'); // Normalize 0304
                if (itemPlant !== userPlant) return;
            }
        }

        if (selectedPlant) {
            const itemPlant = String(item['Plant'] || '').replace(/^0+/, '');
            if (itemPlant !== selectedPlant) return;
        }

        const receiverVal = item['Claim Receiver'] || item.person || 'Unknown';

        if (selectedReceiver) {
            if (receiverVal !== selectedReceiver) return;
        }

        const key = slip + '|' + receiverVal;
        const hasRecripte = item['Recripte'] && item['Recripte'].trim() !== '';
        const recripteInc = hasRecripte ? 1 : 0;

        if (!uniqueMap.has(key)) {
            uniqueMap.set(key, { ...item, _effectiveReceiver: receiverVal, _count: 1, _recripteCount: recripteInc });
        } else {
            const existing = uniqueMap.get(key);
            existing._count = (existing._count || 1) + 1;
            existing._recripteCount = (existing._recripteCount || 0) + recripteInc;
        }
    });

    // Convert to array for custom sorting
    let cardsArray = Array.from(uniqueMap.values());

    // FILTER BY STATUS BUTTON (New Logic)
    const currentFilter = deckFilterState[tabKey] || 'All';
    if (currentFilter === 'Pending') {
        cardsArray = cardsArray.filter(item => {
            const count = item._count || 1;
            const recCount = item._recripteCount || 0;
            // Keep if NOT completed (i.e. recCount < count)
            // This covers "In Transit" (recCount=0) and "Receiving" (0 < recCount < count)
            return recCount < count;
        });
    }

    // Sort: Status Priority (ระหว่างตรวจรับ -> ระหว่างขนส่ง -> เสร็จสิ้น) then Date Descending
    cardsArray.sort((a, b) => {
        const getStatusPriority = (item) => {
            const count = item._count || 1;
            const recCount = item._recripteCount || 0;
            if (recCount > 0 && recCount < count) return 1; // ระหว่างตรวจรับ
            if (recCount === 0) return 2; // ระหว่างขนส่ง
            if (recCount === count) return 3; // เสร็จสิ้น
            return 4;
        };

        const priorityA = getStatusPriority(a);
        const priorityB = getStatusPriority(b);

        if (priorityA !== priorityB) return priorityA - priorityB;

        const parseDate = (dateStr) => {
            if (!dateStr) return 0;
            const [d, m, y] = dateStr.split('/');
            return new Date(y, m - 1, d).getTime();
        };
        return parseDate(b['Booking Date']) - parseDate(a['Booking Date']);
    });

    if (cardsArray.length === 0) {
        deckContainer.innerHTML = '<p style="text-align: center; color: var(--text-secondary); grid-column: 1/-1;">No booking records found for this plant.</p>';
        return;
    }

    cardsArray.forEach(item => {
        const slip = item['Booking Slip'];
        const date = item['Booking Date'];
        const plantCode = item['Plant'] || '';
        const normalizedPlantCode = String(plantCode).replace(/^0+/, '');
        const plantName = PLANT_MAPPING[normalizedPlantCode] || plantCode;
        const receiver = item._effectiveReceiver;
        const count = item._count || 1;
        const recCount = item._recripteCount || 0;
        const key = slip + '|' + receiver;

        let statusText = 'ระหว่างขนส่ง';
        let statusBg = '#fef3c7'; // yellow
        let statusColor = '#92400e';

        if (recCount === 0) { statusText = '🚚 ระหว่างขนส่ง'; statusBg = '#fef3c7'; statusColor = '#92400e'; }
        else if (recCount < count) { statusText = '⏳ ระหว่างตรวจรับ'; statusBg = '#fee2e2'; statusColor = '#991b1b'; }
        else if (recCount === count) { statusText = '✅ เสร็จสิ้น'; statusBg = '#dbeafe'; statusColor = '#1e40af'; }

        const card = document.createElement('div');
        card.className = 'deck-card';
        card.style.cursor = 'pointer';
        card.setAttribute('data-key', key);
        card.onclick = function () { toggleDetailView(this, tabKey, slip, receiver); };

        card.innerHTML = `
            <div class="card-header" style="display: block;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
                    <div class="card-title">📄 ${slip}</div>
                    <div class="card-badge" style="background-color: #f3f4f6; color: #374151;">📅 ${date}</div>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <span class="card-badge" style="background-color: ${statusBg}; color: ${statusColor}; font-weight:bold; font-size: 0.75rem;">${statusText}</span>
                    <div class="card-badge" style="background-color: #dbeafe; color: #1e40af;">รับแล้ว ${recCount}/${count} item</div>
                </div>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; font-weight: 700; font-size: 1rem; color: #1e293b;">
                <div>${plantName}</div>
                <div style="color: #475569;">${receiver}</div>
            </div>
        `;
        deckContainer.appendChild(card);
    });
}

function backToDeck(tabKey) {
    const deckContainer = document.getElementById(tabKey === 'navanakorn' ? 'navaNakornDeck' : 'vibhavadiDeck');
    if (!deckContainer) return;

    const cards = deckContainer.getElementsByClassName('deck-card');
    Array.from(cards).forEach(c => c.style.display = 'flex');

    const tableWrapper = document.getElementById(tabKey + 'TableWrapper');
    if (tableWrapper) tableWrapper.style.display = 'none';

    const deckWrapper = document.getElementById(tabKey + 'DeckWrapper');
    if (deckWrapper) {
        deckWrapper.style.height = '';
        deckWrapper.style.flex = '1';
    }

    // Remove active state from cards
    Array.from(cards).forEach(c => {
        c.classList.remove('active-card');
        c.style.border = '1px solid var(--border-color)';
    });

    currentDetailContext = null;
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
        tableWrapper.style.height = '0';
        tableWrapper.style.flex = '1';

        // Set Context and Render
        currentDetailContext = { tabKey, slip, targetReceiver, currentPage: 1, searchTerm: '' };
        renderTopLevelDetailTable(tabKey, slip, targetReceiver);
    }
}

function renderTopLevelDetailTable(tabKey, slip, targetReceiver) {
    const tableWrapper = document.getElementById(tabKey + 'TableWrapper');
    const header = document.getElementById(tabKey + 'ReviewHeader');
    const thead = document.getElementById(tabKey + 'TableHeader');
    const tbody = document.getElementById(tabKey + 'TableBody');

    // Filter Data
    let detailData = globalBookingData.filter(item => {
        const itemReceiver = item['Claim Receiver'] || item.person || 'Unknown';
        // SECURITY: Check User Plant if valid
        if (typeof getEffectiveUserPlant === 'function') {
            const userPlant = getEffectiveUserPlant();
            const userStatus = getCurrentUserStatus();
            if (userPlant) {
                // Allow if User Status matches Claim Receiver
                if (userStatus && item['Claim Receiver'] === userStatus) {
                    // Allow
                } else {
                    const itemPlant = String(item['Plant'] || '').trim().padStart(4, '0');
                    if (itemPlant !== userPlant) return false;
                }
            }
        }
        return item['Booking Slip'] === slip && itemReceiver === targetReceiver;
    });

    // Apply Search Filter
    const searchTerm = (currentDetailContext && currentDetailContext.searchTerm) ? currentDetailContext.searchTerm.toLowerCase() : '';
    if (searchTerm) {
        detailData = detailData.filter(item => {
            return Object.values(item).some(val => String(val).toLowerCase().includes(searchTerm));
        });
    }

    // Pagination Setup
    const ITEMS_PER_PAGE_DECK = 20;
    const currentPage = (currentDetailContext && currentDetailContext.currentPage) ? currentDetailContext.currentPage : 1;
    const totalPages = Math.ceil(detailData.length / ITEMS_PER_PAGE_DECK);

    // Slice Data
    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE_DECK;
    const endIndex = startIndex + ITEMS_PER_PAGE_DECK;
    const pageData = detailData.slice(startIndex, endIndex);

    // Update Header (Only if context changed to prevent losing input focus)
    const currentHeaderKey = header.getAttribute('data-key');
    const newHeaderKey = `${slip}|${targetReceiver}`;

    if (currentHeaderKey !== newHeaderKey || header.innerHTML.trim() === '') {
        header.setAttribute('data-key', newHeaderKey);
        header.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
                <span style="font-weight: bold; font-size: 1.1rem; color: #1e293b;">รายการอะไหล่ส่งเคลม: ${slip} (${targetReceiver})</span>
                <div style="display: flex; align-items: center; gap: 0.5rem;">
                    <input type="text" placeholder="Search..." 
                        style="padding: 0.35rem 0.5rem; border: 1px solid var(--border-color); border-radius: 4px; font-size: 0.875rem;"
                        value="${currentDetailContext.searchTerm || ''}" 
                        oninput="handleDeckDetailSearch(this)">
                    <button class="btn btn-primary bulk-save-btn" style="display: none;" onclick="saveBulkReviewItems(this)">
                        ยืนยันรับ
                    </button>
                </div>
            </div>
        `;
    } else {
        // Ensure Save button is hidden on re-render (since checkboxes are cleared)
        const saveBtn = header.querySelector('.bulk-save-btn');
        if (saveBtn) saveBtn.style.display = 'none';
    }

    // Render Table Header
    thead.innerHTML = '';
    BOOKING_COLUMNS.forEach(col => {
        const th = document.createElement('th');
        if (col.key === 'checkbox') {
            th.innerHTML = `<input type="checkbox" class="select-all-review" onclick="toggleAllReviewCheckboxes(this)" ${isReviewSelectAll ? 'checked' : ''}>`;
        } else {
            th.innerHTML = col.header;
        }
        thead.appendChild(th);
    });

    // Render Table Body
    tbody.innerHTML = '';

    if (pageData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="' + BOOKING_COLUMNS.length + '" style="text-align:center;">No data found.</td></tr>';
    } else {
        pageData.forEach((item, index) => {
            const tr = document.createElement('tr');
            BOOKING_COLUMNS.forEach(col => {
                const td = document.createElement('td');
                let value = item[col.key] || '';

                if (col.key === 'checkbox') {
                    const hasBookingSlip = item['Booking Slip'] && String(item['Booking Slip']).trim() !== '';
                    const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
                    const disabledAttr = hasRecripte ? 'disabled' : '';
                    // If Global Select All is active, and item is not received (In Transit), check it.
                    const contextChecked = (isReviewSelectAll && !hasRecripte) ? 'checked' : '';

                    td.innerHTML = `<input type="checkbox" class="review-checkbox" value="${startIndex + index}" onchange="handleReviewCheckboxChange(this)" ${disabledAttr} ${contextChecked}>`;
                    td.style.textAlign = 'center';
                } else if (col.key === 'CustomStatus') {
                    td.innerHTML = getComputedStatus(item);
                    const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
                    const hasClaimSup = item['ClaimSup'] && String(item['ClaimSup']).trim() !== '';

                    if (!hasRecripte) {
                        td.style.cursor = 'pointer';
                        td.title = 'กดเพื่อยืนยันรับ';
                        td.onclick = function (e) {
                            e.stopPropagation();
                            confirmReceiveItem(item);
                        };
                    } else if (hasRecripte && !hasClaimSup) {
                        td.style.cursor = 'pointer';
                        td.title = 'กดเพื่อยกเลิกการรับ';
                        td.onclick = function (e) {
                            e.stopPropagation();
                            cancelReceiveItem(item);
                        };
                    }
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

    const saveBtn = header.querySelector('.bulk-save-btn');
    if (saveBtn) {
        if (isReviewSelectAll) {
            saveBtn.style.display = 'block';
            saveBtn.textContent = 'ยืนยันรับ (ทั้งหมด / All)';
        } else {
            const hasChecked = tbody.querySelector('.review-checkbox:checked');
            if (hasChecked) {
                const count = tbody.querySelectorAll('.review-checkbox:checked').length;
                saveBtn.style.display = 'block';
                saveBtn.textContent = `ยืนยันรับ (${count})`;
            } else {
                saveBtn.style.display = 'none';
            }
        }
    }

    const paginationId = tabKey === 'navanakorn' ? 'navanakornPagination' : 'vibhavadiPagination';
    renderGenericPagination(paginationId, currentPage, totalPages, changeDeckDetailPage);
}

let isReviewSelectAll = false;

function toggleAllReviewCheckboxes(source) {
    isReviewSelectAll = source.checked;
    const table = source.closest('table');
    if (!table) return;
    const checkboxes = table.querySelectorAll('.review-checkbox:not(:disabled)');
    checkboxes.forEach(cb => cb.checked = source.checked);
    handleReviewCheckboxChange(source);
}

function handleReviewCheckboxChange(source) {
    const wrapper = source.closest('.detail-content-wrapper') || source.closest('[id$="TableWrapper"]');
    if (!wrapper) return;
    const table = wrapper.querySelector('table');
    const checkboxes = table.querySelectorAll('.review-checkbox:checked');
    const allEnabledCheckboxes = table.querySelectorAll('.review-checkbox:not(:disabled)');
    const saveBtn = wrapper.querySelector('.bulk-save-btn');

    // If user unchecks a single item, disable "Select All Mode"
    if (!source.checked && source.classList.contains('review-checkbox')) {
        isReviewSelectAll = false;
        const headerCheckbox = table.querySelector('.select-all-review');
        if (headerCheckbox) headerCheckbox.checked = false;
    }

    // Recalculate if Select All matches visual state (for current page)
    if (checkboxes.length === allEnabledCheckboxes.length && allEnabledCheckboxes.length > 0) {
        // If all on current page are checked, we might be in Select All mode,
        // but to truly be "All Pages", the user must have clicked the header.
        // We keep isReviewSelectAll as is (true if set by header). 
    }

    if (saveBtn) {
        if (isReviewSelectAll) {
            saveBtn.style.display = 'block';
            saveBtn.textContent = `ยืนยันรับ (ทั้งหมด / All)`;
        } else {
            saveBtn.style.display = checkboxes.length > 0 ? 'block' : 'none';
            saveBtn.textContent = `ยืนยันรับ (${checkboxes.length})`;
        }
    }
}

async function saveBulkReviewItems(btnElement) {
    if (!currentDetailContext) return;
    const { slip, targetReceiver, tabKey } = currentDetailContext;
    const detailData = globalBookingData.filter(item => {
        const itemReceiver = item['Claim Receiver'] || item.person || 'Unknown';
        return item['Booking Slip'] === slip && itemReceiver === targetReceiver;
    });

    let selectedItems = [];

    if (isReviewSelectAll) {
        // Select ALL items that are "In Transit" (i.e. not yet received)
        selectedItems = detailData.filter(item => {
            const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
            return !hasRecripte; // Only select pending items
        });
    } else {
        const wrapper = btnElement.closest('.detail-content-wrapper') || btnElement.closest('[id$="TableWrapper"]');
        if (!wrapper) return;
        const checkboxes = wrapper.querySelectorAll('.review-checkbox:checked');

        // We need to map visual indices (which are paginated) back to detailData indices
        // The checkboxes values were set as "startIndex + index", so they are absolute indices relative to detailData
        checkboxes.forEach(cb => {
            const idx = parseInt(cb.value);
            if (detailData[idx]) selectedItems.push(detailData[idx]);
        });
    }

    if (selectedItems.length === 0) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const recripteName = currentUser.name || currentUser.IDRec || 'Unknown';
    const recripteDate = new Date();
    const recripteDateStr = recripteDate.toLocaleString('en-GB');

    selectedItems.forEach(item => {
        const payload = { ...item, 'Recripte': recripteName, 'RecripteDate': recripteDateStr, 'user': recripteName };

        // Optimistic Update
        item['Recripte'] = recripteName;
        item['RecripteDate'] = recripteDateStr;

        // Queue
        SaveQueue.add(payload);
    });

    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: `Updated ${selectedItems.length} items` });

    renderTopLevelDetailTable(tabKey, slip, targetReceiver);
    if (tabKey === 'navanakorn') renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
    if (tabKey === 'vibhavadi') renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');

    const activeKey = slip + '|' + targetReceiver;
    const deckContainer = document.getElementById(tabKey === 'navanakorn' ? 'navaNakornDeck' : 'vibhavadiDeck');
    if (deckContainer) {
        const cards = deckContainer.getElementsByClassName('deck-card');
        let activeCard = null;
        Array.from(cards).forEach(c => {
            if (c.getAttribute('data-key') === activeKey) activeCard = c;
        });
        if (activeCard) {
            activeCard.classList.add('active-card');
            Array.from(cards).forEach(c => {
                if (c !== activeCard) c.style.display = 'none';
            });
        }
    }
}

function changeDeckDetailPage(newPage) {
    if (!currentDetailContext) return;
    currentDetailContext.currentPage = newPage;
    renderTopLevelDetailTable(currentDetailContext.tabKey, currentDetailContext.slip, currentDetailContext.targetReceiver);
}

function populateDeckPlantFilter(targetPlantCode, filterId) {
    const filterSelect = document.getElementById(filterId);
    if (!filterSelect) return;

    const currentSelection = filterSelect.value;

    // Clear existing options except first
    while (filterSelect.options.length > 1) {
        filterSelect.remove(1);
    }

    const plants = new Set();
    globalBookingData.forEach(item => {
        const pcCode = String(item['Plantcenter'] || '').trim();
        const normalize = (s) => s.replace(/^0+/, '');
        if (normalize(pcCode) === normalize(targetPlantCode)) {
            if (item['Plant']) {
                const p = String(item['Plant']).replace(/^0+/, '');
                plants.add(p);
            }
        }
    });

    const sortedPlants = Array.from(plants).sort();

    sortedPlants.forEach(plant => {
        const option = document.createElement('option');
        option.value = plant;
        const normalizedPlant = String(plant).replace(/^0+/, '');
        const plantName = (typeof PLANT_MAPPING !== 'undefined' ? PLANT_MAPPING[normalizedPlant] : plant) || plant;
        option.textContent = plantName;
        filterSelect.appendChild(option);
    });

    if (currentSelection && Array.from(filterSelect.options).some(o => o.value === currentSelection)) {
        filterSelect.value = currentSelection;
    }
}

function handleDeckDetailSearch(input) {
    if (!currentDetailContext) return;
    currentDetailContext.searchTerm = input.value;
    currentDetailContext.currentPage = 1;
    renderTopLevelDetailTable(currentDetailContext.tabKey, currentDetailContext.slip, currentDetailContext.targetReceiver);
}

async function cancelReceiveItem(item) {
    const result = await Swal.fire({
        title: 'ยกเลิกการรับ?',
        text: `คุณต้องการยกเลิกการรับรายการ: ${item['Spare Part Name'] || ''}?`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ยกเลิกการรับ',
        cancelButtonText: 'ไม่',
        confirmButtonColor: '#d33'
    });

    if (!result.isConfirmed) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');

    // Payload to clear values
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
            const getVal = (obj, keyName) => { if (!obj) return ''; const found = Object.keys(obj).find(k => k.toLowerCase().trim() === keyName.toLowerCase()); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...item, 'Recripte': '', 'RecripteDate': '', 'user': currentUser.IDRec || 'Unknown',
        'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || ''
    };

    // Optimistic Update
    item['Recripte'] = '';
    item['RecripteDate'] = '';

    // Queue
    SaveQueue.add(payload);

    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'ยกเลิกการรับเรียบร้อย' });

    if (currentDetailContext) {
        const { slip, targetReceiver, tabKey } = currentDetailContext;
        renderTopLevelDetailTable(tabKey, slip, targetReceiver);
        if (tabKey === 'navanakorn') renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
        if (tabKey === 'vibhavadi') renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');

        const activeKey = slip + '|' + targetReceiver;
        const deckContainer = document.getElementById(tabKey === 'navanakorn' ? 'navaNakornDeck' : 'vibhavadiDeck');
        if (deckContainer) {
            const cards = deckContainer.getElementsByClassName('deck-card');
            let activeCard = null;
            Array.from(cards).forEach(c => {
                if (c.getAttribute('data-key') === activeKey) activeCard = c;
            });
            if (activeCard) {
                activeCard.classList.add('active-card');
                Array.from(cards).forEach(c => {
                    if (c !== activeCard) c.style.display = 'none';
                });
            }
        }
    }
}

async function confirmReceiveItem(item) {
    const result = await Swal.fire({
        title: 'ยืนยันรับ?',
        text: `คุณต้องการยืนยันรับรายการ: ${item['Spare Part Name'] || ''}?`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'ยืนยันรับ',
        cancelButtonText: 'ยกเลิก'
    });

    if (!result.isConfirmed) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const recripteName = currentUser.name || currentUser.IDRec || 'Unknown';
    const recripteDate = new Date();
    const recripteDateStr = recripteDate.toLocaleString('en-GB');

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
            const getVal = (obj, keyName) => { if (!obj) return ''; const found = Object.keys(obj).find(k => k.toLowerCase().trim() === keyName.toLowerCase()); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...item, 'Recripte': recripteName, 'RecripteDate': recripteDateStr, 'user': recripteName,
        'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || ''
    };

    // Optimistic Update
    item['Recripte'] = recripteName;
    item['RecripteDate'] = recripteDateStr;

    // Queue
    SaveQueue.add(payload);

    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'รับรายการเรียบร้อย' });

    if (currentDetailContext) {
        const { slip, targetReceiver, tabKey } = currentDetailContext;
        renderTopLevelDetailTable(tabKey, slip, targetReceiver);
        if (tabKey === 'navanakorn') renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
        if (tabKey === 'vibhavadi') renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');

        const activeKey = slip + '|' + targetReceiver;
        const deckContainer = document.getElementById(tabKey === 'navanakorn' ? 'navaNakornDeck' : 'vibhavadiDeck');
        if (deckContainer) {
            const cards = deckContainer.getElementsByClassName('deck-card');
            let activeCard = null;
            Array.from(cards).forEach(c => {
                if (c.getAttribute('data-key') === activeKey) activeCard = c;
            });
            if (activeCard) {
                activeCard.classList.add('active-card');
                Array.from(cards).forEach(c => {
                    if (c !== activeCard) c.style.display = 'none';
                });
            }
        }
    }
}

function populateDeckReceiverFilter(targetPlantCenter, filterId) {
    const filterSelect = document.getElementById(filterId);
    if (!filterSelect) return;

    const currentSelection = filterSelect.value;
    const receivers = new Set();
    globalBookingData.forEach(item => {
        const pcCode = String(item['Plantcenter'] || '').trim();
        const normalize = (s) => s.replace(/^0+/, '');
        if (normalize(pcCode) !== normalize(targetPlantCenter)) return;
        const receiver = item['Claim Receiver'] || item.person;
        if (receiver) receivers.add(receiver);
    });

    const sortedReceivers = Array.from(receivers).sort();
    filterSelect.innerHTML = '<option value="">All Receivers</option>';
    sortedReceivers.forEach(receiver => {
        const option = document.createElement('option');
        option.value = receiver;
        option.textContent = receiver;
        filterSelect.appendChild(option);
    });

    if (sortedReceivers.includes(currentSelection)) {
        filterSelect.value = currentSelection;
    }
}