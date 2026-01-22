// Global context for detail view
let currentDetailContext = null;

function renderDeckView(targetPlantCode, containerId, tabKey) {
    const deckContainer = document.getElementById(containerId);
    if (!deckContainer) return;
    deckContainer.innerHTML = '';

    const uniqueMap = new Map();

    // Sort by Booking Date Descending
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
        const normalize = (s) => s.replace(/^0+/, '');
        
        if (!slip) return;
        if (normalize(pcCode) !== normalize(targetPlantCode)) return;

        const receiverVal = item['Claim Receiver'] || item.person || 'Unknown';
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

    if (uniqueMap.size === 0) {
        deckContainer.innerHTML = '<p style="text-align: center; color: var(--text-secondary); grid-column: 1/-1;">No booking records found for this plant.</p>';
        return;
    }

    uniqueMap.forEach((item, key) => {
        const slip = item['Booking Slip'];
        const date = item['Booking Date'];
        const plantCode = item['Plant'] || '';
        const plantName = PLANT_MAPPING[plantCode] || plantCode;
        const receiver = item._effectiveReceiver;
        const count = item._count || 1;
        const recCount = item._recripteCount || 0;

        let statusText = 'ระหว่างขนส่ง';
        let statusBg = '#fef3c7'; // yellow
        let statusColor = '#92400e';

        if (recCount === 0) { statusText = 'ระหว่างขนส่ง'; statusBg = '#fef3c7'; statusColor = '#92400e'; }
        else if (recCount < count) { statusText = 'ระหว่างตรวจรับ'; statusBg = '#e0f2fe'; statusColor = '#0369a1'; }
        else if (recCount === count) { statusText = 'เสร็จสิ้น'; statusBg = '#dcfce7'; statusColor = '#166534'; }

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
        deckWrapper.style.height = '100%';
        deckWrapper.style.flex = '';
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
        tableWrapper.style.height = 'auto';
        tableWrapper.style.flex = '1';

        // Set Context and Render
        currentDetailContext = { tabKey, slip, targetReceiver };
        renderTopLevelDetailTable(tabKey, slip, targetReceiver);
    }
}

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
    header.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
            <span>Booking Details: ${slip} (${targetReceiver})</span>
            <button class="btn btn-primary bulk-save-btn" style="display: none;" onclick="saveBulkReviewItems(this)">
                ยืนยันรับ
            </button>
        </div>
    `;

    // Render Table Header
    thead.innerHTML = '';
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
                const hasBookingSlip = item['Booking Slip'] && String(item['Booking Slip']).trim() !== '';
                const hasRecripte = item['Recripte'] && String(item['Recripte']).trim() !== '';
                const disabledAttr = hasRecripte ? 'disabled' : '';

                td.innerHTML = `<input type="checkbox" class="review-checkbox" value="${index}" onchange="handleReviewCheckboxChange(this)" ${disabledAttr}>`;
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

function toggleAllReviewCheckboxes(source) {
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
    const saveBtn = wrapper.querySelector('.bulk-save-btn');
    if (saveBtn) {
        saveBtn.style.display = checkboxes.length > 0 ? 'block' : 'none';
        saveBtn.textContent = `ยืนยันรับ (${checkboxes.length})`;
    }
}

async function saveBulkReviewItems(btnElement) {
    if (!currentDetailContext) return;
    const { slip, targetReceiver, tabKey } = currentDetailContext;
    const detailData = globalBookingData.filter(item => {
        const itemReceiver = item['Claim Receiver'] || item.person || 'Unknown';
        return item['Booking Slip'] === slip && itemReceiver === targetReceiver;
    });
    const wrapper = btnElement.closest('.detail-content-wrapper') || btnElement.closest('[id$="TableWrapper"]');
    if (!wrapper) return;
    const checkboxes = wrapper.querySelectorAll('.review-checkbox:checked');
    const selectedItems = [];
    checkboxes.forEach(cb => {
        const idx = parseInt(cb.value);
        if (detailData[idx]) selectedItems.push(detailData[idx]);
    });
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
    let dateReceived = item['Date Received'];
    let receiver = item['Receiver'] || item['receiver'];
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
             if (!dateReceived) dateReceived = getVal(match.scrap, 'Date Received');
             if (!receiver) receiver = getVal(match.scrap, 'Receiver');
             if (!keep) keep = getVal(match.scrap, 'Keep');
             if (!ciName) ciName = match.fullRow['CI Name'] || '';
             if (!problem) problem = match.fullRow['Problem'] || '';
             if (!productType) productType = match.fullRow['Product Type'] || '';
         }
    }
    
    const payload = { ...item, 'Recripte': '', 'RecripteDate': '', 'user': currentUser.IDRec || 'Unknown',
        'Date Received': dateReceived || '', 'Receiver': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || '' };

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
    let dateReceived = item['Date Received'];
    let receiver = item['Receiver'] || item['receiver'];
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
             if (!dateReceived) dateReceived = getVal(match.scrap, 'Date Received');
             if (!receiver) receiver = getVal(match.scrap, 'Receiver');
             if (!keep) keep = getVal(match.scrap, 'Keep');
             if (!ciName) ciName = match.fullRow['CI Name'] || '';
             if (!problem) problem = match.fullRow['Problem'] || '';
             if (!productType) productType = match.fullRow['Product Type'] || '';
         }
    }

    const payload = { ...item, 'Recripte': recripteName, 'RecripteDate': recripteDateStr, 'user': recripteName,
        'Date Received': dateReceived || '', 'Receiver': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || '' };
    
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