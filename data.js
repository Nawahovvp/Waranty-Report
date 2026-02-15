function processPartsData(partsData) {
    const partsMap = new Map();
    const partsNameMap = new Map();
    if (partsData && partsData.length > 0) {
        const headers = Object.keys(partsData[0]);
        const codeKey = headers.find(h => h.toLowerCase().includes('spare') && h.toLowerCase().includes('code')) || 'Spare Part Code';
        const personKey = headers.find(h => h.toLowerCase().includes('person')) || 'Person';
        const nameKey = headers.find(h => h.toLowerCase().includes('spare') && h.toLowerCase().includes('name')) || 'Spare Part Name';

        partsData.forEach(row => {
            const code = row[codeKey] ? String(row[codeKey]).trim() : '';
            const person = row[personKey] ? String(row[personKey]).trim() : '';
            const name = row[nameKey] ? String(row[nameKey]).trim() : '';
            if (code) {
                partsMap.set(code, person);
                if (name) partsNameMap.set(code, name);
            }
        });
    }
    return { partsMap, partsNameMap };
}

function processWorkFilterData(workFilterData) {
    const workFilterMap = new Map();
    const workFilterDigitMap = new Map();
    let wfCodeKey = 'Work Order Number 1';
    let wfProductKey = 'Product';
    let wfCINameKey = 'CI Name';
    let wfProblemKey = 'Problem';
    let wfProductTypeKey = 'Product Type';

    if (workFilterData.length > 0) {
        const headers = Object.keys(workFilterData[0]);
        wfCodeKey = headers.find(h => h.toLowerCase().includes('work') && (h.toLowerCase().includes('order') || h.toLowerCase().includes('no'))) || headers.find(h => h.toLowerCase().includes('key')) || 'Work Order Number 1';
        wfProductKey = headers.find(h => h.toLowerCase().includes('product') && !h.toLowerCase().includes('type')) || 'Product';
        
        wfCINameKey = headers.find(h => h.trim().toLowerCase() === 'ci name') || wfCINameKey;
        wfProblemKey = headers.find(h => h.trim().toLowerCase() === 'problem') || wfProblemKey;
        wfProductTypeKey = headers.find(h => h.trim().toLowerCase() === 'product type') || wfProductTypeKey;
    }

    workFilterData.forEach(row => {
        const rawKey = row[wfCodeKey];
        const key = rawKey?.trim().toLowerCase();
        if (key) workFilterMap.set(key, row);
        if (rawKey) {
            const digits = String(rawKey).replace(/\D/g, '');
            if (digits.length > 3) workFilterDigitMap.set(digits, row);
        }
    });
    return { workFilterMap, workFilterDigitMap, wfProductKey, wfCINameKey, wfProblemKey, wfProductTypeKey };
}

function processTechnicianData(technicianData) {
    const technicianMap = new Map();
    if (technicianData) {
        technicianData.forEach(row => {
            let id = row['IDEmployee'];
            if (id) {
                id = String(id).trim();
                if (!technicianMap.has(id)) {
                    technicianMap.set(id, row['Phone'] || '');
                }
            }
        });
    }
    return technicianMap;
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

        const { partsMap, partsNameMap } = processPartsData(partsData);
        const { workFilterMap, workFilterDigitMap, wfProductKey, wfCINameKey, wfProblemKey, wfProductTypeKey } = processWorkFilterData(workFilterData);
        const technicianMap = processTechnicianData(technicianData);

        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
        const userPlant = currentUser['Plant'] ? currentUser['Plant'].toString().trim().padStart(4, '0') : null;

        // Detect Scrap Headers early for enrichment
        let scrapCodeKey = 'Spare Part Code';
        let scrapWOKey = 'work order';
        let scrapDateReceivedKey = 'วันที่รับซาก';
        let scrapReceiverKey = 'ผู้รับซาก';
        let scrapKeepKey = 'Keep';

        if (scrapData && scrapData.length > 0) {
            const scrapHeaders = Object.keys(scrapData[0]);
            scrapCodeKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'spare part code') || scrapCodeKey;
            scrapWOKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'work order') || scrapWOKey;
            
            scrapDateReceivedKey = scrapHeaders.find(h => h.includes('วันที่รับซาก') || h.replace(/\s+/g, '').toLowerCase() === 'datereceived') || scrapDateReceivedKey;
            scrapReceiverKey = scrapHeaders.find(h => h.includes('ผู้รับซาก') || h.replace(/\s+/g, '').toLowerCase() === 'receiver') || scrapReceiverKey;
            scrapKeepKey = scrapHeaders.find(h => h.replace(/\s+/g, '').toLowerCase() === 'keep') || scrapKeepKey;
        }

        const scrapMap = new Map();
        if (scrapData) {
            scrapData.forEach(row => {
                const key = ((row[scrapWOKey] || '') + (row[scrapCodeKey] || '')).replace(/\s/g, '').toLowerCase();
                scrapMap.set(key, row);
            });
        }

        let warrantyWOKey = 'Work Order';
        let warrantyClaimReceiverKey = 'Claim Receiver';
        if (warrantyData.length > 0) {
            const wHeaders = Object.keys(warrantyData[0]);
            warrantyWOKey = wHeaders.find(h => h.toLowerCase().includes('work') && h.toLowerCase().includes('order')) || 'Work Order';
            warrantyClaimReceiverKey = wHeaders.find(h => h.toLowerCase().includes('claim') && h.toLowerCase().includes('receiver')) || 'Claim Receiver';
        }

        globalBookingData = warrantyData.map(item => {
            const wo = String(item[warrantyWOKey] || '').trim();
            const code = String(item['Spare Part Code'] || '').trim();
            
            if (warrantyClaimReceiverKey !== 'Claim Receiver' && item[warrantyClaimReceiverKey]) {
                item['Claim Receiver'] = item[warrantyClaimReceiverKey];
            }
            
            // Enrich from WorkFilter (fullRow)
            let wfRow = workFilterMap.get(wo.toLowerCase());
            if (!wfRow) {
                const digits = wo.replace(/\D/g, '');
                if (digits.length > 3) wfRow = workFilterDigitMap.get(digits);
            }

            if (wfRow) {
                if (!item['CI Name']) item['CI Name'] = wfRow[wfCINameKey];
                if (!item['Problem']) item['Problem'] = wfRow[wfProblemKey];
                if (!item['Product Type']) item['Product Type'] = wfRow[wfProductTypeKey];
                if (!item['Product']) item['Product'] = wfRow[wfProductKey];
            }

            // Enrich from Scrap Data
            const scrapKey = (wo + code).replace(/\s/g, '').toLowerCase();
            const scrapRow = scrapMap.get(scrapKey);
            if (scrapRow) {
                if (!item['วันที่รับซาก']) item['วันที่รับซาก'] = scrapRow[scrapDateReceivedKey];
                if (!item['ผู้รับซาก']) item['ผู้รับซาก'] = scrapRow[scrapReceiverKey];
                if (!item['Keep']) item['Keep'] = scrapRow[scrapKeepKey];
            }

            if (!item['Product']) {
                // Fallback logic if still missing (already handled above via wfRow, but keeping safe)
                // ... (Logic merged into wfRow block above)
            }
            return item;
        });

        populateBookingFilter();
        renderBookingTable();

        const savedStatusMap = new Map();
        warrantyData.forEach(row => {
            const key = row['KEY'] ? row['KEY'].toString().trim() : '';
            if (key) {
                savedStatusMap.set(key, row['ActionStatus'] || row['Warranty Action'] || 'บันทึกแล้ว');
            }
        });

        let scrapNameKey = 'Spare Part Name';

        if (scrapData && scrapData.length > 0) {
            const scrapHeaders = Object.keys(scrapData[0]);
            scrapNameKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'spare part name') || scrapNameKey;
        }

        fullData = scrapData.map(scrapRow => {
            if (scrapRow['plant']) {
                scrapRow['plant'] = scrapRow['plant'].toString().trim().padStart(4, '0');
            }
            const workOrderKey = scrapRow[scrapWOKey]?.trim();
            const codeVal = scrapRow[scrapCodeKey] || '';
            const uniqueKey = (workOrderKey || '') + codeVal;
            const statusValue = savedStatusMap.get(uniqueKey) || '';
            const partCode = codeVal.trim();
            const personValue = partsMap.get(partCode) || '';

            if (partCode && partsNameMap.has(partCode)) {
                scrapRow[scrapNameKey] = partsNameMap.get(partCode);
            }
            if (scrapCodeKey !== 'Spare Part Code') scrapRow['Spare Part Code'] = scrapRow[scrapCodeKey];
            if (scrapNameKey !== 'Spare Part Name') scrapRow['Spare Part Name'] = scrapRow[scrapNameKey];
            if (scrapWOKey !== 'work order') scrapRow['work order'] = scrapRow[scrapWOKey];

            // Normalize enrichment keys for easier access in tab-main.js
            if (scrapDateReceivedKey !== 'วันที่รับซาก') scrapRow['วันที่รับซาก'] = scrapRow[scrapDateReceivedKey];
            if (scrapReceiverKey !== 'ผู้รับซาก') scrapRow['ผู้รับซาก'] = scrapRow[scrapReceiverKey];
            if (scrapKeepKey !== 'Keep') scrapRow['Keep'] = scrapRow[scrapKeepKey];

            const wfRow = workFilterMap.get(workOrderKey) || {};
            if (wfCINameKey !== 'CI Name') wfRow['CI Name'] = wfRow[wfCINameKey];
            if (wfProblemKey !== 'Problem') wfRow['Problem'] = wfRow[wfProblemKey];
            if (wfProductTypeKey !== 'Product Type') wfRow['Product Type'] = wfRow[wfProductTypeKey];

            return {
                scrap: scrapRow,
                fullRow: wfRow,
                status: statusValue,
                technicianPhone: technicianMap.get(String(scrapRow['รหัสช่าง'] || '').trim()) || '',
                person: personValue
            };
        });

        if (userPlant) {
            fullData = fullData.filter(item => item.scrap['plant'] === userPlant);
        }

        globalBookingData = globalBookingData.map(item => {
            const partCode = item['Spare Part Code']?.trim();
            const personValue = partCode ? (partsMap.get(partCode) || '') : '';
            return { ...item, person: personValue };
        });

        displayedData = [...fullData];
        populateFilters();
        sortDisplayedData();
        populateFilters();

        try { if (typeof populateBookingFilter === 'function') populateBookingFilter(); renderBookingTable(); } catch (e) { console.error("Error init Booking:", e); }
        try { if (typeof populateSupplierFilter === 'function') populateSupplierFilter(); renderSupplierTable(); } catch (e) { console.error("Error init Supplier:", e); }
        try {
            populateClaimSentFilter();
            renderClaimSentTable();
            renderHistoryTable();
            renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
            renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
        } catch (e) { console.error("Error init rest:", e); }

        sortDisplayedData();
        updateTotalQty();

        const tableHeader = document.getElementById('tableHeader');
        tableHeader.innerHTML = '';
        COLUMNS.forEach(col => {
            const th = document.createElement('th');
            if (col.header.includes('<')) th.innerHTML = col.header;
            else th.textContent = col.header;
            tableHeader.appendChild(th);
        });

        document.getElementById('loadingIndicator').style.display = 'none';
        currentPage = 1;
        const statusFilter = document.getElementById('statusFilter');
        if (statusFilter) statusFilter.value = 'Pending';
        applyFilters();

    } catch (err) {
        console.error(err);
        document.getElementById('loadingIndicator').innerHTML = `<div class="error">Error loading data: ${err.message}</div>`;
    }
}