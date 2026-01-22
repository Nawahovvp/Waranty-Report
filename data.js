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

    if (workFilterData.length > 0) {
        const headers = Object.keys(workFilterData[0]);
        wfCodeKey = headers.find(h => h.toLowerCase().includes('work') && (h.toLowerCase().includes('order') || h.toLowerCase().includes('no'))) || headers.find(h => h.toLowerCase().includes('key')) || 'Work Order Number 1';
        wfProductKey = headers.find(h => h.toLowerCase().includes('product')) || 'Product';
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
    return { workFilterMap, workFilterDigitMap, wfProductKey };
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
        const { workFilterMap, workFilterDigitMap, wfProductKey } = processWorkFilterData(workFilterData);
        const technicianMap = processTechnicianData(technicianData);

        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
        const userPlant = currentUser['Plant'] ? currentUser['Plant'].toString().trim().padStart(4, '0') : null;

        let warrantyWOKey = 'Work Order';
        if (warrantyData.length > 0) {
            const wHeaders = Object.keys(warrantyData[0]);
            warrantyWOKey = wHeaders.find(h => h.toLowerCase().includes('work') && h.toLowerCase().includes('order')) || 'Work Order';
        }

        globalBookingData = warrantyData.filter(row => {
            const action = row['ActionStatus'] || row['Warranty Action'] || '';
            return action === 'เคลมประกัน';
        }).map(item => {
            if (!item['Product']) {
                const rawWO = String(item[warrantyWOKey] || '');
                const wo = rawWO.trim().toLowerCase();
                if (wo && workFilterMap.has(wo)) {
                    item['Product'] = workFilterMap.get(wo)[wfProductKey];
                } else {
                    const digits = rawWO.replace(/\D/g, '');
                    if (digits.length > 3 && workFilterDigitMap.has(digits)) {
                        item['Product'] = workFilterDigitMap.get(digits)[wfProductKey];
                    }
                }
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

        let scrapCodeKey = 'Spare Part Code';
        let scrapNameKey = 'Spare Part Name';
        let scrapWOKey = 'work order';

        if (scrapData && scrapData.length > 0) {
            const scrapHeaders = Object.keys(scrapData[0]);
            scrapCodeKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'spare part code') || scrapCodeKey;
            scrapNameKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'spare part name') || scrapNameKey;
            scrapWOKey = scrapHeaders.find(h => h.trim().toLowerCase() === 'work order') || scrapWOKey;
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

            return {
                scrap: scrapRow,
                fullRow: workFilterMap.get(workOrderKey) || {},
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
        renderTable();

    } catch (err) {
        console.error(err);
        document.getElementById('loadingIndicator').innerHTML = `<div class="error">Error loading data: ${err.message}</div>`;
    }
}