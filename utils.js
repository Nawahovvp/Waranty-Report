async function postToGAS(payload) {
    if (!GAS_API_URL) throw new Error("GAS_API_URL is not configured.");
    return fetch(GAS_API_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });
}

async function fetchData(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Failed to fetch ${url}`);
        const csvText = await response.text();
        return new Promise((resolve, reject) => {
            Papa.parse(csvText, {
                header: true,
                skipEmptyLines: true,
                complete: (results) => resolve(results.data),
                error: (err) => reject(err)
            });
        });
    } catch (e) {
        console.error("Fetch Error:", e);
        return [];
    }
}

function renderGenericPagination(containerId, currentPage, totalPages, onPageChange) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    container.innerHTML = '';
    if (totalPages === 0 && containerId === 'paginationControls') return; 
    if (containerId !== 'paginationControls') container.style.display = 'flex';

    const createButton = (text, onClick, disabled = false, isActive = false) => {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.onclick = onClick;
        btn.disabled = disabled;
        if (isActive) btn.classList.add('active');
        return btn;
    };

    container.appendChild(createButton('<<', () => onPageChange(1), currentPage === 1));
    container.appendChild(createButton('<', () => onPageChange(currentPage - 1), currentPage === 1));
    container.appendChild(createButton(currentPage, () => { }, false, true));
    container.appendChild(createButton('>', () => onPageChange(currentPage + 1), currentPage >= totalPages));
    container.appendChild(createButton('>>', () => onPageChange(totalPages), currentPage >= totalPages));
}

function getComputedStatus(item) {
    if (item['ClaimSup'] && String(item['ClaimSup']).trim() !== '') {
        return '<span class="status-badge status-claim-sent">ส่งเคลมแล้ว</span>';
    }
    if (item['Recripte'] && String(item['Recripte']).trim() !== '') {
        return '<span class="status-badge status-waiting-claim">รอส่งเคลม</span>';
    }
    if (item['Booking Slip'] && String(item['Booking Slip']).trim() !== '') {
        return '<span class="status-badge status-transit">ระหว่างขนส่ง</span>';
    }
    return '<span class="status-badge status-local">คลังพื้นที่</span>';
}

function formatPhoneNumber(phone) {
    if (!phone) return '';
    let cleaned = ('' + phone).replace(/\D/g, '');
    if (cleaned.length === 10 && cleaned.startsWith('0')) {
        return cleaned.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }
    if (cleaned.length === 9) {
        cleaned = '0' + cleaned;
        return cleaned.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }
    return phone;
}

function getBookingKey(row) {
    return (row['Work Order'] || '') + '_' + (row['Spare Part Code'] || '') + '_' + (row['Old Material Code'] || '');
}