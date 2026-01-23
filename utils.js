async function postToGAS(payload) {
    if (!GAS_API_URL) throw new Error("GAS_API_URL is not configured.");
    return fetch(GAS_API_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });
}

// Background Save Queue System
const SaveQueue = {
    queue: [],
    failedQueue: [],
    isProcessing: false,

    add: function(payload) {
        if (!payload._retries) payload._retries = 0;
        const key = payload['Key'] || payload['Work Order'] || payload['work order'] || 'Unknown';
        console.log(`[SaveQueue] 📥 ADD: Queueing item. Total Pending: ${this.queue.length + 1}. Key: ${key}`, payload);
        this.queue.push(payload);
        this.save();
        this.updateUI();
        this.process();
    },

    process: async function() {
        if (this.isProcessing || this.queue.length === 0) return;

        this.isProcessing = true;
        const payload = this.queue[0]; // Peek first item
        const key = payload['Key'] || payload['Work Order'] || payload['work order'] || 'Unknown';
        console.log(`[SaveQueue] 🔄 PROCESS: Sending item to GAS... (Attempt ${payload._retries + 1}) Key: ${key}`);

        try {
            await postToGAS(payload);
            console.log(`[SaveQueue] ✅ SUCCESS: Item sent. Key: ${key}`);
            this.queue.shift(); // Remove only after success
            this.save();
        } catch (error) {
            console.error("Background Save Error:", error);
            
            payload._retries = (payload._retries || 0) + 1;
            console.warn(`[SaveQueue] ⚠️ FAIL: Error sending item. Key: ${key}. Retry count: ${payload._retries}`);
            
            if (payload._retries < 3) {
                // Retry logic: Wait 2 seconds then try again
                await new Promise(resolve => setTimeout(resolve, 2000));
            } else {
                // Max retries reached. Move to failed queue.
                console.error(`[SaveQueue] ❌ ABORT: Max retries reached. Moving to failed queue. Key: ${key}`);
                this.failedQueue.push(this.queue.shift());
                this.save();
                const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 4000, timerProgressBar: true });
                Toast.fire({ icon: 'error', title: 'Save Failed', text: 'Moved to failed queue.' });
            }
        } finally {
            this.isProcessing = false;
            this.updateUI();
            if (this.queue.length > 0) {
                this.process(); // Process next item
            }
        }
    },

    retryFailed: function() {
        while (this.failedQueue.length > 0) {
            const item = this.failedQueue.shift();
            item._retries = 0;
            this.queue.push(item);
        }
        this.save();
        this.updateUI();
        this.process();
    },

    save: function() {
        try {
            localStorage.setItem('WARRANTY_SAVE_QUEUE', JSON.stringify({
                queue: this.queue,
                failedQueue: this.failedQueue
            }));
        } catch (e) { console.error("Error saving queue to storage:", e); }
    },

    init: function() {
        try {
            const stored = localStorage.getItem('WARRANTY_SAVE_QUEUE');
            if (stored) {
                const data = JSON.parse(stored);
                this.queue = data.queue || [];
                this.failedQueue = data.failedQueue || [];
                if (this.queue.length > 0 || this.failedQueue.length > 0) {
                    console.log(`[SaveQueue] Restored ${this.queue.length} pending items from storage.`);
                    this.updateUI();
                    if (this.queue.length > 0) this.process();
                }
            }
        } catch (e) { console.error("Error loading queue from storage:", e); }
    },

    updateUI: function() {
        const container = document.getElementById('saveQueueStatus');
        const text = document.getElementById('saveQueueText');
        if (!container || !text) return;

        const pending = this.queue.length;
        const failed = this.failedQueue.length;

        if (pending === 0 && failed === 0 && !this.isProcessing) {
            container.style.display = 'none';
        } else {
            container.style.display = 'flex'; // Handled by classes mostly, but ensure visibility
            
            if (failed > 0) {
                container.className = 'header-save-status failed';
                text.textContent = failed;
                container.title = `Failed: ${failed} items. Click to Retry.`;
                container.onclick = () => SaveQueue.retryFailed();
            } else {
                container.className = 'header-save-status saving';
                text.textContent = pending;
                container.title = `Saving... ${pending} items pending`;
                container.onclick = null;
            }
        }
    }
};

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
    const action = item['Warranty Action'] || item['ActionStatus'];
    if (action && action !== 'เคลมประกัน' && action !== 'Pending' && action !== 'บันทึกแล้ว') {
        let style = 'background-color: #f3f4f6; color: #1f2937;';
        if (action === 'Sworp') style = 'background-color: #e0e7ff; color: #3730a3;';
        else if (action === 'หมดประกัน') style = 'background-color: #ffedd5; color: #9a3412;';
        else if (action === 'ชำรุด') style = 'background-color: #fce7f3; color: #9d174d;';
        return `<span class="status-badge" style="${style}">${action}</span>`;
    }

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