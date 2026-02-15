let editingItem = null;

function openWorkOrderModal(item, context = 'edit') {
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

    const cards = document.querySelectorAll('#workOrderModal .status-card');
    cards.forEach(card => {
        card.classList.remove('active');
        if (card.getAttribute('onclick').includes(`'${currentAction}'`)) {
            card.classList.add('active');
        }
    });
    document.getElementById('woDetail_Note').value = item['Note'] || '';

    // Context Logic
    const btnSave = document.getElementById('btnSaveWorkOrder');
    const btnDecision = document.getElementById('warrantyDecisionButtons');
    const btnCancel = document.getElementById('btnCancelWarranty');
    const btnDeleteBooking = document.getElementById('btnDeleteBookingSlip');

    if (btnDeleteBooking) {
        const userStatus = getCurrentUserStatus();
        const hasBookingSlip = item['Booking Slip'] && String(item['Booking Slip']).trim() !== '';
        // Show only if User is None AND Item has Booking Slip
        if (userStatus === 'None' && hasBookingSlip) btnDeleteBooking.style.display = 'inline-block';
        else btnDeleteBooking.style.display = 'none';
    }

    if (context === 'warrantyDecision') {
        if (btnSave) btnSave.style.display = 'none';
        if (btnDecision) btnDecision.style.display = 'flex';

        // Show Cancel button ONLY if already decided ('เคลมสำเร็จ' or 'ไม่รับประกัน')
        if (currentAction === 'เคลมสำเร็จ' || currentAction === 'ไม่รับประกัน') {
            if (btnCancel) btnCancel.style.display = 'inline-flex';
        } else {
            if (btnCancel) btnCancel.style.display = 'none';
        }
    } else {
        if (btnSave) btnSave.style.display = 'inline-block';
        if (btnDecision) btnDecision.style.display = 'none';
    }

    document.getElementById('workOrderModal').style.display = 'flex';
}

async function handleWarrantyDecision(decision) {
    if (!editingItem) return;

    const note = document.getElementById('woDetail_Note').value.trim();

    if (decision === 'ไม่รับประกัน' && !note) {
        Swal.fire({
            icon: 'warning',
            title: 'Required',
            text: 'กรุณาระบุหมายเหตุ (Note) สำหรับกรณีไม่รับประกัน',
            confirmButtonColor: '#f59e0b'
        });
        return;
    }

    // Cancel Logic
    if (decision === 'ยกเลิก') {
        document.getElementById('woDetail_ActionValue').value = 'เคลมประกัน';
        document.getElementById('woDetail_Note').value = ''; // Clear UI Note
        editingItem['Warranty Action'] = 'เคลมประกัน';
        editingItem['Note'] = ''; // Clear Item Note
    } else {
        document.getElementById('woDetail_ActionValue').value = decision;
    }

    // Call Save
    await saveWorkOrderDetail(decision === 'ยกเลิก'); // Pass flag
}

function closeWorkOrderModal() {
    document.getElementById('workOrderModal').style.display = 'none';
    editingItem = null;
}

async function deleteBookingSlipFromModal() {
    if (!editingItem) return;

    const result = await Swal.fire({
        title: 'ยืนยันการลบใบจองรถ?',
        text: "คุณต้องการลบข้อมูลใบจองรถและวันที่จองรถออกจากรายการนี้ใช่หรือไม่?",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d33',
        confirmButtonText: 'ใช่, ลบเลย',
        cancelButtonText: 'ยกเลิก'
    });

    if (!result.isConfirmed) return;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const userName = currentUser.name || currentUser.IDRec || 'Unknown';

    // Robust Data Lookup
    let dateReceived = editingItem['วันที่รับซาก'];
    let receiver = editingItem['ผู้รับซาก'];
    let keep = editingItem['Keep'];
    let ciName = editingItem['CI Name'];
    let problem = editingItem['Problem'];
    let productType = editingItem['Product Type'];

    if ((!dateReceived || !receiver || !keep || !ciName || !problem || !productType) && typeof fullData !== 'undefined') {
        const targetKey = ((editingItem['Work Order'] || '') + (editingItem['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
        const match = fullData.find(d => {
            const dKey = ((d.scrap['work order'] || '') + (d.scrap['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
            return dKey === targetKey;
        });
        if (match) {
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...editingItem,
        'Booking Slip': '',
        'Booking Date': '',
        'user': userName,
        'วันที่รับซาก': dateReceived || '',
        'ผู้รับซาก': receiver || '',
        'Keep': keep || '',
        'CI Name': ciName || '',
        'Problem': problem || '',
        'Product Type': productType || ''
    };

    // Optimistic Update
    editingItem['Booking Slip'] = '';
    editingItem['Booking Date'] = '';

    SaveQueue.add(payload);

    // Refresh Views
    if (typeof currentDetailContext !== 'undefined' && currentDetailContext) {
        if (typeof renderTopLevelDetailTable === 'function') {
             renderTopLevelDetailTable(currentDetailContext.tabKey, currentDetailContext.slip, currentDetailContext.targetReceiver);
        }
        if (currentDetailContext.tabKey === 'navanakorn' && typeof renderDeckView === 'function') renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
        if (currentDetailContext.tabKey === 'vibhavadi' && typeof renderDeckView === 'function') renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
    }
    
    if (typeof renderBookingTable === 'function') renderBookingTable();

    closeWorkOrderModal();

    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'ลบใบจองรถเรียบร้อย' });
}

async function saveWorkOrderDetail(isCancel = false) {
    if (!editingItem) return;
    const newSerial = document.getElementById('woDetail_SerialNumber').value;
    const newNote = document.getElementById('woDetail_Note').value;
    const newAction = document.getElementById('woDetail_ActionValue').value;
    const newQty = document.getElementById('woDetail_Qty').value;

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const recripteName = currentUser.name || currentUser.IDRec || 'Unknown';
    const recripteDate = new Date();

    // Robust Data Lookup for missing fields
    let dateReceived = editingItem['วันที่รับซาก'];
    let receiver = editingItem['ผู้รับซาก'];
    let keep = editingItem['Keep'];
    let ciName = editingItem['CI Name'];
    let problem = editingItem['Problem'];
    let productType = editingItem['Product Type'];

    if ((!dateReceived || !receiver || !keep || !ciName || !problem || !productType) && typeof fullData !== 'undefined') {
        const targetKey = ((editingItem['Work Order'] || '') + (editingItem['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
        const match = fullData.find(d => {
            const dKey = ((d.scrap['work order'] || '') + (d.scrap['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
            return dKey === targetKey;
        });

        if (match) {
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...editingItem,
        'Serial Number': newSerial,
        'Qty': newQty,
        'Note': newNote,
        'Warranty Action': newAction,
        'Recripte': recripteName,
        'RecripteDate': recripteDate.toLocaleString('en-GB'),
        'user': recripteName,
        'ActionStatus': newAction,
        'วันที่รับซาก': dateReceived || '',
        'ผู้รับซาก': receiver || '',
        'Keep': keep || '',
        'CI Name': ciName || '',
        'Problem': problem || '',
        'Product Type': productType || ''
    };

    // Logic for Finish and Datefinish
    if (newAction === 'เคลมสำเร็จ' || newAction === 'ไม่รับประกัน') {
        // payload['Finish'] = newAction; // REMOVED
        payload['Datefinish'] = new Date().toLocaleString('en-GB');
    } else if (isCancel) {
        payload['Datefinish'] = '';
        payload['Note'] = ''; // Explicitly clear Note in payload
    }

    // Update Claim Date if status changes to a final state (other than 'เคลมประกัน')
    if (newAction && newAction !== 'เคลมประกัน' && newAction !== 'Pending' && newAction !== 'บันทึกแล้ว') {
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        const formattedDate = `${day}/${month}/${year}`;
        payload['Claim Date'] = formattedDate;
        payload['ClaimSup'] = currentUser.IDRec || 'Unknown';
    }

    // Optimistic Update
    editingItem['Serial Number'] = newSerial;
    editingItem['Note'] = newNote;
    editingItem['Warranty Action'] = newAction;
    editingItem['Qty'] = newQty;
    editingItem['Recripte'] = recripteName;
    editingItem['RecripteDate'] = payload['RecripteDate'];
    // if (payload['Finish']) editingItem['Finish'] = payload['Finish']; // REMOVED
    if (payload['Datefinish'] !== undefined) editingItem['Datefinish'] = payload['Datefinish'];
    if (payload['Claim Date']) editingItem['Claim Date'] = payload['Claim Date'];
    if (payload['ClaimSup']) editingItem['ClaimSup'] = payload['ClaimSup'];

    // Queue
    SaveQueue.add(payload);

    if (currentDetailContext) {
        renderTopLevelDetailTable(currentDetailContext.tabKey, currentDetailContext.slip, currentDetailContext.targetReceiver);
        if (currentDetailContext.tabKey === 'navanakorn') {
            renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
        } else if (currentDetailContext.tabKey === 'vibhavadi') {
            renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
        }
    }

    // Refresh all main tables that might display this data
    if (typeof renderClaimSentTable === 'function') renderClaimSentTable();
    if (typeof renderSupplierTable === 'function') renderSupplierTable();
    if (typeof renderHistoryTable === 'function') renderHistoryTable();
    if (typeof renderBookingTable === 'function') renderBookingTable();
    closeWorkOrderModal();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Work Order updated' });
}

function selectWorkOrderAction(element, action) {
    const container = element.parentElement;
    const cards = container.querySelectorAll('.status-card');
    cards.forEach(card => card.classList.remove('active'));
    element.classList.add('active');
    document.getElementById('woDetail_ActionValue').value = action;
}

function openEditModal(item) {
    editingItem = item;
    const currentSerial = item.fullRow['Serial Number'] || '';
    document.getElementById('editSerialInput').value = currentSerial;

    const isLE = item.fullRow['Product'] === 'L&E';
    const headerTitle = document.getElementById('editModalTitle');
    const label = document.querySelector('#editModal label[for="editSerialInput"]');
    const input = document.getElementById('editSerialInput');

    if (headerTitle) headerTitle.textContent = isLE ? 'Edit LES Code' : 'Edit Serial Number';
    if (label) label.textContent = isLE ? 'LES Code' : 'Serial Number';
    if (input) input.placeholder = isLE ? 'Enter new LES Code' : 'Enter new serial number';

    const btnExample = document.getElementById('btnExampleLE');
    if (btnExample) {
        btnExample.style.display = isLE ? 'inline-block' : 'none';
        if (isLE) {
            btnExample.textContent = 'ตัวอย่าง';
            btnExample.style.fontSize = '1rem';
            btnExample.style.cursor = 'help';
            btnExample.style.textDecoration = 'underline';
            btnExample.style.color = '#3b82f6';
            btnExample.onclick = (e) => e.preventDefault();
            btnExample.onmouseenter = showLETooltip;
            btnExample.onmousemove = moveLETooltip;
            btnExample.onmouseleave = hideLETooltip;
        }
    }

    document.getElementById('editModal').style.display = 'flex';
}

function closeEditModal() {
    document.getElementById('editModal').style.display = 'none';
    editingItem = null;
}

function saveSerialNumber() {
    if (!editingItem) return;
    const newSerial = document.getElementById('editSerialInput').value.trim();
    editingItem.fullRow['Serial Number'] = newSerial;
    renderTable();
    closeEditModal();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Serial Number updated' });
}

function showLEExample() {
    Swal.fire({
        imageUrl: 'LEimage.jpg',
        imageAlt: 'L&E Example',
        width: '900px',
        showConfirmButton: false,
        showCloseButton: false,
        backdrop: false,
        timer: 0
    });
}
let tooltipImage = null;

function showLETooltip(e) {
    if (!tooltipImage) {
        tooltipImage = document.createElement('img');
        tooltipImage.src = 'LEimage.jpg';
        tooltipImage.style.position = 'fixed';
        tooltipImage.style.zIndex = '10000';
        tooltipImage.style.maxWidth = '600px';
        tooltipImage.style.boxShadow = '0 10px 15px -3px rgba(0, 0, 0, 0.1)';
        tooltipImage.style.borderRadius = '8px';
        tooltipImage.style.pointerEvents = 'none';
        document.body.appendChild(tooltipImage);
    }
    tooltipImage.style.display = 'block';
    moveLETooltip(e);
}

function moveLETooltip(e) {
    if (!tooltipImage) return;
    const offset = 15;
    let left = e.clientX + offset;
    let top = e.clientY + offset;

    // Adjust if overflowing right edge
    if (left + tooltipImage.offsetWidth > window.innerWidth) {
        left = e.clientX - tooltipImage.offsetWidth - offset;
    }
    // Adjust if overflowing bottom edge
    if (top + tooltipImage.offsetHeight > window.innerHeight) {
        top = e.clientY - tooltipImage.offsetHeight - offset;
    }

    tooltipImage.style.left = `${left}px`;
    tooltipImage.style.top = `${top}px`;
}

function hideLETooltip() {
    if (tooltipImage) {
        tooltipImage.style.display = 'none';
    }
}

function selectReceiver(btn, name) {
    document.querySelectorAll('.receiver-btn').forEach(b => {
        b.style.backgroundColor = '#f1f5f9';
        b.style.color = '#64748b';
        b.style.border = '1px solid #e2e8f0';
    });
    btn.style.backgroundColor = '#3b82f6';
    btn.style.color = 'white';
    btn.style.border = 'none';
    document.getElementById('store_receiver').value = name;
}

function openStoreModal(item) {
    editingItem = item;
    const isLE = item.fullRow['Product'] === 'L&E';
    document.getElementById('store_receiver').value = '';
    document.querySelectorAll('.receiver-btn').forEach(b => {
        b.style.backgroundColor = '#f1f5f9';
        b.style.color = '#64748b';
        b.style.border = '1px solid #e2e8f0';
    });

    const receiverSection = document.getElementById('receiverSection');
    if (isLE) {
        receiverSection.style.display = 'block';
        const savedReceiver = item.fullRow['Claim Receiver'] || item['Claim Receiver'];
        const personReceiver = item.person;
        const defaultReceiver = savedReceiver || personReceiver || 'Poom';
        document.getElementById('store_receiver').value = defaultReceiver;
        document.querySelectorAll('.receiver-btn').forEach(b => {
            if (b.textContent === defaultReceiver) {
                b.style.backgroundColor = '#3b82f6';
                b.style.color = 'white';
                b.style.border = 'none';
            }
        });
    } else {
        receiverSection.style.display = 'none';
        const defaultReceiver = item.person || '';
        document.getElementById('store_receiver').value = defaultReceiver;
    }

    document.getElementById('store_workOrder').value = item.scrap['work order'] || '';
    document.getElementById('store_partCode').value = item.scrap['Spare Part Code'] || '';
    document.getElementById('store_partName').value = item.scrap['Spare Part Name'] || '';
    document.getElementById('store_qty').value = item.scrap['qty'] || '';
    document.getElementById('store_serial').value = item.fullRow['Serial Number'] || '';
    document.getElementById('store_code').value = item.fullRow['Store Code'] || '';
    document.getElementById('store_name').value = item.fullRow['Store Name'] || '';
    document.getElementById('store_techId').value = item.scrap['รหัสช่าง'] || '';
    document.getElementById('store_techName').value = item.scrap['ชื่อช่าง'] || '';
    document.getElementById('store_mobile').value = item.technicianPhone || '';
    const ciNameEl = document.getElementById('store_ciName');
    if (ciNameEl) ciNameEl.value = item.fullRow['CI Name'] || '';
    const productTypeEl = document.getElementById('store_productType');
    if (productTypeEl) productTypeEl.value = item.fullRow['Product Type'] || '';
    const problemEl = document.getElementById('store_problem');
    if (problemEl) problemEl.value = item.fullRow['Problem'] || '';

    const currentStatus = item.status || item.fullRow['ActionStatus'] || 'เคลมประกัน';
    setButtonActive(currentStatus);

    const btnSave = document.getElementById('btnSaveStore');
    const btnUpdate = document.getElementById('btnUpdateStore');
    const btnDelete = document.getElementById('btnDeleteStore');
    const btnCancelBooking = document.getElementById('btnCancelBookingSlipStore');

    if (btnCancelBooking) {
        // Show button if item has a booking slip (passed from handleBookingStatusClick)
        if (item.bookingSlip && String(item.bookingSlip).trim() !== '') {
            btnCancelBooking.style.display = 'inline-block';
        } else {
            btnCancelBooking.style.display = 'none';
        }
    }

    if (item.status) {
        btnSave.style.display = 'none';
        btnUpdate.style.display = 'inline-block';
        btnDelete.style.display = 'inline-block';
    } else {
        btnSave.style.display = 'inline-block';
        btnUpdate.style.display = 'none';
        btnDelete.style.display = 'none';
    }
    document.getElementById('storeModal').style.display = 'flex';
}

function selectAction(btn, value) {
    document.querySelectorAll('.status-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('store_actionStatus').value = value;
}

function setButtonActive(value) {
    document.getElementById('store_actionStatus').value = value;
    document.querySelectorAll('.status-btn').forEach(btn => {
        if (btn.textContent.trim() === value) btn.classList.add('active');
        else btn.classList.remove('active');
    });
}

function closeStoreModal() {
    document.getElementById('storeModal').style.display = 'none';
    editingItem = null;
}

function saveStoreDetail() {
    if (!editingItem) return;
    const serialInput = document.getElementById('store_serial').value.trim();
    const qtyInput = document.getElementById('store_qty').value;
    const mobileInput = document.getElementById('store_mobile').value;
    const actionStatus = document.getElementById('store_actionStatus').value;

    if (actionStatus === 'เคลมประกัน') {
        const isLE = editingItem.fullRow['Product'] === 'L&E';
        const minLength = isLE ? 6 : 8;
        const input = document.getElementById('store_serial');
        if (!serialInput || serialInput.length < minLength) {
            input.classList.add('input-flash-error');
            setTimeout(() => input.classList.remove('input-flash-error'), 1500);
            Swal.fire({ icon: 'warning', title: 'ตรวจสอบข้อมูล', text: `กรุณาตรวจสอบ Serial Number ให้ถูกต้อง (L&E > 5 หลัก, อื่นๆ >= 8 หลัก)` });
            return;
        }
        if (/[#$%\u0E00-\u0E7F\s]/.test(serialInput)) {
            input.classList.add('input-flash-error');
            setTimeout(() => input.classList.remove('input-flash-error'), 1500);
            Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `Serial Number ห้ามมีอักขระพิเศษ, ภาษาไทย หรือช่องว่าง` });
            return;
        }
        const keep = editingItem.scrap['Keep'] || '';
        if (keep === 'SCOTMAN' && !String(serialInput).startsWith('IDP020') && !String(serialInput).startsWith('NW508')) {
            input.classList.add('input-flash-error');
            setTimeout(() => input.classList.remove('input-flash-error'), 1500);
            Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `Keep เป็น SCOTMAN Serial Number ต้องขึ้นต้นด้วย IDP020 หรือ NW508` });
            return;
        }
        const product = editingItem.fullRow['Product'] || '';
        if (product === 'SCOTMAN' && !String(serialInput).startsWith('IDP020') && !String(serialInput).startsWith('NW508')) {
            input.classList.add('input-flash-error');
            setTimeout(() => input.classList.remove('input-flash-error'), 1500);
            Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `Product เป็น SCOTMAN Serial Number ต้องขึ้นต้นด้วย IDP020 หรือ NW508` });
            return;
        }
    }

    editingItem.fullRow['Serial Number'] = serialInput;
    editingItem.scrap['qty'] = qtyInput;
    editingItem.technicianPhone = mobileInput;
    editingItem.fullRow['ActionStatus'] = actionStatus;

    // Sync with globalBookingData if it exists
    renderTable();
    sendDatatoGAS(editingItem);

    // Update buttons to reflect saved state
    const btnSave = document.getElementById('btnSaveStore');
    const btnUpdate = document.getElementById('btnUpdateStore');
    const btnDelete = document.getElementById('btnDeleteStore');

    if (editingItem.status) {
        btnSave.style.display = 'none';
        btnUpdate.style.display = 'inline-block';
        btnDelete.style.display = 'inline-block';
    }

    // Refresh Booking Table if data was updated from there
    if (typeof renderBookingTable === 'function' && typeof globalBookingData !== 'undefined') {
        renderBookingTable();
    }

    closeStoreModal();
}

async function deleteStoreDetail() {
    if (!editingItem) return;
    const result = await Swal.fire({
        title: 'ยืนยันการลบ?', text: "คุณต้องการลบข้อมูลนี้ออกจากระบบใช่หรือไม่?", icon: 'warning',
        showCancelButton: true, confirmButtonColor: '#ef4444', confirmButtonText: 'ลบข้อมูล (Delete)'
    });
    if (!result.isConfirmed) return;

    const payload = {
        'work order': editingItem.scrap['work order'],
        'Spare Part Code': editingItem.scrap['Spare Part Code'],
        'operation': 'delete'
    };

    // Optimistic Update
    editingItem.status = '';
    if (editingItem.fullRow) editingItem.fullRow['ActionStatus'] = '';

    // Sync removal with globalBookingData
    if (typeof globalBookingData !== 'undefined' && globalBookingData.length > 0) {
        const targetKey = ((editingItem.scrap?.['work order'] || editingItem['Work Order'] || '') + (editingItem.scrap?.['Spare Part Code'] || editingItem['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();

        const bookingIndex = globalBookingData.findIndex(item => {
            const k = ((item['Work Order'] || '') + (item['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
            return k === targetKey;
        });

        if (bookingIndex !== -1) {
            globalBookingData.splice(bookingIndex, 1);
        }
    }

    SaveQueue.add(payload);
    renderTable();

    // Refresh Booking Table if data was deleted from there
    if (typeof renderBookingTable === 'function') {
        renderBookingTable();
    }

    closeStoreModal();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Record deleted' });
}

async function cancelBookingSlipFromStoreModal() {
    if (!editingItem) return;

    const result = await Swal.fire({
        title: 'ยกเลิกใบจองรถ?',
        text: "คุณต้องการลบข้อมูลใบจองรถและวันที่จองรถออกจากรายการนี้ใช่หรือไม่?",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d33',
        confirmButtonText: 'ใช่, ยกเลิก',
        cancelButtonText: 'ไม่'
    });

    if (!result.isConfirmed) return;

    // Find the item in globalBookingData to update
    const workOrder = (editingItem.scrap && editingItem.scrap['work order']) || editingItem['Work Order'] || '';
    const partCode = (editingItem.scrap && editingItem.scrap['Spare Part Code']) || editingItem['Spare Part Code'] || '';
    const targetKey = (workOrder + partCode).replace(/\s/g, '').toLowerCase();

    const globalItem = globalBookingData.find(row => {
        const rowKey = ((row['Work Order'] || '') + (row['Spare Part Code'] || '')).replace(/\s/g, '').toLowerCase();
        return rowKey === targetKey;
    });

    if (globalItem) {
        const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
        const payload = {
            ...globalItem,
            'Booking Slip': '',
            'Booking Date': '',
            'Plantcenter': '',
            'user': currentUser.IDRec || 'Unknown'
        };

        // Optimistic Update
        globalItem['Booking Slip'] = '';
        globalItem['Booking Date'] = '';
        globalItem['Plantcenter'] = '';

        SaveQueue.add(payload);

        if (typeof renderBookingTable === 'function') renderBookingTable();
        
        closeStoreModal();
        const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
        Toast.fire({ icon: 'success', title: 'ยกเลิกใบจองรถเรียบร้อย' });
    }
}

async function openBookingSlipModal(item) {
    const currentSlip = item['Booking Slip'] || '';
    const result = await Swal.fire({
        title: 'จัดการใบจองรถ',
        html: `<p>Current Slip: <b>${currentSlip || '-'}</b></p><p>เลือกการทำงาน:</p>`,
        showDenyButton: true, showCancelButton: true, confirmButtonText: 'แก้ไข (Edit)', denyButtonText: 'ลบข้อมูล (Delete)', cancelButtonText: 'ยกเลิก', denyButtonColor: '#d33'
    });

    if (result.isConfirmed) {
        const { value: newSlip } = await Swal.fire({
            title: 'แก้ไขเลขใบจองรถ', input: 'text', inputValue: currentSlip, showCancelButton: true, confirmButtonText: 'บันทึก',
            inputValidator: (value) => { if (!value) return 'กรุณากรอกข้อมูล!'; }
        });
        if (newSlip) {
            const d = new Date();
            const formattedDate = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
            updateBookingSlip(item, newSlip, formattedDate);
        }
    } else if (result.isDenied) {
        const confirmDelete = await Swal.fire({
            title: 'ยืนยันการลบ?', text: "ข้อมูลใบจองรถและวันที่จองจะถูกลบออก", icon: 'warning', showCancelButton: true, confirmButtonText: 'ลบข้อมูล', confirmButtonColor: '#d33'
        });
        if (confirmDelete.isConfirmed) updateBookingSlip(item, '', '', '');
    }
}

async function updateBookingSlip(item, newSlip, newDate, newPlantCenter = null) {
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
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...item,
        'Booking Slip': newSlip,
        'Booking Date': newDate,
        'Plantcenter': newPlantCenter !== null ? newPlantCenter : item['Plantcenter'],
        'user': JSON.parse(localStorage.getItem('currentUser') || '{}').IDRec || 'Unknown',
        'วันที่รับซาก': dateReceived || '',
        'ผู้รับซาก': receiver || '',
        'Keep': keep || '',
        'CI Name': ciName || '',
        'Problem': problem || '',
        'Product Type': productType || ''
    };
    payload['Key'] = (item['Work Order'] || '') + (item['Spare Part Code'] || '');

    // Optimistic
    item['Booking Slip'] = newSlip;
    item['Booking Date'] = newDate;
    if (newPlantCenter !== null) item['Plantcenter'] = newPlantCenter;

    SaveQueue.add(payload);
    renderBookingTable();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Booking Slip updated' });
}

async function openClaimReceiverModal(item) {
    const currentReceiver = item['Claim Receiver'] || '';
    const receiverOptions = ['Mai', 'Mon', 'Poom', 'Not'];
    let buttonsHtml = '<div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">';
    receiverOptions.forEach(opt => {
        let btnColor = '#f1f5f9'; let txtColor = '#64748b';
        if (opt === 'Mai') { btnColor = '#dbeafe'; txtColor = '#1e40af'; }
        if (opt === 'Mon') { btnColor = '#dcfce7'; txtColor = '#166534'; }
        if (opt === 'Poom') { btnColor = '#fae8ff'; txtColor = '#6b21a8'; }
        if (opt === 'Not') { btnColor = '#ffedd5'; txtColor = '#c2410c'; }
        const isSelected = (currentReceiver.toLowerCase() === opt.toLowerCase());
        const borderStyle = isSelected ? '2px solid var(--primary-color)' : '1px solid transparent';
        buttonsHtml += `<button type="button" class="swal2-confirm swal2-styled receiver-opt-btn" style="background-color: ${btnColor}; color: ${txtColor}; border: ${borderStyle}; margin: 5px; flex: 1 0 40%;" onclick="window.selectedClaimReceiver = '${opt}'; document.querySelectorAll('.receiver-opt-btn').forEach(b => b.style.border='1px solid transparent'); this.style.border='2px solid var(--primary-color)';">${opt}</button>`;
    });
    buttonsHtml += `<button type="button" class="swal2-confirm swal2-styled receiver-opt-btn" style="background-color: #fee2e2; color: #991b1b; border: 1px solid transparent; margin: 5px; flex: 1 0 40%;" onclick="window.selectedClaimReceiver = ''; document.querySelectorAll('.receiver-opt-btn').forEach(b => b.style.border='1px solid transparent'); this.style.border='2px solid var(--primary-color)';">ล้างค่า (Clear)</button></div>`;
    window.selectedClaimReceiver = currentReceiver;
    const result = await Swal.fire({ title: 'เลือกผู้รับงาน (Claim Receiver)', html: buttonsHtml, showCancelButton: true, confirmButtonText: 'บันทึก (Save)', cancelButtonText: 'ยกเลิก', preConfirm: () => window.selectedClaimReceiver });
    if (result.isConfirmed && result.value !== currentReceiver) updateClaimReceiver(item, result.value);
}

async function updateClaimReceiver(item, newReceiver) {
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
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...item, 'Claim Receiver': newReceiver,
        'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || ''
    };

    // Optimistic
    item['Claim Receiver'] = newReceiver;

    SaveQueue.add(payload);
    renderBookingTable();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Receiver updated' });
}

async function openMobileModal(item) {
    const { value: newMobile } = await Swal.fire({ title: 'เพิ่มเบอร์โทรศัพท์ (Add Mobile)', input: 'tel', inputLabel: 'Phone Number', inputPlaceholder: 'Enter phone number...', showCancelButton: true, confirmButtonText: 'บันทึก (Save)', cancelButtonText: 'ยกเลิก', inputValidator: (value) => { if (!value) return 'You need to write something!'; } });
    if (newMobile) updateMobile(item, newMobile);
}

async function updateMobile(item, newMobile) {
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
            const getVal = (obj, keyName) => { if (!obj) return ''; const cleanKey = keyName.replace(/\s+/g, '').toLowerCase(); const found = Object.keys(obj).find(k => k.replace(/\s+/g, '').toLowerCase() === cleanKey); return found ? obj[found] : ''; };
            if (!dateReceived) dateReceived = getVal(match.scrap, 'วันที่รับซาก') || getVal(match.scrap, 'Date Received');
            if (!receiver) receiver = getVal(match.scrap, 'ผู้รับซาก') || getVal(match.scrap, 'Receiver');
            if (!keep) keep = getVal(match.scrap, 'Keep');
            if (!ciName) ciName = match.fullRow['CI Name'] || '';
            if (!problem) problem = match.fullRow['Problem'] || '';
            if (!productType) productType = match.fullRow['Product Type'] || '';
        }
    }

    const payload = {
        ...item, 'Mobile': newMobile,
        'วันที่รับซาก': dateReceived || '', 'ผู้รับซาก': receiver || '', 'Keep': keep || '', 'CI Name': ciName || '', 'Problem': problem || '', 'Product Type': productType || ''
    };

    // Optimistic
    item['Mobile'] = newMobile;

    SaveQueue.add(payload);
    renderBookingTable();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Mobile updated' });
}

async function openMainPersonModal(item) {
    const currentReceiver = item.person || '';
    const receiverOptions = ['Mai', 'Mon', 'Poom', 'Not'];
    let buttonsHtml = '<div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">';
    receiverOptions.forEach(opt => {
        let btnColor = '#f1f5f9'; let txtColor = '#64748b';
        if (opt === 'Mai') { btnColor = '#dbeafe'; txtColor = '#1e40af'; }
        if (opt === 'Mon') { btnColor = '#dcfce7'; txtColor = '#166534'; }
        if (opt === 'Poom') { btnColor = '#fae8ff'; txtColor = '#6b21a8'; }
        if (opt === 'Not') { btnColor = '#ffedd5'; txtColor = '#c2410c'; }
        const isSelected = (currentReceiver.toLowerCase() === opt.toLowerCase());
        const borderStyle = isSelected ? '2px solid var(--primary-color)' : '1px solid transparent';
        buttonsHtml += `<button type="button" class="swal2-confirm swal2-styled receiver-opt-btn" style="background-color: ${btnColor}; color: ${txtColor}; border: ${borderStyle}; margin: 5px; flex: 1 0 40%;" onclick="window.selectedMainPerson = '${opt}'; document.querySelectorAll('.receiver-opt-btn').forEach(b => b.style.border='1px solid transparent'); this.style.border='2px solid var(--primary-color)';">${opt}</button>`;
    });
    buttonsHtml += '</div>';
    window.selectedMainPerson = currentReceiver;
    const result = await Swal.fire({ title: 'Select Person (Claim Receiver)', html: buttonsHtml, showCancelButton: true, confirmButtonText: 'Save', cancelButtonText: 'Cancel', preConfirm: () => window.selectedMainPerson });
    if (result.isConfirmed && result.value && result.value !== currentReceiver) saveMainPerson(item, result.value);
}

async function saveMainPerson(item, newPerson) {
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const payload = { ...item.scrap, ...item.fullRow, 'user': currentUser.IDRec || 'Unknown', 'ActionStatus': item.status || item.fullRow['ActionStatus'] || 'เคลมประกัน', 'Claim Receiver': newPerson, 'Person': newPerson };
    delete payload['Values'];

    // Optimistic
    item.person = newPerson;
    if (item.fullRow) item.fullRow['Claim Receiver'] = newPerson;

    SaveQueue.add(payload);
    renderTable();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Person saved' });
}

async function openMainMobileModal(item) {
    const currentMobile = item.technicianPhone || '';
    const { value: newMobile } = await Swal.fire({ title: 'Edit Mobile Number', input: 'tel', inputValue: currentMobile, inputLabel: 'Phone Number', inputPlaceholder: 'Enter phone number...', showCancelButton: true, confirmButtonText: 'Save', cancelButtonText: 'Cancel' });
    if (newMobile !== undefined && newMobile !== currentMobile) saveMainMobile(item, newMobile);
}

function saveMainMobile(item, newMobile) {
    item.technicianPhone = newMobile;
    renderTable();
    const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 1500, timerProgressBar: true });
    Toast.fire({ icon: 'success', title: 'Mobile updated' });
}