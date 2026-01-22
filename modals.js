let editingItem = null;

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

    const buttons = document.querySelectorAll('#workOrderModal .status-btn');
    buttons.forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('onclick').includes(`'${currentAction}'`)) {
            btn.classList.add('active');
        }
    });
    document.getElementById('woDetail_Note').value = item['Note'] || '';
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

    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const recripteName = currentUser.name || currentUser.IDRec || 'Unknown';
    const recripteDate = new Date();

    const payload = {
        ...editingItem,
        'Serial Number': newSerial,
        'Qty': newQty,
        'Note': newNote,
        'Warranty Action': newAction,
        'Recripte': recripteName,
        'RecripteDate': recripteDate.toLocaleString('en-GB'),
        'user': recripteName,
        'ActionStatus': newAction
    };

    Swal.fire({ title: 'Saving...', text: 'Updating Work Order details...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

    try {
        await postToGAS(payload);
        editingItem['Serial Number'] = newSerial;
        editingItem['Note'] = newNote;
        editingItem['Warranty Action'] = newAction;
        editingItem['Qty'] = newQty;
        editingItem['Recripte'] = recripteName;
        editingItem['RecripteDate'] = payload['RecripteDate'];

        if (currentDetailContext) {
            renderTopLevelDetailTable(currentDetailContext.tabKey, currentDetailContext.slip, currentDetailContext.targetReceiver);
            if (currentDetailContext.tabKey === 'navanakorn') {
                renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
            } else if (currentDetailContext.tabKey === 'vibhavadi') {
                renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
            }
        }
        closeWorkOrderModal();
        Swal.fire({ icon: 'success', title: 'Saved!', text: 'Work Order updated successfully.', timer: 1500, showConfirmButton: false });
    } catch (error) {
        console.error('Error saving Work Order:', error);
        Swal.fire({ icon: 'error', title: 'Error', text: 'Failed to save changes to Google Sheet.' });
    }
}

function selectWorkOrderAction(element, action) {
    const container = element.parentElement;
    const buttons = container.querySelectorAll('.status-btn');
    buttons.forEach(btn => btn.classList.remove('active'));
    element.classList.add('active');
    document.getElementById('woDetail_ActionValue').value = action;
}

function openEditModal(item) {
    editingItem = item;
    const currentSerial = item.fullRow['Serial Number'] || '';
    document.getElementById('editSerialInput').value = currentSerial;
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
        if (keep === 'SCOTMAN' && !String(serialInput).startsWith('NW508')) {
            input.classList.add('input-flash-error');
            setTimeout(() => input.classList.remove('input-flash-error'), 1500);
            Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ถูกต้อง', text: `Keep เป็น SCOTMAN Serial Number ต้องขึ้นต้นด้วย NW508` });
            return;
        }
    }

    editingItem.fullRow['Serial Number'] = serialInput;
    editingItem.scrap['qty'] = qtyInput;
    editingItem.technicianPhone = mobileInput;
    editingItem.fullRow['ActionStatus'] = actionStatus;
    renderTable();
    sendDatatoGAS(editingItem);
    closeStoreModal();
}

async function deleteStoreDetail() {
    if (!editingItem) return;
    const result = await Swal.fire({
        title: 'ยืนยันการลบ?', text: "คุณต้องการลบข้อมูลนี้ออกจากระบบใช่หรือไม่?", icon: 'warning',
        showCancelButton: true, confirmButtonColor: '#ef4444', confirmButtonText: 'ลบข้อมูล (Delete)'
    });
    if (!result.isConfirmed) return;

    Swal.fire({ title: 'Deleting...', text: 'Updating Google Sheet...', didOpen: () => Swal.showLoading() });
    const payload = {
        'work order': editingItem.scrap['work order'],
        'Spare Part Code': editingItem.scrap['Spare Part Code'],
        'operation': 'delete'
    };

    try {
        await postToGAS(payload);
        Swal.fire('Deleted!', 'Record has been deleted.', 'success');
        editingItem.status = '';
        if (editingItem.fullRow) editingItem.fullRow['ActionStatus'] = '';
        renderTable();
        closeStoreModal();
    } catch (error) {
        console.error(error);
        Swal.fire('Error', 'Failed to delete.', 'error');
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
    Swal.fire({ title: 'Updating...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const payload = { ...item, 'Booking Slip': newSlip, 'Booking Date': newDate, 'Plantcenter': newPlantCenter !== null ? newPlantCenter : item['Plantcenter'] };
    try {
        await postToGAS(payload);
        item['Booking Slip'] = newSlip;
        item['Booking Date'] = newDate;
        if (newPlantCenter !== null) item['Plantcenter'] = newPlantCenter;
        renderBookingTable();
        Swal.fire({ icon: 'success', title: 'Updated', timer: 1500, showConfirmButton: false });
    } catch (error) {
        console.error('Error updating receiver:', error);
        Swal.fire('Error', 'Failed to update', 'error');
    }
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
    Swal.fire({ title: 'Updating...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const payload = { ...item, 'Claim Receiver': newReceiver };
    try {
        await postToGAS(payload);
        item['Claim Receiver'] = newReceiver;
        renderBookingTable();
        Swal.fire({ icon: 'success', title: 'Updated', timer: 1500, showConfirmButton: false });
    } catch (error) {
        console.error('Error updating receiver:', error);
        Swal.fire('Error', 'Failed to update receiver', 'error');
    }
}

async function openMobileModal(item) {
    const { value: newMobile } = await Swal.fire({ title: 'เพิ่มเบอร์โทรศัพท์ (Add Mobile)', input: 'tel', inputLabel: 'Phone Number', inputPlaceholder: 'Enter phone number...', showCancelButton: true, confirmButtonText: 'บันทึก (Save)', cancelButtonText: 'ยกเลิก', inputValidator: (value) => { if (!value) return 'You need to write something!'; } });
    if (newMobile) updateMobile(item, newMobile);
}

async function updateMobile(item, newMobile) {
    Swal.fire({ title: 'Updating...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const payload = { ...item, 'Mobile': newMobile };
    try {
        await postToGAS(payload);
        item['Mobile'] = newMobile;
        renderBookingTable();
        Swal.fire({ icon: 'success', title: 'Updated', timer: 1500, showConfirmButton: false });
    } catch (error) {
        console.error('Error updating mobile:', error);
        Swal.fire('Error', 'Failed to update mobile', 'error');
    }
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
    Swal.fire({ title: 'Saving...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const currentUser = JSON.parse(localStorage.getItem('currentUser') || '{}');
    const payload = { ...item.scrap, ...item.fullRow, 'user': currentUser.IDRec || 'Unknown', 'ActionStatus': item.status || item.fullRow['ActionStatus'] || 'เคลมประกัน', 'Claim Receiver': newPerson, 'Person': newPerson };
    delete payload['Values'];
    try {
        await postToGAS(payload);
        item.person = newPerson;
        if (item.fullRow) item.fullRow['Claim Receiver'] = newPerson;
        renderTable();
        Swal.fire({ icon: 'success', title: 'Saved', timer: 1500, showConfirmButton: false });
    } catch (e) { console.error(e); Swal.fire('Error', 'Failed to save person', 'error'); }
}

async function openMainMobileModal(item) {
    const currentMobile = item.technicianPhone || '';
    const { value: newMobile } = await Swal.fire({ title: 'Edit Mobile Number', input: 'tel', inputValue: currentMobile, inputLabel: 'Phone Number', inputPlaceholder: 'Enter phone number...', showCancelButton: true, confirmButtonText: 'Save', cancelButtonText: 'Cancel' });
    if (newMobile !== undefined && newMobile !== currentMobile) saveMainMobile(item, newMobile);
}

function saveMainMobile(item, newMobile) {
    item.technicianPhone = newMobile;
    renderTable();
}