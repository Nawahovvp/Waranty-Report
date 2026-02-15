// Sidebar Functions
function openSidebar() {
    document.getElementById('sidebar').classList.add('active');
    document.getElementById('sidebarOverlay').classList.add('active');
    document.body.style.overflow = 'hidden';
}

function closeSidebar() {
    document.getElementById('sidebar').classList.remove('active');
    document.getElementById('sidebarOverlay').classList.remove('active');
    document.body.style.overflow = '';
}

function switchTab(tabId) {
    // Update sidebar menu items
    document.querySelectorAll('.sidebar-item').forEach(item => {
        item.classList.remove('active');
    });

    // Update old tab buttons (for backward compatibility)
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });

    const tabs = document.querySelectorAll('.tab-btn');
    document.querySelectorAll('.tab-content').forEach(el => el.style.display = 'none');
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));

    // Function to update page title
    function updatePageTitle(title) {
        const pageTitleElement = document.getElementById('currentPageTitle');
        if (pageTitleElement) {
            pageTitleElement.textContent = title;
        }
    }

    if (tabId === 'main') {
        if (tabs[0]) tabs[0].classList.add('active');
        document.querySelectorAll('.sidebar-item')[0]?.classList.add('active');
        document.getElementById('tab-content-main').style.display = 'flex';
        document.getElementById('tab-content-main').classList.add('active');
        updatePageTitle('บันทึกข้อมูล (Save Data)');
    } else if (tabId === 'booking') {
        if (tabs[1]) tabs[1].classList.add('active');
        document.querySelectorAll('.sidebar-item')[1]?.classList.add('active');
        document.getElementById('tab-content-booking').style.display = 'flex';
        document.getElementById('tab-content-booking').classList.add('active');
        const statusFilter = document.getElementById('bookingStatusFilter');
        if (statusFilter) statusFilter.value = 'TodayOrLocal';
        if (typeof currentBookingPage !== 'undefined') currentBookingPage = 1;
        populateBookingFilter(); renderBookingTable();
        updatePageTitle('บันทึกจองรถส่งคลังกลาง');
    } else if (tabId === 'navanakorn') {
        if (tabs[2]) tabs[2].classList.add('active');
        document.querySelectorAll('.sidebar-item')[2]?.classList.add('active');
        document.getElementById('tab-content-navanakorn').style.display = 'flex';
        document.getElementById('tab-content-navanakorn').style.flexDirection = 'column';
        document.getElementById('tab-content-navanakorn').classList.add('active');
        backToDeck('navanakorn'); renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
        updatePageTitle('ใบจองรถนวนคร');
    } else if (tabId === 'vibhavadi') {
        if (tabs[3]) tabs[3].classList.add('active');
        document.querySelectorAll('.sidebar-item')[3]?.classList.add('active');
        document.getElementById('tab-content-vibhavadi').style.display = 'flex';
        document.getElementById('tab-content-vibhavadi').style.flexDirection = 'column';
        document.getElementById('tab-content-vibhavadi').classList.add('active');
        backToDeck('vibhavadi'); renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
        updatePageTitle('ใบจองรถวิภาวดี');
    } else if (tabId === 'supplier') {
        if (tabs[4]) tabs[4].classList.add('active');
        document.querySelectorAll('.sidebar-item')[4]?.classList.add('active');
        document.getElementById('tab-content-supplier').style.display = 'flex';
        document.getElementById('tab-content-supplier').classList.add('active');
        populateSupplierFilter(); renderSupplierTable();
        updatePageTitle('อะไหล่รอส่งSupplier');
    } else if (tabId === 'claimSent') {
        if (tabs[5]) tabs[5].classList.add('active');
        document.querySelectorAll('.sidebar-item')[5]?.classList.add('active');
        document.getElementById('tab-content-claimSent').style.display = 'flex';
        document.getElementById('tab-content-claimSent').classList.add('active');
        populateClaimSentFilter(); renderClaimSentTable();
        updatePageTitle('ส่งเคลมแล้ว');
    } else if (tabId === 'claimSummary') {
        if (tabs[6]) tabs[6].classList.add('active');
        document.querySelectorAll('.sidebar-item')[6]?.classList.add('active');
        document.getElementById('tab-content-claimSummary').style.display = 'flex';
        document.getElementById('tab-content-claimSummary').classList.add('active');
        renderClaimSummaryTable();
        updatePageTitle('สรุปเคลมประกัน');
    } else if (tabId === 'history') {
        if (tabs[7]) tabs[7].classList.add('active');
        document.querySelectorAll('.sidebar-item')[7]?.classList.add('active');
        document.getElementById('tab-content-history').style.display = 'flex';
        document.getElementById('tab-content-history').classList.add('active');
        renderHistoryTable();
        updatePageTitle('ประวัติการเคลม');
    }
}

initAuth();