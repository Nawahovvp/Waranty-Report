function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(btn => { btn.classList.remove('active'); });
    const tabs = document.querySelectorAll('.tab-btn');
    document.querySelectorAll('.tab-content').forEach(el => el.style.display = 'none');
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));

    if (tabId === 'main') {
        tabs[0].classList.add('active');
        document.getElementById('tab-content-main').style.display = 'flex';
        document.getElementById('tab-content-main').classList.add('active');
    } else if (tabId === 'booking') {
        tabs[1].classList.add('active');
        document.getElementById('tab-content-booking').style.display = 'flex';
        document.getElementById('tab-content-booking').classList.add('active');
        populateBookingFilter(); renderBookingTable();
    } else if (tabId === 'navanakorn') {
        tabs[2].classList.add('active');
        document.getElementById('tab-content-navanakorn').style.display = 'block';
        document.getElementById('tab-content-navanakorn').classList.add('active');
        backToDeck('navanakorn'); renderDeckView('0301', 'navaNakornDeck', 'navanakorn');
    } else if (tabId === 'vibhavadi') {
        tabs[3].classList.add('active');
        document.getElementById('tab-content-vibhavadi').style.display = 'block';
        document.getElementById('tab-content-vibhavadi').classList.add('active');
        backToDeck('vibhavadi'); renderDeckView('0326', 'vibhavadiDeck', 'vibhavadi');
    } else if (tabId === 'supplier') {
        tabs[4].classList.add('active');
        document.getElementById('tab-content-supplier').style.display = 'flex';
        document.getElementById('tab-content-supplier').classList.add('active');
        populateSupplierFilter(); renderSupplierTable();
    } else if (tabId === 'claimSent') {
        tabs[5].classList.add('active');
        document.getElementById('tab-content-claimSent').style.display = 'flex';
        document.getElementById('tab-content-claimSent').classList.add('active');
        populateClaimSentFilter(); renderClaimSentTable();
    } else if (tabId === 'history') {
        tabs[6].classList.add('active');
        document.getElementById('tab-content-history').style.display = 'flex';
        document.getElementById('tab-content-history').classList.add('active');
        renderHistoryTable();
    }
}

initAuth();