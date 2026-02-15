async function initAuth() {
    allEmployees = await fetchData(EMPLOYEE_SHEET_URL);

    const savedUser = localStorage.getItem('currentUser');
    if (savedUser) {
        let user = JSON.parse(savedUser);
        // Refresh user data from fresh fetch
        if (allEmployees && allEmployees.length > 0) {
            const freshUser = allEmployees.find(emp => String(emp['IDRec']) === String(user['IDRec']));
            if (freshUser) {
                console.log("[Auth] Refreshing user data from sheet...", freshUser);
                user = freshUser;
                localStorage.setItem('currentUser', JSON.stringify(user));
            }
        }
        showApp(user);
    } else {
        document.getElementById('loginOverlay').style.display = 'flex';
    }
}

function handleLogin() {
    const idInput = document.getElementById('idInput').value.trim();
    const passInput = document.getElementById('passInput').value.trim();
    const errorMsg = document.getElementById('loginError');

    if (!idInput || !passInput) {
        errorMsg.textContent = "Please enter ID and Password";
        errorMsg.style.display = 'block';
        return;
    }

    const user = allEmployees.find(emp => emp['IDRec'] === idInput);

    if (user) {
        const expectedPass = user['IDRec'].slice(-4);
        if (passInput === expectedPass) {
            localStorage.setItem('currentUser', JSON.stringify(user));
            errorMsg.style.display = 'none';
            showApp(user);
        } else {
            errorMsg.textContent = "Incorrect Password";
            errorMsg.style.display = 'block';
        }
    } else {
        errorMsg.textContent = "User ID not found";
        errorMsg.style.display = 'block';
    }
}

function handleLogout() {
    localStorage.removeItem('currentUser');
    location.reload();
}

function showApp(user) {
    document.getElementById('loginOverlay').style.display = 'none';
    document.getElementById('mainApp').style.display = 'flex';

    const displayUserName = document.getElementById('displayUserName');
    if (displayUserName) displayUserName.textContent = user['Name'] || user['IDRec'];

    // Debug User Object
    console.log("Current User Object:", user);
    console.log("User Keys:", Object.keys(user));

    const findKey = (obj, keyName) => Object.keys(obj).find(k => k.toLowerCase().trim() === keyName.toLowerCase());
    const statusKey = findKey(user, 'status') || findKey(user, 'สถานะ') || findKey(user, 'state');
    const userStatus = statusKey ? user[statusKey] : '';

    const details = [
        `ID: ${user['IDRec']}`,
        user['ตำแหน่ง'],
        user['หน่วยงาน'],
        user['Team'],
        user['Plant'],
        userStatus ? `Status: ${userStatus}` : ''
    ].filter(Boolean).join(' | ');
    document.getElementById('displayUserDetail').textContent = details;

    loadTableData();
}

function openUserProfile() {
    const savedUser = localStorage.getItem('currentUser');
    if (!savedUser) return;
    const user = JSON.parse(savedUser);

    const setText = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.textContent = val || '-';
    };

    setText('profileId', user['IDRec']);
    setText('profileName', user['Name'] || user['IDRec']);
    setText('profilePosition', user['ตำแหน่ง']);
    setText('profileUnit', (user['หน่วยงาน'] || '') + ' ' + (user['Team'] || ''));
    setText('profilePlant', user['Plant']);

    const findKey = (obj, keyName) => Object.keys(obj).find(k => k.toLowerCase().trim() === keyName.toLowerCase());
    const statusKey = findKey(user, 'status') || findKey(user, 'สถานะ') || findKey(user, 'state');
    const userStatus = statusKey ? user[statusKey] : 'Active';

    // Update Badge
    const badge = document.getElementById('profileStatusBadge');
    if (badge) {
        badge.textContent = userStatus;
        // Optional: dynamic color based on status
        if (String(userStatus).toLowerCase() === 'admin') {
            badge.style.background = '#fef3c7';
            badge.style.color = '#d97706';
            badge.style.borderColor = '#fcd34d';
        } else {
            badge.style.background = '#ecfdf5';
            badge.style.color = '#059669';
            badge.style.borderColor = '#d1fae5';
        }
    }

    // Avatar
    const name = user['Name'] || user['IDRec'] || 'U';
    const initials = String(name).charAt(0).toUpperCase();
    const avatarEl = document.getElementById('userAvatar');
    if (avatarEl) avatarEl.textContent = initials;

    const modal = document.getElementById('userProfileModal');
    if (modal) modal.style.display = 'flex';
}

function closeUserProfile() {
    const modal = document.getElementById('userProfileModal');
    if (modal) modal.style.display = 'none';
}

function toggleMenu() {
    openUserProfile();
}

window.onclick = function (event) {
    const modal = document.getElementById('userProfileModal');
    if (event.target === modal) {
        closeUserProfile();
    }
}

document.getElementById('passInput').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        handleLogin();
    }
});