async function initAuth() {
    allEmployees = await fetchData(EMPLOYEE_SHEET_URL);
    allEmployees = await fetchData(EMPLOYEE_SHEET_URL); // Double fetch in original code, kept for consistency

    const savedUser = localStorage.getItem('currentUser');
    if (savedUser) {
        const user = JSON.parse(savedUser);
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

    document.getElementById('displayUserName').textContent = user['Name'] || user['IDRec'];
    const menuUserDisplay = document.getElementById('menuUserDisplay');
    if (menuUserDisplay) menuUserDisplay.textContent = user['Name'] || user['IDRec'];

    const details = [
        `ID: ${user['IDRec']}`,
        user['ตำแหน่ง'],
        user['หน่วยงาน'],
        user['Team'],
        user['Plant']
    ].filter(Boolean).join(' | ');
    document.getElementById('displayUserDetail').textContent = details;

    loadTableData();
}

function toggleMenu() {
    const menu = document.getElementById('userDropdown');
    menu.classList.toggle('show');
}

window.onclick = function (event) {
    if (!event.target.closest('.menu-container')) {
        const dropdowns = document.getElementsByClassName("dropdown-menu");
        for (let i = 0; i < dropdowns.length; i++) {
            const openDropdown = dropdowns[i];
            if (openDropdown.classList.contains('show')) {
                openDropdown.classList.remove('show');
            }
        }
    }
}

document.getElementById('passInput').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        handleLogin();
    }
});