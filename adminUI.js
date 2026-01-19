const { assignTagToItem, removeTagFromItem, getMenuItems } = require('./admin');

function populateItems() {
    const select = document.getElementById('itemSelect');
    const items = getMenuItems();
    items.forEach(item => {
        const option = document.createElement('option');
        option.value = item.name;
        option.textContent = item.name;
        select.appendChild(option);
    });
    displayMenu();
}

function addTag() {
    const itemName = document.getElementById('itemSelect').value;
    const tag = document.getElementById('tagInput').value.trim();
    if (tag) {
        assignTagToItem(itemName, tag);
        displayMenu();
    }
}

function removeTag() {
    const itemName = document.getElementById('itemSelect').value;
    const tag = document.getElementById('tagInput').value.trim();
    if (tag) {
        removeTagFromItem(itemName, tag);
        displayMenu();
    }
}

function displayMenu() {
    const display = document.getElementById('menuDisplay');
    display.innerHTML = '<h2>Menu Items</h2>';
    const items = getMenuItems();
    items.forEach(item => {
        display.innerHTML += `<p>${item.name} - $${item.price} - Tags: ${item.tags.join(', ')}</p>`;
    });
}

window.onload = populateItems;
