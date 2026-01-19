const MenuItem = require('./menuItem');

// Example menu items storage (replace with DB in production)
let menuItems = [
    new MenuItem('Burger', 5.99),
    new MenuItem('Salad', 4.99)
];

function assignTagToItem(itemName, tag) {
    const item = menuItems.find(i => i.name === itemName);
    if (item) {
        item.addTag(tag);
    }
}

function removeTagFromItem(itemName, tag) {
    const item = menuItems.find(i => i.name === itemName);
    if (item) {
        item.removeTag(tag);
    }
}

function getMenuItems() {
    return menuItems;
}

module.exports = { assignTagToItem, removeTagFromItem, getMenuItems };
