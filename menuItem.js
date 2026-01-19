class MenuItem {
    constructor(name, price) {
        this.name = name;
        this.price = price;
        this.tags = []; // Array of strings for optional tags
    }

    addTag(tag) {
        if (!this.tags.includes(tag)) {
            this.tags.push(tag);
        }
    }

    removeTag(tag) {
        this.tags = this.tags.filter(t => t !== tag);
    }
}

module.exports = MenuItem;
