const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

function applyUpdateEventCategories(newItem, categories, changes) {
  if (Array.isArray(categories)) {
    newItem.setCategories(categories);
    changes.push("categories");
  }
}

function createMockItem(initialCategories = ["Existing"]) {
  return {
    categories: initialCategories.slice(),
    setCategoriesCalls: [],
    setCategories(categories) {
      this.setCategoriesCalls.push(categories);
      this.categories = categories;
    },
  };
}

describe("updateEvent category normalization", () => {
  it("treats categories: null as no-op", () => {
    const item = createMockItem();
    const changes = [];

    applyUpdateEventCategories(item, null, changes);

    assert.deepEqual(item.categories, ["Existing"]);
    assert.deepEqual(item.setCategoriesCalls, []);
    assert.deepEqual(changes, []);
  });

  it("treats categories: undefined as no-op", () => {
    const item = createMockItem();
    const changes = [];

    applyUpdateEventCategories(item, undefined, changes);

    assert.deepEqual(item.categories, ["Existing"]);
    assert.deepEqual(item.setCategoriesCalls, []);
    assert.deepEqual(changes, []);
  });

  it("clears all categories with an empty array", () => {
    const item = createMockItem();
    const changes = [];
    const categories = [];

    applyUpdateEventCategories(item, categories, changes);

    assert.deepEqual(item.categories, []);
    assert.strictEqual(item.setCategoriesCalls[0], categories);
    assert.deepEqual(changes, ["categories"]);
  });

  it("sets a multi-item category list", () => {
    const item = createMockItem();
    const changes = [];
    const categories = ["Work", "Personal"];

    applyUpdateEventCategories(item, categories, changes);

    assert.deepEqual(item.categories, ["Work", "Personal"]);
    assert.strictEqual(item.setCategoriesCalls[0], categories);
    assert.deepEqual(changes, ["categories"]);
  });

  it("sets a single-item category list", () => {
    const item = createMockItem();
    const changes = [];
    const categories = ["Work"];

    applyUpdateEventCategories(item, categories, changes);

    assert.deepEqual(item.categories, ["Work"]);
    assert.strictEqual(item.setCategoriesCalls[0], categories);
    assert.deepEqual(changes, ["categories"]);
  });
});
