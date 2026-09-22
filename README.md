# How to export a data grid to Excel from JavaScript

Calls `grid.export.excel()` from one button to write a real .xlsx workbook,
with a checkbox for all rows versus only the filtered view, and a checkbox
for whether a hidden column comes with it.

Live demo: https://toclocoinc.github.io/lattice-grid-howto-export-to-excel/

**Read the how-to:** https://www.latticegrid.dev/docs/how-to/export-to-excel/

## The full source

Two files: `index.html` loads the grid and declares the mount point, `demo.js` configures and creates it. Copy both as they are below and it runs.

### index.html

```html
<!doctype html>
<html lang="en-GB">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>How to export a data grid to Excel from JavaScript</title>
    <meta
      name="description"
      content="Call grid.export.excel() to write a real .xlsx workbook straight from a JavaScript data grid, with a choice of all rows or only the filtered view, and whether hidden columns are included. Built with Lattice Grid loaded by script tag, no install and no build."
    />
    <link rel="icon" href="data:," />
    <!--
      The grid's stylesheet, from jsDelivr. The address names the exact
      release, 1.68.2, and carries the hash of the file it expects, so the
      page can never quietly pick up a different build than the one it was
      checked against.
    -->
    <link
      rel="stylesheet"
      href="https://cdn.jsdelivr.net/npm/@toclocoinc/lattice-grid@1.68.2/lattice-grid.min.css"
      integrity="sha384-mcpd7S8C5nz58bZDAXdYH6rzezEhfN7B4u2SlW426dSe20GnkxTu4TygyOILnCth"
      crossorigin="anonymous"
    />
    <style>
      body { margin: 0; font-family: system-ui, sans-serif; background: #f4f6f9; color: #131a24; }
      header { padding: 1.5rem 1.5rem 0.5rem; max-width: 960px; margin: 0 auto; }
      header p { color: #4a5568; }
      header a { color: #2d6bff; }
      main { max-width: 960px; margin: 0 auto; padding: 0 1.5rem 2.5rem; }
      #grid { height: 420px; }
      .toolbar { display: flex; align-items: center; gap: 1.25rem; flex-wrap: wrap; margin: 0 0 0.75rem; }
      .toolbar input[type="text"] { padding: 0.4rem 0.6rem; border: 1px solid #cbd5e1; border-radius: 6px; }
      .toolbar label { display: flex; align-items: center; gap: 0.35rem; font-size: 0.9rem; }
      .toolbar button { padding: 0.5rem 0.9rem; border: 0; border-radius: 6px; background: #2d6bff; color: #fff; font-weight: 600; cursor: pointer; }
    </style>
  </head>
  <body>
    <header>
      <h1>How to export a data grid to Excel from JavaScript</h1>
      <p>
        Type in the box to filter the rows, then download. Read the
        <a href="https://www.latticegrid.dev/docs/how-to/export-to-excel/">full how-to</a>
        on latticegrid.dev.
      </p>
    </header>
    <main>
      <div class="toolbar">
        <input type="text" id="quick-filter" placeholder="Filter, e.g. Peripherals" />
        <label><input type="checkbox" id="all-rows" /> All rows, ignore the filter</label>
        <label><input type="checkbox" id="include-hidden" /> Include hidden columns</label>
        <button type="button" id="download">Download .xlsx</button>
      </div>
      <div id="grid"></div>
    </main>

    <!--
      The library, as a classic script tag. No npm install, no bundler, no
      type="module": the file runs as it arrives and leaves the LatticeGrid
      global behind.
    -->
    <script
      src="https://cdn.jsdelivr.net/npm/@toclocoinc/lattice-grid@1.68.2/lattice-grid.min.js"
      integrity="sha384-vCzLyFYn0T0lz/vkdH4x0JpJZkOazZgI2LiGui7lm5uerdZd0Z46G9hr3Aq1FFPS"
      crossorigin="anonymous"
    ></script>
    <script src="./demo.js"></script>
  </body>
</html>
```

### demo.js

```js
/**
 * Export a data grid to Excel from JavaScript.
 *
 * One button calls grid.export.excel() to write a real .xlsx workbook, with
 * two checkboxes for the two questions a reader asks first: export every row
 * or only the filtered view, and whether the Margin column, hidden by
 * default, comes with it.
 */

// Tied to toclocoinc.github.io only; has no effect anywhere else and needs
// no key at all to run this page from a local copy.
LatticeGrid.setLicence(
  'LG1.eyJ2IjoxLCJwIjoibGF0dGljZS1ncmlkIiwidCI6IlRPQ0xPQ08gSW5jIC0gcHVibGljIGRlbW9zIiwiZSI6IjIwMzAtMDEtMDEiLCJkIjpbInRvY2xvY29pbmMuZ2l0aHViLmlvIl19.9De42ua3aCGpiMB6EVRP7Tv-upUlDI-0T07rlSPzvCrsqg8t4YJi7SRnStEpAg48uzmcG7il1fR_TfwkUE7iCA'
);

const rows = [
  { sku: 'SKU-1001', product: 'Wireless Mouse', category: 'Peripherals', units: 240, price: 18.99, margin: 6.2 },
  { sku: 'SKU-1002', product: 'Mechanical Keyboard', category: 'Peripherals', units: 96, price: 74.5, margin: 21.4 },
  { sku: 'SKU-1003', product: '27" Monitor', category: 'Displays', units: 58, price: 219.0, margin: 47.1 },
  { sku: 'SKU-1004', product: 'USB-C Dock', category: 'Accessories', units: 130, price: 42.75, margin: 12.9 },
  { sku: 'SKU-1005', product: 'Laptop Stand', category: 'Accessories', units: 172, price: 27.0, margin: 9.4 },
  { sku: 'SKU-1006', product: 'Webcam 1080p', category: 'Peripherals', units: 84, price: 39.99, margin: 11.6 },
  { sku: 'SKU-1007', product: '34" Ultrawide Monitor', category: 'Displays', units: 21, price: 389.0, margin: 66.3 },
  { sku: 'SKU-1008', product: 'Noise-cancelling Headset', category: 'Audio', units: 65, price: 89.0, margin: 24.8 },
  { sku: 'SKU-1009', product: 'Desk Lamp', category: 'Accessories', units: 110, price: 22.5, margin: 7.9 },
  { sku: 'SKU-1010', product: 'Bluetooth Speaker', category: 'Audio', units: 47, price: 54.0, margin: 15.2 },
  { sku: 'SKU-1011', product: 'Ergonomic Chair', category: 'Furniture', units: 18, price: 249.0, margin: 58.0 },
  { sku: 'SKU-1012', product: 'Standing Desk', category: 'Furniture', units: 12, price: 429.0, margin: 91.4 },
];

const money = { style: 'currency', currency: 'USD', decimals: 2 };

const grid = LatticeGrid.createGrid(document.getElementById('grid'), {
  rowKey: 'sku',
  columns: [
    { field: 'sku', title: 'SKU', layout: { width: 100 } },
    { field: 'product', title: 'Product', layout: { width: 210 } },
    { field: 'category', title: 'Category', layout: { width: 130 } },
    { field: 'units', title: 'Units', type: 'number', layout: { width: 90 } },
    { field: 'price', title: 'Price', type: 'number', format: money, layout: { width: 110 } },
    // Hidden by default: on screen for a manager, out of a customer's export
    // unless "Include hidden columns" is checked.
    { field: 'margin', title: 'Margin', type: 'number', format: money, layout: { width: 110, hidden: true } },
  ],
  rows,
});
window.__demoGrid = grid; // read by tools/verify.mjs to check the export

document.getElementById('quick-filter').addEventListener('input', (e) => {
  grid.filters.quick(e.target.value);
});

document.getElementById('download').addEventListener('click', () => {
  grid.export.excel({
    fileName: 'product-catalogue',
    // 'all' exports every row regardless of the quick filter above;
    // 'visible' (the default) exports only what the filter leaves.
    rows: document.getElementById('all-rows').checked ? 'all' : 'visible',
    // 'omit' (the default) drops the Margin column; 'hidden' carries it into
    // the workbook, itself marked hidden there for a round trip back in.
    hiddenColumns: document.getElementById('include-hidden').checked ? 'hidden' : 'omit',
  });
});
```

`rows: 'all'` exports every row regardless of the current filter; the
default, `'visible'`, exports only what the filter leaves. `hidden: true`
carries a column the grid has hidden into the file as well; it stays out by
default.

## Running it yourself

Open `index.html` in a browser, or serve the folder with any static file
server. The grid loads from jsDelivr by script tag, so there is no install
and no build step. It runs keyless on `localhost`; the licence key in
`demo.js` is bound to `toclocoinc.github.io` and has no effect anywhere else.

## Licence

MIT, see [LICENSE](./LICENSE). Lattice Grid itself is licensed separately
per domain: https://www.latticegrid.dev/pricing/
