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
