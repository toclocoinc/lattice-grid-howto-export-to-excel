# How to export a data grid to Excel from JavaScript

Calls `grid.export.excel()` from one button to write a real .xlsx workbook,
with a checkbox for all rows versus only the filtered view, and a checkbox
for whether a hidden column comes with it.

Live demo: https://toclocoinc.github.io/lattice-grid-howto-export-to-excel/

**Read the how-to:** https://www.latticegrid.dev/docs/how-to/export-to-excel/

## The snippet

```js
document.getElementById('download').addEventListener('click', () => {
  grid.export.excel({
    fileName: 'product-catalogue',
    rows: allRowsChecked ? 'all' : 'visible',
    hidden: includeHiddenChecked,
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
