/**
 * Load the demo in a real headless browser and check that it works.
 *
 * Serves the project locally and drives it with the same minimal Chrome
 * driver the grid's own test suite uses (tools/browser.js), so this needs no
 * dependency beyond Chrome itself.
 *
 * Checks:
 *   - the library arrived and left the LatticeGrid global behind;
 *   - the grid mounted rows, with the Margin column starting hidden;
 *   - the download button and both checkboxes are on the page;
 *   - grid.export.excel() actually produces a valid .xlsx: a zip whose
 *     first four bytes are the local-file-header signature and which
 *     contains an xl/worksheets/sheet1.xml entry;
 *   - hiddenColumns:'hidden' really does add the Margin column's data over
 *     hiddenColumns:'omit' (the exported file is larger, not just "valid");
 *   - after the quick filter narrows the grid, rows:'visible' exports a
 *     smaller file than rows:'all' does (byte size stands in for row count,
 *     since the worksheet itself is DEFLATE-compressed inside the zip);
 *   - nothing logged a console error or threw while the page ran.
 *
 * Exits non-zero on any failure, so it can gate a deployment.
 *
 * Usage: node tools/verify.mjs
 */

import { startServer } from './serve.mjs';
import { Browser, available } from './browser.js';

if (!available()) {
  console.log('No headless browser on this machine; skipping verify.mjs.');
  process.exit(0);
}

const { server, port } = await startServer();
const browser = new Browser();

/**
 * Check a base64-encoded .xlsx for the zip signature and the worksheet entry
 * the WO asks for, without needing a zip library.
 */
/*
 * A zip's local-file-header names are stored uncompressed, so a plain text
 * search finds them; the worksheet's own XML is DEFLATE-compressed inside
 * that entry and cannot be read this way. That is enough to prove the file
 * is a real .xlsx with the entry the WO asks for; how much a given export
 * option changed is checked by comparing the whole file's byte size instead
 * (bigger with the hidden column included, smaller once a filter narrows
 * the rows), which needs no zip or inflate library.
 */
function looksLikeXlsx(base64) {
  const bytes = Buffer.from(base64, 'base64');
  const isZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
  const text = bytes.toString('latin1');
  const hasSheet = text.includes('xl/worksheets/sheet1.xml');
  const hasContentTypes = text.includes('[Content_Types].xml');
  return { bytes: bytes.length, isZip, hasSheet, hasContentTypes };
}

/** Produce a workbook in-page and bring its bytes back as base64. */
async function exportBase64(browser, opts) {
  return browser.evaluate(`(async () => {
    const g = window.__demoGrid;
    const blob = await g.export.excel(${JSON.stringify(opts)});
    const buf = await blob.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let binary = '';
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  })()`);
}

try {
  await browser.start();

  // Collect console errors and thrown errors from the very first script the
  // page runs, before LatticeGrid or demo.js execute.
  await browser.send('Page.addScriptToEvaluateOnNewDocument', {
    source: `
      window.__errs = [];
      addEventListener('error', (e) => window.__errs.push(String(e.message || e)));
      addEventListener('unhandledrejection', (e) => window.__errs.push('unhandledrejection: ' + String(e.reason)));
      const origError = console.error.bind(console);
      console.error = (...args) => { window.__errs.push('console.error: ' + args.map(String).join(' ')); origError(...args); };
    `,
  });

  await browser.open(`http://127.0.0.1:${port}/`);

  const hasGrid = await browser.evaluate(`typeof LatticeGrid !== 'undefined' && typeof LatticeGrid.createGrid === 'function'`);
  const rowCount = await browser.evaluate(`document.querySelectorAll('#grid .lat-row').length`);
  const marginHidden = await browser.evaluate(`!document.querySelector('#grid [data-col="margin"]')`);
  const hasButton = await browser.evaluate(`!!document.getElementById('download')`);
  const hasCheckboxes = await browser.evaluate(`!!document.getElementById('all-rows') && !!document.getElementById('include-hidden')`);

  // Two exports with every row visible: one omits the hidden Margin column,
  // one includes it. If the option works the second file is bigger.
  const omitBase64 = await exportBase64(browser, { download: false, hiddenColumns: 'omit', rows: 'all' });
  const includeBase64 = await exportBase64(browser, { download: false, hiddenColumns: 'hidden', rows: 'all' });

  // Narrow the grid with the quick filter, then export both ways: 'visible'
  // should carry only the matching rows, 'all' should ignore the filter.
  await browser.evaluate(`document.getElementById('quick-filter').value = 'Audio';
    document.getElementById('quick-filter').dispatchEvent(new Event('input'));`);
  // The filter coalesces keystrokes before it takes effect; give it a beat.
  await new Promise((r) => setTimeout(r, 300));
  const filteredVisibleBase64 = await exportBase64(browser, { download: false, hiddenColumns: 'omit', rows: 'visible' });
  const filteredAllBase64 = await exportBase64(browser, { download: false, hiddenColumns: 'omit', rows: 'all' });

  const errors = await browser.evaluate('window.__errs');

  const omit = looksLikeXlsx(omitBase64);
  const include = looksLikeXlsx(includeBase64);
  const filteredVisible = looksLikeXlsx(filteredVisibleBase64);
  const filteredAll = looksLikeXlsx(filteredAllBase64);

  const failures = [];
  if (!hasGrid) failures.push('LatticeGrid.createGrid was not found on the page');
  if (!(rowCount > 0)) failures.push(`expected rendered rows, found ${rowCount}`);
  if (!marginHidden) failures.push('the Margin column is not hidden by default');
  if (!hasButton) failures.push('the download button is missing');
  if (!hasCheckboxes) failures.push('the all-rows or include-hidden checkbox is missing');
  for (const [name, r] of [['omit', omit], ['include', include], ['filteredVisible', filteredVisible], ['filteredAll', filteredAll]]) {
    if (!r.isZip || !r.hasSheet || !r.hasContentTypes) {
      failures.push(`${name} export is not a valid .xlsx (${JSON.stringify(r)})`);
    }
  }
  if (!(include.bytes > omit.bytes)) {
    failures.push(`hiddenColumns:'hidden' (${include.bytes}B) did not add data over 'omit' (${omit.bytes}B)`);
  }
  if (!(filteredVisible.bytes < filteredAll.bytes)) {
    failures.push(
      `rows:'visible' (${filteredVisible.bytes}B) did not export less than rows:'all' (${filteredAll.bytes}B) after filtering`
    );
  }
  if (errors.length) failures.push(`console/window errors: ${errors.join(' | ')}`);

  if (failures.length) {
    console.error('FAILED:\n' + failures.map((f) => `  - ${f}`).join('\n'));
    process.exitCode = 1;
  } else {
    console.log(
      `OK: grid loaded, ${rowCount} rows rendered, Margin hidden by default, hidden columns add ` +
      `${include.bytes - omit.bytes}B when included, the filter shrinks the 'visible' export by ` +
      `${filteredAll.bytes - filteredVisible.bytes}B versus 'all', 0 console errors.`
    );
  }
} finally {
  await browser.close();
  server.close();
}
