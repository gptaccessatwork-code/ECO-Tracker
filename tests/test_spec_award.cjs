// Run with Node 22.13+ (built-in TypeScript stripping; no dependencies).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { stripTypeScriptTypes } = require('node:module');

const jasmine = 'siewyeejasmine_lim@amat.com';
function table(headers, rows) {
  return '<table><tr>' + headers.map(h => `<th>${h}</th>`).join('') +
    '</tr>' + rows.map(row => '<tr>' + row.map(c => `<td>${c}</td>`).join('') + '</tr>').join('') + '</table>';
}
const newRow = ['704902R04', 'SEMNSO-KA-GP', '704902R04-KA-GP', '10/15/2026'];
const newHeaders = ['Slot#', 'Material', 'Slot#+Part#', 'NBD'];
const newBody = '<p>Hi Ichor Team. Spec award as follows:</p>' + table(newHeaders, [newRow, newRow]);
const oldBody = '<p>Qtr 4; Need By Date</p>' + table(['Slot Number', 'Part Number'], [['710854', 'SEMSYS-XA2T-GP']]);

function workbook(existing = []) {
  const added = [];
  const range = {
    getTexts: () => existing.map(s => [s]),
    getCell() { return this; }, getResizedRange() { return this; },
    setNumberFormat() {}, getFormat() { return this; }, setHorizontalAlignment() {}
  };
  const target = {
    getName: () => 'Table1',
    getHeaderRowRange: () => ({ getTexts: () => [['System Number', 'Spec Award Date']] }),
    getRowCount: () => existing.length,
    getColumnByName: () => ({ getRangeBetweenHeaderAndTotal: () => range }),
    getWorksheet: () => ({ getProtection: () => ({ getProtected: () => false }) }),
    addRows: (_, rows) => added.push(...rows)
  };
  return { getTables: () => [target], added };
}

for (const filename of ['spec_award_office_script.ts', 'spec_award_combined_office_script.ts']) {
  const source = fs.readFileSync(path.join(__dirname, '../power_automate', filename), 'utf8');
  const context = vm.createContext({ console: { log() {} }, ExcelScript: { HorizontalAlignment: { left: 'left' } } });
  vm.runInContext(stripTypeScriptTypes(source), context);
  const combined = filename.includes('combined');
  const run = (html, sender = jasmine, existing = [], live = true) => {
    const book = workbook(existing);
    const args = [book, html, '2026-10-08T17:00:00Z'];
    if (combined) args.push(live);
    args.push(sender);
    const result = context.main(...args);
    return { result, added: book.added };
  };
  assert.equal(run(newBody).added[0][0], '704902R04-KA-GP');
  assert.equal(run(newBody).added.length, 1, 'deduplicates rows');
  assert.equal(run(newBody).result.skippedDuplicates, 1);
  assert.equal(run(newBody, ' SIEWYEEJASMINE_LIM@AMAT.COM ').added.length, 1);
  assert.equal(run(oldBody).added[0][0], '710854-XA2T-GP', 'Jasmine retains legacy support');
  assert.equal(run(oldBody, 'faith_fan@amat.com').added.length, 1);
  assert.equal(run(newBody, 'faith_fan@amat.com').added.length, 0);
  assert.equal(run(newBody, '').added.length, 0);
  assert.equal(run(newBody.replace('Spec award', 'Please review')).added.length, 0);
  assert.equal(run(newBody.replaceAll('704902R04-KA-GP', 'WRONG')).added.length, 0);
  assert.equal(run('<p>Thanks</p><div id="divRplyFwdMsg">' + newBody).added.length, 0);
  assert.equal(run('<p>Thanks</p><div style="border-top:solid"><p>From: Jasmine Sent: today To: Ichor Subject: Spec award</p>' + newBody).added.length, 0);
  assert.equal(run('<p>Spec award</p>' + table(newHeaders.slice(0, 3), [newRow.slice(0, 3)])).added.length, 0);
  assert.equal(run(newBody, jasmine, ['704902R04-KA-GP']).added.length, 0);
  assert.equal(run(newBody).added[0][1], Date.UTC(2026, 9, 9) / 86400000 + 25569, 'uses Singapore received date');
  const book = workbook();
  const defaultArgs = [book, oldBody, '2026-10-08T17:00:00Z'];
  if (combined) defaultArgs.push(true);
  context.main(...defaultArgs);
  assert.equal(book.added.length, 1, 'omitting optional sender preserves legacy calls');
  if (combined) {
    const result = context.main({ getTables() { throw Error('TEST mode accessed workbook'); } }, newBody, '', false, jasmine);
    assert.equal(result.detected, true);
    assert.deepEqual(Array.from(result.systemNumbers), ['704902R04-KA-GP']);
    assert.equal(result.skippedDuplicates, 1);
    assert.equal(result.forwardHtml.includes('704902R04-KA-GP'), true);
    assert.equal(run('<p>Thanks</p>', jasmine, [], false).result.detected, false);
  }
  console.log(`${filename}: passed`);
}
