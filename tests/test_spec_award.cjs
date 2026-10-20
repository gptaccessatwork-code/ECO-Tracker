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
  assert.equal(run(newBody.replaceAll('704902R04-KA-GP', 'WRONG')).added[0][0], '704902R04-KA-GP',
    'Material is authoritative even when Slot#+Part# differs');
  const lowerTable = table(['Material', 'Slot#', 'CM NBD'], [
    ['0244-05708', '710143', '11/30/2026'],
    ['0244-05708', '710148', '12/02/2026'],
    ['0244-05717', '710148', '12/02/2026']
  ]);
  assert.deepEqual(run(newBody + lowerTable).added.map(row => row[0]),
    ['704902R04-KA-GP', '0244-05708', '0244-05717'],
    'all Material tables contribute and repeated materials across slots deduplicate');
  assert.equal(run(newBody + lowerTable).result.skippedDuplicates, 2);
  assert.equal(run('<p>Spec award</p>' + lowerTable).added.length, 2);
  assert.equal(run('<p>Spec award</p>' + lowerTable, 'faith_fan@amat.com').added.length, 0);
  assert.deepEqual(run(oldBody + newBody + lowerTable).added.map(row => row[0]),
    ['710854-XA2T-GP', '704902R04-KA-GP', '0244-05708', '0244-05717'],
    'legacy and Material layouts both contribute in the same message');
  assert.deepEqual(run(oldBody + oldBody.replaceAll('710854', '710855')).added.map(row => row[0]),
    ['710854-XA2T-GP', '710855-XA2T-GP'], 'all legacy tables are collected');
  assert.equal(run('<p>Thanks</p><div id="divRplyFwdMsg">' + newBody).added.length, 0);
  assert.equal(run('<p>Thanks</p><div style="border-top:solid"><p>From: Jasmine Sent: today To: Ichor Subject: Spec award</p>' + newBody).added.length, 0);
  const mixedHistory = newBody + '<div style="border-top:solid"><p>From: Jasmine Sent: today To: Ichor Subject: Spec award</p>' +
    oldBody + '<div id="divRplyFwdMsg">' + oldBody;
  assert.equal(run(mixedHistory).added[0][0], '704902R04-KA-GP',
    'earlier Outlook header takes precedence over a later reply marker');
  assert.equal(run('<p>Thanks</p>' + mixedHistory.slice(newBody.length)).added.length, 0,
    'mixed reply markers do not expose historical awards');
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
  // Optional real-run output; keep private email HTML outside the repository.
  if (process.argv[2]) {
    const payload = JSON.parse(fs.readFileSync(process.argv[2], 'utf8').replace(/^\uFEFF/, ''));
    const captured = payload.result;
    const emailHtml = payload.parameters
      ? payload.parameters['ScriptParameters/emailHtml'] : captured.forwardHtml;
    const existingSystems = captured ? captured.systemNumbers : ['431757R01-XP-GP',
      'B09332-PJ-ZGP2', 'B09332-PJ-ZGP3', '708695-XA3T-GPA', '708695-XA3T-GPB',
      '710419-XA3T-GPA', '710419-XA3T-GPB', '710419-XA3T-GPC'];
    const expected = ['710897-XA1T-GP', '710290-XXT-GPA', '710290-XXT-GPB',
      '710290-XXT-GPC', 'B10067-PJ-ZGP2', 'B10067-PJ-ZGP3',
      '0244-05708', '0244-05717', '0244-05709', '0244-05441'];
    const actual = run(emailHtml, jasmine, existingSystems);
    assert.deepEqual(actual.added.map(row => row[0]), expected,
      'current award is added even when quoted historical systems already exist');
    assert.equal(actual.result.skippedDuplicates, 23);
    assert.equal(actual.result.added, 10);
    if (combined) {
      const preview = run(emailHtml, jasmine, [], false).result;
      assert.deepEqual(Array.from(preview.systemNumbers), expected);
      assert.equal(preview.skippedDuplicates, 23);
    }
    if (combined) {
      assert.equal(actual.result.forwardHtml.includes('431757R01'), false,
        'quoted historical award is excluded from forwarding');
    }
    console.log(`${filename}: supplied email regression passed (33 source rows, 10 added, 23 duplicates)`);
  }
  console.log(`${filename}: passed`);
}
