#!/usr/bin/env node
/** Boundaries of the office tool dispatcher. Run after `npm run build`. */
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const s = require('../dist/index.js');

let failed = 0;
function assert(cond, msg) {
  if (!cond) {
    console.error('FAIL:', msg);
    failed += 1;
  } else {
    console.log('OK  ', msg);
  }
}

function calls() {
  return { slide: 0, office: 0, officeType: null };
}

async function dispatch(docType, name, args = {}) {
  const seen = calls();
  const result = await s.dispatchOfficeTool({
    docType,
    name,
    args,
    runSlide: async () => {
      seen.slide += 1;
      return { ok: true, via: 'slide' };
    },
    runOffice: async (toolName, toolArgs, officeType) => {
      seen.office += 1;
      seen.officeType = officeType;
      return JSON.stringify({ ok: true, via: 'office', toolName, officeType });
    },
  });
  return { result, seen };
}

const unknown = await dispatch('SHEET', 'not_a_tool');
assert(
  unknown.seen.slide === 0 && unknown.seen.office === 0,
  'unknown tool runs nowhere',
);
assert(
  JSON.parse(unknown.result).error === 'unknown tool not_a_tool',
  'unknown tool names itself',
);

const slideSheet = await dispatch('SLIDE', 'set_cells', { values: [['a']] });
assert(
  slideSheet.seen.office === 0 && slideSheet.seen.slide === 0,
  'sheet tool on a presentation does not reach Univer',
);
assert(
  JSON.parse(slideSheet.result).error ===
    'tool is not available for this document',
  'sheet tool on a presentation is rejected',
);

const slideOnSheet = await dispatch('SHEET', 'listSlides');
assert(
  slideOnSheet.seen.office === 0 && slideOnSheet.seen.slide === 0,
  'slide tool does not run inside a spreadsheet',
);

const listed = await dispatch('SLIDE', 'listSlides');
assert(
  listed.seen.slide === 1 && listed.seen.office === 0,
  'listSlides stays in the slide editor',
);
assert(
  JSON.parse(listed.result).via === 'slide',
  'listSlides result comes from the slide editor',
);

const planned = await dispatch('SLIDE', 'create_plan', { name: 'План' });
assert(
  planned.seen.slide === 0 && planned.seen.office === 1,
  'plan tools still run while a presentation is open',
);

const cells = await dispatch('SHEET', 'set_cells', { values: [['a']] });
assert(
  cells.seen.office === 1 && cells.seen.officeType === 'SHEET',
  'set_cells runs on a sheet',
);

const docWrite = await dispatch('DOC', 'replace_text', {
  search: 'a',
  replacement: 'b',
});
assert(
  docWrite.seen.officeType === 'DOC',
  'document tools keep the DOC target',
);

assert(
  s.needsApprove('set_cells', { values: [['одна']] }) === false,
  'a single cell write does not ask for confirmation',
);
assert(
  s.needsApprove('set_cells', {
    values: [Array.from({ length: 21 }, () => 'x')],
  }) === true,
  'a write past 20 cells asks for confirmation',
);
assert(s.needsApprove('delete_sheet', {}) === true, 'delete_sheet always asks');
assert(s.needsApprove('get_selection', {}) === false, 'read tools do not ask');
assert(
  s.needsApprove('setShapeText', { text: 'Заголовок' }) === true,
  'slide writes always ask',
);
assert(
  s.needsApprove('format_range', { range: 'A1:B2' }, { area: () => 4 }) ===
    false,
  'a small format does not ask',
);
assert(
  s.needsApprove('format_range', { range: 'A1:Z100' }, { area: () => 60 }) ===
    true,
  'a large format asks',
);
assert(s.estimateA1Area('A1:C10') === 30, 'A1:C10 is 30 cells');
assert(
  s.resolveToolTarget('SLIDE', 'addSlide') === 'slide',
  'addSlide targets the slide editor',
);
assert(
  s.resolveToolTarget('DOC', 'setShapeText') === 'reject',
  'slide write is rejected on a document',
);

if (failed) {
  console.error(`\n${failed} failed`);
  process.exit(1);
}
console.log('\nall tool policy checks passed');
