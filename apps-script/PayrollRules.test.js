// Run with: node apps-script/PayrollRules.test.js
//
// Unit test PayrollRules.gs against real June 2026 Salary Sheet rows
// read from the actual signed PDFs on 2026-07-09. Per the plan's
// verification section: sample spanning MD/TGT/PRT/Peon designations
// across HES, LMS1, LMS3 (LMS3 has the different 25% ADA rate).
//
// MD rows (e.g. HES's Lalita Kanwar) are intentionally excluded from
// the compute-chain assertions — real data shows MDs are EPF-exempt,
// a designation-based exception PayrollRules.gs does not yet model
// (documented in computeEPF's JSDoc). They ARE used for the
// reconciliation check below, since that just sums Net Pay regardless
// of how each row got there.

const path = require('path');
const constants = require(path.join(__dirname, 'PayrollConstants.gs'));
Object.assign(global, constants);
const rules = require(path.join(__dirname, 'PayrollRules.gs'));
const { computeADA, computeDA, computeEPF, computeESI, runPayrollForEmployee, validateReconciliation } = rules;

let pass = 0, fail = 0;
function check(label, actual, expected) {
  const ok = actual === expected;
  if (ok) pass++; else fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}: got ${actual}, expected ${expected}`);
}

console.log('=== HES (35% ADA, simple schema, no INCENTIVE/TRAV/TF) ===');
[
  { name: 'Raj Bahadur (Gardener)', school: 'HES', basic: 11965, ada: 4188, da: 598, epf: 1800, esi: 126, netPay: 14825 },
  { name: 'Tilak Raj Sharma (Inventory Mgr)', school: 'HES', basic: 19309, ada: 6758, da: 965, epf: 1800, esi: 0, netPay: 25232 },
  { name: 'Vikas Thakur (System Coord)', school: 'HES', basic: 23056, ada: 8070, da: 1153, epf: 1800, esi: 0, netPay: 30479 },
  { name: 'Rakesh Kumar (Peon cum Driver)', school: 'HES', basic: 11462, ada: 4012, da: 573, epf: 1800, esi: 120, netPay: 14127 },
].forEach(e => {
  const ada = computeADA(e.basic, e.school);
  const da = computeDA(e.basic);
  const ns = e.basic + da + ada;
  const { epfEmployee } = computeEPF(ns);
  const esi = computeESI(ns);
  const netPay = (e.basic + ada + da) - epfEmployee - esi; // HES schema has no LENC/RRF/TDS for these 4 rows
  check(`${e.name} ADA`, ada, e.ada);
  check(`${e.name} DA`, da, e.da);
  check(`${e.name} EPF`, epfEmployee, e.epf);
  check(`${e.name} ESI`, esi, e.esi);
  check(`${e.name} NetPay`, netPay, e.netPay);
});

console.log('\n=== LMS 1 / Kullu (35% ADA) ===');
[
  { name: 'Veena Devi (PRT)', school: 'LMS 1', basic: 17072, ada: 5975, da: 854, epf: 1800, esi: 0, netPay: 22101 },
  { name: 'Jai Chand (Lab Asst)', school: 'LMS 1', basic: 19647, ada: 6876, da: 982, epf: 1800, esi: 0, netPay: 25705 },
  { name: 'Shiv Dassi (Helper)', school: 'LMS 1', basic: 7684, ada: 2689, da: 384, epf: 1291, esi: 81, netPay: 9385 },
  { name: 'Sunny Devi (Helper)', school: 'LMS 1', basic: 8086, ada: 2830, da: 404, epf: 1358, esi: 85, netPay: 9877 },
].forEach(e => {
  const ada = computeADA(e.basic, e.school);
  const da = computeDA(e.basic);
  const ns = e.basic + da + ada;
  const { epfEmployee } = computeEPF(ns);
  const esi = computeESI(ns);
  const netPay = (e.basic + ada + da) - epfEmployee - esi;
  check(`${e.name} ADA`, ada, e.ada);
  check(`${e.name} DA`, da, e.da);
  check(`${e.name} EPF`, epfEmployee, e.epf);
  check(`${e.name} ESI`, esi, e.esi);
  check(`${e.name} NetPay`, netPay, e.netPay);
});

console.log('\n=== LMS 3 / Dunkhra (25% ADA — the edge case) ===');
[
  { name: 'Seema Sharma (Helper)', school: 'LMS 3', basic: 6863, ada: 1716, da: 343, epf: 1071, esi: 67, netPay: 7784 },
  { name: 'Monika Behal (PRT)', school: 'LMS 3', basic: 12610, ada: 3153, da: 631, epf: 1800, esi: 123, netPay: 14471 },
  { name: 'Tirth Ram (Peon cum Driver)', school: 'LMS 3', basic: 10949, ada: 2737, da: 547, epf: 1708, esi: 107, netPay: 12418 },
  { name: 'Suresh Kumar (PRT)', school: 'LMS 3', basic: 13164, ada: 3291, da: 658, epf: 1800, esi: 128, netPay: 15185 },
].forEach(e => {
  const ada = computeADA(e.basic, e.school);
  const da = computeDA(e.basic);
  const ns = e.basic + da + ada;
  const { epfEmployee } = computeEPF(ns);
  const esi = computeESI(ns);
  const netPay = (e.basic + ada + da) - epfEmployee - esi;
  check(`${e.name} ADA`, ada, e.ada);
  check(`${e.name} DA`, da, e.da);
  check(`${e.name} EPF`, epfEmployee, e.epf);
  check(`${e.name} ESI`, esi, e.esi);
  check(`${e.name} NetPay`, netPay, e.netPay);
});

console.log('\n=== runPayrollForEmployee() full-chain wrapper (LMS3 Suresh Kumar) ===');
{
  const r = runPayrollForEmployee({ monthBasic: 13164, school: 'LMS 3' });
  check('full-chain ADA', r.ada, 3291);
  check('full-chain DA', r.da, 658);
  check('full-chain EPF', r.epfEmployee, 1800);
  check('full-chain ESI', r.esi, 128);
  check('full-chain GrSal', r.grossPay, 17113);
  check('full-chain CTI', r.cti, 18913);
  check('full-chain NetPay', r.netPay, 15185);
}

console.log('\n=== validateReconciliation() ===');
{
  // Real June 2026 totals — every one of these reconciled exactly by hand this session.
  const real = [
    ['HES', 307982, 307982],
    ['LMS1 Kullu', 446442, 446442],
    ['LMS2 Kelheli', 801305, 801305],
    ['LMS3 Dunkhra', 353783, 353783],
    ['LMS4 NerChowk', 478315, 478315],
    ['LMS5 Sayoli', 66234, 66234],
    ['SES', 101283, 101283],
  ];
  real.forEach(([label, bankTotal, sheetTotal]) => {
    const r = validateReconciliation(bankTotal, sheetTotal);
    check(`${label} reconciles`, r.ok, true);
  });
  // Deliberately corrupted row — must fail, not silently pass.
  const bad = validateReconciliation(307982, 307982 - 1);
  check('corrupted row correctly FAILS (diff=1)', bad.ok, false);
  check('corrupted row reports correct diff', bad.diff, 1);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
