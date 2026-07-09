// ═══════════════════════════════════════════════════════════════════
// PayrollRules.gs — PURE LOGIC ONLY.
//
// No SpreadsheetApp, no Session, no UrlFetchApp calls in this file.
// Every function takes plain numbers/objects in and returns plain
// numbers/objects out, so it can be (a) unit-tested with plain Node
// against real values read from the actual June 2026 Salary Sheet +
// Bank Letter PDFs, and (b) called identically from Apps Script once
// PayrollSheetIO.gs (the I/O layer) exists. See PayrollConstants.gs
// for every tunable number this file uses, and the plan file
// (/Users/udaykanwar/.claude/plans/staged-watching-heron.md) for the
// full reverse-engineering trail this reimplements.
// ═══════════════════════════════════════════════════════════════════

/**
 * @param {number} fullBasic - the employee's full-month Basic from their structuring record
 * @param {number} presentDays
 * @param {number} workingDaysInMonth - caller-supplied; see PRESENT_DAYS_DENOMINATOR_MODE
 *   in PayrollConstants.gs — this function deliberately does NOT hardcode a
 *   denominator since that rule is unconfirmed. Passing a wrong denominator
 *   here silently produces a wrong Month Basic for every employee, so the
 *   I/O layer must not call this until PRESENT_DAYS_DENOMINATOR_MODE is resolved.
 * @returns {number} Month Basic, rounded (matches the real sheets' integer values)
 */
function computeMonthBasic(fullBasic, presentDays, workingDaysInMonth) {
  if (!workingDaysInMonth) throw new Error('computeMonthBasic: workingDaysInMonth is required (PRESENT_DAYS_DENOMINATOR_MODE is unconfirmed — do not default this).');
  return Math.round(fullBasic * (presentDays / workingDaysInMonth));
}

/**
 * @param {number} monthBasic
 * @param {string} school - short key matching PAYROLL_SCHOOL_NAMES / ADA_RATE, e.g. 'LMS 3'
 * @returns {number} Additional Duty Allowance, rounded
 * Verified against real data with Math.round (not floor/ceil) — e.g.
 * Dunkhra's Seema Sharma: 6863 * 0.25 = 1715.75 -> actual ADA 1716.
 * LMS3 = 25% of Month Basic, all other checked schools = 35%.
 */
function computeADA(monthBasic, school) {
  const rate = ADA_RATE.hasOwnProperty(school) ? ADA_RATE[school] : 0.35;
  return Math.round(rate * monthBasic);
}

/**
 * @param {number} monthBasic
 * @returns {number} Dearness Allowance, rounded
 * Verified: 5% of Month Basic, no school-specific variation observed.
 * Uses Math.round — e.g. Kelheli's Mitter Bhushan: 17373 * 0.05 = 868.65
 * -> actual DA 869.
 */
function computeDA(monthBasic) {
  return Math.round(DA_RATE_DEFAULT * monthBasic);
}

/**
 * @param {{monthBasic:number, ada:number, da:number, lenc?:number, incentive?:number,
 *          travelAllowance?:number, tfA?:number}} parts
 * @returns {number} GrSal
 * Verified composition: MonthBasic + ADA + DA + LENC + INCENTIVE + TRAV.ALLOW + TF(A).
 */
function computeGrossPay(parts) {
  const lenc = parts.lenc || 0;
  const incentive = parts.incentive || 0;
  const travelAllowance = parts.travelAllowance || 0;
  const tfA = parts.tfA || 0;
  return parts.monthBasic + parts.ada + parts.da + lenc + incentive + travelAllowance + tfA;
}

/**
 * @param {number} ns - "net salary" base = monthBasic + da + ada (matches calcSalary()'s ns)
 * @returns {{epfEmployee:number, epfEmployer:number}}
 * Verified: 12% of ns (8.33% + 3.67% components), capped at Rs.1800 total
 * once ns exceeds Rs.15,000. Employer share is numerically identical to
 * the employee deduction (confirmed across every non-MD row checked) —
 * this is what the real sheets' single "EPF" column represents, used
 * both to inflate CTI and to deduct from GrSal. Uses Math.round on each
 * component — e.g. Kelheli's Nimme Ram: ns=14883, 14883*0.12=1785.96
 * -> actual EPF 1786.
 *
 * IMPORTANT: the 8.33%/3.67% split is rounded ONCE as a combined 12%,
 * not as two independently-rounded sub-components summed — those two
 * approaches disagree by Rs.1 whenever the fractional parts round
 * opposite ways (caught by Dunkhra's Seema Sharma: ns=8922, combined
 * round(8922*0.12)=1071 matches the real sheet; round(8922*0.0833)
 * + round(8922*0.0367) = 743+327 = 1070 does not).
 *
 * KNOWN EXCEPTION, not yet modeled: MDs (e.g. HES's Lalita Kanwar,
 * ns=77703) show EPF=0 in the real sheets despite ns far exceeding the
 * threshold — MDs/top-tier roles appear EPF-exempt. This function does
 * NOT apply that exemption; callers must special-case it by designation
 * until the exact exemption rule (which designations, why) is confirmed.
 */
function computeEPF(ns) {
  const over = ns > EPF_THRESHOLD;
  const epf = over
    ? (EPF_EMPLOYEE_CAP_COMPONENT_1 + EPF_EMPLOYEE_CAP_COMPONENT_2)
    : Math.round(ns * (EPF_EMPLOYEE_RATE + EPF_PENSION_RATE));
  return { epfEmployee: epf, epfEmployer: epf };
}

/**
 * @param {number} ns - same base as computeEPF
 * @returns {number} employee ESI deduction, rounded, 0 if ns >= threshold
 * Verified: 0.75% of ns when ns < Rs.21,000, else 0. Uses Math.round —
 * e.g. HES's Raj Bahadur: ns=16751, 16751*0.0075=125.6325 -> actual ESI 126.
 */
function computeESI(ns) {
  if (ns >= ESI_THRESHOLD) return 0;
  return Math.round(ns * ESI_EMPLOYEE_RATE);
}

/**
 * @param {number} grossPay
 * @param {number} epfEmployer
 * @returns {number} CTI (Cost to Institution)
 * Verified: CTI = GrSal + EPF (employer's matching share).
 */
function computeCTI(grossPay, epfEmployer) {
  return grossPay + epfEmployer;
}

/**
 * @param {{epf?:number, esi?:number, tds?:number, rrf?:number, otherRecovery?:number,
 *          security?:number, tfD?:number}} deductions
 * @returns {number} Total Ded
 * Verified composition: EPF + ESI + TDS + RRF + OTHERRECOVERY + SECURITY + TF(D).
 * tfD should equal the same run's tfA (the tuition-fee wash) — callers
 * are responsible for passing tfD === tfA; this function does not enforce it.
 */
function computeTotalDeductions(deductions) {
  const epf = deductions.epf || 0;
  const esi = deductions.esi || 0;
  const tds = deductions.tds || 0;
  const rrf = deductions.rrf || 0;
  const otherRecovery = deductions.otherRecovery || 0;
  const security = deductions.security || 0;
  const tfD = deductions.tfD || 0;
  return epf + esi + tds + rrf + otherRecovery + security + tfD;
}

/**
 * @param {number} grossPay
 * @param {number} totalDeductions
 * @returns {number} Net Pay
 * Verified: NetPay = GrSal - TotalDed on every sampled row.
 */
function computeNetPay(grossPay, totalDeductions) {
  return grossPay - totalDeductions;
}

/**
 * Runs the full per-employee chain in one call, given a fully-resolved
 * set of inputs (i.e. Month Basic proration and any manual entries
 * already decided by the caller/I-O layer). Convenience wrapper only —
 * each step is independently testable above.
 * @param {{monthBasic:number, school:string, lenc?:number, incentive?:number,
 *          travelAllowance?:number, tfA?:number, tds?:number, rrf?:number,
 *          otherRecovery?:number, security?:number}} input
 */
function runPayrollForEmployee(input) {
  const ada = computeADA(input.monthBasic, input.school);
  const da = computeDA(input.monthBasic);
  const ns = input.monthBasic + da + ada;
  const grossPay = computeGrossPay({
    monthBasic: input.monthBasic, ada, da,
    lenc: input.lenc, incentive: input.incentive,
    travelAllowance: input.travelAllowance, tfA: input.tfA,
  });
  const { epfEmployee, epfEmployer } = computeEPF(ns);
  const esi = computeESI(ns);
  const cti = computeCTI(grossPay, epfEmployer);
  const totalDeductions = computeTotalDeductions({
    epf: epfEmployee, esi, tds: input.tds || 0, rrf: input.rrf || 0,
    otherRecovery: input.otherRecovery || 0, security: input.security || 0,
    tfD: input.tfA || 0, // the wash — tfD mirrors tfA
  });
  const netPay = computeNetPay(grossPay, totalDeductions);
  return { ada, da, ns, grossPay, epfEmployee, epfEmployer, esi, cti, totalDeductions, netPay };
}

/**
 * @param {number} bankLetterTotal - sum of the Bank Letter's Net Payable column
 * @param {number} salarySheetGrandTotalNetPay - the Salary Sheet's Grand Total Net Pay
 * @returns {{ok:boolean, diff:number}}
 * This is the reconciliation check done by hand on 2026-07-09 across all
 * 7 entities in the June 2026 batch (all passed exactly). The Payroll
 * Portal must run this automatically and refuse to mark a run
 * "ready to sign" if it fails — see the plan's "hard validation gate."
 */
function validateReconciliation(bankLetterTotal, salarySheetGrandTotalNetPay) {
  const diff = bankLetterTotal - salarySheetGrandTotalNetPay;
  return { ok: diff === 0, diff: diff };
}

// Apps Script has no module system — when this file is copied into the
// Apps Script editor, remove this block (or leave it; typeof module
// is 'undefined' there, so it's inert). Kept only so this file can be
// unit-tested directly with plain Node (see the accompanying test script).
if (typeof module !== 'undefined') {
  module.exports = {
    computeMonthBasic, computeADA, computeDA, computeGrossPay,
    computeEPF, computeESI, computeCTI, computeTotalDeductions,
    computeNetPay, runPayrollForEmployee, validateReconciliation,
  };
}
