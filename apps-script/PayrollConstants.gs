// ═══════════════════════════════════════════════════════════════════
// PayrollConstants.gs — pure config, no SpreadsheetApp/Session calls.
//
// Every tunable number the monthly Payroll Engine uses lives here, not
// scattered through PayrollRules.gs, mirroring Constants.gs's pattern
// for the Incentive/Leave engines. Rates below were verified directly
// against 13 real HES/LMS Salary Sheet + Bank Letter PDFs (June 2026
// run, 158 employees across HES + LMS 1-5 + SES) read and cross-checked
// on 2026-07-09 — see the plan file
// (/Users/udaykanwar/.claude/plans/staged-watching-heron.md) for the
// full reverse-engineering trail. Where a figure could NOT be confirmed
// from that sample, it is marked UNCONFIRMED below rather than guessed.
// ═══════════════════════════════════════════════════════════════════

// ── School/entity canonical keys ────────────────────────────────────
// The actual documents use human names ("La Montessori School Kullu"),
// dashboard.html's D.base pay-grade table uses short codes ("LMS 1").
// This is the single mapping both the offer calculator and the monthly
// engine should read from — do not duplicate it inline elsewhere.
const PAYROLL_SCHOOL_NAMES = {
  HES:    'La Montessori School HES',
  'LMS 1': 'La Montessori School Kullu',    // Dhalpur
  'LMS 2': 'La Montessori School Kelheli',
  'LMS 3': 'La Montessori School Dunkhra',
  'LMS 4': 'La Montessori School NerChowk',
  'LMS 5': 'LMS 5 - Sayoli',
  'LMS 6': null, // Jogindernagar — salary sheet not yet supplied, name TBD when it arrives
  SES:    'Swadhyaya Educational Services',
};

// ── Additional Duty Allowance rate, by school ───────────────────────
// Verified exactly against real data: e.g. LMS3/Dunkhra's Seema Sharma
// (Basic 6863, ADA 1716 = 25.0%) and LMS1/Kullu's Chandra Devi
// (Basic 12848, ADA 4497 = 35.0%). Copied from dashboard.html's ADA_R —
// this is now the single source of truth; dashboard.html should read
// from here once PayrollConstants.gs is loaded in that project too.
// LMS 5/6 and SES rates NOT independently verified (LMS5's sample was
// too small/uniform to distinguish 35% from another rate with confidence;
// SES had no salary sheet at all this run) — kept at the dashboard.html
// default (0.35) but flagged.
const ADA_RATE = {
  HES: 0.35,
  'LMS 1': 0.35,
  'LMS 2': 0.35,
  'LMS 3': 0.25,
  'LMS 4': 0.35,
  'LMS 5': 0.35,   // UNVERIFIED for this campus specifically — inherited default
  'LMS 6': 0.25,   // UNVERIFIED — no data yet, matches dashboard.html's existing guess
  SES: 0.35,       // UNVERIFIED — no salary sheet available to check
};

// ── Dearness Allowance rate ──────────────────────────────────────────
// Verified exactly against every row checked across HES/LMS1/LMS2/LMS3:
// DA = 5% of Month Basic, no school-specific variation observed.
const DA_RATE_DEFAULT = 0.05;

// ── EPF (employee + matching employer share) ────────────────────────
// Verified: 12% of (Month Basic + DA + ADA), capped at Rs.1800/month
// once that base exceeds Rs.15,000 — matches dashboard.html's
// epfE (8.33%, capped 1250) + pfE (3.67%, capped 550) split exactly.
// CTI = GrSal + EPF (employer's matching share, same number).
const EPF_EMPLOYEE_RATE = 0.0833;
const EPF_PENSION_RATE = 0.0367; // combined with above = 12.00%
const EPF_EMPLOYEE_CAP_COMPONENT_1 = 1250; // cap on the 8.33% component above the threshold
const EPF_EMPLOYEE_CAP_COMPONENT_2 = 550;  // cap on the 3.67% component above the threshold
const EPF_THRESHOLD = 15000; // base (Basic+DA+ADA) above which the flat caps apply instead of %

// ── ESI (employee share only — this engine's scope) ─────────────────
// Verified: 0.75% of (Month Basic + DA + ADA) when that figure is
// under Rs.21,000; 0 otherwise. Matches dashboard.html's esiEe exactly.
const ESI_EMPLOYEE_RATE = 0.0075;
const ESI_THRESHOLD = 21000;

// ── Present-days proration ──────────────────────────────────────────
// KNOWN UNRESOLVED — do not silently pick a side. The real sheets show
// `Present` as a decimal (e.g. 17.84, 28.59) prorating Month Basic
// against SOME denominator, but the sample didn't disambiguate between
// calendar days in the month and a working-days-only calendar (Sunday +
// 2nd Saturday + GH holiday events, campus-specific — the same rule
// WhatsApp Hub's isHolidayDate() already encodes, per
// whatsapp-hub/server.js:1062-1068). computeMonthBasic() below takes
// the denominator as an explicit parameter rather than hardcoding one —
// callers MUST supply it; PRESENT_DAYS_DENOMINATOR_MODE exists only to
// make the open question visible in code, not to silently resolve it.
const PRESENT_DAYS_DENOMINATOR_MODE = 'UNCONFIRMED'; // one of: 'CALENDAR_DAYS', 'WORKING_DAYS' — set once confirmed with Uday

// ── TDS ───────────────────────────────────────────────────────────────
// No consistent %-of-salary pattern found across the sample (e.g. MDs
// showed flat Rs.25,000/Rs.10,000 figures with no visible slab logic).
// Default: manual per-employee entry, NOT a computed slab. Revisit only
// if Uday confirms a real income-tax-slab rule HR is actually applying.
const TDS_MODE = 'MANUAL_ENTRY';

// ── RRF (monthly installment shown on the payroll register) ─────────
// KNOWN UNRESOLVED — the monthly RRF figures seen (Rs.5,000-25,000
// range) did not cleanly match dashboard.html's annual rrf1/rrf2/rrf3
// (12%/9%/6% of CTI, from calcSalary()) divided by 12 on the rows
// checked. RRF_MONTHLY_SUGGESTS_DIVIDE_ANNUAL_BY_12 controls whether
// the Payroll Engine pre-fills a suggested RRF value for accounts to
// review/override (true) or leaves the field fully blank (false) until
// the real rule is confirmed — default false until confirmed, since a
// wrong pre-fill that accounts doesn't catch is worse than an empty box.
const RRF_MONTHLY_SUGGESTS_DIVIDE_ANNUAL_BY_12 = false;

// ── Tuition Fee wash (TF(A) / TF(D)) ─────────────────────────────────
// Confirmed: every TF(A) value in the real sheets reappears exactly as
// TF(D) in the same row (added to Gross, deducted back — net-neutral
// on take-home). Structurally identical to calcSalary()'s tuiFee.
// No separate rate needed here — TF(A) is a direct pass-through of the
// employee's staff-kid tuition fee value from Employee_Master.

// ── Ref No. sequence ──────────────────────────────────────────────────
// Confirmed from the real June 2026 batch: HES/LMS/SAL/{n}/{FY} uses
// ONE shared incrementing counter across ALL entities in a run (HES=18,
// LMS1=19, LMS2=20, LMS3=21, LMS4=22, LMS5=23), not a per-campus counter.
// The I/O layer (PayrollSheetIO.gs) must draw the next {n} from a single
// shared counter cell/row, not one per entity.
const PAYROLL_REF_NO_TEMPLATE = 'HES/LMS/SAL/{n}/{fy}({fyRange})';

// ── Total Deduction / Net Pay component list ─────────────────────────
// Confirmed column order and composition from every sampled row:
//   GrSal    = MonthBasic + ADA + DA + LENC + INCENTIVE + TRAV.ALLOW + TF(A)
//   CTI      = GrSal + EPF
//   TotalDed = EPF + ESI + TDS + RRF + OTHERRECOVERY + SECURITY + TF(D)
//   NetPay   = GrSal - TotalDed
// LENC (leave encashment), INCENTIVE, TRAV.ALLOW, TDS, OTHERRECOVERY and
// SECURITY are manual/derived-elsewhere inputs, not formulas of this file.

// Apps Script has no module system — inert there (typeof module is
// 'undefined'), present only so PayrollRules.gs's Node unit tests can
// pull these constants in via require(). See PayrollRules.gs's own
// trailing block for the matching pattern.
if (typeof module !== 'undefined') {
  module.exports = {
    PAYROLL_SCHOOL_NAMES, ADA_RATE, DA_RATE_DEFAULT,
    EPF_EMPLOYEE_RATE, EPF_PENSION_RATE, EPF_EMPLOYEE_CAP_COMPONENT_1,
    EPF_EMPLOYEE_CAP_COMPONENT_2, EPF_THRESHOLD,
    ESI_EMPLOYEE_RATE, ESI_THRESHOLD,
    PRESENT_DAYS_DENOMINATOR_MODE, TDS_MODE,
    RRF_MONTHLY_SUGGESTS_DIVIDE_ANNUAL_BY_12, PAYROLL_REF_NO_TEMPLATE,
  };
}
