# LMCS Bye Laws 2026

Static HTML website containing the Bye Laws for **La Montessori Chain of Schools (LMCS)** — a confidential internal document for staff and parents.

## Deployment

This is a pure static site with no build step. Upload the entire folder to any static host:

- **GitHub Pages**: Push to a repo, enable Pages from the repo settings, and set the source to the root of `main`.
- **Netlify / Vercel**: Drag and drop the folder, or connect the GitHub repo.

The root `index.html` is the entry point.

## Structure

```
lmcs-bye-laws/
├── index.html                          # Main landing page
├── style.css                           # Shared stylesheet
├── abbreviations.html
├── vision-mission.html
│
├── student-code-of-conduct/
│   ├── index.html
│   ├── faq.html
│   ├── starting-at-lms/
│   │   ├── use-of-almanac.html
│   │   ├── communication-with-school.html
│   │   └── school-calendars.html
│   ├── transport-attendance-timing/
│   │   ├── absence-from-school.html
│   │   ├── late-arrival.html
│   │   ├── transport-facility.html
│   │   └── school-timing-vacations.html
│   ├── fees-administrative-procedures/
│   │   ├── general-fees-rules.html
│   │   ├── fees-2026.html
│   │   ├── parent-discounts-incentives.html
│   │   ├── new-admission-special-cases.html
│   │   ├── information-changes.html
│   │   └── withdrawal-transfer-certificate.html
│   ├── behaviour-student-culture/
│   │   ├── code-of-conduct.html
│   │   ├── school-rules.html
│   │   ├── misconduct.html
│   │   ├── school-uniform.html
│   │   ├── rewards-punishments.html
│   │   └── house-system.html
│   └── tips-successful-academic-life/
│       ├── reading-habits.html
│       ├── extra-curricular-activities.html
│       ├── role-of-homework.html
│       ├── private-tuition.html
│       └── examination-cwa-rules.html
│
├── employee-code-of-conduct/
│   ├── index.html
│   ├── appointments-hr/
│   │   ├── process-of-hiring.html
│   │   ├── rrf-security.html
│   │   ├── fines.html
│   │   └── resignation-rules.html
│   ├── service-rules-leaves-job-roles/
│   │   ├── job-roles.html
│   │   ├── leave-vacation-rules.html
│   │   └── incentives-bonuses.html
│   ├── reports/
│   │   ├── registers-at-school.html
│   │   └── evaluation-of-performance.html
│   ├── inventory-infrastructure/
│   │   └── index.html
│   ├── transport-logistics/
│   │   ├── outstation-driver-rules.html
│   │   └── transport-fees-2026.html
│   ├── school-rules/
│   │   ├── code-of-conduct.html
│   │   └── communication-at-lms.html
│   └── academics-examination/
│       └── examination-process.html
│
└── administrative-code-of-conduct/
    └── index.html
```

## Design

- Brand colour: `#CE0000` (LMCS Red)
- Responsive sidebar + main content layout
- Breadcrumb navigation
- Card grid for section landing pages
- Office Order callout boxes for formal school orders
- No JavaScript, no dependencies — works offline

## Source

Content extracted from the **LMCS Bye Laws 2026** Notion workspace. This document is **Confidential — Internal Use Only**.
