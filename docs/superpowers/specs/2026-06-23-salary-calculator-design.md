# Salary Calculator Web App — Design Spec

## Overview

Single-file HTML salary calculator for Chinese individual income tax, replicating the `SalaryCalculator.xlsx` logic. Mobile-first, all inputs adjustable, real-time computation. Zero dependencies.

## Architecture

```
Single HTML file (index.html)
├── <style> — CSS variables, mobile-first layout, dark theme
├── <body> — semantic HTML structure
│   ├── Result bar (sticky top) — 月到手, 年到手, 年总包
│   ├── Section ① Salary inputs (expanded by default)
│   ├── Section ② City & insurance/housing fund (expanded by default)
│   ├── Section ③ Advanced insurance config (collapsed)
│   ├── Section ④ Special deductions (collapsed)
│   └── Section ⑤ Monthly breakdown (collapsed)
└── <script> — all logic inline
    ├── Constants (tax tables, city data, deduction options)
    ├── State object (all input values)
    ├── Compute functions (exact Excel formula replication)
    └── Render/event binding
```

## Data Flow

```
Input change → update state → recompute → update DOM
                                    ├── Result bar
                                    ├── Breakdown details
                                    └── Monthly table (if expanded)
```

No virtual DOM, no state management library — direct DOM updates on change.

## Key Design Decisions

- **Single file**: HTML + inline CSS + inline JS, open directly in browser
- **Mobile-first**: Single column, sticky result bar, collapsible sections
- **All inputs adjustable**: Every Excel parameter exposed, secondary params collapsed by default
- **City presets**: Selecting a city auto-fills insurance/housing fund base limits; manual override supported
- **Deduction dropdowns**: Options match Info sheet lookup table exactly

## Computation (from Excel)

1. Annual salary = monthly × 12 + allowance
2. Year-end bonus = monthly × bonus_months
3. Insurance base = MEDIAN(monthly, lower, upper)
4. Taxable income (salary) = MAX(0, salary - exemption - deductions_total - insurance - housing_fund)
5. Tax (salary) = progressive rate from salary tax table
6. Tax (bonus) = bonus × rate from bonus tax table - quick deduction
7. Monthly tax = cumulative withholding method (12 months)
8. Take-home = total_package - insurance - tax(salary) - tax(bonus) + housing_fund

## Data Tables (embedded JS)

- Salary progressive tax rate table (7 brackets)
- Year-end bonus tax rate table (7 brackets)
- City insurance/housing fund base limits (Beijing, Shanghai, Guangzhou, Shenzhen, Hangzhou)
- Special deduction options (6 categories × multiple tiers)

## Scope

**In scope**: All Excel inputs and outputs, mobile-responsive layout, real-time computation
**Out of scope**: Data persistence, export, multi-year comparison, tax optimization suggestions

## Testing

- Manual verification against Excel values for sample inputs
- Edge cases: salary below insurance lower bound, above upper bound, zero bonus, all deductions enabled
