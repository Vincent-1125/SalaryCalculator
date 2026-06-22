# Salary Calculator Web App — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single-file HTML salary calculator replicating SalaryCalculator.xlsx logic, mobile-first, all inputs adjustable.

**Architecture:** Single `index.html` with inline CSS + JS. Mobile-first single column layout with sticky result bar, collapsible input sections, and real-time DOM updates on any input change.

**Tech Stack:** Plain HTML5, CSS3 (custom properties), vanilla JavaScript (ES2020+). Zero dependencies, zero build step.

**Design spec:** `docs/superpowers/specs/2026-06-23-salary-calculator-design.md`

---

### Task 1: HTML Skeleton & Semantic Structure

**Files:**
- Create: `index.html`

- [ ] **Step 1: Write the HTML structure**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>薪资计算器</title>
  <style>
    /* CSS will go here — placeholder for now */
    :root {
      --color-bg: #0f1117;
      --color-surface: #1a1d27;
      --color-surface-raised: #242836;
      --color-text: #e8eaed;
      --color-text-secondary: #9aa0a6;
      --color-accent: #8ab4f8;
      --color-accent-dim: #4a6fa5;
      --color-positive: #34a853;
      --color-warning: #fbbc04;
      --color-border: #2d3140;
      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 16px;
      --text-xs: 0.75rem;
      --text-sm: 0.875rem;
      --text-base: 1rem;
      --text-lg: 1.25rem;
      --text-xl: 1.5rem;
      --text-2xl: 2rem;
      --duration-fast: 150ms;
      --duration-normal: 250ms;
      --ease-out: cubic-bezier(0.16, 1, 0.3, 1);
    }

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    html { font-size: 16px; -webkit-text-size-adjust: 100%; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
      background: var(--color-bg);
      color: var(--color-text);
      line-height: 1.5;
      min-height: 100dvh;
      padding-bottom: env(safe-area-inset-bottom);
    }

    /* Result Bar — sticky top */
    .result-bar {
      position: sticky;
      top: 0;
      z-index: 10;
      background: var(--color-surface);
      border-bottom: 1px solid var(--color-border);
      padding: 16px 20px;
      display: flex;
      flex-direction: column;
      gap: 4px;
    }
    .result-bar__label { font-size: var(--text-xs); color: var(--color-text-secondary); }
    .result-bar__row { display: flex; justify-content: space-between; align-items: baseline; }
    .result-bar__main {
      font-size: var(--text-2xl);
      font-weight: 700;
      color: var(--color-accent);
      letter-spacing: -0.02em;
    }
    .result-bar__sub {
      font-size: var(--text-sm);
      color: var(--color-text-secondary);
    }

    /* Main content */
    .main { padding: 16px 20px; display: flex; flex-direction: column; gap: 12px; }

    /* Section (collapsible) */
    .section {
      background: var(--color-surface);
      border-radius: var(--radius-md);
      border: 1px solid var(--color-border);
      overflow: hidden;
    }
    .section__header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 16px;
      cursor: pointer;
      user-select: none;
      -webkit-tap-highlight-color: transparent;
    }
    .section__header:active { background: var(--color-surface-raised); }
    .section__title { font-size: var(--text-base); font-weight: 600; }
    .section__arrow {
      width: 20px; height: 20px;
      transition: transform var(--duration-fast) var(--ease-out);
      color: var(--color-text-secondary);
    }
    .section--open .section__arrow { transform: rotate(180deg); }
    .section__body { display: none; padding: 0 16px 16px; }
    .section--open .section__body { display: block; }

    /* Form groups */
    .form-group { margin-bottom: 12px; }
    .form-group:last-child { margin-bottom: 0; }
    .form-label { display: block; font-size: var(--text-sm); color: var(--color-text-secondary); margin-bottom: 6px; }
    .form-row { display: flex; gap: 8px; align-items: center; }
    .form-input, .form-select {
      flex: 1;
      background: var(--color-surface-raised);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-sm);
      padding: 10px 12px;
      color: var(--color-text);
      font-size: var(--text-base);
      font-family: inherit;
      -webkit-appearance: none;
      appearance: none;
      min-width: 0;
    }
    .form-input:focus, .form-select:focus {
      outline: none;
      border-color: var(--color-accent);
      box-shadow: 0 0 0 2px rgba(138, 180, 248, 0.15);
    }
    .form-unit {
      font-size: var(--text-sm);
      color: var(--color-text-secondary);
      white-space: nowrap;
      flex-shrink: 0;
    }

    /* Deduction item */
    .deduction-item {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 0;
      border-bottom: 1px solid var(--color-border);
    }
    .deduction-item:last-child { border-bottom: none; }
    .deduction-item__label { font-size: var(--text-sm); }
    .deduction-item__amount { font-size: var(--text-sm); color: var(--color-accent); font-weight: 600; }

    /* Monthly table */
    .monthly-table {
      width: 100%;
      border-collapse: collapse;
      font-size: var(--text-xs);
    }
    .monthly-table th, .monthly-table td {
      padding: 8px 6px;
      text-align: right;
      border-bottom: 1px solid var(--color-border);
    }
    .monthly-table th {
      color: var(--color-text-secondary);
      font-weight: 500;
      position: sticky;
      top: 0;
      background: var(--color-surface);
    }
    .monthly-table td:first-child, .monthly-table th:first-child { text-align: center; }
    .monthly-table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  </style>
</head>
<body>
  <!-- Result Bar -->
  <div class="result-bar" id="resultBar">
    <span class="result-bar__label">月到手（含公积金）</span>
    <div class="result-bar__row">
      <span class="result-bar__main" id="monthlyTakeHome">¥0</span>
      <span class="result-bar__sub">年到手 <span id="yearlyTakeHome">¥0</span></span>
    </div>
    <div class="result-bar__row">
      <span class="result-bar__sub">年总包 <span id="totalPackage">¥0</span></span>
      <span class="result-bar__sub">年个税 <span id="totalTax">¥0</span></span>
    </div>
  </div>

  <div class="main">
    <!-- Section 1: Salary -->
    <div class="section section--open" id="secSalary">
      <div class="section__header" onclick="toggleSection('secSalary')">
        <span class="section__title">💰 薪资</span>
        <svg class="section__arrow" viewBox="0 0 20 20" fill="currentColor"><path d="M5 7l5 5 5-5" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round"/></svg>
      </div>
      <div class="section__body">
        <div class="form-group">
          <label class="form-label">月薪</label>
          <div class="form-row">
            <input class="form-input" type="number" id="monthlySalary" value="20000" min="0" step="100" oninput="recalc()">
            <span class="form-unit">元/月</span>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">津贴（年）</label>
          <div class="form-row">
            <input class="form-input" type="number" id="annualAllowance" value="14400" min="0" step="100" oninput="recalc()">
            <span class="form-unit">元/年</span>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">年终期数</label>
          <div class="form-row">
            <input class="form-input" type="number" id="bonusMonths" value="4" min="0" max="12" step="0.5" oninput="recalc()">
            <span class="form-unit">月</span>
          </div>
        </div>
      </div>
    </div>

    <!-- Section 2: City & Insurance -->
    <div class="section section--open" id="secCity">
      <div class="section__header" onclick="toggleSection('secCity')">
        <span class="section__title">🏙️ 城市与五险一金</span>
        <svg class="section__arrow" viewBox="0 0 20 20" fill="currentColor"><path d="M5 7l5 5 5-5" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round"/></svg>
      </div>
      <div class="section__body">
        <div class="form-group">
          <label class="form-label">城市</label>
          <select class="form-select" id="citySelect" onchange="onCityChange()">
            <option value="hangzhou">杭州</option>
            <option value="beijing">北京</option>
            <option value="shanghai">上海</option>
            <option value="guangzhou">广州</option>
            <option value="shenzhen">深圳</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">公积金缴存比例</label>
          <div class="form-row">
            <input class="form-input" type="number" id="housingFundRate" value="12" min="5" max="12" step="1" oninput="recalc()">
            <span class="form-unit">%</span>
          </div>
        </div>
      </div>
    </div>

    <!-- Section 3: Advanced Insurance Config -->
    <div class="section" id="secAdvancedInsurance">
      <div class="section__header" onclick="toggleSection('secAdvancedInsurance')">
        <span class="section__title">⚙️ 五险一金高级配置</span>
        <svg class="section__arrow" viewBox="0 0 20 20" fill="currentColor"><path d="M5 7l5 5 5-5" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round"/></svg>
      </div>
      <div class="section__body" id="advancedInsuranceBody">
        <!-- Filled by JS -->
      </div>
    </div>

    <!-- Section 4: Special Deductions -->
    <div class="section" id="secDeductions">
      <div class="section__header" onclick="toggleSection('secDeductions')">
        <span class="section__title">📋 专项附加扣除</span>
        <svg class="section__arrow" viewBox="0 0 20 20" fill="currentColor"><path d="M5 7l5 5 5-5" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round"/></svg>
      </div>
      <div class="section__body" id="deductionsBody">
        <!-- Filled by JS -->
      </div>
    </div>

    <!-- Section 5: Monthly Breakdown -->
    <div class="section" id="secMonthly">
      <div class="section__header" onclick="toggleSection('secMonthly')">
        <span class="section__title">📅 每月个税明细</span>
        <svg class="section__arrow" viewBox="0 0 20 20" fill="currentColor"><path d="M5 7l5 5 5-5" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round"/></svg>
      </div>
      <div class="section__body">
        <div class="monthly-table-wrap">
          <table class="monthly-table">
            <thead>
              <tr>
                <th>月份</th><th>月薪</th><th>累计应税所得额</th><th>累计应纳税额</th><th>本月纳税</th><th>本月到手</th>
              </tr>
            </thead>
            <tbody id="monthlyTableBody">
              <!-- Filled by JS -->
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <script>
    // JS placeholder — all logic goes here
    'use strict';
  </script>
</body>
</html>
```

- [ ] **Step 2: Verify HTML loads without errors**

Open `index.html` in a browser. Verify dark background, result bar visible, two expanded sections (Salary, City), three collapsed sections, no JS errors in console.

---

### Task 2: JavaScript Data Constants

**Files:**
- Modify: `index.html` — add constants inside `<script>` tag

- [ ] **Step 3: Add all static data constants**

```javascript
// --- Tax rate tables ---

// Salary progressive tax table (annual cumulative)
const SALARY_TAX_TABLE = [
  { threshold: 0, rate: 0.03, quickDeduction: 0 },
  { threshold: 36000, rate: 0.10, quickDeduction: 2520 },
  { threshold: 144000, rate: 0.20, quickDeduction: 16920 },
  { threshold: 300000, rate: 0.25, quickDeduction: 31920 },
  { threshold: 420000, rate: 0.30, quickDeduction: 52920 },
  { threshold: 660000, rate: 0.35, quickDeduction: 85920 },
  { threshold: 960000, rate: 0.45, quickDeduction: 181920 },
];

// Year-end bonus tax table (separate calculation)
const BONUS_TAX_TABLE = [
  { threshold: 0, rate: 0.03, quickDeduction: 0 },
  { threshold: 36000, rate: 0.10, quickDeduction: 210 },
  { threshold: 144000, rate: 0.20, quickDeduction: 1410 },
  { threshold: 300000, rate: 0.25, quickDeduction: 2660 },
  { threshold: 420000, rate: 0.30, quickDeduction: 4410 },
  { threshold: 660000, rate: 0.35, quickDeduction: 7160 },
  { threshold: 960000, rate: 0.45, quickDeduction: 15160 },
];

// --- City insurance/housing fund base limits ---
const CITY_DATA = {
  beijing: { name: '北京', siUpper: 35811, siLower: 7162, hfUpper: 35811, hfLower: 2540 },
  shanghai: { name: '上海', siUpper: 37302, siLower: 7460, hfUpper: 37302, hfLower: 2690 },
  guangzhou: { name: '广州', siUpper: 27549, siLower: 5510, hfUpper: 39828, hfLower: 2500 },
  shenzhen: { name: '深圳', siUpper: 27549, siLower: 4775, hfUpper: 44265, hfLower: 2360 },
  hangzhou: { name: '杭州', siUpper: 25300, siLower: 4986, hfUpper: 40694, hfLower: 2490 },
};

// --- Social insurance rates ---
const SI_RATE_TOTAL = 0.105;   // Total social insurance
const SI_RATE_PENSION = 0.08;  // Pension
const SI_RATE_MEDICAL = 0.02;  // Medical
const SI_RATE_UNEMPLOY = 0.005; // Unemployment

// Annual tax exemption
const TAX_EXEMPTION = 60000; // 5000/month * 12
const MONTHLY_BASIC_EXEMPTION = 5000;

// --- Special deduction options ---
const DEDUCTION_OPTIONS = {
  childrenEducation: {
    label: '子女教育扣除',
    options: [
      { label: '不符合，扣除0元', value: 0 },
      { label: '一个孩子，夫妻各扣除1000元', value: 1000 },
      { label: '一个孩子，仅有一方扣除2000元', value: 2000 },
      { label: '两个孩子，夫妻各扣除2000元', value: 2000 },
      { label: '两个孩子，仅有一方扣除4000元', value: 4000 },
    ],
  },
  continuingEducation: {
    label: '继续教育扣除',
    options: [
      { label: '不符合，扣除0元', value: 0 },
      { label: '接受学历教育中，扣除400元', value: 400 },
    ],
  },
  elderlySupport: {
    label: '赡养老人扣除',
    options: [
      { label: '不符合，扣除0元', value: 0 },
      { label: '独生子女，扣除3000元', value: 3000 },
      { label: '两兄弟姐妹分摊，扣除1500元', value: 1500 },
      { label: '三兄弟姐妹分摊，扣除1000元', value: 1000 },
    ],
  },
  housingLoanInterest: {
    label: '住房贷款利息扣除',
    options: [
      { label: '不符合，扣除0元', value: 0 },
      { label: '夫妻各扣除500元', value: 500 },
      { label: '仅由一方扣除1000元', value: 1000 },
    ],
  },
  housingRent: {
    label: '住房租金扣除',
    options: [
      { label: '不符合，扣除0元', value: 0 },
      { label: '所在城市户籍人口少于100万，扣除800元', value: 800 },
      { label: '所在城市户籍人口超过100万，扣除1100元', value: 1100 },
      { label: '直辖市、省会城市、计划单列市等，扣除1500元', value: 1500 },
    ],
  },
  infantCare: {
    label: '3岁以下婴幼儿照护',
    options: [
      { label: '不符合，扣除0元', value: 0 },
      { label: '一方扣除2000元', value: 2000 },
      { label: '夫妻各扣除1000元', value: 1000 },
    ],
  },
};

// --- Helper: look up tax from a tax table ---
function calcTax(taxableIncome, table) {
  if (taxableIncome <= 0) return 0;
  let bracket = table[0];
  for (const b of table) {
    if (taxableIncome > b.threshold) bracket = b;
  }
  return taxableIncome * bracket.rate - bracket.quickDeduction;
}

// --- Helper: MEDIAN for base clamping ---
function clamp(value, lower, upper) {
  return Math.max(lower, Math.min(upper, value));
}
```

---

### Task 3: State Management & Computation Engine

**Files:**
- Modify: `index.html` — add state + compute functions inside `<script>` tag

- [ ] **Step 4: Add state object and read/write helpers**

```javascript
// --- State ---
// All inputs are read from DOM directly; this is for computed intermediates
let state = {
  // Inputs (synced from DOM)
  monthlySalary: 20000,
  annualAllowance: 14400,
  bonusMonths: 4,
  city: 'hangzhou',
  housingFundRate: 0.12,

  // Advanced insurance (synced from DOM, default from city)
  siUpper: 25300,
  siLower: 4986,
  hfUpper: 40694,
  hfLower: 2490,
  siRateTotal: 0.105,
  siRatePension: 0.08,
  siRateMedical: 0.02,
  siRateUnemploy: 0.005,
  taxExemption: 60000,

  // Deductions (synced from DOM)
  deductions: {
    childrenEducation: 0,
    continuingEducation: 0,
    elderlySupport: 0,
    housingLoanInterest: 0,
    housingRent: 1500,
    infantCare: 0,
    personalPension: 0,
  },

  // Computed results
  annualSalary: 0,
  bonus: 0,
  totalPackage: 0,
  siBase: 0,
  hfBase: 0,
  annualSI: 0,
  annualPension: 0,
  annualMedical: 0,
  annualUnemploy: 0,
  annualHF: 0,
  totalDeductions: 0,
  taxableIncomeSalary: 0,
  taxableIncomeBonus: 0,
  taxSalary: 0,
  taxBonus: 0,
  annualTakeHome: 0,
  monthlyTakeHome: 0,
  monthlyRows: [],
};

function readInputs() {
  const s = state;
  s.monthlySalary = parseFloat(document.getElementById('monthlySalary').value) || 0;
  s.annualAllowance = parseFloat(document.getElementById('annualAllowance').value) || 0;
  s.bonusMonths = parseFloat(document.getElementById('bonusMonths').value) || 0;
  s.city = document.getElementById('citySelect').value;
  s.housingFundRate = (parseFloat(document.getElementById('housingFundRate').value) || 12) / 100;

  // Advanced insurance (will be populated by renderAdvancedInsurance)
  s.siRateTotal = (parseFloat(document.getElementById('siRateTotal')?.value) || 10.5) / 100;
  s.siRatePension = (parseFloat(document.getElementById('siRatePension')?.value) || 8) / 100;
  s.siRateMedical = (parseFloat(document.getElementById('siRateMedical')?.value) || 2) / 100;
  s.siRateUnemploy = (parseFloat(document.getElementById('siRateUnemploy')?.value) || 0.5) / 100;

  // Deductions
  for (const key of Object.keys(s.deductions)) {
    if (key === 'personalPension') {
      s.deductions.personalPension = parseFloat(document.getElementById('deduction_personalPension')?.value) || 0;
    } else {
      const sel = document.getElementById('deduction_' + key);
      if (sel) s.deductions[key] = parseFloat(sel.value) || 0;
    }
  }
}
```

- [ ] **Step 5: Add core computation function**

```javascript
function compute() {
  const s = state;

  // 1. Income items
  s.annualSalary = s.monthlySalary * 12 + s.annualAllowance;
  s.bonus = s.monthlySalary * s.bonusMonths;
  s.totalPackage = s.annualSalary + s.bonus;

  // 2. Insurance & housing fund bases (MEDIAN clamping)
  s.siBase = clamp(s.monthlySalary, s.siLower, s.siUpper);
  s.hfBase = clamp(s.monthlySalary, s.hfLower, s.hfUpper);

  // 3. Annual insurance & housing fund
  s.annualPension = s.siBase * s.siRatePension * 12;
  s.annualMedical = s.siBase * s.siRateMedical * 12;
  s.annualUnemploy = s.siBase * s.siRateUnemploy * 12;
  s.annualSI = s.annualPension + s.annualMedical + s.annualUnemploy;
  s.annualHF = s.hfBase * s.housingFundRate * 12;

  // 4. Total special deductions (monthly * 12)
  s.totalDeductions = 0;
  for (const [key, val] of Object.entries(s.deductions)) {
    s.totalDeductions += val;
  }

  // 5. Taxable income
  s.taxableIncomeSalary = Math.max(0, s.annualSalary - s.taxExemption - s.totalDeductions * 12 - s.annualSI - s.annualHF);
  s.taxableIncomeBonus = Math.max(0, s.bonus);

  // 6. Tax
  s.taxSalary = calcTax(s.taxableIncomeSalary, SALARY_TAX_TABLE);
  s.taxBonus = calcTax(s.taxableIncomeBonus, BONUS_TAX_TABLE);

  // 7. Take-home
  s.annualTakeHome = s.totalPackage - s.annualSI + s.annualHF - s.taxSalary - s.taxBonus;
  s.monthlyTakeHome = s.annualTakeHome / 12;

  // 8. Monthly cumulative tax calculation
  s.monthlyRows = [];
  const monthlyDeduction = MONTHLY_BASIC_EXEMPTION + s.totalDeductions + (s.annualSI + s.annualHF) / 12;
  let cumulativeTaxable = 0;
  let prevCumulativeTax = 0;

  for (let month = 1; month <= 12; month++) {
    cumulativeTaxable += s.monthlySalary - monthlyDeduction;
    const cumulativeTax = calcTax(Math.max(0, cumulativeTaxable), SALARY_TAX_TABLE);
    const monthTax = cumulativeTax - prevCumulativeTax;
    const monthTakeHome = s.monthlySalary - monthTax - (s.annualSI + s.annualHF) / 12;
    s.monthlyRows.push({
      month,
      salary: s.monthlySalary,
      cumulativeTaxable: Math.max(0, cumulativeTaxable),
      cumulativeTax,
      monthTax,
      monthTakeHome,
    });
    prevCumulativeTax = cumulativeTax;
  }
}
```

---

### Task 4: Dynamic Section Rendering (Advanced Insurance, Deductions, Monthly Table)

**Files:**
- Modify: `index.html` — add render functions

- [ ] **Step 6: Add render functions for dynamic content**

```javascript
// --- Render advanced insurance config ---
function renderAdvancedInsurance() {
  const container = document.getElementById('advancedInsuranceBody');
  const c = CITY_DATA[state.city] || CITY_DATA.hangzhou;
  container.innerHTML = `
    <div class="form-group">
      <label class="form-label">社保缴费基数上限</label>
      <div class="form-row">
        <input class="form-input" type="number" id="siUpper" value="${c.siUpper}" min="0" step="100" oninput="recalc()">
        <span class="form-unit">元/月</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">社保缴费基数下限</label>
      <div class="form-row">
        <input class="form-input" type="number" id="siLower" value="${c.siLower}" min="0" step="100" oninput="recalc()">
        <span class="form-unit">元/月</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">公积金缴存基数上限</label>
      <div class="form-row">
        <input class="form-input" type="number" id="hfUpper" value="${c.hfUpper}" min="0" step="100" oninput="recalc()">
        <span class="form-unit">元/月</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">公积金缴存基数下限</label>
      <div class="form-row">
        <input class="form-input" type="number" id="hfLower" value="${c.hfLower}" min="0" step="100" oninput="recalc()">
        <span class="form-unit">元/月</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">社保总缴费比例</label>
      <div class="form-row">
        <input class="form-input" type="number" id="siRateTotal" value="${state.siRateTotal * 100}" min="0" max="100" step="0.1" oninput="recalc()">
        <span class="form-unit">%</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">养老保险个人缴纳比例</label>
      <div class="form-row">
        <input class="form-input" type="number" id="siRatePension" value="${state.siRatePension * 100}" min="0" max="100" step="0.1" oninput="recalc()">
        <span class="form-unit">%</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">医疗保险个人缴纳比例</label>
      <div class="form-row">
        <input class="form-input" type="number" id="siRateMedical" value="${state.siRateMedical * 100}" min="0" max="100" step="0.1" oninput="recalc()">
        <span class="form-unit">%</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">失业保险个人缴纳比例</label>
      <div class="form-row">
        <input class="form-input" type="number" id="siRateUnemploy" value="${state.siRateUnemploy * 100}" min="0" max="100" step="0.1" oninput="recalc()">
        <span class="form-unit">%</span>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">免征额（年）</label>
      <div class="form-row">
        <input class="form-input" type="number" id="taxExemption" value="${state.taxExemption}" min="0" step="1000" oninput="recalc()">
        <span class="form-unit">元/年</span>
      </div>
    </div>
  `;
}

// --- Render deduction options ---
function renderDeductions() {
  const container = document.getElementById('deductionsBody');
  let html = '';
  for (const [key, def] of Object.entries(DEDUCTION_OPTIONS)) {
    const currentVal = state.deductions[key] || 0;
    html += '<div class="form-group"><label class="form-label">' + def.label + '</label>';
    html += '<select class="form-select" id="deduction_' + key + '" onchange="recalc()">';
    for (const opt of def.options) {
      const sel = opt.value === currentVal ? ' selected' : '';
      html += '<option value="' + opt.value + '"' + sel + '>' + opt.label + '</option>';
    }
    html += '</select></div>';
  }
  // Personal pension is a number input
  html += '<div class="form-group"><label class="form-label">个人养老金（年）</label>';
  html += '<div class="form-row">';
  html += '<input class="form-input" type="number" id="deduction_personalPension" value="' + (state.deductions.personalPension || 0) + '" min="0" step="100" oninput="recalc()">';
  html += '<span class="form-unit">元/年</span></div></div>';
  container.innerHTML = html;
}

// --- Render monthly table ---
function renderMonthlyTable() {
  const tbody = document.getElementById('monthlyTableBody');
  let html = '';
  for (const row of state.monthlyRows) {
    html += '<tr>';
    html += '<td>' + row.month + '</td>';
    html += '<td>' + fmtMoney(row.salary) + '</td>';
    html += '<td>' + fmtMoney(row.cumulativeTaxable) + '</td>';
    html += '<td>' + fmtMoney(row.cumulativeTax) + '</td>';
    html += '<td>' + fmtMoney(row.monthTax) + '</td>';
    html += '<td>' + fmtMoney(row.monthTakeHome) + '</td>';
    html += '</tr>';
  }
  tbody.innerHTML = html;
}

// --- Format currency ---
function fmtMoney(n) {
  return '¥' + Math.round(n).toLocaleString('zh-CN');
}
```

---

### Task 5: Result Bar Update & Recalc Orchestration

**Files:**
- Modify: `index.html` — add update functions and toggle logic

- [ ] **Step 7: Add result bar update and recalc orchestration**

```javascript
// --- Update result bar ---
function updateResultBar() {
  document.getElementById('monthlyTakeHome').textContent = fmtMoney(state.monthlyTakeHome);
  document.getElementById('yearlyTakeHome').textContent = fmtMoney(state.annualTakeHome);
  document.getElementById('totalPackage').textContent = fmtMoney(state.totalPackage);
  document.getElementById('totalTax').textContent = fmtMoney(state.taxSalary + state.taxBonus);
}

// --- Section toggle ---
function toggleSection(id) {
  document.getElementById(id).classList.toggle('section--open');
}

// --- City change handler ---
function onCityChange() {
  state.city = document.getElementById('citySelect').value;
  renderAdvancedInsurance();
  recalc();
}

// --- Main recalc ---
function recalc() {
  readInputs();
  compute();
  updateResultBar();
  renderMonthlyTable();
}

// --- Init ---
function init() {
  renderAdvancedInsurance();
  renderDeductions();
  recalc();
}

document.addEventListener('DOMContentLoaded', init);
```

---

### Task 6: Visual Polish & Mobile Refinements

**Files:**
- Modify: `index.html` — add extra CSS for states, transitions, and print-like refinements

- [ ] **Step 8: Add refined CSS for interaction and mobile**

Add inside the existing `<style>` block, after the base styles:

```css
/* --- Result bar details --- */
.result-bar__details {
  display: flex;
  justify-content: space-between;
  font-size: var(--text-xs);
  color: var(--color-text-secondary);
}

/* --- Input number spinners --- */
.form-input[type="number"] {
  -moz-appearance: textfield;
}
.form-input[type="number"]::-webkit-inner-spin-button,
.form-input[type="number"]::-webkit-outer-spin-button {
  -webkit-appearance: none;
}

/* --- Section transition --- */
.section__body {
  animation: fadeSlideIn var(--duration-normal) var(--ease-out);
}

@keyframes fadeSlideIn {
  from { opacity: 0; transform: translateY(-8px); }
  to { opacity: 1; transform: translateY(0); }
}

/* --- Highlight on change --- */
.result-bar__main {
  transition: color var(--duration-fast);
}
.result-bar__main--changed {
  color: var(--color-positive);
}

/* --- Select dropdown styling --- */
.form-select {
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath d='M2 4l4 4 4-4' stroke='%239aa0a6' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: right 12px center;
  padding-right: 32px;
}

/* --- Monthly table sticky header --- */
.monthly-table thead th {
  position: sticky;
  top: 0;
  background: var(--color-surface);
  z-index: 2;
  box-shadow: 0 1px 0 var(--color-border);
}

/* --- Safe area for notched phones --- */
.main { padding-bottom: calc(16px + env(safe-area-inset-bottom, 0px)); }

/* --- Tap target sizing --- */
.section__header { min-height: 48px; }
.form-input, .form-select { min-height: 44px; }

/* --- Active/focus states --- */
.form-input:active, .form-select:active {
  border-color: var(--color-accent);
}

/* --- Currency result emphasis --- */
.result-bar__main { font-variant-numeric: tabular-nums; }
```

- [ ] **Step 9: Add flash effect on recalc**

```javascript
// Add to recalc() function, after updateResultBar():
function recalc() {
  readInputs();
  compute();
  updateResultBar();
  renderMonthlyTable();

  // Flash the main number
  const el = document.getElementById('monthlyTakeHome');
  el.classList.add('result-bar__main--changed');
  setTimeout(() => el.classList.remove('result-bar__main--changed'), 200);
}
```

---

### Task 7: Verification Against Excel

**Files:**
- Test against: `SalaryCalculator.xlsx`

- [ ] **Step 10: Verify with default values**

Open `index.html` in browser. With defaults (monthly=20000, allowance=14400, bonusMonths=4, city=杭州, housingFundRate=12%, housingRent=1500), verify:

| Metric | Expected (from Excel) |
|--------|----------------------|
| 年总包 | ~254,400 |
| 年到手（含公积金） | ~218,000 (approximate, verify exactly) |

Read the Excel data-only values to confirm exact numbers.

- [ ] **Step 11: Verify edge cases**

Test these scenarios, spot-check against Excel:
1. Monthly salary = 5000 (below insurance lower bound) — insurance base should be 4986
2. Monthly salary = 50000 (above insurance upper bound) — insurance base should be 25300
3. Bonus months = 0 — no bonus tax
4. All deductions enabled at max — tax should reduce
5. Zero salary — all outputs should be 0

---

### Task 8: Add Summary Breakdown Below Result Bar

**Files:**
- Modify: `index.html`

- [ ] **Step 12: Add expandable detail card in result bar**

After the result bar's main row, add a collapsible detail section:

```html
<!-- Inside result-bar, after the last row -->
<div class="result-bar__details" id="resultDetails" style="display:none;">
</div>
```

And a tap handler on the result bar to toggle detail visibility.

```javascript
// Add to updateResultBar():
function updateResultBar() {
  document.getElementById('monthlyTakeHome').textContent = fmtMoney(state.monthlyTakeHome);
  document.getElementById('yearlyTakeHome').textContent = fmtMoney(state.annualTakeHome);
  document.getElementById('totalPackage').textContent = fmtMoney(state.totalPackage);
  document.getElementById('totalTax').textContent = fmtMoney(state.taxSalary + state.taxBonus);

  // Detail breakdown
  const details = document.getElementById('resultDetails');
  details.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;padding-top:12px;border-top:1px solid var(--color-border);margin-top:8px;">
      <div><span style="color:var(--color-text-secondary)">年总包</span><br>${fmtMoney(state.totalPackage)}</div>
      <div><span style="color:var(--color-text-secondary)">月薪</span><br>${fmtMoney(state.monthlySalary)}</div>
      <div><span style="color:var(--color-text-secondary)">工资个税</span><br>${fmtMoney(state.taxSalary)}</div>
      <div><span style="color:var(--color-text-secondary)">年终个税</span><br>${fmtMoney(state.taxBonus)}</div>
      <div><span style="color:var(--color-text-secondary)">年五险</span><br>${fmtMoney(state.annualSI)}</div>
      <div><span style="color:var(--color-text-secondary)">年公积金</span><br>${fmtMoney(state.annualHF)}</div>
      <div><span style="color:var(--color-text-secondary)">年专项扣除</span><br>${fmtMoney(state.totalDeductions * 12)}</div>
      <div><span style="color:var(--color-text-secondary)">年到手</span><br>${fmtMoney(state.annualTakeHome)}</div>
    </div>
  `;
}

// Toggle detail on result bar tap
document.getElementById('resultBar').addEventListener('click', function(e) {
  if (e.target.closest('.section__header')) return;
  const details = document.getElementById('resultDetails');
  details.style.display = details.style.display === 'none' ? 'block' : 'none';
});
```
