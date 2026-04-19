// ============================================================
// EXPENSE TRACKER - Main Entry Point
// ============================================================

const SHEET_NAMES = {
  DASHBOARD: 'Dashboard',
  EXPENSES: 'Expenses',
  CATEGORIES: 'Categories',
  BUDGET: 'Budget',
  ASSETS: 'Assets',
  SETTINGS: 'Settings'
};

// ── On Open: add menu ──────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Expense Tracker')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Send Test Reminder Email', 'sendMonthlyReminderEmail')
    .addSeparator()
    .addItem('Setup Sheets (First Run)', 'setupSheets')
    .addItem('Schedule Monthly Reminder', 'scheduleMonthlyReminder')
    .addToUi();
}

// ── Show Sidebar ───────────────────────────────────────────
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('💰 Expense Tracker')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Initial Setup ──────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheet(ss, SHEET_NAMES.EXPENSES,   setupExpensesSheet);
  _ensureSheet(ss, SHEET_NAMES.CATEGORIES, setupCategoriesSheet);
  _ensureSheet(ss, SHEET_NAMES.BUDGET,     setupBudgetSheet);
  _ensureSheet(ss, SHEET_NAMES.ASSETS,     setupAssetsSheet);
  _ensureSheet(ss, SHEET_NAMES.SETTINGS,   setupSettingsSheet);
  _ensureSheet(ss, SHEET_NAMES.DASHBOARD,  setupDashboardSheet);

  refreshDashboard();
  SpreadsheetApp.getUi().alert('✅ Setup complete! Open the sidebar to start tracking.');
}

function _ensureSheet(ss, name, setupFn) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    setupFn(sheet);
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────
// SHEET SETUP FUNCTIONS
// ─────────────────────────────────────────────────────────────

function setupExpensesSheet(sheet) {
  sheet.clearContents();
  const headers = ['Date', 'Month', 'Year', 'Category', 'Description', 'Amount (₦)', 'Payment Method', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 180);
  sheet.setFrozenRows(1);
}

function setupCategoriesSheet(sheet) {
  sheet.clearContents();
  const headers = ['Category Name', 'Icon', 'Color', 'Type', 'Created Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);

  const defaults = [
    ['Housing', '🏠', '#4CAF50', 'Fixed', new Date()],
    ['Food & Dining', '🍽️', '#FF9800', 'Variable', new Date()],
    ['Transport', '🚗', '#2196F3', 'Variable', new Date()],
    ['Utilities', '💡', '#9C27B0', 'Fixed', new Date()],
    ['Healthcare', '🏥', '#F44336', 'Variable', new Date()],
    ['Entertainment', '🎭', '#E91E63', 'Discretionary', new Date()],
    ['Shopping', '🛍️', '#FF5722', 'Discretionary', new Date()],
    ['Savings', '💰', '#4CAF50', 'Fixed', new Date()],
    ['Education', '📚', '#00BCD4', 'Fixed', new Date()],
    ['Personal Care', '💆', '#8BC34A', 'Variable', new Date()],
    ['Subscriptions', '📱', '#607D8B', 'Fixed', new Date()],
    ['Miscellaneous', '📦', '#9E9E9E', 'Variable', new Date()]
  ];
  sheet.getRange(2, 1, defaults.length, 5).setValues(defaults);
  sheet.setFrozenRows(1);
}

function setupBudgetSheet(sheet) {
  sheet.clearContents();
  const headers = ['Month', 'Year', 'Category', 'Budgeted Amount (₦)', 'Created Date', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);
}

function setupAssetsSheet(sheet) {
  sheet.clearContents();
  const headers = ['Date Recorded', 'Month', 'Year', 'Asset Name', 'Asset Class', 'Value (₦)', 'Currency', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);

  const defaults = [
    [new Date(), _monthName(new Date()), new Date().getFullYear(), 'Cash / Bank Account', 'Cash', 0, 'NGN', ''],
    [new Date(), _monthName(new Date()), new Date().getFullYear(), 'Cryptocurrency', 'Digital Assets', 0, 'NGN', ''],
    [new Date(), _monthName(new Date()), new Date().getFullYear(), 'Stocks / ETFs', 'Equities', 0, 'NGN', ''],
    [new Date(), _monthName(new Date()), new Date().getFullYear(), 'Real Estate', 'Property', 0, 'NGN', ''],
    [new Date(), _monthName(new Date()), new Date().getFullYear(), 'Savings Account', 'Cash', 0, 'NGN', ''],
    [new Date(), _monthName(new Date()), new Date().getFullYear(), 'Other Assets', 'Other', 0, 'NGN', '']
  ];
  sheet.getRange(2, 1, defaults.length, 8).setValues(defaults);
  sheet.setFrozenRows(1);
}

function setupSettingsSheet(sheet) {
  sheet.clearContents();
  sheet.getRange(1,1).setValue('Setting').setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
  sheet.getRange(1,2).setValue('Value').setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');

  const settings = [
    ['Email', Session.getActiveUser().getEmail()],
    ['Currency', '₦'],
    ['Reminder Day', '25'],
    ['Name', 'User']
  ];
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
}

function setupDashboardSheet(sheet) {
  sheet.clearContents();
  sheet.getRange('A1').setValue('📊 Dashboard — Auto-generated. Do not edit manually.')
    .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold').setFontSize(13);
  sheet.getRange(1,1,1,12).merge().setBackground('#1a1a2e');
}

// ─────────────────────────────────────────────────────────────
// SIDEBAR API — called via google.script.run
// ─────────────────────────────────────────────────────────────

// ── Categories ────────────────────────────────────────────
function getCategories() {
  const sheet = _getSheet(SHEET_NAMES.CATEGORIES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({
    name: String(r[0]),
    icon: String(r[1] || '📦'),
    color: String(r[2] || '#9E9E9E'),
    type: String(r[3] || 'Variable')
  }));
}

function addCategory(name, icon, color, type) {
  if (!name) return { success: false, message: 'Category name required.' };
  const sheet = _getSheet(SHEET_NAMES.CATEGORIES);
  const existing = getCategories().map(c => c.name.toLowerCase());
  if (existing.includes(name.toLowerCase())) return { success: false, message: 'Category already exists.' };
  sheet.appendRow([name, icon || '📦', color || '#9E9E9E', type || 'Variable', new Date()]);
  return { success: true, message: `Category "${name}" added.` };
}

function deleteCategory(name) {
  const sheet = _getSheet(SHEET_NAMES.CATEGORIES);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase() === name.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: `Category "${name}" deleted.` };
    }
  }
  return { success: false, message: 'Category not found.' };
}

function updateCategory(oldName, newName, icon, color, type) {
  const sheet = _getSheet(SHEET_NAMES.CATEGORIES);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === oldName.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[newName, icon, color, type, data[i][4]]]);
      return { success: true, message: `Category updated.` };
    }
  }
  return { success: false, message: 'Category not found.' };
}

// ── Budget ────────────────────────────────────────────────
function getBudgetsForMonth(month, year) {
  const sheet = _getSheet(SHEET_NAMES.BUDGET);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0] === month && Number(r[1]) === Number(year))
    .map(r => ({ month: r[0], year: r[1], category: r[2], amount: r[3], notes: r[5] || '' }));
}

function setBudget(month, year, category, amount, notes) {
  const sheet = _getSheet(SHEET_NAMES.BUDGET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === month && Number(data[i][1]) === Number(year) && data[i][2] === category) {
      sheet.getRange(i + 1, 4, 1, 3).setValues([[Number(amount), new Date(), notes || '']]);
      return { success: true, message: `Budget updated for ${category}.` };
    }
  }
  sheet.appendRow([month, year, category, Number(amount), new Date(), notes || '']);
  return { success: true, message: `Budget set for ${category}.` };
}

function deleteBudget(month, year, category) {
  const sheet = _getSheet(SHEET_NAMES.BUDGET);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === month && Number(data[i][1]) === Number(year) && data[i][2] === category) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Budget entry deleted.' };
    }
  }
  return { success: false, message: 'Budget entry not found.' };
}

// ── Expenses ──────────────────────────────────────────────
function addExpense(date, category, description, amount, paymentMethod, notes) {
  const sheet = _getSheet(SHEET_NAMES.EXPENSES);
  const d = new Date(date);
  sheet.appendRow([d, _monthName(d), d.getFullYear(), category, description, Number(amount), paymentMethod || '', notes || '']);
  return { success: true, message: 'Expense logged.' };
}

function getExpensesForMonth(month, year) {
  const sheet = _getSheet(SHEET_NAMES.EXPENSES);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[1] === month && Number(r[2]) === Number(year))
    .map((r, i) => ({ row: i + 2, date: r[0], month: r[1], year: r[2], category: r[3], description: r[4], amount: r[5], paymentMethod: r[6], notes: r[7] }));
}

function deleteExpense(rowIndex) {
  const sheet = _getSheet(SHEET_NAMES.EXPENSES);
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Expense deleted.' };
}

function updateExpense(rowIndex, date, category, description, amount, paymentMethod, notes) {
  const sheet = _getSheet(SHEET_NAMES.EXPENSES);
  const d = new Date(date);
  sheet.getRange(rowIndex, 1, 1, 8).setValues([[d, _monthName(d), d.getFullYear(), category, description, Number(amount), paymentMethod || '', notes || '']]);
  return { success: true, message: 'Expense updated.' };
}

// ── Assets ────────────────────────────────────────────────
function getAssetsForMonth(month, year) {
  const sheet = _getSheet(SHEET_NAMES.ASSETS);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[1] === month && Number(r[2]) === Number(year))
    .map((r, i) => ({ row: i + 2, assetName: r[3], assetClass: r[4], value: r[5], currency: r[6], notes: r[7] }));
}

function getAllAssetNames() {
  const sheet = _getSheet(SHEET_NAMES.ASSETS);
  const data = sheet.getDataRange().getValues();
  const names = [...new Set(data.slice(1).map(r => r[3]).filter(Boolean))];
  return names;
}

function upsertAsset(month, year, assetName, assetClass, value, currency, notes) {
  const sheet = _getSheet(SHEET_NAMES.ASSETS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === month && Number(data[i][2]) === Number(year) && data[i][3] === assetName) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([[new Date(), month, year, assetName, assetClass, Number(value), currency || 'NGN', notes || '']]);
      return { success: true, message: `Asset "${assetName}" updated.` };
    }
  }
  sheet.appendRow([new Date(), month, year, assetName, assetClass, Number(value), currency || 'NGN', notes || '']);
  return { success: true, message: `Asset "${assetName}" added.` };
}

function deleteAsset(rowIndex) {
  const sheet = _getSheet(SHEET_NAMES.ASSETS);
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Asset deleted.' };
}

// ── Settings ──────────────────────────────────────────────
function getSettings() {
  const sheet = _getSheet(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const settings = {};
  data.slice(1).forEach(r => { settings[r[0]] = r[1]; });
  return settings;
}

function saveSettings(email, currency, reminderDay, name) {
  const sheet = _getSheet(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const map = { 'Email': email, 'Currency': currency, 'Reminder Day': reminderDay, 'Name': name };
  for (let i = 1; i < data.length; i++) {
    if (map[data[i][0]] !== undefined) {
      sheet.getRange(i + 1, 2).setValue(map[data[i][0]]);
    }
  }
  return { success: true, message: 'Settings saved.' };
}

// ── Dashboard Data ─────────────────────────────────────────
function getDashboardData(month, year) {
  const expenses = getExpensesForMonth(month, year);
  const budgets  = getBudgetsForMonth(month, year);
  const assets   = getAssetsForMonth(month, year);
  const cats     = getCategories();

  // Spending by category
  const byCat = {};
  expenses.forEach(e => {
    byCat[e.category] = (byCat[e.category] || 0) + Number(e.amount);
  });

  // Budget vs actual
  const budgetMap = {};
  budgets.forEach(b => { budgetMap[b.category] = Number(b.amount); });

  // Totals
  const totalSpent    = expenses.reduce((s, e) => s + Number(e.amount), 0);
  const totalBudgeted = budgets.reduce((s, b) => s + Number(b.amount), 0);
  const totalAssets   = assets.reduce((s, a) => s + Number(a.value), 0);

  // Quarterly (find current quarter months)
  const qStart = Math.floor((new Date(`${month} 1, ${year}`).getMonth()) / 3) * 3;
  const qMonths = [0,1,2].map(i => _monthName(new Date(year, qStart + i, 1)));
  const qExpenses = _getSheet(SHEET_NAMES.EXPENSES).getDataRange().getValues().slice(1)
    .filter(r => qMonths.includes(r[1]) && Number(r[2]) === Number(year));
  const quarterlySpend = qExpenses.reduce((s, r) => s + Number(r[5]), 0);

  // Category breakdown with budget
  const breakdown = cats.map(c => ({
    name: c.name,
    icon: c.icon,
    color: c.color,
    type: c.type,
    spent: byCat[c.name] || 0,
    budgeted: budgetMap[c.name] || 0
  })).filter(c => c.spent > 0 || c.budgeted > 0);

  // Monthly trend (last 6 months)
  const trend = _getLast6MonthsTrend(month, year);

  // Savings rate
  const savingsRate = totalBudgeted > 0 ? Math.max(0, ((totalBudgeted - totalSpent) / totalBudgeted) * 100).toFixed(1) : 0;

  // Top spends
  const topSpends = Object.entries(byCat).sort((a,b) => b[1]-a[1]).slice(0,5).map(([cat, amt]) => ({ category: cat, amount: amt }));

  return { totalSpent, totalBudgeted, totalAssets, quarterlySpend, breakdown, trend, savingsRate, topSpends, expenses, assets };
}

function _getLast6MonthsTrend(currentMonth, year) {
  const sheet = _getSheet(SHEET_NAMES.EXPENSES);
  const data  = sheet.getDataRange().getValues().slice(1);
  const months = [];
  let d = new Date(`${currentMonth} 1, ${year}`);
  for (let i = 5; i >= 0; i--) {
    const m = new Date(d.getFullYear(), d.getMonth() - i, 1);
    months.push({ month: _monthName(m), year: m.getFullYear(), label: `${_monthName(m).slice(0,3)} ${m.getFullYear()}` });
  }
  return months.map(m => {
    const total = data.filter(r => r[1] === m.month && Number(r[2]) === m.year).reduce((s, r) => s + Number(r[5]), 0);
    return { label: m.label, total };
  });
}

// ── Email Reminder ─────────────────────────────────────────
function sendMonthlyReminderEmail() {
  const settings = getSettings();
  const email = settings['Email'] || Session.getActiveUser().getEmail();
  const name  = settings['Name'] || 'there';
  const nextMonth = _monthName(new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1));
  const year  = new Date().getMonth() === 11 ? new Date().getFullYear() + 1 : new Date().getFullYear();
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const subject = `Budget Reminder: Set your ${nextMonth} ${year} budget`;
  const body = `
    <div style="font-family: 'Helvetica Neue', sans-serif; max-width: 600px; margin: 0 auto; background: #f8f9fa; padding: 20px;">
      <div style="background: #1a1a2e; color: white; padding: 30px; border-radius: 12px 12px 0 0; text-align: center;">
        <h1 style="margin:0; font-size: 28px;">&#9733; Monthly Budget Reminder</h1>
        <p style="margin: 8px 0 0; opacity: 0.8;">${nextMonth} ${year} is approaching</p>
      </div>
      <div style="background: white; padding: 30px; border-radius: 0 0 12px 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <p style="color: #333; font-size: 16px;">Hi <strong>${name}</strong>,</p>
        <p style="color: #555; line-height: 1.6;">It's time to set your budget for <strong>${nextMonth} ${year}</strong>. Taking a few minutes now to plan your finances will help you stay on track and reach your savings goals.</p>
        <div style="background: #f0f4ff; border-left: 4px solid #1a1a2e; padding: 16px; border-radius: 4px; margin: 20px 0;">
          <p style="margin:0; color: #333; font-weight: bold;">Quick Checklist</p>
          <ul style="color: #555; margin: 8px 0 0; padding-left: 20px;">
            <li>Review last month's spending</li>
            <li>Adjust budgets for changing expenses</li>
            <li>Update your asset values</li>
            <li>Set savings targets</li>
          </ul>
        </div>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${ssUrl}" style="background: #1a1a2e; color: white; padding: 14px 32px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; display: inline-block;">Open Expense Tracker →</a>
        </div>
        <p style="color: #888; font-size: 13px; text-align: center;">This reminder was sent automatically from your Expense Tracker.</p>
      </div>
    </div>
  `;
  GmailApp.sendEmail(email, subject, '', { htmlBody: body, name: 'Expense Tracker' });
  return { success: true, message: `Reminder sent to ${email}` };
}

function scheduleMonthlyReminder() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyReminderEmail') ScriptApp.deleteTrigger(t);
  });
  const settings = getSettings();
  const day = parseInt(settings['Reminder Day'] || '25');
  ScriptApp.newTrigger('sendMonthlyReminderEmail')
    .timeBased().onMonthDay(day).atHour(9).create();
  return { success: true, message: `Monthly reminder scheduled for day ${day} of each month.` };
}

// ── Refresh Dashboard Sheet ────────────────────────────────
function refreshDashboard() {
  const now = new Date();
  const month = _monthName(now);
  const year  = now.getFullYear();
  const data  = getDashboardData(month, year);
  const sheet = _getSheet(SHEET_NAMES.DASHBOARD);
  sheet.clearContents();
  sheet.clearFormats();

  const dark = '#1a1a2e', accent = '#e94560', green = '#4CAF50', amber = '#FF9800';
  const white = '#ffffff', light = '#f8f9ff';

  // Title row
  sheet.getRange(1, 1, 1, 10).merge()
    .setValue(`📊 EXPENSE TRACKER DASHBOARD — ${month} ${year}`)
    .setBackground(dark).setFontColor(white).setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // KPI Cards row 3
  const kpiRow = 3;
  const kpis = [
    ['💸 Total Spent', _fmt(data.totalSpent), data.totalSpent > data.totalBudgeted ? '#FFEBEE' : '#E8F5E9'],
    ['📋 Total Budgeted', _fmt(data.totalBudgeted), '#E3F2FD'],
    ['💼 Total Assets', _fmt(data.totalAssets), '#F3E5F5'],
    ['📈 Quarterly Spend', _fmt(data.quarterlySpend), '#FFF8E1'],
    ['💚 Savings Rate', data.savingsRate + '%', Number(data.savingsRate) >= 20 ? '#E8F5E9' : '#FFF3E0']
  ];

  kpis.forEach((kpi, i) => {
    const col = i * 2 + 1;
    sheet.getRange(kpiRow, col, 1, 2).merge().setValue(kpi[0])
      .setBackground(dark).setFontColor(white).setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center');
    sheet.getRange(kpiRow + 1, col, 1, 2).merge().setValue(kpi[1])
      .setBackground(kpi[2]).setFontColor(dark).setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(kpiRow + 1, 45);
  });
  sheet.setRowHeight(kpiRow, 30);

  // Budget vs Actual header
  const budRow = 6;
  sheet.getRange(budRow, 1, 1, 6).merge()
    .setValue('💰 BUDGET vs ACTUAL')
    .setBackground(dark).setFontColor(white).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.getRange(budRow + 1, 1, 1, 6).setValues([['Category', 'Type', 'Budgeted (₦)', 'Spent (₦)', 'Remaining (₦)', 'Status']])
    .setBackground('#2d2d44').setFontColor(white).setFontWeight('bold');

  let r = budRow + 2;
  data.breakdown.forEach(c => {
    const rem    = c.budgeted - c.spent;
    const status = c.budgeted === 0 ? '⚪ No Budget' : rem >= 0 ? '✅ On Track' : '🚨 Over Budget';
    const bg     = rem < 0 ? '#FFEBEE' : (r % 2 === 0 ? light : white);
    sheet.getRange(r, 1, 1, 6).setValues([[c.icon + ' ' + c.name, c.type, c.budgeted, c.spent, rem, status]])
      .setBackground(bg);
    sheet.getRange(r, 3, 1, 3).setNumberFormat('#,##0');
    r++;
  });

  // Monthly Trend
  const trendRow = r + 2;
  sheet.getRange(trendRow, 1, 1, 6).merge()
    .setValue('📈 6-MONTH SPENDING TREND')
    .setBackground(dark).setFontColor(white).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.getRange(trendRow + 1, 1, 1, 6)
    .setValues([data.trend.map(t => t.label)])
    .setBackground('#2d2d44').setFontColor(white).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(trendRow + 2, 1, 1, 6)
    .setValues([data.trend.map(t => t.total)])
    .setBackground(light).setNumberFormat('#,##0').setHorizontalAlignment('center');

  // Assets
  const assRow = trendRow + 5;
  sheet.getRange(assRow, 1, 1, 4).merge()
    .setValue('💼 ASSET PORTFOLIO — ' + month + ' ' + year)
    .setBackground(dark).setFontColor(white).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.getRange(assRow + 1, 1, 1, 4).setValues([['Asset', 'Class', 'Value (₦)', '% of Portfolio']])
    .setBackground('#2d2d44').setFontColor(white).setFontWeight('bold');

  let ar = assRow + 2;
  data.assets.forEach((a, i) => {
    const pct = data.totalAssets > 0 ? ((a.value / data.totalAssets) * 100).toFixed(1) + '%' : '0%';
    sheet.getRange(ar, 1, 1, 4).setValues([[a.assetName, a.assetClass, a.value, pct]])
      .setBackground(i % 2 === 0 ? light : white);
    sheet.getRange(ar, 3).setNumberFormat('#,##0');
    ar++;
  });
  sheet.getRange(ar, 1).setValue('TOTAL').setFontWeight('bold');
  sheet.getRange(ar, 3).setValue(data.totalAssets).setNumberFormat('#,##0').setFontWeight('bold').setBackground('#E8F5E9');

  // Top spending categories
  const topRow = assRow;
  sheet.getRange(topRow, 6, 1, 4).merge()
    .setValue('🔥 TOP SPENDING AREAS')
    .setBackground(dark).setFontColor(white).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.getRange(topRow + 1, 6, 1, 4).setValues([['#', 'Category', 'Amount (₦)', '% of Spend']])
    .setBackground('#2d2d44').setFontColor(white).setFontWeight('bold');

  let tr = topRow + 2;
  data.topSpends.forEach((s, i) => {
    const pct = data.totalSpent > 0 ? ((s.amount / data.totalSpent) * 100).toFixed(1) + '%' : '0%';
    const tip = s.amount > (data.totalBudgeted * 0.25) ? '⚠️ High' : '';
    sheet.getRange(tr, 6, 1, 4).setValues([[i+1, s.category + (tip ? ' ' + tip : ''), s.amount, pct]])
      .setBackground(i % 2 === 0 ? light : white);
    sheet.getRange(tr, 8).setNumberFormat('#,##0');
    tr++;
  });

  // Last refreshed
  sheet.getRange(ar + 2, 1, 1, 6).merge()
    .setValue(`Last refreshed: ${new Date().toLocaleString()}`)
    .setFontColor('#999').setFontSize(9).setHorizontalAlignment('right');

  // Auto-resize columns
  sheet.autoResizeColumns(1, 10);
  return { success: true };
}

// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────
function _getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found. Please run Setup first.`);
  return sheet;
}

function _monthName(d) {
  return ['January','February','March','April','May','June',
          'July','August','September','October','November','December'][d.getMonth()];
}

function _fmt(n) {
  return '₦' + Number(n).toLocaleString('en-NG', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}