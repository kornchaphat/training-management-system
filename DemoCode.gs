/**
 * Training Management System (TMS) - Demo Backend
 * Full-stack Google Apps Script web application.
 *
 * DEMO: Auth and email notifications removed.
 * Production version uses OTP email auth + role-based access control.
 */

// ============================================
// CONFIGURATION
// ============================================
// SPREADSHEET_ID is now stored in Script Properties
// Go to Project Settings > Script Properties > Add: SPREADSHEET_ID = your-sheet-id

const CONFIG = {
  SESSION_TIMEOUT_MINUTES: 120,
  OTP_EXPIRY_MINUTES: 10,
  CACHE_PREFIX: 'TMS_',
  TIMEZONE: 'Asia/Bangkok'
};

// Email Configuration
// To use a group email as sender, you MUST add it as a "Send mail as" alias in Gmail:
// 1. Go to Gmail Settings > Accounts > Send mail as > Add another email address
// 2. Add your group email and verify it
// 3. Then set it below
const EMAIL_CONFIG = {
  SENDER_EMAIL: 'demo@acmecorp.com',
  SENDER_NAME: 'Acme Corp TMS'
};

// Company Logo - removed for demo
const COMPANY_LOGO = '';

// Forms Storage Folder ID - Retrieved from Script Properties

// forceReauthorization() removed for demo
// Set in: Project Settings > Script Properties > Add property
// Property: FORMS_FOLDER_ID
// Value: Your Google Drive folder ID (e.g., 1CHQe5BwcE86iIU3yX45aVfrmdSzAyZAe)
function getFormsFolderId_() {
  const folderId = PropertiesService.getScriptProperties().getProperty('FORMS_FOLDER_ID');
  return folderId || null;  // Returns null if not set (forms will stay in root Drive)
}

const SHEET_NAMES = {
  USERS: 'Users',
  PROGRAMS: 'Programs',
  SESSIONS: 'Sessions',
  ENROLLMENTS: 'Enrollments',
  EMPLOYEES: 'Employees',
  OTP_SESSIONS: 'OTP_Sessions',
  ACTIVE_SESSIONS: 'Active_Sessions'
};

// ============================================
// UTILITIES
// ============================================
function getSpreadsheet_() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('SPREADSHEET_ID not set in Script Properties. Go to Project Settings > Script Properties > Add property.');
    }
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    throw new Error('Cannot open spreadsheet. Check SPREADSHEET_ID in Script Properties. Error: ' + e.message);
  }
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Sheet "${name}" not found. Please run setupDatabase() first.`);
  }
  return sheet;
}

function getSheetSafe_(name) {
  // Returns null instead of throwing if sheet doesn't exist
  try {
    const ss = getSpreadsheet_();
    return ss.getSheetByName(name);
  } catch (e) {
    return null;
  }
}

function parseList_(value) {
  if (!value) return [];
  if (value === 'All') return ['All'];
  return value.toString().split(',').map(s => s.trim()).filter(s => s);
}

function formatDateForSheet_(date) {
  if (!date) return '';
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

// ============================================
// WEB APP ENTRY POINTS
// ============================================
function doGet(e) {
  // DEMO: Bypass authentication, serve app directly
  const demoUser = {
    token: 'demo-session',
    email: 'demo@acmecorp.com',
    name: 'Demo Admin',
    role: 'Developer',
    countries: ['All'],
    primaryEntity: 'Global'
  };
  return serveApp_(demoUser);
}

// serveLogin_() removed for demo

function serveApp_(sessionData) {
  const template = HtmlService.createTemplateFromFile('TmsApp');
  template.userEmail = sessionData.email;
  template.userName = sessionData.name;
  template.userRole = sessionData.role;
  template.userCountries = JSON.stringify(sessionData.countries || []);
  template.primaryEntity = sessionData.primaryEntity || '';
  template.sessionToken = sessionData.token;
  
  return template.evaluate()
    .setTitle('Training Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ============================================
// AUTHENTICATION - OTP
// ============================================
function sendOTP(email) {
  return { success: false, message: 'Disabled in demo.' };
}

function verifyOTP(email, otp) {
  return { success: false, message: 'Disabled in demo.' };
}

function logout(sessionToken) {
  return { success: true };
}

// ============================================
// SESSION MANAGEMENT
// ============================================
function createSession_(user) { return null; }

function validateSession_(token) {
  // DEMO: Return demo user for any token
  return { token: 'demo-session', email: 'demo@acmecorp.com', name: 'Demo Admin', role: 'Developer', countries: ['All'], primaryEntity: 'Global' };
}

function extendSession_(token) { }

function invalidateSession_(token) { }

// ============================================
// OTP STORAGE
// ============================================
function storeOTP_(email, otp, expiry) { }

function getStoredOTP_(email) { return null; }

function clearOTP_(email) { }

// ============================================
// USER MANAGEMENT
// ============================================
function getUserByEmail_(email) { return null; }

// ============================================
// ACCESS CONTROL
// ============================================
function withSession_(sessionToken, callback) {
  // DEMO: Bypass session validation
  const demoUser = {
    token: 'demo-session',
    email: 'demo@acmecorp.com',
    name: 'Demo Admin',
    role: 'Developer',
    countries: ['All'],
    primaryEntity: 'Global'
  };
  return callback(demoUser);
}

function hasFullAccess_(role) {
  // Only Developer and Global Talent CoE have unrestricted access
  // Regional Talent CoE is now scoped by their Countries field
  return role === 'Developer' || role === 'Global Talent CoE';
}

function canManagePrograms_(role) {
  // Allow all user roles to create/edit programs
  return true;
}

/**
 * Check if user can VIEW all sessions (used for filtering)
 * Global HRBP can see all but edit restrictions apply separately
 */
function canViewAllSessions_(userData) {
  const role = userData.role || '';
  
  // Full access roles see everything
  if (hasFullAccess_(role)) return true;
  
  // Global HRBP can VIEW all sessions (but edit is restricted)
  if (role === 'Global HRBP') return true;
  
  // Check if user has "All" in countries
  const userCountries = userData.countries || [];
  return userCountries.some(c => c.toLowerCase() === 'all');
}

function canEditSession_(userData, sessionEntity, sessionCreatedBy) {
  const role = userData.role || '';
  
  // Global HRBP: Can only edit sessions they created (regardless of entity)
  if (role === 'Global HRBP') {
    return sessionCreatedBy && sessionCreatedBy.toLowerCase() === userData.email.toLowerCase();
  }
  
  // Full access roles can edit anything
  if (hasFullAccess_(role)) {
    return true;
  }
  
  // Creator can always edit their own session
  if (sessionCreatedBy && sessionCreatedBy.toLowerCase() === userData.email.toLowerCase()) {
    return true;
  }
  
  // Entity-based access check for Regional Talent CoE, Country HRBP, Manager
  const userCountries = userData.countries || [];
  const hasAllCountries = userCountries.some(c => c.toLowerCase() === 'all');
  
  // Global sessions: only users with "All" access can edit (unless they created it)
  if (sessionEntity === 'Global') {
    return hasAllCountries;
  }
  
  // Other sessions: check if entity is in user's access list
  if (hasAllCountries) return true;
  return userCountries.includes(sessionEntity);
}

/**
 * Fetch exchange rates from Finance's SYS_ExchangeRates spreadsheet
 * Returns a Map with key "YYYY-MM-CURRENCY" -> rate
 * Example: "2024-03-THB" -> 35.12
 */
function getExchangeRatesMap_() {
  try {
    const EXCHANGE_RATES_SPREADSHEET_ID = '';  // Set your exchange rates sheet ID
    const externalSS = SpreadsheetApp.openById(EXCHANGE_RATES_SPREADSHEET_ID);
    const sheet = externalSS.getSheetByName('SYS_ExchangeRates');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      console.log('Exchange rates sheet not found or empty');
      return new Map();
    }
    
    // Columns: A=Date, B=CurrencyCode, C=ExchangeRate, D=SourceType
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const ratesMap = new Map();
    
    data.forEach(row => {
      const dateVal = row[0];
      const currency = String(row[1] || '').trim().toUpperCase();
      const rate = parseFloat(row[2]) || 0;
      
      if (!dateVal || !currency || rate === 0) return;
      
      // Convert date to YYYY-MM format
      let dateKey = '';
      if (dateVal instanceof Date) {
        const yyyy = dateVal.getFullYear();
        const mm = String(dateVal.getMonth() + 1).padStart(2, '0');
        dateKey = `${yyyy}-${mm}`;
      } else {
        // Try parsing string date
        const d = new Date(dateVal);
        if (!isNaN(d.getTime())) {
          const yyyy = d.getFullYear();
          const mm = String(d.getMonth() + 1).padStart(2, '0');
          dateKey = `${yyyy}-${mm}`;
        }
      }
      
      if (dateKey) {
        const key = `${dateKey}-${currency}`;
        ratesMap.set(key, rate);
      }
    });
    
    console.log('Exchange rates loaded:', ratesMap.size, 'entries');
    return ratesMap;
  } catch (e) {
    console.error('Error fetching exchange rates:', e);
    return new Map();
  }
}

/**
 * Convert local currency to USD using exchange rate for the given month
 * @param {number} localAmount - Amount in local currency
 * @param {string} currency - Currency code (e.g., "THB", "PHP")
 * @param {string} dateStr - Date string in YYYY-MM-DD format
 * @param {Map} ratesMap - Exchange rates map from getExchangeRatesMap_()
 * @returns {number|null} - USD amount or null if no rate found
 */
function convertToUSD_(localAmount, currency, dateStr, ratesMap) {
  if (!localAmount || !currency || !dateStr || !ratesMap) return null;
  
  // Extract YYYY-MM from date string
  const monthKey = dateStr.substring(0, 7); // "2024-03-15" -> "2024-03"
  const lookupKey = `${monthKey}-${currency.toUpperCase()}`;
  
  const rate = ratesMap.get(lookupKey);
  if (!rate || rate === 0) {
    // Try to find the most recent rate for this currency as fallback
    let fallbackRate = null;
    let fallbackKey = null;
    
    for (const [key, value] of ratesMap.entries()) {
      if (key.endsWith(`-${currency.toUpperCase()}`)) {
        if (!fallbackKey || key > fallbackKey) {
          fallbackKey = key;
          fallbackRate = value;
        }
      }
    }
    
    if (fallbackRate) {
      return Math.round((localAmount / fallbackRate) * 100) / 100;
    }
    return null;
  }
  
  return Math.round((localAmount / rate) * 100) / 100;
}

/**
 * TEST: Run this in Apps Script Editor to debug Analytics
 */
function testAnalyticsDirect() {
  const ss = getSpreadsheet_();
  
  // Get sheets
  const sessSheet = ss.getSheetByName('Sessions');
  const enrollSheet = ss.getSheetByName('Enrollments');
  const tm1Sheet = ss.getSheetByName('TM1');
  const progSheet = ss.getSheetByName('Programs');
  
  console.log('=== SHEET DATA ===');
  console.log('Sessions rows:', sessSheet ? sessSheet.getLastRow() - 1 : 0);
  console.log('Enrollments rows:', enrollSheet ? enrollSheet.getLastRow() - 1 : 0);
  console.log('TM1 rows:', tm1Sheet ? tm1Sheet.getLastRow() - 1 : 0);
  console.log('Programs rows:', progSheet ? progSheet.getLastRow() - 1 : 0);
  
  // Sample Session
  if (sessSheet && sessSheet.getLastRow() > 1) {
    const row = sessSheet.getRange(2, 1, 1, 16).getValues()[0];
    console.log('Sample Session:', {
      programId: row[0],
      sessionId: row[1],
      status: row[3],
      date: row[5],
      hours: row[6],
      satisfaction: row[8],
      entity: row[13]
    });
  }

/**
 * DIAGNOSTIC: Run this function to debug getInitialData issues
 * Check the Execution Log for output
 */
function debugGetInitialData() {
  const ss = getSpreadsheet_();
  const progSheet = ss.getSheetByName('Programs');
  const sessSheet = ss.getSheetByName('Sessions');
  
  console.log('=== PROGRAMS SHEET DEBUG ===');
  if (!progSheet) {
    console.log('ERROR: Programs sheet not found!');
    return;
  }
  
  const progLastRow = progSheet.getLastRow();
  const progLastCol = progSheet.getLastColumn();
  console.log('Programs - LastRow:', progLastRow, 'LastCol:', progLastCol);
  
  if (progLastRow > 1) {
    // Get headers
    const headers = progSheet.getRange(1, 1, 1, progLastCol).getValues()[0];
    console.log('Programs Headers:', headers);
    
    // Get first data row
    const firstRow = progSheet.getRange(2, 1, 1, progLastCol).getValues()[0];
    console.log('Programs First Row:', firstRow);
    console.log('  - Column A (ID):', firstRow[0]);
    console.log('  - Column B (Name):', firstRow[1]);
    console.log('  - Column C (Category):', firstRow[2]);
    console.log('  - Column D (Description):', firstRow[3]);
    console.log('  - Column E (ModifiedBy):', firstRow[4]);
    console.log('  - Column F (ModifiedOn):', firstRow[5]);
    console.log('  - Column F type:', typeof firstRow[5], firstRow[5] instanceof Date ? 'isDate' : 'notDate');
  }
  
  console.log('\n=== SESSIONS SHEET DEBUG ===');
  if (!sessSheet) {
    console.log('ERROR: Sessions sheet not found!');
    return;
  }
  
  const sessLastRow = sessSheet.getLastRow();
  const sessLastCol = sessSheet.getLastColumn();
  console.log('Sessions - LastRow:', sessLastRow, 'LastCol:', sessLastCol);
  
  if (sessLastRow > 1) {
    // Get headers
    const headers = sessSheet.getRange(1, 1, 1, Math.min(sessLastCol, 19)).getValues()[0];
    console.log('Sessions Headers:', headers);
    
    // Get first data row
    const firstRow = sessSheet.getRange(2, 1, 1, Math.min(sessLastCol, 19)).getValues()[0];
    console.log('Sessions First Row Values:');
    console.log('  - Column L (ModifiedBy) [index 11]:', firstRow[11]);
    console.log('  - Column M (ModifiedOn) [index 12]:', firstRow[12]);
    console.log('  - Column N (Entity) [index 13]:', firstRow[13]);
    console.log('  - Column O (TrackQR) [index 14]:', firstRow[14]);
  }
  
  console.log('\n=== TEST getInitialData LOGIC ===');
  try {
    // Test program reading
    const lastCol = Math.min(progSheet.getLastColumn(), 6);
    console.log('Reading', lastCol, 'columns from Programs');
    const progData = progSheet.getRange(2, 1, progSheet.getLastRow() - 1, lastCol).getValues();
    console.log('Read', progData.length, 'program rows');
    
    const programs = progData.map((r, idx) => {
      let modifiedOnStr = '';
      if (lastCol >= 6 && r.length >= 6 && r[5]) {
        const modifiedOn = r[5];
        if (modifiedOn instanceof Date) {
          modifiedOnStr = Utilities.formatDate(modifiedOn, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm');
        } else {
          modifiedOnStr = String(modifiedOn);
        }
      }
      return { 
        id: String(r[0] || '').trim(), 
        name: r[1] || '', 
        category: r[2] || '', 
        description: r[3] || '',
        modifiedBy: (lastCol >= 5 && r.length >= 5) ? (r[4] || '') : '',
        modifiedOn: modifiedOnStr
      };
    });
    
    console.log('Successfully created', programs.length, 'program objects');
    console.log('First program:', JSON.stringify(programs[0]));
    
  } catch (e) {
    console.log('ERROR in program reading:', e.message);
    console.log('Stack:', e.stack);
  }
  
  console.log('\n=== SUCCESS ===');
}
  if (enrollSheet && enrollSheet.getLastRow() > 1) {
    const row = enrollSheet.getRange(2, 1, 1, 13).getValues()[0];
    console.log('Sample Enrollment:', {
      sessionId: row[1],
      entity: row[5],
      jobBand: row[6],
      function: row[7],
      subFunction: row[8],
      status: row[10],
      statusExact: "'" + row[10] + "'"
    });
  }
  
  // Count attended
  if (enrollSheet && enrollSheet.getLastRow() > 1) {
    const allStatuses = enrollSheet.getRange(2, 11, enrollSheet.getLastRow() - 1, 1).getValues().flat();
    const attended = allStatuses.filter(s => String(s).trim() === 'Attended').length;
    const uniqueStatuses = [...new Set(allStatuses.map(s => String(s).trim()))];
    console.log('Enrollment statuses found:', uniqueStatuses);
    console.log('Attended count:', attended, 'of', allStatuses.length);
  }
  
  // Sample TM1
  if (tm1Sheet && tm1Sheet.getLastRow() > 1) {
    const row = tm1Sheet.getRange(2, 1, 1, 6).getValues()[0];
    console.log('Sample TM1:', {
      year: row[0],
      quarter: row[1],
      month: row[2],
      entity: row[3],
      ytg: row[4],
      actual: row[5]
    });
  }
}

/**
 * TEST: Run this to test getAnalyticsData with a real session
 */
function testAnalyticsWithSession() {
  // Get a session token from Active_Sessions
  const ss = getSpreadsheet_();
  const activeSheet = ss.getSheetByName('Active_Sessions');
  
  if (!activeSheet || activeSheet.getLastRow() < 2) {
    console.log('No active sessions. Please log in first.');
    return;
  }
  
  const sessionToken = activeSheet.getRange(2, 1).getValue();
  console.log('Using session token:', sessionToken);
  
  const result = getAnalyticsData(sessionToken, {});
  
  console.log('=== ANALYTICS RESULT ===');
  console.log('Has error:', !!result.error);
  if (result.error) {
    console.log('Error:', result.error);
  }
  console.log('KPIs:', JSON.stringify(result.kpis, null, 2));
  console.log('Debug:', JSON.stringify(result.debug, null, 2));
}

/**
 * TEST: Bypass session and test analytics calculation directly
 */
function testAnalyticsNoSession() {
  console.log('Testing analytics WITHOUT session validation...');
  
  const ss = getSpreadsheet_();
  
  // Simulate user data for Developer role
  const fakeUserData = {
    email: 'test@test.com',
    role: 'Developer',
    countries: ['All'],
    functions: ['All']
  };
  
  try {
    const sessSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
    const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
    const progSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
    
    console.log('Sessions:', sessSheet ? sessSheet.getLastRow() - 1 : 0);
    console.log('Enrollments:', enrollSheet ? enrollSheet.getLastRow() - 1 : 0);
    console.log('Programs:', progSheet ? progSheet.getLastRow() - 1 : 0);
    
    // Get sessions
    const sessData = sessSheet && sessSheet.getLastRow() > 1 
      ? sessSheet.getRange(2, 1, sessSheet.getLastRow() - 1, 15).getValues() 
      : [];
    console.log('Session data rows:', sessData.length);
    
    // Get enrollments
    const enrollData = enrollSheet && enrollSheet.getLastRow() > 1
      ? enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 13).getValues()
      : [];
    console.log('Enrollment data rows:', enrollData.length);
    
    // Calculate simple KPIs
    let totalHours = 0;
    const sessionDurations = new Map();
    
    sessData.forEach(s => {
      const sessId = String(s[1]).trim();
      const duration = parseFloat(s[6]) || 0;
      sessionDurations.set(sessId, duration);
    });
    
    const attendedCount = enrollData.filter(e => String(e[10] || '').trim() === 'Attended').length;
    console.log('Attended enrollments:', attendedCount);
    
    enrollData.forEach(e => {
      const sessId = String(e[1]).trim();
      const status = String(e[10] || '').trim();
      if (status === 'Attended') {
        totalHours += sessionDurations.get(sessId) || 0;
      }
    });
    
    const uniqueLearners = new Set(enrollData.map(e => String(e[2]).trim()).filter(Boolean)).size;
    
    console.log('=== CALCULATED KPIs ===');
    console.log('Total Hours:', totalHours);
    console.log('Unique Learners:', uniqueLearners);
    console.log('Avg Hours/Employee:', uniqueLearners > 0 ? totalHours / uniqueLearners : 0);
    
    // Check satisfaction
    const satisfactionScores = sessData.map(s => parseFloat(s[8]) || 0).filter(s => s > 0);
    const avgSatisfaction = satisfactionScores.length > 0 
      ? satisfactionScores.reduce((a, b) => a + b, 0) / satisfactionScores.length 
      : 0;
    console.log('Avg Satisfaction:', avgSatisfaction);
    
    console.log('SUCCESS - Analytics calculation works!');
    
  } catch (e) {
    console.log('ERROR:', e.message);
    console.log('Stack:', e.stack);
  }
}

// Comprehensive debug function for analytics issues
function debugAnalyticsData() {
  const ss = getSpreadsheet_();
  
  console.log('========================================');
  console.log('ANALYTICS DEBUG - TM1 SHEET');
  console.log('========================================');
  
  const tm1Sheet = ss.getSheetByName('TM1');
  if (tm1Sheet) {
    const lastRow = tm1Sheet.getLastRow();
    console.log('TM1 last row:', lastRow);
    
    if (lastRow > 1) {
      const headers = tm1Sheet.getRange(1, 1, 1, 6).getValues()[0];
      console.log('TM1 headers:', JSON.stringify(headers));
      
      const data = tm1Sheet.getRange(2, 1, Math.min(5, lastRow-1), 6).getValues();
      data.forEach((row, i) => {
        console.log(`TM1 Row ${i+2}: Year=${row[0]}, Quarter=${row[1]}, Month=${row[2]}, Entity="${row[3]}", YTG=${row[4]}, Actual=${row[5]}`);
      });
      
      // Calculate totals
      const allData = tm1Sheet.getRange(2, 1, lastRow-1, 6).getValues();
      let totalActual = 0;
      const entitiesWithSpend = {};
      allData.forEach(row => {
        const entity = String(row[3] || '').trim();
        const actual = parseFloat(row[5]) || 0;
        totalActual += actual;
        if (entity) {
          entitiesWithSpend[entity] = (entitiesWithSpend[entity] || 0) + actual;
        }
      });
      console.log('Total Actual (should match KPI):', totalActual);
      console.log('Entities with spend:', JSON.stringify(entitiesWithSpend));
      console.log('Entity count:', Object.keys(entitiesWithSpend).length);
    }
  } else {
    console.log('TM1 sheet NOT FOUND!');
  }
  
  console.log('========================================');
  console.log('ANALYTICS DEBUG - SESSIONS IsIndivDev');
  console.log('========================================');
  
  const sessSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
  if (sessSheet) {
    const lastRow = sessSheet.getLastRow();
    console.log('Sessions last row:', lastRow);
    
    if (lastRow > 1) {
      const headers = sessSheet.getRange(1, 1, 1, 15).getValues()[0];
      console.log('Column K (index 10) header:', headers[10]);
      
      const data = sessSheet.getRange(2, 1, lastRow-1, 15).getValues();
      let yesCount = 0;
      let noCount = 0;
      let emptyCount = 0;
      let otherValues = [];
      
      data.forEach((row, i) => {
        const val = row[10];
        const valStr = String(val || '').trim().toLowerCase();
        if (valStr === 'yes' || valStr === 'true' || val === true) {
          yesCount++;
          if (yesCount <= 3) {
            console.log(`Session with IsIndivDev=Yes: ID=${row[1]}, Name="${row[2]}", RawValue="${val}"`);
          }
        } else if (valStr === 'no' || valStr === 'false' || val === false) {
          noCount++;
        } else if (!val) {
          emptyCount++;
        } else {
          otherValues.push(val);
        }
      });
      
      console.log(`IsIndivDev counts: Yes=${yesCount}, No=${noCount}, Empty=${emptyCount}`);
      if (otherValues.length > 0) {
        console.log('Other IsIndivDev values:', JSON.stringify([...new Set(otherValues)]));
      }
    }
  }
  
  console.log('========================================');
  console.log('DEBUG COMPLETE - Check logs above');
  console.log('========================================');
  
  return 'Debug complete - check Execution Log';
}

// ============================================
// COMBINED API FOR FAST INITIAL LOAD
// ============================================
function getInitialData(sessionToken, centerYear, centerMonth, filterEntity) {
  try {
    return withSession_(sessionToken, (userData) => {
      try {
        const programSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
        const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
        const enrollmentSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      
      // User access info
      const userCountries = userData.countries || [];
      const hasAllCountries = userCountries.some(c => c.toLowerCase() === 'all');
      const primaryEntity = userData.primaryEntity || '';
      const role = userData.role || '';
      
      // Build entity options for dropdown (what user can CREATE sessions for)
      let entityOptions = [];
      if (hasFullAccess_(role)) {
        // Full access: can create for any entity
        entityOptions = ['Global'];
        if (primaryEntity && primaryEntity !== 'Global') {
          entityOptions.push(primaryEntity);
        }
      } else if (role === 'Global HRBP') {
        // Global HRBP: can only create for their assigned entities (even though they can VIEW all)
        if (hasAllCountries) {
          entityOptions = ['Global'];
          if (primaryEntity && primaryEntity !== 'Global') {
            entityOptions.push(primaryEntity);
          }
        } else {
          entityOptions = [...userCountries];
        }
      } else {
        // Regional CoE, Country HRBP, Manager: only their assigned entities
        entityOptions = [...userCountries];
      }
      
      // Build available entities for filter dropdowns (what user can SEE)
      let availableEntities = [];
      if (canViewAllSessions_(userData)) {
        // Full access + Global HRBP: can see all entities
        if (sessionSheet && sessionSheet.getLastRow() > 1) {
          const entData = sessionSheet.getRange(2, 14, sessionSheet.getLastRow() - 1, 1).getValues();
          availableEntities = [...new Set(entData.flat().filter(e => e))].sort();
        }
      } else {
        // Regional CoE, Country HRBP, Manager: only their entities + Global
        availableEntities = ['Global', ...userCountries];
      }
      
      // Programs
      let programs = [];
      if (programSheet && programSheet.getLastRow() > 1) {
        // Read columns: ID, Name, Category, Description, LastModifiedBy, LastModifiedOn (max 6)
        const lastCol = Math.min(programSheet.getLastColumn(), 6);
        const progData = programSheet.getRange(2, 1, programSheet.getLastRow() - 1, lastCol).getValues();
        programs = progData.map(r => {
          let modifiedOnStr = '';
          // Only access r[5] if we have 6 columns and r[5] has a value
          if (lastCol >= 6 && r.length >= 6 && r[5]) {
            const modifiedOn = r[5];
            if (modifiedOn instanceof Date) {
              modifiedOnStr = Utilities.formatDate(modifiedOn, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm');
            } else {
              modifiedOnStr = String(modifiedOn);
            }
          }
          return { 
            id: String(r[0] || '').trim(), 
            name: r[1] || '', 
            category: r[2] || '', 
            description: r[3] || '',
            modifiedBy: (lastCol >= 5 && r.length >= 5) ? (r[4] || '') : '',
            modifiedOn: modifiedOnStr,
            sessionCount: 0 // Will be updated below
          };
        });
      }
      
      // Build program map for lookups
      const programMap = new Map(programs.map(p => [p.id, p.name]));
      
      // Build session count map
      const sessionCountMap = new Map();
      if (sessionSheet && sessionSheet.getLastRow() > 1) {
        sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 1).getValues().flat().forEach(pId => {
          const id = String(pId).trim();
          sessionCountMap.set(id, (sessionCountMap.get(id) || 0) + 1);
        });
      }
      
      // Update programs with session counts
      programs.forEach(p => {
        p.sessionCount = sessionCountMap.get(p.id) || 0;
      });
      
      const participantMap = new Map();
      
      if (enrollmentSheet && enrollmentSheet.getLastRow() > 1) {
        enrollmentSheet.getRange(2, 2, enrollmentSheet.getLastRow() - 1, 1).getValues().flat().forEach(id => {
          const s = String(id).trim();
          if (s) participantMap.set(s, (participantMap.get(s) || 0) + 1);
        });
      }
      
      // Generate 5 months calendar data
      const offsets = [-2, -1, 0, 1, 2];
      const months = offsets.map(o => {
        const d = new Date(centerYear, centerMonth - 1 + o, 1);
        return {
          year: d.getFullYear(),
          month: d.getMonth() + 1,
          data: new Map(),
          monthName: d.toLocaleString('default', { month: 'long' })
        };
      });
      
      // Process sessions
      let totalSessions = 0, completed = 0, upcoming = 0;
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const monthlyEvents = [];
      let filteredSessionIds = new Set();
      
      if (sessionSheet && sessionSheet.getLastRow() >= 2) {
        let values = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 15).getValues();
        
        // Apply visibility filtering
        // Column indices: L(11)=LastModifiedBy, M(12)=LastModifiedOn, N(13)=Entity, O(14)=TrackQR
        if (!canViewAllSessions_(userData)) {
          values = values.filter(r => {
            const sessionEntity = String(r[13] || '').trim();
            const modifiedBy = String(r[11] || '').trim();
            
            // User can always see their own sessions
            if (modifiedBy.toLowerCase() === userData.email.toLowerCase()) return true;
            
            // Everyone can see Global sessions
            if (sessionEntity === 'Global') return true;
            
            // Check if session entity is in user's country list
            return userCountries.includes(sessionEntity);
          });
        }
        
        // Apply calendar filter if provided
        if (filterEntity && filterEntity !== 'all') {
          values = values.filter(r => r[13] === filterEntity);
        }
        
        totalSessions = values.length;
        
        values.forEach(r => {
          const date = r[5];
          const status = (r[3] || '').toString().toLowerCase();
          const sid = String(r[1]).trim();
          const sessionEntity = String(r[13] || '').trim();
          const modifiedBy = String(r[11] || '').trim();
          
          // Track filtered session IDs for participant count
          filteredSessionIds.add(sid);
          
          if (status === 'completed' || status === 'closed') completed++;
          else if (date instanceof Date && date >= today) upcoming++;
          
          // Determine if user can edit this session
          const canEdit = canEditSession_(userData, sessionEntity, modifiedBy);
          
          // Monthly calendar events
          if (date instanceof Date) {
            monthlyEvents.push({
              sessionId: String(r[1]).trim(),
              programId: String(r[0]).trim(),
              name: r[2] || '',
              date: formatDateForSheet_(date),
              day: date.getDate(),
              month: date.getMonth() + 1,
              year: date.getFullYear(),
              status: r[3] || 'Open',
              programName: programMap.get(String(r[0]).trim()) || '',
              entity: sessionEntity,
              isGlobal: sessionEntity === 'Global',
              canEdit: canEdit
            });
          }
          
          if (!date || !(date instanceof Date)) return;
          
          const bucket = months.find(m => m.year === date.getFullYear() && m.month === (date.getMonth() + 1));
          if (!bucket) return;
          
          const pid = String(r[0]).trim();
          
          if (!bucket.data.has(pid)) {
            bucket.data.set(pid, {
              programId: pid,
              programName: programMap.get(pid) || pid,
              sessionCount: 0,
              participantCount: 0,
              minDate: new Date(date),
              maxDate: new Date(date)
            });
          }
          
          const p = bucket.data.get(pid);
          p.sessionCount++;
          p.participantCount += (participantMap.get(sid) || 0);
          if (date < p.minDate) p.minDate = new Date(date);
          if (date > p.maxDate) p.maxDate = new Date(date);
        });
      }
      
      // Calculate unique participants from filtered sessions only - count only "Attended" status
      let uniqueParticipants = 0;
      if (enrollmentSheet && enrollmentSheet.getLastRow() > 1 && filteredSessionIds && filteredSessionIds.size > 0) {
        // Read columns B (SessionID), C (EmployeeID), K (Status)
        const enrollData = enrollmentSheet.getRange(2, 2, enrollmentSheet.getLastRow() - 1, 10).getValues();
        const filteredEmployees = new Set();
        enrollData.forEach(row => {
          const sessionId = String(row[0]).trim();
          const employeeId = String(row[1]).trim();
          const status = String(row[9] || '').trim(); // Column K (index 9 from column B)
          if (filteredSessionIds.has(sessionId) && employeeId && status === 'Attended') {
            filteredEmployees.add(employeeId);
          }
        });
        uniqueParticipants = filteredEmployees.size;
      } else if (enrollmentSheet && enrollmentSheet.getLastRow() > 1 && (!filteredSessionIds || filteredSessionIds.size === 0)) {
        // Read columns C (EmployeeID), K (Status)
        const enrollData = enrollmentSheet.getRange(2, 3, enrollmentSheet.getLastRow() - 1, 9).getValues();
        const attendedEmployees = new Set();
        enrollData.forEach(row => {
          const employeeId = String(row[0]).trim();
          const status = String(row[8] || '').trim(); // Column K (index 8 from column C)
          if (employeeId && status === 'Attended') {
            attendedEmployees.add(employeeId);
          }
        });
        uniqueParticipants = attendedEmployees.size;
      }
      
      const calendarData = {
        months: months.map(m => ({
          year: m.year,
          month: m.month,
          monthName: `${m.monthName} ${m.year}`,
          programs: Array.from(m.data.values())
            .map(p => ({ ...p, minDate: p.minDate.toISOString(), maxDate: p.maxDate.toISOString() }))
            .sort((a, b) => new Date(a.minDate) - new Date(b.minDate))
        }))
      };
      
      return {
        programs,
        calendar: calendarData,
        monthlyEvents,
        stats: { totalSessions, completed, upcoming, ongoing: totalSessions - completed - upcoming, totalParticipants: uniqueParticipants },
        entityOptions,
        primaryEntity,
        availableEntities,
        canManagePrograms: canManagePrograms_(userData.role)
      };
    } catch (e) {
      console.error('getInitialData error:', e);
      return { programs: [], calendar: { months: [] }, stats: {}, entityOptions: [], primaryEntity: '', availableEntities: [], canManagePrograms: false, monthlyEvents: [] };
    }
  });
  } catch (outerError) {
    console.error('getInitialData outer error:', outerError);
    return { programs: [], calendar: { months: [] }, stats: {}, entityOptions: [], primaryEntity: '', availableEntities: [], canManagePrograms: false, monthlyEvents: [] };
  }
}

// Fast session list with entity, type, participant count
function getSessionsList(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      if (!sessionSheet || sessionSheet.getLastRow() <= 1) return [];
      
      const programSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      const enrollmentSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      
      const programMap = new Map();
      if (programSheet && programSheet.getLastRow() > 1) {
        programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 2).getValues()
          .forEach(r => programMap.set(String(r[0]).trim(), r[1]));
      }
      
      // Build participant count map
      const participantCountMap = new Map();
      if (enrollmentSheet && enrollmentSheet.getLastRow() > 1) {
        enrollmentSheet.getRange(2, 2, enrollmentSheet.getLastRow() - 1, 1).getValues().flat().forEach(id => {
          const s = String(id).trim();
          if (s) participantCountMap.set(s, (participantCountMap.get(s) || 0) + 1);
        });
      }
      
      let data = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 15).getValues();
      
      // Apply visibility filtering
      // Column indices: L(11)=LastModifiedBy, M(12)=LastModifiedOn, N(13)=Entity, O(14)=TrackQR
      
      // Global HRBP can VIEW all but canEdit is restricted in canEditSession_
      if (!canViewAllSessions_(userData)) {
        const userCountries = userData.countries || [];
        
        data = data.filter(r => {
          const sessionEntity = String(r[13] || '').trim();
          const modifiedBy = String(r[11] || '').trim();
          
          // User can always see their own sessions
          if (modifiedBy.toLowerCase() === userData.email.toLowerCase()) return true;
          
          // Everyone can see Global sessions
          if (sessionEntity === 'Global') return true;
          
          // Check if session entity is in user's country list
          return userCountries.includes(sessionEntity);
        });
      }
      
      return data.map(r => {
        const sessionId = String(r[1]).trim();
        const date = r[5];
        const sessionEntity = String(r[13] || '').trim();
        const modifiedBy = String(r[11] || '').trim();
        
        return {
          sessionId,
          programId: String(r[0]).trim(),
          programName: programMap.get(String(r[0]).trim()) || '',
          name: r[2] || '',
          status: r[3] || 'Open',
          type: r[4] || 'Classroom',
          date: date ? formatDateForSheet_(date) : '',
          year: date instanceof Date ? date.getFullYear() : '',
          month: date instanceof Date ? date.getMonth() + 1 : '',
          entity: sessionEntity,
          participantCount: participantCountMap.get(sessionId) || 0,
          isGlobal: sessionEntity === 'Global',
          canEdit: canEditSession_(userData, sessionEntity, modifiedBy),
          modifiedBy: modifiedBy
        };
      });
    } catch (e) {
      return [];
    }
  });
}

// Fast save - returns new session data
function createSessionFast(sessionToken, sessionData) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.SESSIONS);
    
    // Determine entity prefix for session ID
    let entityPrefix = sessionData.entity || '';
    if (!entityPrefix || entityPrefix === 'Global') {
      entityPrefix = 'GLOB';
    }
    // Ensure max 4 chars
    entityPrefix = entityPrefix.substring(0, 4).toUpperCase();
    
    // Generate session ID: ENTITY + YYMM + SEQ
    const targetDate = sessionData.date ? new Date(sessionData.date) : new Date();
    const yy = String(targetDate.getFullYear()).slice(-2);
    const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
    const base = entityPrefix + yy + mm;
    
    let maxSeq = 0;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().forEach(id => {
        const sId = String(id || "").trim();
        if (sId.startsWith(base)) {
          const n = parseInt(sId.substring(base.length), 10);
          if (!isNaN(n) && n > maxSeq) maxSeq = n;
        }
      });
    }
    const sessionId = base + (maxSeq + 1).toString().padStart(3, '0');
    
    // Write row (15 columns)
    // Columns: A=ProgramID, B=SessionID, C=Name, D=Status, E=Type, F=Date, G=Duration, H=Location, I=Score, J=Provider, K=IsIndivDev, L=LastModifiedBy, M=LastModifiedOn, N=Entity, O=TrackQR
    const row = [
      String(sessionData.programId).trim(), sessionId, sessionData.name || '',
      sessionData.status || 'Open', sessionData.type || 'Classroom',
      sessionData.date ? new Date(sessionData.date) : new Date(),
      sessionData.duration || 1, sessionData.location || '', sessionData.score || 0,
      sessionData.provider || 'Internal', sessionData.isIndivDev || 'No',
      userData.email,                     // Column L - Last Modified By
      new Date(),                         // Column M - Last Modified On
      sessionData.entity || '',           // Column N - Entity
      sessionData.trackQR || 'No'         // Column O - Track QR
    ];
    
    const target = sheet.getLastRow() + 1;
    sheet.getRange(target, 1, 1, 15).setValues([row]);
    sheet.getRange(target, 2).setNumberFormat("@");
    
    return { success: true, sessionId: sessionId };
  });
}

// Participants summary with full employee details
function getParticipantsSummary(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const empSheet = getSheetSafe_(SHEET_NAMES.EMPLOYEES);
      const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      
      if (!empSheet || empSheet.getLastRow() <= 1) return { participants: [], filters: {} };
      
      // Auto-update enrollments to "Absent" for past QR-tracked sessions
      autoUpdateAbsentEnrollments_();
      
      // Determine user's entity access
      const userCountries = userData.countries || [];
      const canViewAll = canViewAllSessions_(userData);
      
      // Build session info map (hours from sessions) - Column B=SessionID, G=Duration
      const sessionInfoMap = new Map();
      if (sessionSheet && sessionSheet.getLastRow() > 1) {
        sessionSheet.getRange(2, 2, sessionSheet.getLastRow() - 1, 6).getValues().forEach(r => {
          sessionInfoMap.set(String(r[0]).trim(), {
            hours: parseFloat(r[5]) || 0
          });
        });
      }
      
      // Build enrollment stats map (aggregate by employee)
      // Count only "Attended" for sessions/hours
      // Columns: A=EnrollmentID, B=SessionID, C=EmployeeID, K=Status (index 10)
      const enrollmentStats = new Map();
      if (enrollSheet && enrollSheet.getLastRow() > 1) {
        const enrollData = enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 11).getValues();
        enrollData.forEach(r => {
          const empId = String(r[2]).trim();
          if (!empId) return; // Skip blank
          
          const sessionId = String(r[1]).trim();
          const status = String(r[10] || '').trim();
          const sessionInfo = sessionInfoMap.get(sessionId) || { hours: 0 };
          
          if (!enrollmentStats.has(empId)) {
            enrollmentStats.set(empId, { sessionCount: 0, totalHours: 0 });
          }
          
          // Only count "Attended" sessions and hours
          if (status === 'Attended') {
            const stat = enrollmentStats.get(empId);
            stat.sessionCount++;
            stat.totalHours += sessionInfo.hours;
          }
        });
      }
      
      // Read ALL employees from Employees sheet
      // Columns: A=Entity, B=EmployeeID, C=Name, D=JobBand, E=Function, F=SubFunction, G=PositionType, H=Email
      const empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 8).getValues();
      
      const participants = [];
      const entities = new Set();
      const functions = new Set();
      const jobBands = new Set();
      
      empData.forEach(r => {
        const empId = String(r[1]).trim();
        const entity = String(r[0] || '').trim();
        
        // Skip blank rows (no Employee ID)
        if (!empId) return;
        
        // Collect filter options from ALL employees (before filtering)
        if (entity) entities.add(entity);
        if (r[4]) functions.add(r[4]);
        if (r[3]) jobBands.add(r[3]);
        
        // Apply entity filtering based on user's Countries
        if (!canViewAll) {
          if (!userCountries.includes(entity)) {
            return; // Skip employees not in user's Countries
          }
        }
        
        // Get enrollment stats (default to 0 if none)
        const stats = enrollmentStats.get(empId) || { sessionCount: 0, totalHours: 0 };
        
        participants.push({
          employeeId: empId,
          entity: entity,
          name: r[2] || '',
          jobBand: r[3] || '',
          function: r[4] || '',
          subFunction: r[5] || '',
          positionType: r[6] || '',
          email: r[7] || '',
          sessionCount: stats.sessionCount,
          totalHours: stats.totalHours
        });
      });
      
      // Sort by session count (descending), then by name (ascending)
      participants.sort((a, b) => {
        if (b.sessionCount !== a.sessionCount) return b.sessionCount - a.sessionCount;
        return (a.name || '').localeCompare(b.name || '');
      });
      
      // Filter options should only include entities the user can see
      let filteredEntities = entities;
      if (!canViewAll) {
        filteredEntities = new Set([...entities].filter(e => userCountries.includes(e)));
      }
      
      return {
        participants: participants,
        filters: {
          entities: Array.from(filteredEntities).sort(),
          functions: Array.from(functions).sort(),
          jobBands: Array.from(jobBands).sort()
        }
      };
    } catch (e) {
      console.error('getParticipantsSummary error:', e);
      return { participants: [], filters: {} };
    }
  });
}

/**
 * Auto-update enrollments to "Absent" for sessions where:
 * - Track via QR = "Yes"
 * - Complete Date < Today (past)
 * - Status = "Enrolled"
 */
function autoUpdateAbsentEnrollments_() {
  try {
    const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
    const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
    
    if (!sessionSheet || !enrollSheet) return;
    if (sessionSheet.getLastRow() <= 1 || enrollSheet.getLastRow() <= 1) return;
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Build set of session IDs that are: Track via QR = Yes AND Date < Today
    // Column structure: B=SessionID (index 1), F=Date (index 5), O=TrackQR (index 14)
    const sessData = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 15).getValues();
    const pastQRSessions = new Set();
    
    sessData.forEach(r => {
      const sessionId = String(r[1]).trim();
      const sessionDate = r[5];
      const trackQR = String(r[14] || '').trim().toLowerCase();
      
      if (trackQR === 'yes' && sessionDate instanceof Date) {
        const sessDay = new Date(sessionDate);
        sessDay.setHours(0, 0, 0, 0);
        if (sessDay < today) {
          pastQRSessions.add(sessionId);
        }
      }
    });
    
    if (pastQRSessions.size === 0) return;
    
    // Find enrollments to update: SessionID in pastQRSessions AND Status = "Enrolled"
    // Column K (index 10) = Status
    const enrollData = enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 11).getValues();
    const rowsToUpdate = [];
    
    enrollData.forEach((r, idx) => {
      const sessionId = String(r[1]).trim();
      const status = String(r[10] || '').trim();
      
      if (pastQRSessions.has(sessionId) && status === 'Enrolled') {
        rowsToUpdate.push(idx + 2); // +2 because data starts at row 2 and idx is 0-based
      }
    });
    
    // Update status to "Absent" for identified rows
    rowsToUpdate.forEach(rowNum => {
      enrollSheet.getRange(rowNum, 11).setValue('Absent'); // Column K = 11
    });
    
    if (rowsToUpdate.length > 0) {
      console.log(`Auto-updated ${rowsToUpdate.length} enrollments to "Absent"`);
    }
  } catch (e) {
    console.error('autoUpdateAbsentEnrollments_ error:', e);
  }
}

// Get participant's sessions
function getParticipantSessions(sessionToken, employeeId) {
  return withSession_(sessionToken, (userData) => {
    try {
      const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      const programSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      
      if (!enrollSheet || enrollSheet.getLastRow() <= 1) return { sessions: [], summary: {} };
      
      // Auto-update enrollments to "Absent" for past QR-tracked sessions
      autoUpdateAbsentEnrollments_();
      
      // Build maps
      const programMap = new Map();
      if (programSheet && programSheet.getLastRow() > 1) {
        programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 2).getValues()
          .forEach(r => programMap.set(String(r[0]).trim(), r[1]));
      }
      
      const sessionMap = new Map();
      if (sessionSheet && sessionSheet.getLastRow() > 1) {
        sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 8).getValues().forEach(r => {
          sessionMap.set(String(r[1]).trim(), {
            programId: String(r[0]).trim(),
            name: r[2] || '',
            status: r[3] || '',
            type: r[4] || '',
            date: r[5] ? formatDateForSheet_(r[5]) : '',
            hours: parseFloat(r[6]) || 0,
            location: r[7] || ''
          });
        });
      }
      
      // Find enrollments for this employee
      const enrollData = enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 11).getValues();
      const sessions = [];
      let attendedSessions = 0;
      let totalHours = 0;
      
      enrollData.forEach(r => {
        if (String(r[2]).trim() !== String(employeeId).trim()) return;
        
        const sessionId = String(r[1]).trim();
        const sess = sessionMap.get(sessionId);
        if (!sess) return;
        
        const enrollStatus = String(r[10] || 'Enrolled').trim();
        
        sessions.push({
          sessionId,
          sessionName: sess.name,
          programName: programMap.get(sess.programId) || '',
          date: sess.date,
          hours: sess.hours,
          type: sess.type,
          status: enrollStatus
        });
        
        // Only count "Attended" for summary stats
        if (enrollStatus === 'Attended') {
          attendedSessions++;
          totalHours += sess.hours;
        }
      });
      
      sessions.sort((a, b) => new Date(b.date) - new Date(a.date));
      
      return {
        sessions,
        summary: {
          totalSessions: attendedSessions,
          totalHours: totalHours
        }
      };
    } catch (e) {
      return { sessions: [], summary: {} };
    }
  });
}

// ============================================
// PROGRAMS API
// ============================================
function getPrograms(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      if (!sheet) return { programs: [], canManage: false };
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return { programs: [], canManage: canManagePrograms_(userData.role) };
      
      // Read columns: ID, Name, Category, Description, LastModifiedBy, LastModifiedOn (max 6)
      const lastCol = Math.min(sheet.getLastColumn(), 6);
      const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      
      // Get session counts per program
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      const sessionCountMap = new Map();
      if (sessionSheet && sessionSheet.getLastRow() > 1) {
        sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 1).getValues().flat().forEach(pId => {
          const id = String(pId).trim();
          sessionCountMap.set(id, (sessionCountMap.get(id) || 0) + 1);
        });
      }
      
      const programs = data.map(row => {
        let modifiedOnStr = '';
        // Only access row[5] if we have 6 columns
        if (lastCol >= 6 && row.length >= 6 && row[5]) {
          const modifiedOn = row[5];
          if (modifiedOn instanceof Date) {
            modifiedOnStr = Utilities.formatDate(modifiedOn, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm');
          } else {
            modifiedOnStr = String(modifiedOn);
          }
        }
        
        return {
          id: String(row[0] || '').trim(),
          name: row[1] || '',
          category: row[2] || '',
          description: row[3] || '',
          modifiedBy: (lastCol >= 5 && row.length >= 5) ? (row[4] || '') : '',
          modifiedOn: modifiedOnStr,
          sessionCount: sessionCountMap.get(String(row[0] || '').trim()) || 0
        };
      });
      
      return { 
        programs, 
        canManage: canManagePrograms_(userData.role)
      };
    } catch (e) {
      console.error('getPrograms error:', e);
      return { programs: [], canManage: false };
    }
  });
}

function generateProgramId(sessionToken) {
  return withSession_(sessionToken, (sessionData) => {
    try {
      const sheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      const existing = new Set();
      
      if (sheet && sheet.getLastRow() > 1) {
        const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
        ids.forEach(id => existing.add(String(id)));
      }
      
      while (true) {
        const id = "P" + Math.floor(100000 + Math.random() * 900000);
        if (!existing.has(id)) return id;
      }
    } catch (e) {
      return "P" + Math.floor(100000 + Math.random() * 900000);
    }
  });
}

function createProgram(sessionToken, programData) {
  return withSession_(sessionToken, (userData) => {
    if (!canManagePrograms_(userData.role)) {
      throw new Error('You do not have permission to create programs.');
    }
    
    const sheet = getSheet_(SHEET_NAMES.PROGRAMS);
    const id = programData.id || ('P' + Math.floor(100000 + Math.random() * 900000));
    
    // Columns: ID, Name, Category, Description, LastModifiedBy, LastModifiedOn
    sheet.appendRow([
      id, 
      programData.name, 
      programData.category, 
      programData.description || '',
      userData.email,
      new Date()
    ]);
    
    return { success: true, id: id };
  });
}

function updateProgram(sessionToken, programData) {
  return withSession_(sessionToken, (userData) => {
    if (!canManagePrograms_(userData.role)) {
      throw new Error('You do not have permission to update programs.');
    }
    
    const sheet = getSheet_(SHEET_NAMES.PROGRAMS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(programData.id).trim()) {
        // Update all 6 columns including Last Modified
        sheet.getRange(i + 1, 1, 1, 6).setValues([[
          programData.id,
          programData.name,
          programData.category,
          programData.description || '',
          userData.email,    // Last Modified By
          new Date()         // Last Modified On
        ]]);
        return { success: true };
      }
    }
    
    throw new Error('Program not found.');
  });
}

function deleteProgram(sessionToken, programId) {
  return withSession_(sessionToken, (userData) => {
    if (!canManagePrograms_(userData.role)) {
      throw new Error('You do not have permission to delete programs.');
    }
    
    const sheet = getSheet_(SHEET_NAMES.PROGRAMS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(programId).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    
    throw new Error('Program not found.');
  });
}

// ============================================
// SESSIONS API
// ============================================
function generateSessionId(sessionToken, dateInput) {
  return withSession_(sessionToken, (sessionData) => {
    const prefix = "EWTH";
    let targetDate = new Date();
    if (dateInput) {
      const parsed = new Date(dateInput);
      if (!isNaN(parsed.getTime())) targetDate = parsed;
    }
    
    const yy = String(targetDate.getFullYear()).slice(-2);
    const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
    const base = prefix + yy + mm;
    
    let maxSeq = 0;
    
    try {
      const sheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      if (sheet && sheet.getLastRow() > 1) {
        const ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
        ids.forEach(id => {
          const sId = String(id || "").trim();
          if (sId.startsWith(base)) {
            const tail = sId.substring(base.length);
            const n = parseInt(tail, 10);
            if (!isNaN(n) && n > maxSeq) maxSeq = n;
          }
        });
      }
    } catch (e) {
      console.error('generateSessionId error:', e);
    }
    
    return base + (maxSeq + 1).toString().padStart(3, '0');
  });
}

function getSessions(sessionToken, programId) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      if (!sheet) return [];
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return [];
      
      // Read 15 columns: A-O
      const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
      
      let sessions = data.map(row => {
        const pId = String(row[0]).trim();
        // Column indices: L(11)=LastModifiedBy, M(12)=LastModifiedOn, N(13)=Entity, O(14)=TrackQR
        const sessionEntity = String(row[13] || '').trim();
        const modifiedBy = String(row[11] || '').trim();
        
        return {
          programId: pId,
          sessionId: String(row[1]).trim(),
          name: row[2] || '',
          status: row[3] || 'Open',
          type: row[4] || 'Classroom',
          date: row[5] ? formatDateForSheet_(row[5]) : '',
          duration: row[6] || 1,
          location: row[7] || '',
          score: row[8] || 0,
          provider: row[9] || 'Internal',
          isIndivDev: row[10] || 'No',
          modifiedBy: modifiedBy,
          entity: sessionEntity,
          trackQR: row[14] || 'No',
          isGlobal: sessionEntity === 'Global',
          canEdit: canEditSession_(userData, sessionEntity, modifiedBy)
        };
      });
      
      // Filter by programId if provided
      if (programId) {
        sessions = sessions.filter(s => s.programId === String(programId).trim());
      }
      
      // Apply visibility filtering
      // Global HRBP can VIEW all but canEdit is already restricted above
      if (canViewAllSessions_(userData)) {
        return sessions;
      }
      
      // Entity-based filtering for Regional Talent CoE, Country HRBP, Manager
      const userCountries = userData.countries || [];
      
      // Filter by entity access
      return sessions.filter(s => {
        // User can always see their own sessions
        if (s.modifiedBy.toLowerCase() === userData.email.toLowerCase()) return true;
        
        // Everyone can see Global sessions
        if (s.entity === 'Global') return true;
        
        // Check if session entity is in user's country list
        return userCountries.includes(s.entity);
      });
      
    } catch (e) {
      console.error('getSessions error:', e);
      return [];
    }
  });
}

function createSession(sessionToken, sessionData) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.SESSIONS);
    
    // Generate session ID if not provided
    let sessionId = sessionData.sessionId;
    if (!sessionId || sessionId === 'Auto' || sessionId === 'New') {
      const prefix = "EWTH";
      const targetDate = sessionData.date ? new Date(sessionData.date) : new Date();
      const yy = String(targetDate.getFullYear()).slice(-2);
      const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
      const base = prefix + yy + mm;
      
      let maxSeq = 0;
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
        ids.forEach(id => {
          const sId = String(id || "").trim();
          if (sId.startsWith(base)) {
            const tail = sId.substring(base.length);
            const n = parseInt(tail, 10);
            if (!isNaN(n) && n > maxSeq) maxSeq = n;
          }
        });
      }
      
      sessionId = base + (maxSeq + 1).toString().padStart(3, '0');
    }
    
    // 15 columns (A-O): ProgramID, SessionID, Name, Status, Type, Date, Duration, Location, Score, Provider, IsIndivDev, LastModifiedBy, LastModifiedOn, Entity, TrackQR
    const row = [
      String(sessionData.programId).trim(),
      sessionId,
      sessionData.name || '',
      sessionData.status || 'Open',
      sessionData.type || 'Classroom',
      sessionData.date ? new Date(sessionData.date) : new Date(),
      sessionData.duration || 1,
      sessionData.location || '',
      sessionData.score || 0,
      sessionData.provider || 'Internal',
      sessionData.isIndivDev || 'No',
      userData.email,              // Column L - Last Modified By
      new Date(),                  // Column M - Last Modified On
      sessionData.entity || '',    // Column N - Entity
      sessionData.trackQR || 'No'  // Column O - Track Attendance via QR
    ];
    
    const target = sheet.getLastRow() + 1;
    sheet.getRange(target, 1, 1, 15).setValues([row]);
    sheet.getRange(target, 2).setNumberFormat("@");
    
    return { success: true, sessionId: sessionId };
  });
}

function updateSession(sessionToken, sessionData) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.SESSIONS);
    const data = sheet.getDataRange().getValues();
    const targetId = String(sessionData.sessionId).trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === targetId) {
        // Check permission using canEditSession_
        // Column indices: L(11)=LastModifiedBy, M(12)=LastModifiedOn, N(13)=Entity
        const sessionEntity = String(data[i][13] || '').trim();
        const modifiedBy = String(data[i][11] || '').trim();
        
        if (!canEditSession_(userData, sessionEntity, modifiedBy)) {
          throw new Error('You do not have permission to edit this session.');
        }
        
        // Update first 11 columns
        sheet.getRange(i + 1, 1, 1, 11).setValues([[
          String(sessionData.programId).trim(),
          targetId,
          sessionData.name || '',
          sessionData.status || 'Open',
          sessionData.type || 'Classroom',
          sessionData.date ? new Date(sessionData.date) : new Date(),
          sessionData.duration || 1,
          sessionData.location || '',
          sessionData.score || 0,
          sessionData.provider || 'Internal',
          sessionData.isIndivDev || 'No'
        ]]);
        
        // Update Last Modified By (column L = 12)
        sheet.getRange(i + 1, 12).setValue(userData.email);
        
        // Update Last Modified On (column M = 13)
        sheet.getRange(i + 1, 13).setValue(new Date());
        
        // Update entity if provided (column N = 14)
        if (sessionData.entity !== undefined) {
          sheet.getRange(i + 1, 14).setValue(sessionData.entity || '');
        }
        
        // Update trackQR if provided (column O = 15)
        if (sessionData.trackQR !== undefined) {
          sheet.getRange(i + 1, 15).setValue(sessionData.trackQR || 'No');
        }
        
        sheet.getRange(i + 1, 2).setNumberFormat("@");
        
        return { success: true };
      }
    }
    
    throw new Error('Session not found.');
  });
}

/**
 * Quick update just the trackQR field for a session
 */
function updateSessionTrackQR(sessionToken, sessionId, trackQR) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.SESSIONS);
    const data = sheet.getDataRange().getValues();
    const targetId = String(sessionId).trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === targetId) {
        // Update trackQR column (O = 15)
        sheet.getRange(i + 1, 15).setValue(trackQR || 'No');
        return { success: true };
      }
    }
    
    throw new Error('Session not found.');
  });
}

function deleteSession(sessionToken, sessionId) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.SESSIONS);
    const data = sheet.getDataRange().getValues();
    const targetId = String(sessionId).trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === targetId) {
        // Check permission using canEditSession_
        // Column indices: L(11)=CreatedBy, M(12)=CreatedOn, N(13)=Entity
        const sessionEntity = String(data[i][13] || '').trim();
        const createdBy = String(data[i][11] || '').trim();
        
        if (!canEditSession_(userData, sessionEntity, createdBy)) {
          throw new Error('You do not have permission to delete this session.');
        }
        
        sheet.deleteRow(i + 1);
        deleteEnrollmentsBySession_(sessionId);
        
        return { success: true };
      }
    }
    
    throw new Error('Session not found.');
  });
}

function getSessionDetails(sessionToken, programId, sessionId) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      if (!sheet) return null;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return null;
      
      const targetProg = String(programId).trim();
      const targetSess = String(sessionId).trim();
      
      // Read 15 columns: A-O
      const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
      
      for (let i = 0; i < data.length; i++) {
        const r = data[i];
        if (String(r[0]).trim() === targetProg && String(r[1]).trim() === targetSess) {
          const modifiedBy = String(r[11] || '').trim();
          const modifiedOn = r[12]; // Column M - LastModifiedOn
          const sessionEntity = String(r[13] || '').trim(); // Column N - Entity
          
          // Format modifiedOn as date/time string
          let modifiedOnStr = '';
          if (modifiedOn instanceof Date) {
            modifiedOnStr = Utilities.formatDate(modifiedOn, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm');
          } else if (modifiedOn) {
            modifiedOnStr = String(modifiedOn);
          }
          
          return {
            programId: String(r[0]).trim(),
            sessionId: String(r[1]).trim(),
            name: r[2] || '',
            status: r[3] || 'Open',
            type: r[4] || 'Classroom',
            date: r[5] ? formatDateForSheet_(r[5]) : '',
            duration: r[6] || 1,
            location: r[7] || '',
            score: r[8] || 0,
            provider: r[9] || 'Internal',
            isIndivDev: r[10] || 'No',
            modifiedBy: modifiedBy,
            modifiedOn: modifiedOnStr,
            entity: sessionEntity,
            trackQR: r[14] || 'No',
            isGlobal: sessionEntity === 'Global',
            canEdit: canEditSession_(userData, sessionEntity, modifiedBy)
          };
        }
      }
    } catch (e) {
      console.error('getSessionDetails error:', e);
    }
    
    return null;
  });
}

// ============================================
// ENROLLMENTS API
// ============================================
function getEnrollments(sessionToken, sessionId) {
  return withSession_(sessionToken, (userData) => {
    try {
      console.log('getEnrollments called with sessionId:', sessionId);
      
      // Auto-update enrollments to "Absent" for past QR-tracked sessions
      autoUpdateAbsentEnrollments_();
      
      const sheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      if (!sheet) {
        console.log('Enrollments sheet not found');
        return [];
      }
      
      const lastRow = sheet.getLastRow();
      console.log('Enrollments sheet lastRow:', lastRow);
      
      if (lastRow <= 1) {
        console.log('No data rows in Enrollments sheet');
        return [];
      }
      
      const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
      console.log('Total enrollment rows:', data.length);
      
      let enrollments = data.map(row => {
        const obj = {
          enrollmentId: String(row[0] || ''),
          sessionId: String(row[1] || '').trim(),
          employeeId: String(row[2] || '').trim(),
          name: String(row[3] || ''),
          email: String(row[4] || ''),
          entity: String(row[5] || ''),
          jobBand: String(row[6] || ''),
          empFunction: String(row[7] || ''),
          subFunction: String(row[8] || ''),
          positionType: String(row[9] || ''),
          status: String(row[10] || 'Enrolled'),
          currency: String(row[11] || ''),
          cost: row[12] || ''
        };
        return obj;
      });
      
      if (sessionId) {
        const searchId = String(sessionId).trim();
        console.log('Filtering for sessionId:', searchId);
        
        // Log all unique session IDs in enrollments for debugging
        const uniqueSessionIds = [...new Set(enrollments.map(e => e.sessionId))];
        console.log('Unique session IDs in enrollments:', uniqueSessionIds.join(', '));
        
        enrollments = enrollments.filter(e => e.sessionId === searchId);
        console.log('Filtered enrollments count:', enrollments.length);
        
        if (enrollments.length > 0) {
          console.log('First enrollment:', JSON.stringify(enrollments[0]));
        }
      }
      
      return enrollments;
    } catch (e) {
      console.error('getEnrollments error:', e);
      return [];
    }
  });
}

function addParticipants(sessionToken, sessionId, employeeIds) {
  return withSession_(sessionToken, (userData) => {
    const enrollSheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
    const empSheet = getSheet_(SHEET_NAMES.EMPLOYEES);
    const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
    
    // Get session date to determine if this is a backlog add
    let isBacklog = false;
    if (sessionSheet) {
      const sessData = sessionSheet.getDataRange().getValues();
      for (let i = 1; i < sessData.length; i++) {
        if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
          const sessionDate = sessData[i][5]; // Column F = Training Date
          if (sessionDate instanceof Date) {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const sessDay = new Date(sessionDate);
            sessDay.setHours(0, 0, 0, 0);
            isBacklog = sessDay < today;
          }
          break;
        }
      }
    }
    
    // Build employee map
    const empLastRow = empSheet.getLastRow();
    const empMap = new Map();
    
    if (empLastRow > 1) {
      const empData = empSheet.getRange(2, 1, empLastRow - 1, 8).getValues();
      empData.forEach(r => {
        empMap.set(String(r[1]).trim(), {
          entity: r[0] || '',
          name: r[2] || '',
          jobBand: r[3] || '',
          function: r[4] || '',
          subFunction: r[5] || '',
          positionType: r[6] || '',
          email: r[7] || ''
        });
      });
    }
    
    // Get existing enrollments
    const enrollLastRow = enrollSheet.getLastRow();
    const existing = new Set();
    
    if (enrollLastRow > 1) {
      const enrollData = enrollSheet.getRange(2, 2, enrollLastRow - 1, 2).getValues();
      enrollData.forEach(r => {
        existing.add(`${String(r[0]).trim()}-${String(r[1]).trim()}`);
      });
    }
    
    // Determine initial status:
    // - If backlog (past date): 'Attended' 
    // - Otherwise: 'Enrolled'
    const initialStatus = isBacklog ? 'Attended' : 'Enrolled';
    
    // Add new enrollments
    const newRows = [];
    let added = 0;
    
    employeeIds.forEach(eid => {
      const key = `${sessionId}-${eid}`;
      if (!existing.has(key)) {
        const emp = empMap.get(String(eid).trim()) || { entity: '', name: '', email: '', function: '', subFunction: '', jobBand: '', positionType: '' };
        newRows.push([
          Utilities.getUuid(),
          sessionId,
          eid,
          emp.name,
          emp.email,
          emp.entity,
          emp.jobBand,
          emp.function,
          emp.subFunction,
          emp.positionType,
          initialStatus,  // Set based on date (backlog = Attended, future/today = Enrolled)
          '',
          ''
        ]);
        existing.add(key);
        added++;
      }
    });
    
    if (newRows.length) {
      enrollSheet.getRange(enrollSheet.getLastRow() + 1, 1, newRows.length, 13).setValues(newRows);
    }
    
    return { success: true, message: `Added ${added} participant(s).`, isBacklog: isBacklog };
  });
}

function removeParticipants(sessionToken, enrollmentIds) {
  return withSession_(sessionToken, (userData) => {
    if (!enrollmentIds || !enrollmentIds.length) {
      return { success: true, message: 'No IDs provided.' };
    }
    
    const sheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, message: 'No participants to remove.' };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
    const toDelete = new Set(enrollmentIds.map(String));
    let deletedCount = 0;
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (toDelete.has(data[i])) {
        sheet.deleteRow(i + 2);
        deletedCount++;
      }
    }
    
    return { success: true, message: `Removed ${deletedCount} participant(s).` };
  });
}

function updateParticipantCost(sessionToken, enrollmentId, currency, cost) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'No enrollments found.' };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i]) === String(enrollmentId)) {
        sheet.getRange(i + 2, 12).setValue(currency || '');
        sheet.getRange(i + 2, 13).setValue(cost || '');
        return { success: true };
      }
    }
    
    return { success: false, message: 'Participant not found.' };
  });
}

/**
 * Update enrollment status (Attended, Enrolled, Absent)
 */
function updateEnrollmentStatus(sessionToken, enrollmentId, newStatus) {
  return withSession_(sessionToken, (userData) => {
    const sheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'No enrollments found.' };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i]) === String(enrollmentId)) {
        // Status is in column 11 (K)
        sheet.getRange(i + 2, 11).setValue(newStatus || 'Enrolled');
        return { success: true };
      }
    }
    
    return { success: false, message: 'Enrollment not found.' };
  });
}

function addManualParticipant(sessionToken, sessionId, person) {
  return withSession_(sessionToken, (userData) => {
    // First add to employees sheet if not exists
    const empSheet = getSheet_(SHEET_NAMES.EMPLOYEES);
    const finder = empSheet.getRange("B:B").createTextFinder(person.id).matchEntireCell(true);
    
    if (!finder.findNext()) {
      empSheet.appendRow([
        person.entity || '',
        person.id,
        person.name,
        person.jobBand || '',
        person.function || '',
        person.subFunction || '',
        person.positionType || '',
        person.email
      ]);
    }
    
    // Then add enrollment
    return addParticipants(sessionToken, sessionId, [person.id]);
  });
}

function deleteEnrollmentsBySession_(sessionId) {
  try {
    const sheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim() === String(sessionId).trim()) {
        sheet.deleteRow(i + 2);
      }
    }
  } catch (e) {
    console.error('deleteEnrollmentsBySession_ error:', e);
  }
}

// ============================================
// EMPLOYEES API
// ============================================
/**
 * Helper function to check if an item passes access filters (entity-only)
 * @param {Object} userData - User data with countries
 * @param {string} itemEntity - The entity of the item
 * @param {string} itemCreatedBy - Who created the item (for sessions)
 * @returns {boolean} - Whether the item passes the filter
 */
function passesAccessFilter_(userData, itemEntity, itemCreatedBy) {
  // Full access users see everything
  if (hasFullAccess_(userData.role)) {
    return true;
  }
  
  // User can always see their own creations
  if (itemCreatedBy && itemCreatedBy.toLowerCase() === userData.email.toLowerCase()) {
    return true;
  }
  
  // Everyone can see Global items
  if (itemEntity === 'Global') {
    return true;
  }
  
  // Get user's access lists
  const userCountries = userData.countries || [];
  const hasAllCountries = userCountries.some(c => c.toLowerCase() === 'all');
  
  // If user has "All" countries, they see everything
  if (hasAllCountries) {
    return true;
  }
  
  // Check entity filter
  return userCountries.includes(itemEntity);
}

/**
 * Helper to filter employees by access (entity-only)
 */
function filterEmployeesByAccess_(employees, userData) {
  // Full access users see everything
  if (hasFullAccess_(userData.role)) {
    return employees;
  }
  
  const userCountries = userData.countries || [];
  const hasAllCountries = userCountries.some(c => c.toLowerCase() === 'all');
  
  // If user has "All" countries, they see all employees
  if (hasAllCountries) {
    return employees;
  }
  
  // Filter by entity
  return employees.filter(emp => {
    return userCountries.includes(emp.entity);
  });
}

function getEmployees(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      // Try cache first (5 min TTL) - cache key includes role for filtered results
      const cache = CacheService.getScriptCache();
      const cacheKey = 'EMPLOYEES_' + userData.email;
      const cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
      
      const sheet = getSheetSafe_(SHEET_NAMES.EMPLOYEES);
      if (!sheet) return [];
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return [];
      
      const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      
      let employees = data
        .filter(row => {
          // Skip blank rows (no Employee ID)
          const empId = String(row[1] || '').trim();
          return empId !== '';
        })
        .map(row => ({
          entity: row[0] || '',
          id: String(row[1]).trim(),
          name: row[2] || '',
          jobBand: row[3] || '',
          function: row[4] || '',
          subFunction: row[5] || '',
          positionType: row[6] || '',
          email: row[7] || ''
        }));
      
      // Apply access filtering
      const filtered = filterEmployeesByAccess_(employees, userData);
      
      // Cache for 5 minutes (may be large, so use try-catch)
      try {
        cache.put(cacheKey, JSON.stringify(filtered), 300);
      } catch (cacheErr) {
        console.log('Could not cache employees (may exceed size limit)');
      }
      
      return filtered;
      
    } catch (e) {
      console.error('getEmployees error:', e);
      return [];
    }
  });
}

/**
 * Fast employee search using TextFinder (server-side)
 * Returns up to 100 matching employees
 */
function searchEmployees(sessionToken, query) {
  return withSession_(sessionToken, (userData) => {
    try {
      if (!query || query.length < 2) return [];
      
      const sheet = getSheetSafe_(SHEET_NAMES.EMPLOYEES);
      if (!sheet) return [];
      
      // Use TextFinder for fast search
      const finder = sheet.createTextFinder(query).matchCase(false);
      const matches = finder.findAll();
      
      // Limit to first 100 matches for performance
      const results = [];
      const seenIds = new Set();
      
      for (const cell of matches) {
        if (results.length >= 100) break;
        
        const rowNum = cell.getRow();
        if (rowNum === 1) continue; // Skip header
        
        const row = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];
        const empId = String(row[1]).trim();
        
        // Skip blank Employee IDs
        if (!empId) continue;
        
        if (seenIds.has(empId)) continue;
        seenIds.add(empId);
        
        const emp = {
          entity: row[0] || '',
          id: empId,
          name: row[2] || '',
          jobBand: row[3] || '',
          function: row[4] || '',
          subFunction: row[5] || '',
          positionType: row[6] || '',
          email: row[7] || ''
        };
        
        // Apply access filtering
        if (passesEmployeeAccessFilter_(emp, userData)) {
          results.push(emp);
        }
      }
      
      return results;
    } catch (e) {
      console.error('searchEmployees error:', e);
      return [];
    }
  });
}

/**
 * Check if single employee passes access filter (entity-only)
 */
function passesEmployeeAccessFilter_(emp, userData) {
  if (hasFullAccess_(userData.role)) return true;
  
  const userCountries = userData.countries || [];
  const hasAllCountries = userCountries.some(c => c.toLowerCase() === 'all');
  
  if (hasAllCountries) return true;
  
  // Check entity filter
  return userCountries.includes(emp.entity);
}

// ============================================
// MASS UPLOAD
// ============================================
function previewMassUpload(sessionToken, base64Data, fileName) {
  return withSession_(sessionToken, (userData) => {
    try {
      const fileBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.CSV, fileName);
      const rows = Utilities.parseCsv(fileBlob.getDataAsString());
      
      if (!rows || rows.length === 0) return { results: [] };
      
      // Find the target column
      const headerRow = rows[0].map(h => h.toString().toLowerCase().trim());
      let targetColIdx = -1;
      for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i].includes("employee id") || headerRow[i].includes("email")) {
          targetColIdx = i;
          break;
        }
      }
      
      if (targetColIdx > -1) rows.shift();
      else targetColIdx = (rows[0].length >= 5) ? 4 : 0;
      
      // Build employee maps
      const empSheet = getSheetSafe_(SHEET_NAMES.EMPLOYEES);
      const empMapById = {};
      const empMapByEmail = {};
      
      if (empSheet && empSheet.getLastRow() > 1) {
        const data = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 8).getValues();
        
        for (let i = 0; i < data.length; i++) {
          const empId = String(data[i][1]).trim();
          const email = String(data[i][7] || '').trim().toLowerCase();
          const name = data[i][2];
          
          if (empId) empMapById[empId] = { id: empId, name: name };
          if (email) empMapByEmail[email] = { id: empId, name: name };
        }
      }
      
      // Process rows
      const results = [];
      rows.forEach(row => {
        if (row.length <= targetColIdx) return;
        const input = String(row[targetColIdx]).trim();
        if (!input || input.startsWith("Paste")) return;
        
        const match = empMapById[input] || empMapByEmail[input.toLowerCase()];
        results.push({
          input: input,
          isValid: !!match,
          name: match ? match.name : null,
          employeeId: match ? match.id : null,
          method: match ? (empMapById[input] ? "ID" : "Email") : null
        });
      });
      
      return { results: results };
    } catch (e) {
      console.error('previewMassUpload error:', e);
      return { results: [], error: e.message };
    }
  });
}

function createUploadTemplate(sessionToken, programId, programName, sessionId, sessionName) {
  return withSession_(sessionToken, (userData) => {
    try {
      const ss = SpreadsheetApp.create(`Upload_Template_${sessionId || "Session"}`);
      const sh = ss.getActiveSheet();
      sh.appendRow(["Program ID", "Program Name", "Session ID", "Session Name", "Employee ID / Email"]);
      sh.getRange(2, 1, 1, 4).setValues([[programId || "", programName || "", sessionId || "", sessionName || ""]]);
      sh.getRange("E2").setValue("Paste Employee IDs OR Emails here...");
      
      // Convert to Excel for download
      const ssFile = DriveApp.getFileById(ss.getId());
      const xlsxBlob = ssFile.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      const base64 = Utilities.base64Encode(xlsxBlob.getBytes());
      const filename = `Upload_Template_${sessionId || "Session"}.xlsx`;
      
      // Delete the temp spreadsheet
      ssFile.setTrashed(true);
      
      return { success: true, base64: base64, filename: filename };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

// ============================================
// CALENDAR DATA (5-Month View)
// ============================================
function getCalendarData(sessionToken, centerYear, centerMonth) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      const enrollmentSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      const programSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      
      // Build participant count map
      const participantMap = new Map();
      if (enrollmentSheet && enrollmentSheet.getLastRow() > 1) {
        enrollmentSheet.getRange(2, 2, enrollmentSheet.getLastRow() - 1, 1).getValues().flat().forEach(id => {
          const s = String(id).trim();
          if (s) participantMap.set(s, (participantMap.get(s) || 0) + 1);
        });
      }
      
      // Build program name map
      const programMap = new Map();
      if (programSheet && programSheet.getLastRow() > 1) {
        programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 2).getValues().forEach(r => {
          programMap.set(String(r[0]).trim(), r[1]);
        });
      }
      
      // Generate 5 months
      const offsets = [-2, -1, 0, 1, 2];
      const months = offsets.map(o => {
        const d = new Date(centerYear, centerMonth - 1 + o, 1);
        return {
          year: d.getFullYear(),
          month: d.getMonth() + 1,
          data: new Map(),
          monthName: d.toLocaleString('default', { month: 'long' })
        };
      });
      
      // Process sessions
      if (sessionSheet && sessionSheet.getLastRow() >= 2) {
        const values = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 16).getValues();
        
        // Apply visibility filtering
        // Column indices: L(11)=LastModifiedBy, M(12)=LastModifiedOn, N(13)=Entity
        
        let filteredValues = values;
        // Global HRBP can VIEW all sessions
        if (!canViewAllSessions_(userData)) {
          const userCountries = userData.countries || [];
          
          filteredValues = values.filter(r => {
            const sessionEntity = String(r[13] || '').trim();
            const modifiedBy = String(r[11] || '').trim();
            
            // User can always see their own sessions
            if (modifiedBy.toLowerCase() === userData.email.toLowerCase()) return true;
            
            // Everyone can see Global sessions
            if (sessionEntity === 'Global') return true;
            
            // Check entity access
            return userCountries.includes(sessionEntity);
          });
        }
        
        filteredValues.forEach(r => {
          const date = r[5];
          if (!date || !(date instanceof Date)) return;
          
          const bucket = months.find(m => m.year === date.getFullYear() && m.month === (date.getMonth() + 1));
          if (!bucket) return;
          
          const pid = String(r[0]).trim();
          const sid = String(r[1]).trim();
          
          if (!bucket.data.has(pid)) {
            bucket.data.set(pid, {
              programId: pid,
              programName: programMap.get(pid) || pid,
              sessionCount: 0,
              participantCount: 0,
              minDate: new Date(date),
              maxDate: new Date(date)
            });
          }
          
          const p = bucket.data.get(pid);
          p.sessionCount++;
          p.participantCount += (participantMap.get(sid) || 0);
          if (date < p.minDate) p.minDate = new Date(date);
          if (date > p.maxDate) p.maxDate = new Date(date);
        });
      }
      
      // Format response
      const result = months.map(m => ({
        year: m.year,
        month: m.month,
        monthName: `${m.monthName} ${m.year}`,
        programs: Array.from(m.data.values())
          .map(p => ({
            ...p,
            minDate: p.minDate.toISOString(),
            maxDate: p.maxDate.toISOString()
          }))
          .sort((a, b) => new Date(a.minDate) - new Date(b.minDate))
      }));
      
      return { months: result };
    } catch (e) {
      console.error('getCalendarData error:', e);
      return { months: [] };
    }
  });
}

// ============================================
// DASHBOARD STATS
// ============================================
function getDashboardStats(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessions = getSessions(sessionToken, null);
      const enrollments = getEnrollments(sessionToken, null);
      
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      
      let totalSessions = sessions.length;
      let completed = 0;
      let upcoming = 0;
      let ongoing = 0;
      
      sessions.forEach(s => {
        const status = (s.status || '').toLowerCase();
        if (status === 'completed' || status === 'closed') completed++;
        else if (status === 'ongoing') ongoing++;
        else {
          const sessionDate = s.date ? new Date(s.date) : null;
          if (sessionDate && sessionDate >= today) upcoming++;
        }
      });
      
      const uniqueParticipants = new Set(enrollments.map(e => e.employeeId)).size;
      
      return {
        totalSessions,
        completed,
        upcoming,
        ongoing,
        totalParticipants: uniqueParticipants,
        totalEnrollments: enrollments.length
      };
    } catch (e) {
      console.error('getDashboardStats error:', e);
      return {
        totalSessions: 0,
        completed: 0,
        upcoming: 0,
        ongoing: 0,
        totalParticipants: 0,
        totalEnrollments: 0
      };
    }
  });
}

// ============================================
// ANALYTICS DATA
// ============================================

/**
 * Get comprehensive analytics data for dashboard
 * Applies same access filtering as Calendar
 */
function getAnalyticsData(sessionToken, filters) {
  console.log('getAnalyticsData called, sessionToken:', sessionToken ? 'exists' : 'missing');
  console.log('getAnalyticsData filters:', JSON.stringify(filters));
  
  try {
    console.log('Validating session...');
    // Validate session first
    const sessionData = validateSession_(sessionToken);
    console.log('Session valid:', !!sessionData);
    
    if (!sessionData) {
      console.log('Returning session expired error');
      return {
        kpis: { totalHours: 0, avgHoursPerEmp: 0, completionRate: 0, uniqueLearners: 0, totalCost: 0, avgCostPerEmp: 0, budgetUtilization: 0, avgSatisfaction: 0 },
        charts: {},
        indivDev: [],
        availableYears: [],
        availableEntities: [],
        availableFunctions: [],
        availableSubFunctions: [],
        availableJobBands: [],
        error: 'Session expired',
        debug: { sessionExpired: true }
      };
    }
    
    extendSession_(sessionToken);
    const userData = sessionData;
    console.log('User:', userData.email, userData.role);
    
    const ss = getSpreadsheet_();
    const year = filters?.year || '';
    const quarter = filters?.quarter || '';
    const month = filters?.month || '';
    const entityFilter = filters?.entity || '';
    const functionFilter = filters?.function || '';
    const subFunctionFilter = filters?.subFunction || '';
    const jobBandFilter = filters?.jobBand || '';
    const programFilter = filters?.program || '';
      
    // Get all data sheets
    const sessSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
    const progSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
    const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
    const tm1Sheet = ss.getSheetByName('TM1');
    
    // Debug info
    const debug = {
      sessSheetExists: !!sessSheet,
      progSheetExists: !!progSheet,
      enrollSheetExists: !!enrollSheet,
      tm1SheetExists: !!tm1Sheet,
      userRole: userData.role,
      userEntities: userData.countries,
      filters: { year, quarter, month, entityFilter, functionFilter, subFunctionFilter, jobBandFilter, programFilter }
    };
    
    // Build program map (id -> {name, category})
    const programMap = new Map();
    const allProgramNames = [];
    if (progSheet && progSheet.getLastRow() > 1) {
      const progData = progSheet.getRange(2, 1, progSheet.getLastRow() - 1, 3).getValues();
      progData.forEach(r => {
        const name = String(r[1] || '').trim();
        programMap.set(String(r[0]).trim(), { name: name, category: r[2] || 'Uncategorized' });
        if (name) allProgramNames.push(name);
      });
    }
    debug.programCount = programMap.size;
    
    // Get sessions with filtering
    const sessData = sessSheet && sessSheet.getLastRow() > 1 
        ? sessSheet.getRange(2, 1, sessSheet.getLastRow() - 1, 16).getValues() 
        : [];
      debug.totalSessions = sessData.length;
      
      // Apply access filtering (entity-only)
      let userEntities = [];
      if (userData.countries) {
        if (typeof userData.countries === 'string') {
          userEntities = userData.countries.split(',').map(c => c.trim()).filter(c => c);
        } else if (Array.isArray(userData.countries)) {
          userEntities = userData.countries;
        }
      }
      
      // Check for view all access (includes Global HRBP)
      const role = (userData.role || '').trim();
      const hasAllEntities = canViewAllSessions_(userData);
      
      debug.userEntitiesParsed = userEntities;
      debug.hasAllEntities = hasAllEntities;
      
      // Filter sessions
      // Column indices: L(11)=LastModifiedBy, M(12)=LastModifiedOn, N(13)=Entity
      let filteredSessions = sessData.filter(row => {
        // Skip undefined or empty rows
        if (!row || !Array.isArray(row) || row.length < 14) return false;
        
        const sessionEntity = String(row[13] || '').trim(); // Column N - Entity
        
        // Access filtering
        // Everyone sees Global sessions
        if (!hasAllEntities && sessionEntity !== 'Global') {
          if (userEntities.length > 0 && !userEntities.includes(sessionEntity)) {
            return false;
          }
        }
        
        // Entity filter from UI dropdown
        if (entityFilter && sessionEntity !== entityFilter) {
          return false;
        }
        
        // Date filtering
        const sessionDate = row[5]; // Column F - Date
        if (year || quarter || month) {
          // If time filter is active, we need a valid date
          if (!sessionDate) return false;
          
          let d;
          if (sessionDate instanceof Date) {
            d = sessionDate;
          } else if (typeof sessionDate === 'string') {
            // Try parsing MM/DD/YYYY format
            const parts = sessionDate.split('/');
            if (parts.length === 3) {
              d = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
            } else {
              d = new Date(sessionDate);
            }
          } else if (typeof sessionDate === 'number') {
            // Excel serial date
            d = new Date((sessionDate - 25569) * 86400 * 1000);
          } else {
            return false; // Invalid date format
          }
          
          // Check if date is valid
          if (isNaN(d.getTime())) return false;
          
          const sessionYear = d.getFullYear();
          const sessionMonth = d.getMonth() + 1;
          const sessionQuarter = sessionMonth <= 3 ? 'Q1' : sessionMonth <= 6 ? 'Q2' : sessionMonth <= 9 ? 'Q3' : 'Q4';
          
          // Yearly filter: include all sessions in that year
          if (year && sessionYear != parseInt(year)) return false;
          
          // Quarterly filter: include sessions in that quarter (only if quarter is specified)
          if (quarter && sessionQuarter !== quarter) return false;
          
          // Monthly filter: include sessions in that month (only if month is specified)
          if (month && sessionMonth != parseInt(month)) return false;
        }
        
        // Program filter
        if (programFilter) {
          const progId = String(row[0] || '').trim();
          const prog = programMap.get(progId);
          const progName = prog?.name || '';
          if (progName !== programFilter) return false;
        }
        
        return true;
      });
      debug.filteredSessions = filteredSessions.length;
      
      // Debug: Log filter results
      console.log('=== SESSION FILTER DEBUG ===');
      console.log('Time filters: year=' + year + ', quarter=' + quarter + ', month=' + month);
      console.log('Other filters: entity=' + entityFilter + ', program=' + programFilter);
      console.log('Total sessions:', sessData.length, '-> Filtered:', filteredSessions.length);
      
      // Log first 3 sessions to understand date format
      if (sessData.length > 0) {
        console.log('First 3 sessions date info:');
        for (let i = 0; i < Math.min(3, sessData.length); i++) {
          const row = sessData[i];
          if (row && row.length > 5) {
            const dateVal = row[5];
            let parsedDate = null;
            if (dateVal instanceof Date) {
              parsedDate = dateVal;
            } else if (typeof dateVal === 'string') {
              const parts = dateVal.split('/');
              if (parts.length === 3) {
                parsedDate = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
              }
            }
            console.log(`  Session ${i}: raw="${dateVal}", type=${typeof dateVal}, parsed=${parsedDate ? parsedDate.toISOString() : 'null'}`);
          }
        }
      }
      
      // Log sample session for debugging - convert Date to string
      if (sessData.length > 0 && sessData[0] && sessData[0].length > 12) {
        const sampleDate = sessData[0][5];
        const isIndivValue = sessData[0][10];
        console.log('Sample session date:', sampleDate, 'type:', typeof sampleDate);
        debug.sampleSession = {
          programId: String(sessData[0][0] || ''),
          sessionId: String(sessData[0][1] || ''),
          name: String(sessData[0][2] || ''),
          status: String(sessData[0][3] || ''),
          type: String(sessData[0][4] || ''),
          date: sampleDate instanceof Date ? sampleDate.toISOString() : String(sampleDate || ''),
          duration: sessData[0][6],
          isIndivDev: String(isIndivValue || ''),
          entity: String(sessData[0][12] || '')
        };
      }
      
      // Get session IDs for enrollment filtering
      const filteredSessionIds = new Set(filteredSessions.map(s => String(s[1]).trim()));
      
      // Get enrollments
      const enrollData = enrollSheet && enrollSheet.getLastRow() > 1
        ? enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 13).getValues()
        : [];
      debug.totalEnrollments = enrollData.length;
      
      // Filter enrollments to only those in filtered sessions AND matching new filters
      const filteredEnrollments = enrollData.filter(e => {
        if (!e || !Array.isArray(e) || e.length < 11) return false;
        
        const sessId = String(e[1]).trim();
        if (!filteredSessionIds.has(sessId)) return false;
        
        // Apply entity filter from enrollment (column 5)
        if (entityFilter) {
          const enrollEntity = String(e[5] || '').trim();
          if (enrollEntity !== entityFilter) return false;
        }
        
        // Apply function filter (column 7)
        if (functionFilter) {
          const enrollFunction = String(e[7] || '').trim();
          if (enrollFunction !== functionFilter) return false;
        }
        
        // Apply sub-function filter (column 8)
        if (subFunctionFilter) {
          const enrollSubFunction = String(e[8] || '').trim();
          if (enrollSubFunction !== subFunctionFilter) return false;
        }
        
        // Apply job band filter (column 6)
        if (jobBandFilter) {
          const enrollJobBand = String(e[6] || '').trim();
          if (enrollJobBand !== jobBandFilter) return false;
        }
        
        return true;
      });
      debug.filteredEnrollments = filteredEnrollments.length;
      
      // Count attended enrollments
      const attendedEnrollments = filteredEnrollments.filter(e => String(e[10] || '').trim() === 'Attended');
      debug.attendedEnrollments = attendedEnrollments.length;
      
      // Sample enrollment for debugging
      if (enrollData.length > 0 && enrollData[0] && enrollData[0].length > 10) {
        debug.sampleEnrollment = {
          id: enrollData[0][0],
          sessionId: enrollData[0][1],
          employeeId: enrollData[0][2],
          name: enrollData[0][3],
          entity: enrollData[0][5],
          jobBand: enrollData[0][6],
          function: enrollData[0][7],
          subFunction: enrollData[0][8],
          status: enrollData[0][10]
        };
      }
      
      // Get TM1 data for cost calculations
      const tm1Data = tm1Sheet && tm1Sheet.getLastRow() > 1
        ? tm1Sheet.getRange(2, 1, tm1Sheet.getLastRow() - 1, 6).getValues()
        : [];
      debug.totalTM1Rows = tm1Data.length;
      
      // Filter TM1 data
      console.log('=== TM1 FILTERING ===');
      console.log('TM1 filter params: year=' + year + ', quarter=' + quarter + ', month=' + month + ', entityFilter=' + entityFilter);
      console.log('Total TM1 rows before filter:', tm1Data.length);
      
      const filteredTM1 = tm1Data.filter(row => {
        if (!row || !Array.isArray(row) || row.length < 6) return false;
        
        const rowYear = String(row[0] || '').trim();
        const rowQuarter = String(row[1] || '').trim();
        const rowMonth = String(row[2] || '').trim();
        const rowEntity = String(row[3] || '').trim();
        
        // Access filtering - skip if user has "All" access
        if (!hasAllEntities && userEntities.length > 0 && rowEntity && !userEntities.includes(rowEntity)) {
          return false;
        }
        
        // Year filter - ONLY filter if year is set, otherwise include all years
        if (year && rowYear !== year) return false;
        if (quarter && rowQuarter !== quarter) return false;
        if (month) {
          const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
          const monthName = monthNames[parseInt(month) - 1];
          if (rowMonth !== monthName) return false;
        }
        if (entityFilter && rowEntity !== entityFilter) return false;
        
        return true;
      });
      
      console.log('Filtered TM1 rows:', filteredTM1.length);
      debug.filteredTM1Rows = filteredTM1.length;
      
      // Sample TM1 for debugging
      if (tm1Data.length > 0 && tm1Data[0] && tm1Data[0].length > 5) {
        debug.sampleTM1 = {
          year: tm1Data[0][0],
          quarter: tm1Data[0][1],
          month: tm1Data[0][2],
          entity: tm1Data[0][3],
          ytg: tm1Data[0][4],
          actual: tm1Data[0][5]
        };
      }
      
      // ============================================
      // CALCULATE KPIs
      // ============================================
      
      // Total Learning Hours (sessions × duration × attended count, or just session duration sum)
      let totalHours = 0;
      const sessionDurations = new Map();
      
      filteredSessions.forEach(s => {
        const sessId = String(s[1]).trim();
        const duration = parseFloat(s[6]) || 0; // Column G - Duration
        sessionDurations.set(sessId, duration);
      });
      
      // Calculate hours from enrollments (attended)
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const status = String(e[10] || '').trim(); // Column K - Status
        if (status === 'Attended') {
          totalHours += sessionDurations.get(sessId) || 0;
        }
      });
      
      // Unique learners - only count those with "Attended" status
      const uniqueLearnerIds = new Set();
      filteredEnrollments.forEach(e => {
        const empId = String(e[2]).trim();
        const status = String(e[10] || '').trim();
        if (empId && status === 'Attended') uniqueLearnerIds.add(empId);
      });
      const uniqueLearners = uniqueLearnerIds.size;
      
      // Avg hours per employee
      const avgHoursPerEmp = uniqueLearners > 0 ? totalHours / uniqueLearners : 0;
      
      // Completion rate
      let openCount = 0, completedCount = 0;
      filteredSessions.forEach(s => {
        const status = String(s[3] || '').trim().toLowerCase();
        if (status === 'completed') completedCount++;
        else if (status === 'open') openCount++;
      });
      const totalStatusCount = openCount + completedCount;
      const completionRate = totalStatusCount > 0 ? (completedCount / totalStatusCount) * 100 : 0;
      
      // Cost calculations from TM1
      // Budget Utilization = SUM(Actual) / SUM(YTG) × 100
      // Both sums are for the filtered period (Yearly/Quarterly/Monthly)
      let totalCost = 0;
      let totalBudget = 0;
      
      // YTG and Actual by Entity - sum all rows in the filtered period
      const ytgByEntity = new Map();
      const spendByEntityMap = new Map();
      
      filteredTM1.forEach(row => {
        const entity = String(row[3]).trim();
        const ytgValue = parseFloat(row[4]) || 0;
        const actualValue = parseFloat(row[5]) || 0;
        
        totalCost += actualValue;
        totalBudget += ytgValue;
        
        // Sum YTG and Actual per entity
        ytgByEntity.set(entity, (ytgByEntity.get(entity) || 0) + ytgValue);
        spendByEntityMap.set(entity, (spendByEntityMap.get(entity) || 0) + actualValue);
      });
      
      console.log('=== BUDGET CALCULATION DEBUG ===');
      console.log('Period filters: year=' + year + ', quarter=' + quarter + ', month=' + month);
      console.log('Filtered TM1 rows: ' + filteredTM1.length);
      console.log('Entities found: ' + ytgByEntity.size);
      
      ytgByEntity.forEach((ytg, entity) => {
        const spent = spendByEntityMap.get(entity) || 0;
        const util = ytg > 0 ? (spent / ytg * 100) : 0;
        console.log(`  ${entity}: YTG=${ytg.toFixed(2)}, Spent=${spent.toFixed(2)}, Util=${util.toFixed(2)}%`);
      });
      
      const avgCostPerEmp = uniqueLearners > 0 ? totalCost / uniqueLearners : 0;
      const budgetUtilization = totalBudget > 0 ? (totalCost / totalBudget) * 100 : 0;
      
      console.log('Total Cost: ' + totalCost.toFixed(2) + ', Total Budget: ' + totalBudget.toFixed(2));
      console.log('Budget Utilization KPI: ' + budgetUtilization.toFixed(1) + '%');
      
      // Avg satisfaction score from sessions
      let totalScore = 0, scoreCount = 0;
      filteredSessions.forEach(s => {
        const score = parseFloat(s[8]) || 0; // Column I - Satisfaction Score
        if (score > 0) {
          totalScore += score;
          scoreCount++;
        }
      });
      const avgSatisfaction = scoreCount > 0 ? totalScore / scoreCount : 0;
      
      // ============================================
      // CALCULATE CHART DATA
      // ============================================
      
      // Create session lookup map for efficient access
      const sessionMap = new Map();
      filteredSessions.forEach(s => {
        const sessId = String(s[1]).trim();
        sessionMap.set(sessId, s);
      });
      
      console.log('Chart data: filteredSessions=' + filteredSessions.length + ', sessionMap size=' + sessionMap.size);
      console.log('Chart data: filteredEnrollments=' + filteredEnrollments.length);
      
      // Count attended for debugging
      const attendedCount = filteredEnrollments.filter(e => String(e[10] || '').trim() === 'Attended').length;
      console.log('Chart data: attendedEnrollments=' + attendedCount);
      
      // Hours by Category (from Program)
      const hoursByCategory = {};
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const status = String(e[10] || '').trim();
        if (status === 'Attended') {
          const session = sessionMap.get(sessId);
          if (session) {
            const progId = String(session[0]).trim();
            const prog = programMap.get(progId);
            const category = prog?.category || 'Uncategorized';
            const hours = sessionDurations.get(sessId) || 0;
            hoursByCategory[category] = (hoursByCategory[category] || 0) + hours;
          }
        }
      });
      console.log('hoursByCategory:', JSON.stringify(hoursByCategory));
      
      // Hours by Channel (Type from Session)
      const hoursByChannel = {};
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const status = String(e[10] || '').trim();
        if (status === 'Attended') {
          const session = sessionMap.get(sessId);
          if (session) {
            const type = String(session[4] || 'Unknown').trim(); // Column E - Type
            const hours = sessionDurations.get(sessId) || 0;
            hoursByChannel[type] = (hoursByChannel[type] || 0) + hours;
          }
        }
      });
      console.log('hoursByChannel:', JSON.stringify(hoursByChannel));
      
      // Hours by Entity (from Enrollment)
      const hoursByEntity = {};
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const status = String(e[10] || '').trim();
        const entity = String(e[5] || 'Unknown').trim(); // Column F - Entity
        if (status === 'Attended') {
          const hours = sessionDurations.get(sessId) || 0;
          hoursByEntity[entity] = (hoursByEntity[entity] || 0) + hours;
        }
      });
      
      // Hours by Function (from Enrollment)
      const hoursByFunction = {};
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const status = String(e[10] || '').trim();
        const func = String(e[7] || 'Unknown').trim(); // Column H - Function
        if (status === 'Attended') {
          const hours = sessionDurations.get(sessId) || 0;
          hoursByFunction[func] = (hoursByFunction[func] || 0) + hours;
        }
      });
      
      // Completion % by Category
      const completionByCategory = {};
      const categoryStats = {};
      filteredSessions.forEach(s => {
        const progId = String(s[0]).trim();
        const prog = programMap.get(progId);
        const category = prog?.category || 'Uncategorized';
        const status = String(s[3] || '').trim().toLowerCase();
        
        if (!categoryStats[category]) categoryStats[category] = { open: 0, completed: 0 };
        if (status === 'completed') categoryStats[category].completed++;
        else if (status === 'open') categoryStats[category].open++;
      });
      Object.keys(categoryStats).forEach(cat => {
        const total = categoryStats[cat].open + categoryStats[cat].completed;
        completionByCategory[cat] = total > 0 ? (categoryStats[cat].completed / total) * 100 : 0;
      });
      
      // Completion % by Entity (from enrollments' entity)
      const completionByEntity = {};
      const entitySessionStatus = {};
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const entity = String(e[5] || 'Unknown').trim();
        const session = sessionMap.get(sessId);
        if (session) {
          const status = String(session[3] || '').trim().toLowerCase();
          if (!entitySessionStatus[entity]) entitySessionStatus[entity] = new Map();
          entitySessionStatus[entity].set(sessId, status);
        }
      });
      Object.keys(entitySessionStatus).forEach(entity => {
        const sessions = entitySessionStatus[entity];
        let open = 0, completed = 0;
        sessions.forEach(status => {
          if (status === 'completed') completed++;
          else if (status === 'open') open++;
        });
        const total = open + completed;
        completionByEntity[entity] = total > 0 ? (completed / total) * 100 : 0;
      });
      
      // Spend by Entity (from TM1)
      const spendByEntity = {};
      console.log('Building spendByEntity from', filteredTM1.length, 'TM1 rows');
      
      // Log first few TM1 rows for debugging
      if (filteredTM1.length > 0) {
        console.log('First TM1 row:', JSON.stringify({
          year: filteredTM1[0][0],
          quarter: filteredTM1[0][1],
          month: filteredTM1[0][2],
          entity: filteredTM1[0][3],
          ytg: filteredTM1[0][4],
          actual: filteredTM1[0][5]
        }));
      }
      
      filteredTM1.forEach((row, idx) => {
        const entity = String(row[3] || '').trim();
        const actual = parseFloat(row[5]) || 0;
        if (entity) {
          spendByEntity[entity] = (spendByEntity[entity] || 0) + actual;
        }
        if (idx < 3) {
          console.log(`TM1 row ${idx}: entity="${entity}", actual=${actual}`);
        }
      });
      
      console.log('spendByEntity result:', JSON.stringify(spendByEntity));
      console.log('spendByEntity keys:', Object.keys(spendByEntity).length);
      
      // Budget Utilization by Entity
      // Use SUM(YTG) for the filtered period as budget
      console.log('=== BUDGET UTILIZATION BY ENTITY CHART ===');
      console.log('ytgByEntity entries:', ytgByEntity.size);
      console.log('spendByEntity entries:', Object.keys(spendByEntity).length);
      
      const budgetByEntity = {};
      ytgByEntity.forEach((budget, entity) => {
        const spent = spendByEntity[entity] || 0;
        const utilization = budget > 0 ? (spent / budget) * 100 : 0;
        budgetByEntity[entity] = utilization;
        console.log(`  ${entity}: spent=${spent.toFixed(2)}, budget=${budget.toFixed(2)}, util=${utilization.toFixed(2)}%`);
      });
      
      console.log('budgetByEntity result:', JSON.stringify(budgetByEntity));
      
      // ============================================
      // INDIVIDUAL DEVELOPMENT DATA
      // ============================================
      
      const indivDev = [];
      let indivDevChecked = 0;
      let matchingEnrollments = 0;
      
      // Fetch exchange rates for USD conversion
      const exchangeRatesMap = getExchangeRatesMap_();
      console.log('Exchange rates map size:', exchangeRatesMap.size);
      
      console.log('Checking IndivDev from', filteredEnrollments.length, 'enrollments');
      console.log('Filtered sessions count:', filteredSessions.length);
      
      // Log first session's IsIndivDev value
      if (filteredSessions.length > 0) {
        const firstSession = filteredSessions[0];
        console.log('First filtered session:', {
          sessionId: firstSession[1],
          sessionName: firstSession[2],
          isIndivDevRaw: firstSession[10],
          isIndivDevType: typeof firstSession[10]
        });
      }
      
      filteredEnrollments.forEach(e => {
        const sessId = String(e[1]).trim();
        const session = sessionMap.get(sessId);
        if (session) {
          matchingEnrollments++;
          const isIndivRaw = session[10];
          const isIndiv = String(isIndivRaw || '').trim().toLowerCase();
          indivDevChecked++;
          
          if (matchingEnrollments <= 3) {
            console.log(`Enrollment ${matchingEnrollments}: sessId=${sessId}, isIndivRaw="${isIndivRaw}", isIndiv="${isIndiv}"`);
          }
          
          if (isIndiv === 'yes' || isIndiv === 'true' || isIndivRaw === true) {
            const progId = String(session[0]).trim();
            const prog = programMap.get(progId);
            
            // Format date
            let dateStr = '';
            const dateVal = session[5];
            if (dateVal instanceof Date) {
              dateStr = Utilities.formatDate(dateVal, CONFIG.TIMEZONE, 'yyyy-MM-dd');
            } else if (dateVal) {
              dateStr = String(dateVal);
            }
            
            // Get cost data
            const currency = String(e[11] || '').trim();
            const costLocal = parseFloat(e[12]) || 0;
            
            // Convert to USD using event month rate
            const costUSD = convertToUSD_(costLocal, currency, dateStr, exchangeRatesMap);
            
            indivDev.push({
              entity: String(e[5] || '').trim(),
              employeeId: String(e[2] || '').trim(),
              name: String(e[3] || '').trim(),
              program: prog?.name || progId,
              session: String(session[2] || '').trim(),
              date: dateStr,
              currency: currency,
              cost: costLocal,
              costUSD: costUSD
            });
          }
        }
      });
      
      console.log('IndivDev: checked', indivDevChecked, 'sessions, found', indivDev.length, 'records');
      console.log('Matching enrollments:', matchingEnrollments);
      
      // ============================================
      // GET AVAILABLE FILTER OPTIONS
      // ============================================
      
      const availableYears = [...new Set(sessData.map(s => {
        const d = s[5];
        if (d instanceof Date) return d.getFullYear();
        if (d) return new Date(d).getFullYear();
        return null;
      }).filter(y => y))].sort((a, b) => b - a);
      
      // Entity filter from Enrollments only (participant entity)
      const availableEntities = [...new Set(enrollData.map(e => String(e[5] || '').trim()).filter(e => e))].sort();
      
      const availableFunctions = [...new Set(enrollData.map(e => String(e[7] || '').trim()).filter(f => f))].sort();
      const availableSubFunctions = [...new Set(enrollData.map(e => String(e[8] || '').trim()).filter(sf => sf))].sort();
      const availableJobBands = [...new Set(enrollData.map(e => String(e[6] || '').trim()).filter(jb => jb))].sort();
      
      // Helper to convert object to chart format
      const toChartData = (obj) => ({
        labels: Object.keys(obj),
        values: Object.values(obj)
      });
      
      console.log('getAnalyticsData SUCCESS - returning data');
      console.log('KPIs:', JSON.stringify({ totalHours, uniqueLearners, avgSatisfaction }));
      
      // Don't include debug in return - can cause serialization issues with Date objects
      console.log('=== ANALYTICS RETURN DATA ===');
      console.log('KPIs: totalHours=' + totalHours + ', uniqueLearners=' + uniqueLearners + ', completionRate=' + completionRate.toFixed(1));
      console.log('Charts hoursByCategory keys:', Object.keys(hoursByCategory).length);
      console.log('Charts hoursByChannel keys:', Object.keys(hoursByChannel).length);
      console.log('Charts hoursByEntity keys:', Object.keys(hoursByEntity).length);
      console.log('Charts hoursByFunction keys:', Object.keys(hoursByFunction).length);
      console.log('Charts completionByCategory keys:', Object.keys(completionByCategory).length);
      console.log('Charts completionByEntity keys:', Object.keys(completionByEntity).length);
      console.log('Charts spendByEntity keys:', Object.keys(spendByEntity).length);
      console.log('Charts budgetByEntity keys:', Object.keys(budgetByEntity).length);
      
      return {
        kpis: {
          totalHours,
          avgHoursPerEmp,
          completionRate,
          uniqueLearners,
          totalCost,
          avgCostPerEmp,
          budgetUtilization,
          avgSatisfaction
        },
        charts: {
          hoursByCategory: toChartData(hoursByCategory),
          hoursByChannel: toChartData(hoursByChannel),
          hoursByEntity: toChartData(hoursByEntity),
          hoursByFunction: toChartData(hoursByFunction),
          completionByCategory: toChartData(completionByCategory),
          completionByEntity: toChartData(completionByEntity),
          spendByEntity: toChartData(spendByEntity),
          budgetByEntity: toChartData(budgetByEntity)
        },
        indivDev,
        availableYears,
        availableEntities,
        availableFunctions,
        availableSubFunctions,
        availableJobBands,
        availablePrograms: allProgramNames.sort()
      };
      
  } catch (e) {
    console.error('getAnalyticsData error:', e);
    return {
      kpis: { totalHours: 0, avgHoursPerEmp: 0, completionRate: 0, uniqueLearners: 0, totalCost: 0, avgCostPerEmp: 0, budgetUtilization: 0, avgSatisfaction: 0 },
      charts: {},
      indivDev: [],
      availableYears: [],
      availableEntities: [],
      availableFunctions: [],
      availableSubFunctions: [],
      availableJobBands: [],
      error: e.message,
      debug: { error: e.message, stack: e.stack }
    };
  }
}

/**
 * Helper function to insert company logo into a Google Doc
 * @param {Body} body - The Google Doc body
 * @returns {void}
 */
function insertPDFLogo_(body) {
  try {
    // Extract base64 data from data URI
    const base64Data = COMPANY_LOGO.split(',')[1];
    const imageBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'logo.jpg');
    
    // Insert logo at position 0 (beginning of document)
    const logo = body.insertImage(0, imageBlob);
    
    // Set logo size (width in points, height proportional)
    const originalWidth = logo.getWidth();
    const originalHeight = logo.getHeight();
    const targetWidth = 120; // 120 points ≈ 1.67 inches
    const targetHeight = (originalHeight / originalWidth) * targetWidth;
    
    logo.setWidth(targetWidth);
    logo.setHeight(targetHeight);
    
    // Get the parent paragraph and center it
    const parent = logo.getParent();
    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      parent.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
    
    // Add a blank line after logo
    body.insertParagraph(1, '');
  } catch (e) {
    console.log('Could not insert logo into PDF:', e.message);
    // Continue without logo if it fails
  }
}

/**
 * Export Individual Development records to PDF
 */
function exportIndivDevPDF(sessionToken, filters) {
  return withSession_(sessionToken, (userData) => {
    try {
      console.log('exportIndivDevPDF called with filters:', JSON.stringify(filters));
      
      // Get analytics data with filters
      const data = getAnalyticsData(sessionToken, filters);
      
      if (!data) {
        console.error('getAnalyticsData returned null');
        return { error: 'Failed to get analytics data' };
      }
      
      const indivDev = data.indivDev || [];
      console.log('IndivDev records found:', indivDev.length);
      
      if (indivDev.length === 0) {
        return { error: 'No individual development records to export. Check that sessions have IsIndivDev=Yes.' };
      }
      
      // Create Google Doc for PDF
      const docName = 'Individual Development Report - ' + new Date().toISOString().split('T')[0];
      console.log('Creating document:', docName);
      const doc = DocumentApp.create(docName);
      const body = doc.getBody();
      
      // Title
      insertPDFLogo_(body); // Add company logo
      body.appendParagraph('Individual Development Report')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      // Filter info
      const filterParts = [];
      if (filters.year) filterParts.push('Year: ' + filters.year);
      if (filters.quarter) filterParts.push('Quarter: ' + filters.quarter);
      if (filters.month) filterParts.push('Month: ' + filters.month);
      if (filters.entity) filterParts.push('Entity: ' + filters.entity);
      const filterText = filterParts.length > 0 ? filterParts.join(' | ') : 'All Data';
      
      body.appendParagraph(filterText)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setForegroundColor('#666666');
      
      body.appendParagraph('Generated: ' + new Date().toLocaleString())
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setForegroundColor('#999999');
      
      body.appendParagraph('');
      
      // Create table
      const tableData = [['Entity', 'Employee', 'Program', 'Session', 'Date', 'Currency', 'Cost', 'Cost (USD)']];
      indivDev.forEach(row => {
        tableData.push([
          row.entity || '-',
          row.name || '-',
          row.program || '-',
          row.session || '-',
          row.date || '-',
          row.currency || '-',
          row.cost ? row.cost.toLocaleString() : '-',
          row.costUSD != null ? '$' + row.costUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '-'
        ]);
      });
      
      const table = body.appendTable(tableData);
      
      // Style header row
      const headerRow = table.getRow(0);
      for (let i = 0; i < headerRow.getNumCells(); i++) {
        headerRow.getCell(i)
          .setBackgroundColor('#3b82f6')
          .getChild(0).asText().setBold(true).setForegroundColor('#ffffff');
      }
      
      // Summary
      body.appendParagraph('');
      body.appendParagraph('Total Records: ' + indivDev.length)
        .setBold(true);
      
      const totalCost = indivDev.reduce((sum, r) => sum + (r.cost || 0), 0);
      body.appendParagraph('Total Cost: ' + totalCost.toLocaleString());
      
      const totalCostUSD = indivDev.reduce((sum, r) => sum + (r.costUSD || 0), 0);
      body.appendParagraph('Total Cost (USD): $' + totalCostUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}));
      
      doc.saveAndClose();
      console.log('Document saved, converting to PDF');
      
      // Convert to PDF
      const docFile = DriveApp.getFileById(doc.getId());
      const pdfBlob = docFile.getAs('application/pdf');
      
      // Convert to base64 for direct download (no access issues)
      const base64 = Utilities.base64Encode(pdfBlob.getBytes());
      const filename = 'IndivDev_Report_' + new Date().toISOString().split('T')[0] + '.pdf';
      
      // Delete the temp doc
      docFile.setTrashed(true);
      
      console.log('PDF generated successfully, size:', base64.length);
      
      return { success: true, base64: base64, filename: filename };
      
    } catch (e) {
      console.error('exportIndivDevPDF error:', e);
      console.error('Error stack:', e.stack);
      return { success: false, error: 'Failed to generate PDF: ' + e.message };
    }
  });
}

/**
 * Export Analytics Dashboard to PDF
 */
function exportAnalyticsPDF(sessionToken, filters) {
  return withSession_(sessionToken, (userData) => {
    try {
      // Get analytics data with filters
      const data = getAnalyticsData(sessionToken, filters);
      
      if (!data || !data.kpis) {
        return { error: 'No analytics data available' };
      }
      
      const kpis = data.kpis;
      const charts = data.charts;
      
      // Format helpers
      const formatNum = (n, decimals = 1) => {
        if (n >= 1000000) return (n/1000000).toFixed(decimals) + 'M';
        if (n >= 1000) return (n/1000).toFixed(decimals) + 'K';
        return n.toFixed(decimals);
      };
      
      // Create Google Doc for PDF
      const doc = DocumentApp.create('Training Analytics Report - ' + new Date().toISOString().split('T')[0]);
      const body = doc.getBody();
      
      // Set page margins
      body.setMarginTop(36);
      body.setMarginBottom(36);
      body.setMarginLeft(36);
      body.setMarginRight(36);
      
      // Title
      insertPDFLogo_(body); // Add company logo
      body.appendParagraph('Training Analytics Report')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setForegroundColor('#1e40af');
      
      // Filter info
      const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const monthName = filters.month ? monthNames[parseInt(filters.month)] : '';
      const filterText = `Period: ${filters.year || 'All Time'} ${filters.quarter || ''} ${monthName} | Entity: ${filters.entity || 'All Entities'}`;
      body.appendParagraph(filterText)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setForegroundColor('#666666');
      
      body.appendParagraph('Generated: ' + new Date().toLocaleString())
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setForegroundColor('#999999')
        .setFontSize(9);
      
      body.appendParagraph('');
      
      // ============================================
      // KPI SUMMARY
      // ============================================
      
      body.appendParagraph('Key Performance Indicators')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setForegroundColor('#1e40af');
      
      const kpiData = [
        ['Metric', 'Value'],
        ['Total Learning Hours', formatNum(kpis.totalHours || 0)],
        ['Avg Hours / Employee', (kpis.avgHoursPerEmp || 0).toFixed(1)],
        ['Completion Rate', (kpis.completionRate || 0).toFixed(1) + '%'],
        ['Unique Learners', String(kpis.uniqueLearners || 0)],
        ['Total Learning Cost (USD)', formatNum(kpis.totalCost || 0)],
        ['Avg Cost / Employee (USD)', formatNum(kpis.avgCostPerEmp || 0)],
        ['Budget Utilization', (kpis.budgetUtilization || 0).toFixed(1) + '%'],
        ['Avg Training Satisfaction', (kpis.avgSatisfaction || 0).toFixed(1) + ' / 5']
      ];
      
      const kpiTable = body.appendTable(kpiData);
      kpiTable.setBorderWidth(1);
      kpiTable.setBorderColor('#e5e7eb');
      
      // Style KPI header
      const kpiHeader = kpiTable.getRow(0);
      for (let i = 0; i < kpiHeader.getNumCells(); i++) {
        kpiHeader.getCell(i)
          .setBackgroundColor('#3b82f6')
          .getChild(0).asText().setBold(true).setForegroundColor('#ffffff');
      }
      
      // Alternate row colors
      for (let i = 1; i < kpiTable.getNumRows(); i++) {
        if (i % 2 === 0) {
          for (let j = 0; j < kpiTable.getRow(i).getNumCells(); j++) {
            kpiTable.getRow(i).getCell(j).setBackgroundColor('#f8fafc');
          }
        }
      }
      
      body.appendParagraph('');
      
      // ============================================
      // LEARNING HOURS BY CATEGORY
      // ============================================
      
      if (charts.hoursByCategory && charts.hoursByCategory.labels.length > 0) {
        body.appendParagraph('Learning Hours by Category')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const catData = [['Category', 'Hours', '% of Total']];
        const totalCatHours = charts.hoursByCategory.values.reduce((a, b) => a + b, 0);
        
        charts.hoursByCategory.labels.forEach((label, i) => {
          const hours = charts.hoursByCategory.values[i];
          const pct = totalCatHours > 0 ? (hours / totalCatHours * 100).toFixed(1) + '%' : '0%';
          catData.push([label, hours.toFixed(1), pct]);
        });
        
        const catTable = body.appendTable(catData);
        styleReportTable_(catTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // LEARNING HOURS BY CHANNEL
      // ============================================
      
      if (charts.hoursByChannel && charts.hoursByChannel.labels.length > 0) {
        body.appendParagraph('Learning Hours by Channel')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const channelData = [['Channel', 'Hours', '% of Total']];
        const totalChannelHours = charts.hoursByChannel.values.reduce((a, b) => a + b, 0);
        
        charts.hoursByChannel.labels.forEach((label, i) => {
          const hours = charts.hoursByChannel.values[i];
          const pct = totalChannelHours > 0 ? (hours / totalChannelHours * 100).toFixed(1) + '%' : '0%';
          channelData.push([label, hours.toFixed(1), pct]);
        });
        
        const channelTable = body.appendTable(channelData);
        styleReportTable_(channelTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // COMPLETION BY CATEGORY
      // ============================================
      
      if (charts.completionByCategory && charts.completionByCategory.labels.length > 0) {
        body.appendParagraph('Completion Rate by Category')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const compCatData = [['Category', 'Completion %']];
        charts.completionByCategory.labels.forEach((label, i) => {
          compCatData.push([label, charts.completionByCategory.values[i].toFixed(1) + '%']);
        });
        
        const compCatTable = body.appendTable(compCatData);
        styleReportTable_(compCatTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // HOURS BY ENTITY
      // ============================================
      
      if (charts.hoursByEntity && charts.hoursByEntity.labels.length > 0) {
        body.appendParagraph('Learning Hours by Entity')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const entityHoursData = [['Entity', 'Hours']];
        charts.hoursByEntity.labels.forEach((label, i) => {
          entityHoursData.push([label, charts.hoursByEntity.values[i].toFixed(1)]);
        });
        
        const entityHoursTable = body.appendTable(entityHoursData);
        styleReportTable_(entityHoursTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // HOURS BY FUNCTION
      // ============================================
      
      if (charts.hoursByFunction && charts.hoursByFunction.labels.length > 0) {
        body.appendParagraph('Learning Hours by Function')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const funcData = [['Function', 'Hours']];
        // Sort by hours descending
        const sorted = charts.hoursByFunction.labels
          .map((l, i) => ({ label: l, value: charts.hoursByFunction.values[i] }))
          .sort((a, b) => b.value - a.value)
          .slice(0, 10); // Top 10
        
        sorted.forEach(item => {
          funcData.push([item.label, item.value.toFixed(1)]);
        });
        
        const funcTable = body.appendTable(funcData);
        styleReportTable_(funcTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // SPEND BY ENTITY
      // ============================================
      
      if (charts.spendByEntity && charts.spendByEntity.labels.length > 0) {
        body.appendParagraph('Training Spend by Entity (USD)')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const spendData = [['Entity', 'Spend (USD)']];
        charts.spendByEntity.labels.forEach((label, i) => {
          spendData.push([label, formatNum(charts.spendByEntity.values[i])]);
        });
        
        const spendTable = body.appendTable(spendData);
        styleReportTable_(spendTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // BUDGET UTILIZATION BY ENTITY
      // ============================================
      
      if (charts.budgetByEntity && charts.budgetByEntity.labels.length > 0) {
        body.appendParagraph('Budget Utilization by Entity')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        const budgetData = [['Entity', 'Utilization %']];
        charts.budgetByEntity.labels.forEach((label, i) => {
          budgetData.push([label, charts.budgetByEntity.values[i].toFixed(1) + '%']);
        });
        
        const budgetTable = body.appendTable(budgetData);
        styleReportTable_(budgetTable);
        body.appendParagraph('');
      }
      
      // ============================================
      // INDIVIDUAL DEVELOPMENT RECORDS
      // ============================================
      
      const indivDev = data.indivDev || [];
      if (indivDev.length > 0) {
        body.appendParagraph('Individual Development Records')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2)
          .setForegroundColor('#1e40af');
        
        body.appendParagraph('Total Records: ' + indivDev.length)
          .setForegroundColor('#666666')
          .setFontSize(10);
        
        const indivDevData = [['Entity', 'Employee', 'Program', 'Session', 'Date', 'Currency', 'Cost', 'Cost (USD)']];
        let totalCost = 0;
        let totalCostUSD = 0;
        
        indivDev.forEach(row => {
          indivDevData.push([
            row.entity || '-',
            row.name || '-',
            row.program || '-',
            row.session || '-',
            row.date || '-',
            row.currency || '-',
            row.cost ? row.cost.toLocaleString() : '-',
            row.costUSD != null ? '$' + row.costUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '-'
          ]);
          totalCost += row.cost || 0;
          totalCostUSD += row.costUSD || 0;
        });
        
        const indivDevTable = body.appendTable(indivDevData);
        styleReportTable_(indivDevTable);
        
        body.appendParagraph('Total Cost: ' + totalCost.toLocaleString())
          .setBold(true)
          .setFontSize(10);
        body.appendParagraph('Total Cost (USD): $' + totalCostUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}))
          .setBold(true)
          .setFontSize(10);
        body.appendParagraph('');
      }
      
      // ============================================
      // FOOTER
      // ============================================
      
      body.appendParagraph('');
      body.appendParagraph('— End of Report —')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setForegroundColor('#999999')
        .setFontSize(9);
      
      doc.saveAndClose();
      
      // Convert to PDF
      const docFile = DriveApp.getFileById(doc.getId());
      const pdfBlob = docFile.getAs('application/pdf');
      
      // Convert to base64 for direct download (no access issues)
      const base64 = Utilities.base64Encode(pdfBlob.getBytes());
      const filename = 'Analytics_Report_' + new Date().toISOString().split('T')[0] + '.pdf';
      
      // Delete the temp doc
      docFile.setTrashed(true);
      
      return { success: true, base64: base64, filename: filename };
      
    } catch (e) {
      console.error('exportAnalyticsPDF error:', e);
      return { success: false, error: e.message };
    }
  });
}

/**
 * Helper to style report tables
 */
function styleReportTable_(table) {
  table.setBorderWidth(1);
  table.setBorderColor('#e5e7eb');
  
  // Style header
  const header = table.getRow(0);
  for (let i = 0; i < header.getNumCells(); i++) {
    header.getCell(i)
      .setBackgroundColor('#10b981')
      .getChild(0).asText().setBold(true).setForegroundColor('#ffffff');
  }
  
  // Alternate row colors
  for (let i = 1; i < table.getNumRows(); i++) {
    if (i % 2 === 0) {
      for (let j = 0; j < table.getRow(i).getNumCells(); j++) {
        table.getRow(i).getCell(j).setBackgroundColor('#f0fdf4');
      }
    }
  }
}

// ============================================
// PDF EXPORT
// ============================================

// Generate Monthly Calendar PDF
function generateCalendarPDF(sessionToken, year, month) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      const programSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      
      const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      const monthName = monthNames[month - 1];
      
      // Build program map
      const programMap = new Map();
      if (programSheet && programSheet.getLastRow() > 1) {
        programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 2).getValues().forEach(r => {
          programMap.set(String(r[0]).trim(), r[1]);
        });
      }
      
      // Get sessions for the month
      let events = [];
      if (sessionSheet && sessionSheet.getLastRow() > 1) {
        sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 15).getValues().forEach(r => {
          const date = r[5];
          if (date instanceof Date && date.getFullYear() === year && date.getMonth() + 1 === month) {
            events.push({
              day: date.getDate(),
              name: r[2] || '',
              programName: programMap.get(String(r[0]).trim()) || '',
              status: r[3] || 'Open'
            });
          }
        });
      }
      
      // Group events by day
      const eventsByDay = {};
      events.forEach(e => {
        if (!eventsByDay[e.day]) eventsByDay[e.day] = [];
        eventsByDay[e.day].push(e);
      });
      
      // Generate HTML for PDF
      const firstDay = new Date(year, month - 1, 1);
      const lastDay = new Date(year, month, 0);
      const startDayOfWeek = firstDay.getDay();
      const daysInMonth = lastDay.getDate();
      
      let calendarRows = '';
      let dayCount = 1;
      let started = false;
      
      for (let week = 0; week < 6; week++) {
        if (dayCount > daysInMonth) break;
        let row = '<tr>';
        for (let d = 0; d < 7; d++) {
          if (week === 0 && d < startDayOfWeek) {
            row += '<td style="background:#f8f9fa;"></td>';
          } else if (dayCount <= daysInMonth) {
            const dayEvents = eventsByDay[dayCount] || [];
            let eventHtml = dayEvents.slice(0, 3).map(e => 
              `<div style="font-size:7px; padding:2px 4px; margin:1px 0; border-radius:2px; background:${e.status === 'Completed' ? '#dcfce7' : '#fef3c7'}; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${e.name}</div>`
            ).join('');
            if (dayEvents.length > 3) {
              eventHtml += `<div style="font-size:6px; color:#666;">+${dayEvents.length - 3} more</div>`;
            }
            row += `<td style="vertical-align:top; height:70px; border:1px solid #e5e7eb; padding:4px;">
              <div style="font-weight:600; font-size:10px; margin-bottom:2px;">${dayCount}</div>
              ${eventHtml}
            </td>`;
            dayCount++;
          } else {
            row += '<td style="background:#f8f9fa;"></td>';
          }
        }
        row += '</tr>';
        calendarRows += row;
      }
      
      const html = `
        <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .logo { text-align: center; margin-bottom: 15px; }
            .logo img { max-width: 120px; height: auto; }
            h1 { text-align: center; color: #1e40af; margin-bottom: 20px; font-size: 18px; }
            table { width: 100%; border-collapse: collapse; }
            th { background: #1e40af; color: white; padding: 8px; font-size: 9px; text-transform: uppercase; }
            td { width: 14.28%; }
          </style>
        </head>
        <body>
          <div class="logo"><img src="${COMPANY_LOGO}" alt="Acme Corp"></div>
          <h1>Training Calendar - ${monthName} ${year}</h1>
          <table>
            <tr>
              <th>Sun</th><th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th>
            </tr>
            ${calendarRows}
          </table>
        </body>
        </html>
      `;
      
      // Create PDF blob
      const blob = Utilities.newBlob(html, 'text/html', 'calendar.html');
      const pdf = blob.getAs('application/pdf');
      const base64 = Utilities.base64Encode(pdf.getBytes());
      
      return {
        success: true,
        base64: base64,
        filename: `Training_Calendar_${monthName}_${year}.pdf`
      };
    } catch (e) {
      console.error('generateCalendarPDF error:', e);
      return { success: false, message: e.message || 'Unknown error' };
    }
  });
}

function generatePDF(sessionToken, year, filterId, mode, entityFilter) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessionSheet = getSheet_(SHEET_NAMES.SESSIONS);
      const programSheet = getSheet_(SHEET_NAMES.PROGRAMS);
      const enrollSheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
      const empSheet = getSheet_(SHEET_NAMES.EMPLOYEES);
      
      // Get user's visibility access
      const userCountries = userData.countries || [];
      const canViewAll = canViewAllSessions_(userData);
      
      // Clean filter ID
      let cleanFilterId = filterId;
      if (filterId && filterId.includes(" - ")) {
        cleanFilterId = filterId.split(" - ")[0].trim();
      }
      
      // Build program map
      const programMap = new Map();
      if (programSheet.getLastRow() > 1) {
        programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 2).getValues().forEach(r => {
          programMap.set(String(r[0]).trim(), r[1]);
        });
      }
      
      // Build participant counts
      const participantCounts = new Map();
      if (enrollSheet.getLastRow() > 1) {
        const enrollIds = enrollSheet.getRange(2, 2, enrollSheet.getLastRow() - 1, 1).getValues().flat();
        enrollIds.forEach(id => {
          const sid = String(id).trim();
          if (sid) participantCounts.set(sid, (participantCounts.get(sid) || 0) + 1);
        });
      }
      
      let rowsToPrint = [];
      let mainHeadline = "";
      let subHeadline = "";
      let fileNameBase = "";
      let summaryStats = { totalHours: 0, totalPrograms: new Set(), attendedCount: 0 };
      
      if (mode === 'employee') {
        // Employee mode
        const empData = empSheet.getLastRow() > 1 ? empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 3).getValues() : [];
        const empObj = empData.find(r => String(r[1]) === String(cleanFilterId));
        const empName = empObj ? empObj[2] : cleanFilterId;
        
        mainHeadline = "Training History Report";
        subHeadline = empName;
        fileNameBase = empName;
        
        // Get all enrollments for this employee (including cost columns L=Currency, M=Cost)
        const allEnrolls = enrollSheet.getLastRow() > 1 ? enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 13).getValues() : [];
        const allSessions = sessionSheet.getLastRow() > 1 ? sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 14).getValues() : [];
        
        const sessionMap = new Map();
        allSessions.forEach(s => sessionMap.set(String(s[1]).trim(), s));
        
        const empEnrolls = allEnrolls.filter(r => String(r[2]) === String(cleanFilterId));
        
        // Collect IndivDev records separately
        const indivDevRecords = [];
        
        empEnrolls.forEach(enr => {
          const sessRow = sessionMap.get(String(enr[1]).trim());
          if (!sessRow) return;
          
          const date = sessRow[5];
          if (year !== "All" && date && new Date(date).getFullYear() != year) return;
          
          const status = String(enr[10] || "").trim();
          let hrs = parseFloat(sessRow[6]) || 0;
          const progName = programMap.get(String(sessRow[0]).trim()) || "Unknown";
          const isIndivDev = String(sessRow[10] || '').toLowerCase() === 'yes';
          const sessionEntity = String(sessRow[13] || '').trim();
          
          if (status === 'Attended' || status === 'Completed') {
            summaryStats.totalHours += hrs;
            summaryStats.attendedCount++;
            summaryStats.totalPrograms.add(progName);
          }
          
          rowsToPrint.push({
            date: date,
            program: progName,
            description: sessRow[2],
            duration: hrs,
            status: status,
            source: 'Local'
          });
          
          // If IndivDev, also collect for separate section
          if (isIndivDev) {
            indivDevRecords.push({
              entity: sessionEntity,
              program: progName,
              session: sessRow[2] || '',
              date: date,
              currency: String(enr[11] || '').trim(),
              cost: parseFloat(enr[12]) || 0
            });
          }
        });
        
        // Store IndivDev records for later use in HTML generation
        summaryStats.indivDevRecords = indivDevRecords;
      } else {
        // Program mode
        const progName = cleanFilterId ? (programMap.get(cleanFilterId) || "Unknown Program") : "All Programs";
        mainHeadline = "Training Calendar Report";
        const entitySuffix = entityFilter && entityFilter !== 'All' ? ` - ${entityFilter}` : '';
        subHeadline = `${progName} (${year === "All" ? "All History" : year})${entitySuffix}`;
        fileNameBase = cleanFilterId ? progName : "All_Programs";
        
        // Get sessions with Entity column (column N = index 13)
        const rawData = sessionSheet.getLastRow() > 1 ? sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 14).getValues() : [];
        
        rawData.forEach(row => {
          const date = row[5];
          if (!date || !(date instanceof Date)) return;
          if (year !== "All" && date.getFullYear() != year) return;
          if (cleanFilterId && String(row[0]).trim() !== String(cleanFilterId).trim()) return;
          
          // Get session entity (column N = index 13)
          const sessionEntity = String(row[13] || '').trim();
          
          // Filter by user's visibility access
          if (!canViewAll && sessionEntity && !userCountries.includes(sessionEntity) && sessionEntity !== 'Global') return;
          
          // Filter by selected entity if specified
          if (entityFilter && entityFilter !== 'All' && sessionEntity !== entityFilter) return;
          
          const sId = String(row[1]).trim();
          rowsToPrint.push({
            date: date,
            entity: sessionEntity || 'N/A',
            program: programMap.get(String(row[0]).trim()) || "Unknown",
            description: row[2],
            duration: parseFloat(row[6]) || 0,
            status: null,
            participantCount: participantCounts.get(sId) || 0
          });
        });
      }
      
      rowsToPrint.sort((a, b) => new Date(a.date) - new Date(b.date));
      
      if (rowsToPrint.length === 0) {
        return { success: false, message: 'No records found.' };
      }
      
      // Generate HTML
      const safeName = fileNameBase.replace(/[^a-zA-Z0-9-_]/g, '_');
      const safeYear = year === "All" ? "All_History" : year;
      const finalFileName = `Report_${safeName}_${safeYear}.pdf`;
      
      const fmtDur = (val) => {
        if (!val) return '0:00';
        const h = Math.floor(val);
        const m = Math.round((val - h) * 60);
        return `${h}:${String(m).padStart(2, '0')}`;
      };
      
      let html = `<html><head><style>
        body { font-family: sans-serif; color: #333; padding: 30px; }
        .logo { text-align: center; margin-bottom: 20px; }
        .logo img { max-width: 120px; height: auto; }
        h1 { color: #0d6efd; border-bottom: 2px solid #0d6efd; padding-bottom: 10px; margin-top: 0; margin-bottom: 10px; font-size: 24px; text-align: center; }
        .sub-headline { font-size: 18px; font-weight: bold; color: #444; margin-top: 0; margin-bottom: 25px; text-align: center; }
        h2 { background-color: #f1f3f4; padding: 8px; border-left: 5px solid #34a853; margin-top: 20px; font-size: 16px; }
        .summary-box { background: #e8f0fe; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: flex; gap: 30px; justify-content: center; }
        .stat { font-weight: bold; color: #1967d2; font-size: 1.1em; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
        th { background-color: #e0e0e0; text-align: left; padding: 8px; border: 1px solid #ccc; }
        td { border: 1px solid #ccc; padding: 8px; }
      </style></head><body>
      <div class="logo"><img src="${COMPANY_LOGO}" alt="Acme Corp"></div>
      <h1>${mainHeadline}</h1>
      <div class="sub-headline">${subHeadline}</div>`;
      
      if (mode === 'employee') {
        html += `<div class="summary-box">
          <div>Total Attended: <span class="stat">${summaryStats.attendedCount}</span></div>
          <div>Total Hours: <span class="stat">${fmtDur(summaryStats.totalHours)}</span></div>
          <div>Unique Programs: <span class="stat">${summaryStats.totalPrograms.size}</span></div>
        </div>`;
      }
      
      // Group by month
      const months = {};
      rowsToPrint.forEach(row => {
        const m = new Date(row.date).toLocaleString('default', { month: 'long', year: 'numeric' });
        if (!months[m]) months[m] = [];
        months[m].push(row);
      });
      
      for (const [monthName, items] of Object.entries(months)) {
        html += `<h2>${monthName}</h2>`;
        if (mode === 'employee') {
          html += `<table><thead><tr>
            <th width="15%">Date</th>
            <th width="25%">Program</th>
            <th width="30%">Session</th>
            <th width="15%">Status</th>
            <th width="15%">Duration</th>
          </tr></thead><tbody>`;
        } else {
          html += `<table><thead><tr>
            <th width="10%">Entity</th>
            <th width="12%">Date</th>
            <th width="25%">Program</th>
            <th width="30%">Description</th>
            <th width="10%">Duration</th>
            <th width="10%">Total Pax</th>
          </tr></thead><tbody>`;
        }
        
        items.forEach(r => {
          html += `<tr>
            ${mode !== 'employee' ? `<td>${r.entity || ''}</td>` : ''}
            <td>${new Date(r.date).toLocaleDateString()}</td>
            <td>${r.program}</td>
            <td>${r.description || ''}</td>
            ${mode === 'employee' ? `<td>${r.status}</td>` : ''}
            <td>${fmtDur(r.duration)}</td>
            ${mode !== 'employee' ? `<td style="text-align:center">${r.participantCount}</td>` : ''}
          </tr>`;
        });
        html += `</tbody></table>`;
      }
      
      // Add IndivDev section for employee mode
      if (mode === 'employee' && summaryStats.indivDevRecords && summaryStats.indivDevRecords.length > 0) {
        const indivDev = summaryStats.indivDevRecords;
        const totalCost = indivDev.reduce((sum, r) => sum + r.cost, 0);
        const currencies = [...new Set(indivDev.map(r => r.currency).filter(c => c))];
        
        // Get exchange rates for USD conversion
        const exchangeRates = getExchangeRatesMap_();
        
        // Calculate USD for each record and total
        let totalCostUSD = 0;
        const indivDevWithUSD = indivDev.map(r => {
          // Convert date to ISO string format for exchange rate lookup
          const dateForLookup = r.date instanceof Date 
            ? r.date.toISOString().split('T')[0] 
            : (r.date ? String(r.date) : '');
          const costUSD = convertToUSD_(r.cost, r.currency, dateForLookup, exchangeRates);
          totalCostUSD += costUSD || 0;
          return { ...r, costUSD };
        });
        
        html += `<h2 style="margin-top: 30px; background-color: #fff3e0; border-left: 5px solid #ff9800;">Individual Development Records</h2>`;
        html += `<table><thead><tr>
          <th width="12%">Date</th>
          <th width="20%">Program</th>
          <th width="25%">Session</th>
          <th width="10%">Currency</th>
          <th width="15%">Cost</th>
          <th width="18%">Cost (USD)</th>
        </tr></thead><tbody>`;
        
        indivDevWithUSD.forEach(r => {
          const dateStr = r.date instanceof Date 
            ? r.date.toLocaleDateString() 
            : (r.date ? new Date(r.date).toLocaleDateString() : '');
          html += `<tr>
            <td>${dateStr}</td>
            <td>${r.program || '-'}</td>
            <td>${r.session || '-'}</td>
            <td>${r.currency || '-'}</td>
            <td style="text-align:right;">${r.cost ? r.cost.toLocaleString() : '-'}</td>
            <td style="text-align:right;">${r.costUSD != null ? '$' + r.costUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '-'}</td>
          </tr>`;
        });
        
        html += `</tbody></table>`;
        
        // Total cost summary
        html += `<div style="margin-top: 15px; padding: 15px; background: #fff3e0; border-radius: 8px; border-left: 5px solid #ff9800;">
          <strong>Total Individual Development Records:</strong> ${indivDev.length}<br>
          <strong>Total Cost:</strong> ${currencies.length === 1 ? currencies[0] + ' ' : ''}${totalCost.toLocaleString()}
          ${currencies.length > 1 ? '<br><em style="font-size:11px; color:#666;">Note: Multiple currencies used in records</em>' : ''}<br>
          <strong>Total Cost (USD):</strong> $${totalCostUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
        </div>`;
      }
      
      html += `</body></html>`;
      
      const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
      const base64 = Utilities.base64Encode(blob.getBytes());
      
      return { success: true, base64: base64, filename: finalFileName };
    } catch (e) {
      console.error('generatePDF error:', e);
      return { success: false, message: e.message };
    }
  });
}

function getEmployeesForExport(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const employees = getEmployees(sessionToken);
      return employees
        .map(e => ({ id: e.id, name: e.name }))
        .sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    } catch (e) {
      return [];
    }
  });
}

function getProgramsForExport(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const result = getPrograms(sessionToken);
      const programs = result.programs || result || [];
      return programs
        .map(p => ({ id: p.id, name: p.name }))
        .sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    } catch (e) {
      return [];
    }
  });
}

/**
 * Get entities for export filter dropdown
 * Returns only entities the user has access to based on their visibility permissions
 */
function getEntitiesForExport(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const userCountries = userData.countries || [];
      const canViewAll = canViewAllSessions_(userData);
      
      // Get unique entities from Sessions sheet (column N)
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      if (!sessionSheet || sessionSheet.getLastRow() < 2) return [];
      
      const entityData = sessionSheet.getRange(2, 14, sessionSheet.getLastRow() - 1, 1).getValues();
      const uniqueEntities = new Set();
      
      entityData.forEach(row => {
        const entity = String(row[0] || '').trim();
        if (entity) {
          // Include entities based on visibility access
          if (canViewAll || userCountries.includes(entity) || entity === 'Global') {
            uniqueEntities.add(entity);
          }
        }
      });
      
      return Array.from(uniqueEntities).sort();
    } catch (e) {
      console.error('getEntitiesForExport error:', e);
      return [];
    }
  });
}

/**
 * Get sessions for a program (for Session Report dropdown)
 * Returns sessions sorted by date (latest first)
 */
function getSessionsForExport(sessionToken, programId) {
  return withSession_(sessionToken, (userData) => {
    try {
      if (!programId) return [];
      
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      if (!sessionSheet || sessionSheet.getLastRow() < 2) return [];
      
      // Read sessions: A=ProgramID, B=SessionID, C=Name, F=Date, N=Entity
      const data = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 14).getValues();
      
      const userCountries = userData.countries || [];
      const canViewAll = canViewAllSessions_(userData);
      
      const sessions = [];
      data.forEach(row => {
        const progId = String(row[0]).trim();
        if (progId !== programId) return;
        
        const sessionEntity = String(row[13] || '').trim();
        
        // Check visibility
        if (!canViewAll && sessionEntity !== 'Global' && !userCountries.includes(sessionEntity)) {
          return;
        }
        
        const dateVal = row[5];
        let dateStr = '';
        if (dateVal instanceof Date) {
          dateStr = Utilities.formatDate(dateVal, CONFIG.TIMEZONE, 'MMM dd, yyyy');
        } else if (dateVal) {
          dateStr = String(dateVal);
        }
        
        sessions.push({
          sessionId: String(row[1]).trim(),
          name: row[2] || '',
          date: dateStr,
          dateObj: dateVal instanceof Date ? dateVal : new Date(dateVal || 0)
        });
      });
      
      // Sort by date (latest first)
      sessions.sort((a, b) => b.dateObj - a.dateObj);
      
      return sessions.map(s => ({
        sessionId: s.sessionId,
        name: s.name,
        date: s.date
      }));
    } catch (e) {
      console.error('getSessionsForExport error:', e);
      return [];
    }
  });
}

/**
 * Generate Session Report PDF
 * Shows session details and participant list with status
 */
function generateSessionPDF(sessionToken, sessionId) {
  return withSession_(sessionToken, (userData) => {
    try {
      if (!sessionId) {
        return { success: false, message: 'Session ID required' };
      }
      
      const sessionSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
      const programSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
      const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
      const empSheet = getSheetSafe_(SHEET_NAMES.EMPLOYEES);
      
      // Find session
      const sessData = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, 15).getValues();
      const sessionRow = sessData.find(r => String(r[1]).trim() === sessionId);
      
      if (!sessionRow) {
        return { success: false, message: 'Session not found' };
      }
      
      // Session details
      const programId = String(sessionRow[0]).trim();
      const sessionName = sessionRow[2] || '';
      const sessionStatus = sessionRow[3] || '';
      const sessionType = sessionRow[4] || '';
      const sessionDate = sessionRow[5];
      const duration = parseFloat(sessionRow[6]) || 0;
      const location = sessionRow[7] || '';
      const isIndivDev = String(sessionRow[10] || '').toLowerCase() === 'yes';
      const sessionEntity = String(sessionRow[13] || '').trim();
      
      // Get program name
      let programName = programId;
      if (programSheet && programSheet.getLastRow() > 1) {
        const progData = programSheet.getRange(2, 1, programSheet.getLastRow() - 1, 2).getValues();
        const progRow = progData.find(r => String(r[0]).trim() === programId);
        if (progRow) programName = progRow[1] || programId;
      }
      
      // Format date
      let dateStr = '';
      if (sessionDate instanceof Date) {
        dateStr = Utilities.formatDate(sessionDate, CONFIG.TIMEZONE, 'MMMM dd, yyyy');
      } else if (sessionDate) {
        dateStr = String(sessionDate);
      }
      
      // Build employee info map
      const empInfoMap = new Map();
      if (empSheet && empSheet.getLastRow() > 1) {
        empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 8).getValues().forEach(r => {
          const empId = String(r[1]).trim();
          if (empId) {
            empInfoMap.set(empId, {
              entity: r[0] || '',
              name: r[2] || '',
              jobBand: r[3] || ''
            });
          }
        });
      }
      
      // Get enrollments for this session
      // Columns: A=EnrollmentID, B=SessionID, C=EmployeeID, D=Name, K=Status, L=Currency, M=Cost
      const enrollData = enrollSheet && enrollSheet.getLastRow() > 1 
        ? enrollSheet.getRange(2, 1, enrollSheet.getLastRow() - 1, 13).getValues()
        : [];
      
      const participants = [];
      const statusCounts = { Enrolled: 0, Attended: 0, Absent: 0, Other: 0 };
      
      enrollData.forEach(r => {
        if (String(r[1]).trim() !== sessionId) return;
        
        const empId = String(r[2]).trim();
        const status = String(r[10] || 'Enrolled').trim();
        const empInfo = empInfoMap.get(empId) || {};
        
        participants.push({
          employeeId: empId,
          name: empInfo.name || r[3] || '',
          entity: empInfo.entity || '',
          status: status,
          currency: isIndivDev ? String(r[11] || '').trim() : '',
          cost: isIndivDev ? (parseFloat(r[12]) || 0) : 0
        });
        
        // Count statuses
        if (status === 'Enrolled') statusCounts.Enrolled++;
        else if (status === 'Attended') statusCounts.Attended++;
        else if (status === 'Absent') statusCounts.Absent++;
        else statusCounts.Other++;
      });
      
      // Sort participants by name
      participants.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
      
      // Format duration helper
      const fmtDur = (val) => {
        if (!val) return '0:00';
        const h = Math.floor(val);
        const m = Math.round((val - h) * 60);
        return `${h}:${String(m).padStart(2, '0')}`;
      };
      
      // Generate HTML for PDF - using same template as Program/Employee Reports
      let html = `<html><head><style>
        body { font-family: sans-serif; color: #333; padding: 30px; }
        .logo { text-align: center; margin-bottom: 20px; }
        .logo img { max-width: 120px; height: auto; }
        h1 { color: #0d6efd; border-bottom: 2px solid #0d6efd; padding-bottom: 10px; margin-top: 0; margin-bottom: 10px; font-size: 24px; text-align: center; }
        .sub-headline { font-size: 18px; font-weight: bold; color: #444; margin-top: 0; margin-bottom: 25px; text-align: center; }
        h2 { background-color: #f1f3f4; padding: 8px; border-left: 5px solid #34a853; margin-top: 20px; font-size: 16px; }
        .summary-box { background: #e8f0fe; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: flex; gap: 30px; justify-content: center; flex-wrap: wrap; }
        .stat { font-weight: bold; color: #1967d2; font-size: 1.1em; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
        th { background-color: #e0e0e0; text-align: left; padding: 8px; border: 1px solid #ccc; }
        td { border: 1px solid #ccc; padding: 8px; }
        .info-table { margin-bottom: 20px; }
        .info-table td { border: none; padding: 5px 10px; }
        .info-table .label { font-weight: bold; color: #555; width: 150px; }
        .status-attended { background: #e8f5e9; color: #2e7d32; padding: 3px 8px; border-radius: 4px; font-size: 11px; }
        .status-absent { background: #ffebee; color: #c62828; padding: 3px 8px; border-radius: 4px; font-size: 11px; }
        .status-enrolled { background: #e3f2fd; color: #1565c0; padding: 3px 8px; border-radius: 4px; font-size: 11px; }
        .indiv-dev-box { background: #fff3e0; padding: 15px; border-radius: 8px; border-left: 5px solid #ff9800; margin-top: 15px; }
      </style></head><body>
      <div class="logo"><img src="${COMPANY_LOGO}" alt="Acme Corp"></div>
      <h1>Session Attendance Report</h1>
      <div class="sub-headline">${sessionName}</div>`;
      
      // Session Info Summary Box
      html += `<div class="summary-box">
        <div>Date: <span class="stat">${dateStr}</span></div>
        <div>Duration: <span class="stat">${fmtDur(duration)}</span></div>
        <div>Participants: <span class="stat">${participants.length}</span></div>
        <div>Location: <span class="stat">${location || 'N/A'}</span></div>
      </div>`;
      
      // Session Details
      html += `<h2>Session Details</h2>
      <table class="info-table">
        <tr><td class="label">Program:</td><td>${programName}</td></tr>
        <tr><td class="label">Session:</td><td>${sessionName}</td></tr>
        <tr><td class="label">Entity:</td><td>${sessionEntity || 'N/A'}</td></tr>
        <tr><td class="label">Type:</td><td>${sessionType}${isIndivDev ? ' (Individual Development)' : ''}</td></tr>
        <tr><td class="label">Status:</td><td>${sessionStatus}</td></tr>
      </table>`;
      
      // Status Summary
      html += `<h2>Attendance Summary</h2>
      <div class="summary-box">
        <div>Attended: <span class="stat">${statusCounts.Attended}</span></div>
        <div>Absent: <span class="stat">${statusCounts.Absent}</span></div>
        <div>Enrolled: <span class="stat">${statusCounts.Enrolled}</span></div>
        ${statusCounts.Other > 0 ? `<div>Other: <span class="stat">${statusCounts.Other}</span></div>` : ''}
      </div>`;
      
      // Participants List
      html += `<h2>Participants List</h2>`;
      
      if (participants.length === 0) {
        html += '<p style="color:#666; font-style:italic;">No participants enrolled in this session.</p>';
      } else {
        // If IndivDev, get exchange rates for USD conversion
        let participantsWithUSD = participants;
        if (isIndivDev) {
          const exchangeRates = getExchangeRatesMap_();
          // Convert sessionDate to ISO string format for exchange rate lookup
          const dateForLookup = sessionDate instanceof Date 
            ? sessionDate.toISOString().split('T')[0] 
            : (sessionDate ? String(sessionDate) : '');
          participantsWithUSD = participants.map(p => ({
            ...p,
            costUSD: convertToUSD_(p.cost, p.currency, dateForLookup, exchangeRates)
          }));
        }
        
        html += `<table>
          <thead><tr>
            <th width="5%">#</th>
            <th width="12%">Employee ID</th>
            <th width="22%">Name</th>
            <th width="12%">Entity</th>
            <th width="12%">Status</th>
            ${isIndivDev ? '<th width="18%">Cost</th><th width="19%">Cost (USD)</th>' : ''}
          </tr></thead>
          <tbody>`;
        
        participantsWithUSD.forEach((p, idx) => {
          const statusClass = p.status === 'Attended' ? 'status-attended' : 
                              p.status === 'Absent' ? 'status-absent' : 'status-enrolled';
          html += `<tr>
            <td>${idx + 1}</td>
            <td>${p.employeeId}</td>
            <td>${p.name}</td>
            <td>${p.entity}</td>
            <td><span class="${statusClass}">${p.status}</span></td>
            ${isIndivDev ? `<td>${p.currency} ${p.cost.toLocaleString()}</td><td>$${p.costUSD != null ? p.costUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '-'}</td>` : ''}
          </tr>`;
        });
        
        html += '</tbody></table>';
        
        // If IndivDev, show total cost with USD
        if (isIndivDev) {
          const totalCost = participantsWithUSD.reduce((sum, p) => sum + p.cost, 0);
          const totalCostUSD = participantsWithUSD.reduce((sum, p) => sum + (p.costUSD || 0), 0);
          const currencies = [...new Set(participantsWithUSD.map(p => p.currency).filter(c => c))];
          html += `<div class="indiv-dev-box">
            <strong>Total Individual Development Cost:</strong> ${currencies.length === 1 ? currencies[0] + ' ' : ''}${totalCost.toLocaleString()}
            ${currencies.length > 1 ? '<br><em style="font-size:11px; color:#666;">Note: Multiple currencies used</em>' : ''}<br>
            <strong>Total Cost (USD):</strong> $${totalCostUSD.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
          </div>`;
        }
      }
      
      html += `</body></html>`;
      
      // Generate PDF
      const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
      const base64 = Utilities.base64Encode(blob.getBytes());
      const filename = `Session_Report_${sessionId}_${new Date().toISOString().split('T')[0]}.pdf`;
      
      return { success: true, base64: base64, filename: filename };
      
    } catch (e) {
      console.error('generateSessionPDF error:', e);
      return { success: false, message: e.message };
    }
  });
}

// ============================================
// SETUP & CLEANUP
// ============================================
function setupDatabase() {
  const ss = getSpreadsheet_();
  
  // Users sheet
  let sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS);
    sheet.appendRow(['Email', 'Name', 'Role', 'Countries', 'Functions', 'Status']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.appendRow(['admin@example.com', 'Admin User', 'Developer', 'All', 'All', 'Active']);
  }
  
  // Programs sheet
  sheet = ss.getSheetByName(SHEET_NAMES.PROGRAMS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PROGRAMS);
    sheet.appendRow(['Program ID', 'Program Name', 'Category', 'Description']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  
  // Sessions sheet (18 columns - includes Form URLs)
  sheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SESSIONS);
    sheet.appendRow(['Program ID', 'Session ID', 'Session Name', 'Session Status', 'Type', 'Complete Date', 'Training Hours', 'Location / Platform', 'Average Satisfaction Score', 'Provider', 'IsIndivDev', 'LastModifiedBy', 'LastModifiedOn', 'Entity', 'Track via QR', 'FeedbackFormURL', 'SurveyFormURL', 'AssessmentURL']);
    sheet.getRange(1, 1, 1, 18).setFontWeight('bold');
    sheet.getRange("B:B").setNumberFormat("@");
  }
  
  // Enrollments sheet
  sheet = ss.getSheetByName(SHEET_NAMES.ENROLLMENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ENROLLMENTS);
    sheet.appendRow(['Enrollment ID', 'Session ID', 'Employee ID', 'Name', 'Email', 'Entity', 'Job Band', 'Function', 'Sub-Function', 'Position Type', 'Enrollment Status', 'Currency', 'Cost']);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
  }
  
  // Employees sheet
  sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.EMPLOYEES);
    sheet.appendRow(['Entity', 'Employee ID', 'Name', 'Job Band', 'Function', 'Sub-Function', 'Position Type', 'Email']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  
  // OTP Sessions sheet
  sheet = ss.getSheetByName(SHEET_NAMES.OTP_SESSIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.OTP_SESSIONS);
    sheet.appendRow(['Email', 'OTP', 'Expiry', 'Created']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  
  // Active Sessions sheet
  sheet = ss.getSheetByName(SHEET_NAMES.ACTIVE_SESSIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ACTIVE_SESSIONS);
    sheet.appendRow(['Token', 'Email', 'Name', 'Role', 'Countries', 'Functions', 'Expiry', 'Created']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  
  // FeedbackScores sheet (for Q1 scores - used for avg calculation)
  sheet = ss.getSheetByName('FeedbackScores');
  if (!sheet) {
    sheet = ss.insertSheet('FeedbackScores');
    sheet.appendRow(['ResponseID', 'SessionID', 'Timestamp', 'Email', 'Q1_Score']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  
  // FormResponses sheet (centralized storage of ALL form responses)
  sheet = ss.getSheetByName('FormResponses');
  if (!sheet) {
    sheet = ss.insertSheet('FormResponses');
    sheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
    sheet.getRange(1, 1, 1, 15).setFontWeight('bold');
  }
  
  // Scan_History sheet (QR attendance scan logs)
  sheet = ss.getSheetByName('Scan_History');
  if (!sheet) {
    sheet = ss.insertSheet('Scan_History');
    sheet.appendRow(['Scan ID', 'Session ID', 'Employee ID', 'Employee Name', 'Action', 'Scanned By', 'Scanned At']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  
  return 'Database setup complete! All 10 sheets created.';
}

function cleanupExpiredData() {
  try {
    const ss = getSpreadsheet_();
    const now = new Date();
    
    // Clean OTP sessions
    let sheet = ss.getSheetByName(SHEET_NAMES.OTP_SESSIONS);
    if (sheet && sheet.getLastRow() > 1) {
      let data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (new Date(data[i][2]) < now) {
          sheet.deleteRow(i + 1);
        }
      }
    }
    
    // Clean active sessions
    sheet = ss.getSheetByName(SHEET_NAMES.ACTIVE_SESSIONS);
    if (sheet && sheet.getLastRow() > 1) {
      let data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (new Date(data[i][6]) < now) {
          sheet.deleteRow(i + 1);
        }
      }
    }
    
    return 'Cleanup complete!';
  } catch (e) {
    return 'Cleanup error: ' + e.message;
  }
}

// Test function to verify setup
function testSetup() {
  try {
    const ss = getSpreadsheet_();
    const sheets = ss.getSheets().map(s => s.getName());
    
    const required = Object.values(SHEET_NAMES);
    const missing = required.filter(name => !sheets.includes(name));
    
    if (missing.length > 0) {
      return 'Missing sheets: ' + missing.join(', ') + '. Run setupDatabase() first.';
    }
    
    return 'All sheets present: ' + sheets.join(', ');
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// ============================================
// QR ATTENDANCE SYSTEM
// ============================================

/**
 * QR_SECRET is used to generate tamper-proof checksums for QR codes.
 * This prevents employees from creating fake QR codes.
 * 
 * IMPORTANT: Change this to your own unique random string!
 * Example: 'MyCompany_Training_2024_xK9mP2vQ'
 * 
 * Keep this secret and don't share it publicly.
 */
const QR_SECRET = 'ACME_TMS_Demo_2025';

// Cache for employee data (speeds up repeated scans)
const EMPLOYEE_CACHE_TTL = 300; // 5 minutes

/**
 * Get employee data with caching
 * Uses CacheService to avoid repeated sheet lookups
 */
function getCachedEmployee_(employeeId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'EMP_' + employeeId;
  
  // Try cache first
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Invalid cache, continue to sheet lookup
    }
  }
  
  // Cache miss - lookup in sheet
  const empSheet = getSheetSafe_(SHEET_NAMES.EMPLOYEES);
  if (!empSheet) return null;
  
  const empFinder = empSheet.getRange('B:B').createTextFinder(employeeId).matchEntireCell(true);
  const empMatch = empFinder.findNext();
  
  if (!empMatch) return null;
  
  const empRowNum = empMatch.getRow();
  const row = empSheet.getRange(empRowNum, 1, 1, 8).getValues()[0];
  
  const employee = {
    entity: row[0] || '',
    id: String(row[1]).trim(),
    name: row[2] || '',
    band: row[3] || '',
    function: row[4] || '',
    subFunction: row[5] || '',
    positionType: row[6] || '',
    email: row[7] || ''
  };
  
  // Cache for next time
  try {
    cache.put(cacheKey, JSON.stringify(employee), EMPLOYEE_CACHE_TTL);
  } catch (e) {
    // Cache put failed, continue anyway
  }
  
  return employee;
}

/**
 * Generate HMAC checksum for QR validation
 */
function generateChecksum_(data) {
  const signature = Utilities.computeHmacSha256Signature(data, QR_SECRET);
  // Take first 4 bytes and convert to hex
  return signature.slice(0, 4).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Generate QR string for an employee
 * Format: EMP|{EmployeeID}|{Name}|{Entity}|{Checksum}
 */
function generateQRString_(employeeId, employeeName, entity) {
  const data = `${employeeId}|${employeeName}|${entity}`;
  const checksum = generateChecksum_(data);
  return `EMP|${data}|${checksum}`;
}

/**
 * Validate and parse QR string
 * Returns { valid: true, employeeId, name, entity } or { valid: false, error }
 */
function validateQRString_(qrString) {
  try {
    if (!qrString || typeof qrString !== 'string') {
      return { valid: false, error: 'Empty or invalid QR code' };
    }
    
    const parts = qrString.split('|');
    if (parts.length !== 5 || parts[0] !== 'EMP') {
      return { valid: false, error: 'Invalid QR format' };
    }
    
    const [prefix, employeeId, name, entity, checksum] = parts;
    const data = `${employeeId}|${name}|${entity}`;
    const expectedChecksum = generateChecksum_(data);
    
    if (checksum !== expectedChecksum) {
      return { valid: false, error: 'Invalid QR checksum - code may be tampered' };
    }
    
    return { valid: true, employeeId, name, entity };
  } catch (e) {
    return { valid: false, error: 'QR validation error: ' + e.message };
  }
}

/**
 * Get active session for scanner (called from scanner web app)
 * Returns session info stored in PropertiesService
 */
function getActiveScannerSession() {
  try {
    // Try cache first (faster for repeated scans)
    const cache = CacheService.getScriptCache();
    const cachedSession = cache.get('ACTIVE_SCANNER_SESSION');
    if (cachedSession) {
      return JSON.parse(cachedSession);
    }
    
    // Fall back to ScriptProperties
    const props = PropertiesService.getScriptProperties();
    const sessionData = props.getProperty('ACTIVE_SCANNER_SESSION');
    if (!sessionData) return null;
    
    // Cache for 5 minutes
    cache.put('ACTIVE_SCANNER_SESSION', sessionData, 300);
    return JSON.parse(sessionData);
  } catch (e) {
    return null;
  }
}

/**
 * Set active session for scanner
 */
function setActiveScannerSession(sessionToken, sessionId, sessionName, programName) {
  return withSession_(sessionToken, (userData) => {
    try {
      const props = PropertiesService.getScriptProperties();
      const sessionData = {
        sessionId,
        sessionName,
        programName,
        setBy: userData.email,
        setAt: new Date().toISOString()
      };
      const sessionJson = JSON.stringify(sessionData);
      props.setProperty('ACTIVE_SCANNER_SESSION', sessionJson);
      
      // Also update cache
      CacheService.getScriptCache().put('ACTIVE_SCANNER_SESSION', sessionJson, 300);
      
      return { success: true, session: sessionData };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });
}

/**
 * Clear active scanner session
 */
function clearActiveScannerSession(sessionToken) {
  return withSession_(sessionToken, (userData) => {
    try {
      const props = PropertiesService.getScriptProperties();
      props.deleteProperty('ACTIVE_SCANNER_SESSION');
      
      // Also clear cache
      CacheService.getScriptCache().remove('ACTIVE_SCANNER_SESSION');
      
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });
}

/**
 * Mark attendance - called from scanner web app
 * OPTIMIZED: Uses TextFinder for fast lookups instead of loading entire sheets
 */
function markAttendance(qrString, sessionId) {
  try {
    // Validate QR
    const qrData = validateQRString_(qrString);
    if (!qrData.valid) {
      return { status: 'error', message: qrData.error };
    }
    
    // Get session ID - either passed directly or from active session
    let activeSessionId = sessionId;
    if (!activeSessionId) {
      const activeSession = getActiveScannerSession();
      if (activeSession) {
        activeSessionId = activeSession.sessionId;
      }
    }
    
    if (!activeSessionId) {
      return { status: 'error', message: 'No active session. Please set a session first.' };
    }
    
    // OPTIMIZED: Use cached employee lookup
    const employee = getCachedEmployee_(qrData.employeeId);
    if (!employee) {
      return { status: 'error', message: `Employee ${qrData.employeeId} not found` };
    }
    
    // OPTIMIZED: Check enrollment using TextFinder on Session ID + Employee ID
    const enrollSheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
    
    // Search for session ID first, then check employee ID in those rows
    const sessionFinder = enrollSheet.getRange('B:B').createTextFinder(activeSessionId).matchEntireCell(true);
    const sessionMatches = sessionFinder.findAll();
    
    for (const match of sessionMatches) {
      const rowNum = match.getRow();
      const empIdInRow = enrollSheet.getRange(rowNum, 3).getValue(); // Column C = Employee ID
      if (String(empIdInRow).trim() === qrData.employeeId) {
        // Already enrolled - log and return quickly
        logScanHistory_(activeSessionId, qrData.employeeId, qrData.name, 'Already Enrolled', 'Scanner');
        return { 
          status: 'success', 
          message: `${qrData.name} is already marked for this session.`,
          name: qrData.name,
          alreadyEnrolled: true
        };
      }
    }
    
    // Add enrollment
    const enrollId = 'ENR' + Date.now();
    const newRow = [
      enrollId,                           // Enrollment ID
      activeSessionId,                    // Session ID
      qrData.employeeId,                  // Employee ID
      qrData.name,                        // Name
      employee.email || '',               // Email
      employee.entity || '',              // Entity
      employee.band || '',                // Job Band
      employee.function || '',            // Function
      employee.subFunction || '',         // Sub-Function
      employee.positionType || '',        // Position Type
      'Attended',                         // Enrollment Status - mark as Attended
      '',                                 // Currency
      ''                                  // Cost
    ];
    
    enrollSheet.appendRow(newRow);
    
    // Log scan history
    logScanHistory_(activeSessionId, qrData.employeeId, qrData.name, 'Enrolled', 'Scanner');
    
    return { 
      status: 'success', 
      message: `${qrData.name} successfully checked in!`,
      name: qrData.name,
      alreadyEnrolled: false
    };
    
  } catch (e) {
    console.error('markAttendance error:', e);
    return { status: 'error', message: 'Server error: ' + e.message };
  }
}

/**
 * Scan attendance with session token (from main app)
 * OPTIMIZED: Uses TextFinder for fast lookups
 */
function scanAttendance(sessionToken, sessionId, qrString) {
  return withSession_(sessionToken, (userData) => {
    try {
      // Validate QR
      const qrData = validateQRString_(qrString);
      if (!qrData.valid) {
        return { status: 'error', message: qrData.error };
      }
      
      // OPTIMIZED: Use TextFinder to verify session exists
      const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
      const sessFinder = sessSheet.getRange('B:B').createTextFinder(sessionId).matchEntireCell(true);
      const sessMatch = sessFinder.findNext();
      
      if (!sessMatch) {
        return { status: 'error', message: 'Session not found: ' + sessionId };
      }
      
      // OPTIMIZED: Use cached employee lookup (avoids sheet reads on repeated scans)
      const employee = getCachedEmployee_(qrData.employeeId);
      if (!employee) {
        return { status: 'error', message: `Employee ${qrData.employeeId} not found` };
      }
      
      // OPTIMIZED: Check enrollment using TextFinder with batch read
      const enrollSheet = getSheet_(SHEET_NAMES.ENROLLMENTS);
      const sessionFinder = enrollSheet.getRange('B:B').createTextFinder(sessionId).matchEntireCell(true);
      const sessionMatches = sessionFinder.findAll();
      
      // OPTIMIZED: Batch read all matching rows at once (columns C and K)
      let foundEnrollment = null;
      let checkedInCount = 0;
      const totalEnrolled = sessionMatches.length;
      
      if (sessionMatches.length > 0) {
        // Read all row data in batch for efficiency
        const rowNums = sessionMatches.map(m => m.getRow());
        
        for (const rowNum of rowNums) {
          // Read columns C (employeeId) and K (status) together
          const rowData = enrollSheet.getRange(rowNum, 3, 1, 9).getValues()[0]; // C to K
          const empIdInRow = String(rowData[0] || '').trim(); // Column C
          const currentStatus = rowData[8] || 'Enrolled'; // Column K (index 8 from C)
          
          // Count attended for stats
          if (currentStatus === 'Attended') {
            checkedInCount++;
          }
          
          // Check if this is our employee
          if (empIdInRow === qrData.employeeId) {
            foundEnrollment = { rowNum, currentStatus };
          }
        }
      }
      
      // Helper to build response with stats
      const buildResponse = (baseResponse, newCheckIn = false) => {
        return {
          ...baseResponse,
          stats: {
            checkedIn: checkedInCount + (newCheckIn ? 1 : 0),
            target: totalEnrolled + (baseResponse.walkIn ? 1 : 0)
          }
        };
      };
      
      if (foundEnrollment) {
        if (foundEnrollment.currentStatus === 'Attended') {
          logScanHistory_(sessionId, qrData.employeeId, qrData.name, 'Already Attended', userData.email);
          return buildResponse({ 
            status: 'success', 
            message: `${qrData.name} already checked in.`,
            name: qrData.name,
            alreadyMarked: true
          });
        } else {
          // Update status to Attended
          enrollSheet.getRange(foundEnrollment.rowNum, 11).setValue('Attended');
          logScanHistory_(sessionId, qrData.employeeId, qrData.name, 'Checked In', userData.email);
          return buildResponse({ 
            status: 'success', 
            message: `${qrData.name} checked in!`,
            name: qrData.name,
            alreadyMarked: false
          }, true);
        }
      }
      
      // Not enrolled - add enrollment with Attended status (walk-in)
      const enrollId = 'ENR' + Date.now();
      const newRow = [
        enrollId,
        sessionId,
        qrData.employeeId,
        qrData.name,
        employee.email || '',
        employee.entity || '',
        employee.band || '',
        employee.function || '',
        employee.subFunction || '',
        employee.positionType || '',
        'Attended',
        '',
        ''
      ];
      
      enrollSheet.appendRow(newRow);
      logScanHistory_(sessionId, qrData.employeeId, qrData.name, 'Walk-in Attended', userData.email);
      
      return buildResponse({ 
        status: 'success', 
        message: `${qrData.name} checked in (walk-in)!`,
        name: qrData.name,
        alreadyMarked: false,
        walkIn: true
      }, true);
      
    } catch (e) {
      console.error('scanAttendance error:', e);
      return { status: 'error', message: 'Server error: ' + e.message };
    }
  });
}

/**
 * Log scan to history sheet
 */
function logScanHistory_(sessionId, employeeId, employeeName, action, scannedBy) {
  try {
    const ss = getSpreadsheet_();
    let histSheet = ss.getSheetByName('Scan_History');
    
    // Create sheet if not exists
    if (!histSheet) {
      histSheet = ss.insertSheet('Scan_History');
      histSheet.appendRow(['Scan ID', 'Session ID', 'Employee ID', 'Employee Name', 'Action', 'Scanned By', 'Scanned At']);
      histSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }
    
    const scanId = 'SCN' + Date.now();
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    
    histSheet.appendRow([scanId, sessionId, employeeId, employeeName, action, scannedBy, timestamp]);
  } catch (e) {
    console.error('logScanHistory error:', e);
  }
}

/**
 * Get scan history for a session
 */
function getScanHistory(sessionToken, sessionId) {
  return withSession_(sessionToken, (userData) => {
    try {
      const ss = getSpreadsheet_();
      const histSheet = ss.getSheetByName('Scan_History');
      
      if (!histSheet) {
        console.log('Scan_History sheet not found');
        return [];
      }
      
      const lastRow = histSheet.getLastRow();
      if (lastRow <= 1) {
        console.log('Scan_History has no data rows');
        return [];
      }
      
      const data = histSheet.getRange(1, 1, lastRow, 7).getValues();
      
      const history = [];
      for (let i = 1; i < data.length; i++) {
        const rowSessionId = String(data[i][1] || '').trim();
        
        // If no filter or matches filter
        if (!sessionId || sessionId === '' || rowSessionId === sessionId) {
          // Format the timestamp properly
          let scannedAt = data[i][6];
          if (scannedAt instanceof Date) {
            scannedAt = Utilities.formatDate(scannedAt, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
          } else if (scannedAt) {
            scannedAt = String(scannedAt);
          } else {
            scannedAt = '';
          }
          
          history.push({
            scanId: String(data[i][0] || ''),
            sessionId: rowSessionId,
            employeeId: String(data[i][2] || ''),
            employeeName: String(data[i][3] || ''),
            action: String(data[i][4] || ''),
            scannedBy: String(data[i][5] || ''),
            scannedAt: scannedAt
          });
        }
      }
      
      // Sort by most recent first
      history.sort((a, b) => new Date(b.scannedAt) - new Date(a.scannedAt));
      
      console.log('getScanHistory returning ' + history.length + ' records');
      return history;
    } catch (e) {
      console.error('getScanHistory error:', e);
      return [];
    }
  });
}

/**
 * DEBUG: Check Scan_History sheet status
 * Run this from Apps Script to verify the sheet exists and has data
 */
function debugScanHistory() {
  const ss = getSpreadsheet_();
  const histSheet = ss.getSheetByName('Scan_History');
  
  if (!histSheet) {
    Logger.log('❌ Scan_History sheet does NOT exist!');
    Logger.log('→ Run setupDatabase() to create it, or do a test scan.');
    return;
  }
  
  const lastRow = histSheet.getLastRow();
  Logger.log('✅ Scan_History sheet exists');
  Logger.log('Total rows (including header): ' + lastRow);
  Logger.log('Data rows: ' + (lastRow - 1));
  
  if (lastRow > 1) {
    const data = histSheet.getRange(2, 1, Math.min(5, lastRow - 1), 7).getValues();
    Logger.log('\nFirst ' + data.length + ' rows:');
    data.forEach((row, i) => {
      Logger.log((i + 1) + '. ' + row[3] + ' | ' + row[4] + ' | ' + row[6]);
    });
  } else {
    Logger.log('\n⚠️ No scan data yet. Try scanning a QR code.');
  }
  
  // Check headers
  const headers = histSheet.getRange(1, 1, 1, 7).getValues()[0];
  Logger.log('\nHeaders: ' + headers.join(' | '));
}

/**
 * DEBUG: Test logging a scan entry manually
 * Run this to verify logging works
 */
function testLogScanHistory() {
  logScanHistory_('TEST-SESSION', 'TEST-EMP-001', 'Test Employee', 'Test Scan', 'debug@test.com');
  Logger.log('✅ Test entry logged. Run debugScanHistory() to verify.');
}

/**
 * Get QR card data for employees
 */
function getQRCardData(sessionToken, filters) {
  return withSession_(sessionToken, (userData) => {
    try {
      let employees = getEmployees(sessionToken);
      
      // Apply filters
      if (filters) {
        if (filters.entity && filters.entity !== 'all') {
          employees = employees.filter(e => e.entity === filters.entity);
        }
        if (filters.function && filters.function !== 'all') {
          employees = employees.filter(e => e.function === filters.function);
        }
        if (filters.search) {
          const search = filters.search.toLowerCase();
          employees = employees.filter(e => 
            e.name.toLowerCase().includes(search) || 
            e.id.toLowerCase().includes(search)
          );
        }
      }
      
      // Generate QR strings for each employee
      return employees.map(emp => ({
        ...emp,
        qrString: generateQRString_(emp.id, emp.name, emp.entity)
      }));
      
    } catch (e) {
      console.error('getQRCardData error:', e);
      return [];
    }
  });
}

function emailQRCard(sessionToken, employeeId, options) {
  return { success: false, message: 'Email disabled in demo.' };
}

/**
 * Get upcoming trainings HTML for an employee
 */
function getUpcomingTrainingsHtml_(employeeId) {
  try {
    const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
    const sessSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
    const progSheet = getSheetSafe_(SHEET_NAMES.PROGRAMS);
    
    if (!enrollSheet || !sessSheet) return '';
    
    // Build session and program maps
    // Sessions sheet columns: A=ProgramID, B=SessionID, C=Name, D=Status, E=Type, F=Date, G=Duration, H=Location
    const sessData = sessSheet.getDataRange().getValues();
    const sessions = {};
    for (let i = 1; i < sessData.length; i++) {
      const sessionId = String(sessData[i][1]).trim(); // Column B = Session ID
      sessions[sessionId] = {
        name: sessData[i][2],           // Column C = Name
        programId: sessData[i][0],      // Column A = Program ID
        date: sessData[i][5],           // Column F = Training Date
        location: sessData[i][7]        // Column H = Location
      };
    }
    
    const programs = {};
    if (progSheet) {
      const progData = progSheet.getDataRange().getValues();
      for (let i = 1; i < progData.length; i++) {
        programs[progData[i][0]] = progData[i][1];
      }
    }
    
    // Find enrollments for this employee
    const enrollData = enrollSheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const upcoming = [];
    for (let i = 1; i < enrollData.length; i++) {
      if (String(enrollData[i][2]).trim() === employeeId) {
        const sessionId = String(enrollData[i][1]).trim();
        const session = sessions[sessionId];
        if (session && session.date && new Date(session.date) >= today) {
          upcoming.push({
            sessionName: session.name,
            programName: programs[session.programId] || session.programId,
            date: Utilities.formatDate(new Date(session.date), CONFIG.TIMEZONE, 'EEE, MMM d, yyyy'),
            location: session.location || 'TBD'
          });
        }
      }
    }
    
    if (upcoming.length === 0) return '';
    
    // Sort by date
    upcoming.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    // Build HTML - mobile responsive with fixed column widths
    const trainingsListHtml = upcoming.slice(0, 5).map(t => `
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; width: 55%; vertical-align: top;">
          <strong style="color: #1e40af;">${t.programName}</strong><br>
          <span style="font-size: 13px;">${t.sessionName}</span>
        </td>
        <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; width: 45%; text-align: right; font-size: 13px; vertical-align: top;">
          ${t.date}<br>
          <span style="color: #64748b;">${t.location}</span>
        </td>
      </tr>
    `).join('');
    
    return `
      <div style="margin-top: 24px;">
        <h3 style="color: #1e40af; font-size: 16px; margin-bottom: 12px;">Your Upcoming Training Sessions</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 14px; table-layout: fixed;">
          ${trainingsListHtml}
        </table>
        ${upcoming.length > 5 ? '<p style="color: #64748b; font-size: 12px; margin-top: 8px;">+ ' + (upcoming.length - 5) + ' more session(s)</p>' : ''}
      </div>
    `;
    
  } catch (e) {
    console.error('getUpcomingTrainingsHtml error:', e);
    return '';
  }
}

function emailQRCardsBulk(sessionToken, employeeIds, options) {
  return { success: false, message: 'Email disabled in demo.' };
}

/**
 * Get attendance stats for a session (for scanner display)
 */
function getSessionAttendanceStats(sessionId) {
  try {
    const enrollSheet = getSheetSafe_(SHEET_NAMES.ENROLLMENTS);
    if (!enrollSheet) return { enrolled: 0, target: 0 };
    
    const data = enrollSheet.getDataRange().getValues();
    let checkedIn = 0;
    let target = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === sessionId) {
        target++; // Count all participants as target
        const status = String(data[i][10] || '').toLowerCase();
        if (status === 'attended') {
          checkedIn++; // Only count Attended as checked in
        }
      }
    }
    
    return { enrolled: checkedIn, target: target };
  } catch (e) {
    console.error('getSessionAttendanceStats error:', e);
    return { enrolled: 0, target: 0 };
  }
}

/**
 * Get sessions list for scanner dropdown
 */
function getSessionsForScanner() {
  try {
    // Get spreadsheet first to ensure connection
    const ss = getSpreadsheet_();
    if (!ss) {
      throw new Error('Cannot connect to spreadsheet');
    }
    
    const sessSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
    if (!sessSheet) {
      console.log('Sessions sheet not found, returning empty array');
      return [];
    }
    
    const progSheet = ss.getSheetByName(SHEET_NAMES.PROGRAMS);
    const programs = {};
    
    if (progSheet && progSheet.getLastRow() > 1) {
      const progData = progSheet.getDataRange().getValues();
      for (let i = 1; i < progData.length; i++) {
        programs[progData[i][0]] = progData[i][1]; // ID -> Name
      }
    }
    
    const lastRow = sessSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('No sessions data, returning empty array');
      return [];
    }
    
    const sessData = sessSheet.getDataRange().getValues();
    const sessions = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Sessions sheet columns (0-indexed):
    // 0=ProgramID, 1=SessionID, 2=SessionName, 3=Status, 4=Type, 
    // 5=CompleteDate, 6=TrainingHours, 7=Location, 8=AvgScore, 
    // 9=Provider, 10=IsIndivDev, 11=LastModifiedBy, 12=LastModifiedOn, 
    // 13=Entity, 14=TrackQR
    
    for (let i = 1; i < sessData.length; i++) {
      const row = sessData[i];
      if (!row || !row[1]) continue; // Skip empty rows or rows without session ID
      
      const sessionDate = row[5]; // Complete Date column (F)
      const status = String(row[3] || '').toLowerCase();
      const trackQR = String(row[14] || '').toLowerCase();
      
      // Show sessions that:
      // 1. Are not cancelled, AND
      // 2. Have a date that is today or in the future, OR no date set yet
      let dateObj = null;
      if (sessionDate) {
        dateObj = sessionDate instanceof Date ? sessionDate : new Date(sessionDate);
        // Check if date is valid
        if (isNaN(dateObj.getTime())) {
          dateObj = null;
        }
      }
      
      const isValidDate = !dateObj || dateObj >= today;
      const isNotCancelled = status !== 'cancelled';
      
      if (isNotCancelled && isValidDate) {
        sessions.push({
          id: String(row[1]).trim(),              // Session ID (column B)
          name: row[2] || '',                     // Session Name (column C)
          programId: String(row[0]).trim(),       // Program ID (column A)
          programName: programs[row[0]] || String(row[0]),
          date: dateObj ? Utilities.formatDate(dateObj, CONFIG.TIMEZONE, 'MMM d, yyyy') : 'TBD',
          status: row[3] || 'Open',
          type: row[4] || 'Classroom',
          location: row[7] || '',                 // Location column (H)
          entity: row[13] || '',                  // Entity column (N)
          trackQR: trackQR === 'yes'
        });
      }
    }
    
    // Sort by date (TBD dates go last)
    sessions.sort((a, b) => {
      if (a.date === 'TBD') return 1;
      if (b.date === 'TBD') return -1;
      return new Date(a.date) - new Date(b.date);
    });
    
    console.log('getSessionsForScanner returning', sessions.length, 'sessions');
    return sessions;
  } catch (e) {
    console.error('getSessionsForScanner error:', e.message, e.stack);
    // Rethrow with more info so frontend can see the error
    throw new Error('Failed to load scanner sessions: ' + e.message);
  }
}

/**
 * Test function - run this from Apps Script editor to verify scanner can load sessions
 * Go to Run > testScannerSessions
 */
function testScannerSessions() {
  try {
    console.log('Testing getSessionsForScanner...');
    const sessions = getSessionsForScanner();
    console.log('Success! Found', sessions.length, 'sessions');
    if (sessions.length > 0) {
      console.log('First session:', JSON.stringify(sessions[0]));
    }
    return { success: true, count: sessions.length, sample: sessions[0] };
  } catch (e) {
    console.error('Test failed:', e.message);
    return { success: false, error: e.message };
  }
}

// ============================================
// SURVEY FORM FUNCTIONS
// ============================================

/**
 * Debug function - run this from Apps Script editor to check configuration
 */
function testFormConfiguration() {
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const formResponsesSheet = ss.getSheetByName('FormResponses');
  
  console.log('=== Form Configuration Debug ===');
  console.log('Spreadsheet ID:', PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  console.log('Forms Folder ID:', PropertiesService.getScriptProperties().getProperty('FORMS_FOLDER_ID'));
  
  if (sessSheet) {
    const headers = sessSheet.getRange(1, 1, 1, sessSheet.getLastColumn()).getValues()[0];
    console.log('Sessions sheet columns:', headers);
    console.log('Total columns:', headers.length);
    
    // Check a sample row
    if (sessSheet.getLastRow() > 1) {
      const sampleRow = sessSheet.getRange(2, 1, 1, sessSheet.getLastColumn()).getValues()[0];
      console.log('Sample row data length:', sampleRow.length);
      console.log('Column Q (index 16) FeedbackURL:', sampleRow[16]);
      console.log('Column R (index 17) SurveyURL:', sampleRow[17]);
      console.log('Column S (index 18) AssessmentURL:', sampleRow[18]);
    }
  } else {
    console.log('Sessions sheet not found!');
  }
  
  if (formResponsesSheet) {
    console.log('FormResponses sheet exists with', formResponsesSheet.getLastRow() - 1, 'responses');
  } else {
    console.log('FormResponses sheet does not exist (will be created on first response)');
  }
  
  // Check triggers
  const triggers = ScriptApp.getProjectTriggers();
  console.log('Active triggers:', triggers.length);
  triggers.forEach(t => {
    console.log('- ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')');
  });
  
  return 'Check the Execution Log for results';
}

/**
 * Utility: Move existing forms from My Drive to the configured folder
 * Run this once after setting up FORMS_FOLDER_ID to move any forms that were created before
 */
function moveExistingFormsToFolder() {
  const formsFolderId = getFormsFolderId_();
  if (!formsFolderId) {
    Logger.log('No FORMS_FOLDER_ID configured. Set it in Script Properties first.');
    return;
  }
  
  const targetFolder = DriveApp.getFolderById(formsFolderId);
  Logger.log('Target folder: ' + targetFolder.getName());
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  if (!sessSheet) {
    Logger.log('Sessions sheet not found');
    return;
  }
  
  const sessData = sessSheet.getDataRange().getValues();
  let movedCount = 0;
  
  // Check columns Q, R, S for form URLs (indices 16, 17, 18)
  for (let i = 1; i < sessData.length; i++) {
    const formUrls = [
      { url: sessData[i][16], type: 'Feedback' },  // Q
      { url: sessData[i][17], type: 'Survey' },    // R
      { url: sessData[i][18], type: 'Assessment' } // S
    ];
    
    formUrls.forEach(f => {
      if (f.url && f.url.includes('docs.google.com/forms')) {
        try {
          // Extract form ID from URL
          const match = f.url.match(/\/d\/([a-zA-Z0-9_-]+)/);
          if (match) {
            const formId = match[1];
            const formFile = DriveApp.getFileById(formId);
            
            // Check if already in target folder
            const parents = formFile.getParents();
            let alreadyInFolder = false;
            const parentFolders = [];
            
            while (parents.hasNext()) {
              const parent = parents.next();
              parentFolders.push(parent.getId());
              if (parent.getId() === formsFolderId) {
                alreadyInFolder = true;
              }
            }
            
            if (!alreadyInFolder) {
              // Add to target folder
              targetFolder.addFile(formFile);
              
              // Remove from old parents
              parentFolders.forEach(parentId => {
                try {
                  DriveApp.getFolderById(parentId).removeFile(formFile);
                } catch (e) {
                  // Ignore errors removing from parent
                }
              });
              
              Logger.log('Moved ' + f.type + ' form: ' + formFile.getName());
              movedCount++;
            }
          }
        } catch (e) {
          Logger.log('Could not move form (' + f.type + '): ' + e.message);
        }
      }
    });
  }
  
  Logger.log('Done! Moved ' + movedCount + ' forms to folder.');
}

/**
 * Fix FormResponses sheet structure - adds RawJSON column if missing
 * Run this once if your FormResponses sheet is missing the RawJSON column
 */
function fixFormResponsesSheet() {
  const ss = getSpreadsheet_();
  let frSheet = ss.getSheetByName('FormResponses');
  
  if (!frSheet) {
    frSheet = ss.insertSheet('FormResponses');
    frSheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
    Logger.log('Created FormResponses sheet with all columns (no Email)');
    return;
  }
  
  const lastCol = frSheet.getLastColumn();
  Logger.log('Current columns: ' + lastCol);
  
  // Check if RawJSON column exists (should be column 15 now - no Email)
  if (lastCol < 15) {
    // Add RawJSON header
    frSheet.getRange(1, 15).setValue('RawJSON');
    Logger.log('Added RawJSON column header');
    
    // For existing rows, generate RawJSON from Q1-Q10
    const lastRow = frSheet.getLastRow();
    if (lastRow > 1) {
      const data = frSheet.getRange(2, 1, lastRow - 1, 14).getValues();
      const rawJsonColumn = [];
      
      data.forEach(row => {
        const rawData = {};
        for (let i = 4; i < 14; i++) { // Q1-Q10 are indices 4-13 (no Email)
          if (row[i]) {
            rawData['Q' + (i - 3)] = row[i];
          }
        }
        rawJsonColumn.push([JSON.stringify(rawData)]);
      });
      
      frSheet.getRange(2, 15, rawJsonColumn.length, 1).setValues(rawJsonColumn);
      Logger.log('Generated RawJSON for ' + rawJsonColumn.length + ' existing rows');
    }
  } else {
    Logger.log('FormResponses sheet already has RawJSON column');
  }
  
  // Show current headers
  const headers = frSheet.getRange(1, 1, 1, Math.max(lastCol, 15)).getValues()[0];
  Logger.log('Headers: ' + headers.join(', '));
}

/**
 * Debug: Check responses for a specific session
 * Run this to verify responses are being stored correctly
 */
function debugSessionResponses() {
  const sessionId = 'EWTH2512001'; // Change this to test different sessions
  
  const ss = getSpreadsheet_();
  const frSheet = ss.getSheetByName('FormResponses');
  
  if (!frSheet) {
    Logger.log('FormResponses sheet not found');
    return;
  }
  
  const data = frSheet.getDataRange().getValues();
  Logger.log('Total rows in FormResponses: ' + (data.length - 1));
  Logger.log('Headers: ' + data[0].join(', '));
  
  let matchCount = 0;
  for (let i = 1; i < data.length; i++) {
    const rowSessionId = String(data[i][1]).trim();
    const formType = String(data[i][2]).trim();
    
    if (rowSessionId === sessionId) {
      matchCount++;
      Logger.log('Found response row ' + (i + 1) + ':');
      Logger.log('  SessionID: ' + rowSessionId);
      Logger.log('  FormType: "' + formType + '"');
      Logger.log('  Email: ' + data[i][4]);
      Logger.log('  Q1: ' + data[i][5]);
      Logger.log('  RawJSON column (15): ' + (data[i][15] || '(empty)'));
    }
  }
  
  Logger.log('Total matches for session ' + sessionId + ': ' + matchCount);
}

/**
 * Setup automatic weekly form archival trigger
 * Run this ONCE to set up automatic 90-day archival
 */
function setupFormArchivalTrigger() {
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'archiveOldForms');
  
  if (existingTrigger) {
    Logger.log('Archive trigger already exists. Skipping setup.');
    Logger.log('Trigger ID: ' + existingTrigger.getUniqueId());
    return 'Trigger already exists';
  }
  
  // Create weekly trigger - runs every Sunday at 2 AM
  ScriptApp.newTrigger('archiveOldForms')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();
  
  Logger.log('✅ Form archival trigger created successfully!');
  Logger.log('archiveOldForms will run every Sunday at 2 AM');
  Logger.log('Forms older than 90 days will be moved to Archive folder');
  
  return 'Trigger created successfully';
}

/**
 * Remove the form archival trigger (if you want to disable it)
 */
function removeFormArchivalTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'archiveOldForms') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
      Logger.log('Removed trigger: ' + trigger.getUniqueId());
    }
  });
  
  if (removed === 0) {
    Logger.log('No archive trigger found to remove');
  } else {
    Logger.log('Removed ' + removed + ' trigger(s)');
  }
  
  return 'Removed ' + removed + ' trigger(s)';
}

/**
 * Auto-update session statuses based on date
 * Sessions with date < today and status = 'Open' → 'Completed'
 * This runs daily via trigger
 */
function autoUpdateSessionStatuses() {
  try {
    const ss = getSpreadsheet_();
    const sessSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
    
    if (!sessSheet || sessSheet.getLastRow() < 2) {
      Logger.log('No sessions to update');
      return { updated: 0 };
    }
    
    const data = sessSheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let updatedCount = 0;
    
    // Headers: ProgramID(0), SessionID(1), Name(2), Status(3), Type(4), Date(5), ...
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][3] || '').trim();
      const dateVal = data[i][5];
      
      if (status === 'Open' && dateVal) {
        let sessionDate;
        if (dateVal instanceof Date) {
          sessionDate = new Date(dateVal);
        } else {
          sessionDate = new Date(dateVal);
        }
        sessionDate.setHours(0, 0, 0, 0);
        
        // If session date is before today, mark as Completed
        if (sessionDate < today) {
          sessSheet.getRange(i + 1, 4).setValue('Completed'); // Column D = Status
          updatedCount++;
          Logger.log('Updated session ' + data[i][1] + ' to Completed (date: ' + dateVal + ')');
        }
      }
    }
    
    Logger.log('✅ Auto-update complete. Updated ' + updatedCount + ' session(s) to Completed.');
    return { updated: updatedCount };
    
  } catch (e) {
    Logger.log('❌ autoUpdateSessionStatuses error: ' + e.message);
    return { error: e.message };
  }
}

/**
 * Setup daily trigger to auto-update session statuses
 * Run this ONCE to enable automatic status updates
 */
function setupDailyStatusUpdateTrigger() {
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'autoUpdateSessionStatuses');
  
  if (existingTrigger) {
    Logger.log('Status update trigger already exists. Skipping setup.');
    Logger.log('Trigger ID: ' + existingTrigger.getUniqueId());
    return 'Trigger already exists';
  }
  
  // Create daily trigger - runs every day at 12:05 AM
  ScriptApp.newTrigger('autoUpdateSessionStatuses')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();
  
  Logger.log('✅ Daily status update trigger created successfully!');
  Logger.log('autoUpdateSessionStatuses will run every day at ~12:05 AM');
  Logger.log('Sessions with past dates will be auto-marked as Completed');
  
  return 'Trigger created successfully';
}

/**
 * Remove the daily status update trigger (if you want to disable it)
 */
function removeDailyStatusUpdateTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoUpdateSessionStatuses') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
      Logger.log('Removed trigger: ' + trigger.getUniqueId());
    }
  });
  
  if (removed === 0) {
    Logger.log('No status update trigger found to remove');
  } else {
    Logger.log('Removed ' + removed + ' trigger(s)');
  }
  
  return 'Removed ' + removed + ' trigger(s)';
}

/**
 * Get survey report data with filters
 * Called from Reports > Survey Result
 */
function getSurveyReportData(sessionToken, filters) {
  return withSession_(sessionToken, (userData) => {
    try {
      console.log('getSurveyReportData called with filters:', JSON.stringify(filters));
      
      const ss = getSpreadsheet_();
      const frSheet = ss.getSheetByName('FormResponses');
      
      if (!frSheet || frSheet.getLastRow() < 2) {
        console.log('FormResponses sheet empty or not found');
        return { success: true, responses: [], sessions: [] };
      }
      
      const frData = frSheet.getDataRange().getValues();
      console.log('FormResponses has ' + (frData.length - 1) + ' rows');
      
      const responses = [];
      
      // Get sessions data for matching
      const sessSheet = ss.getSheetByName('Sessions');
      const sessData = sessSheet.getDataRange().getValues();
      
      // Build session lookup
      const sessionMap = {};
      for (let i = 1; i < sessData.length; i++) {
        const sessId = String(sessData[i][1]).trim();
        sessionMap[sessId] = {
          programId: sessData[i][0],
          sessionName: sessData[i][2],
          date: sessData[i][5],
          status: sessData[i][3]
        };
      }
      
      // Filter responses
      for (let i = 1; i < frData.length; i++) {
        const sessionId = String(frData[i][1]).trim();
        const formType = String(frData[i][2] || '').toLowerCase();
        const sessionInfo = sessionMap[sessionId] || {};
        
        // Apply filters - skip row if filter doesn't match
        if (filters && filters.programId && sessionInfo.programId !== filters.programId) continue;
        if (filters && filters.sessionId && sessionId !== filters.sessionId) continue;
        if (filters && filters.formType && formType !== filters.formType.toLowerCase()) continue;
        if (filters && filters.date) {
          const filterDate = new Date(filters.date).toDateString();
          const sessionDate = sessionInfo.date ? new Date(sessionInfo.date).toDateString() : '';
          if (filterDate !== sessionDate) continue;
        }
        
        // Generate rawJSON if missing
        // Columns: ResponseID(0), SessionID(1), FormType(2), Timestamp(3), Q1(4)...Q10(13), RawJSON(14)
        let rawJSON = frData[i][14];
        if (!rawJSON) {
          const generatedData = {};
          for (let q = 0; q < 10; q++) {
            if (frData[i][4 + q]) {
              generatedData['Q' + (q + 1)] = frData[i][4 + q];
            }
          }
          rawJSON = JSON.stringify(generatedData);
        }
        
        responses.push({
          responseId: frData[i][0],
          sessionId: sessionId,
          sessionName: sessionInfo.sessionName || sessionId,
          formType: formType,
          timestamp: frData[i][3],
          q1: frData[i][4],
          rawJSON: rawJSON
        });
      }
      
      // Get unique sessions for the dropdown
      const sessions = [];
      const seenSessions = new Set();
      for (let i = 1; i < sessData.length; i++) {
        const sessId = String(sessData[i][1]).trim();
        if (!seenSessions.has(sessId)) {
          seenSessions.add(sessId);
          sessions.push({
            sessionId: sessId,
            sessionName: sessData[i][2],
            programId: sessData[i][0],
            date: sessData[i][5]
          });
        }
      }
      
      console.log('getSurveyReportData: Found ' + responses.length + ' matching responses');
      return { success: true, responses: responses, sessions: sessions };
    } catch (e) {
      console.error('getSurveyReportData error:', e);
      return { success: false, message: e.message, responses: [], sessions: [] };
    }
  });
}

/**
 * Debug function - run this to check FormResponses data
 */
function debugFormResponses() {
  const ss = getSpreadsheet_();
  const frSheet = ss.getSheetByName('FormResponses');
  
  if (!frSheet) {
    Logger.log('FormResponses sheet not found');
    return;
  }
  
  const data = frSheet.getDataRange().getValues();
  Logger.log('FormResponses has ' + (data.length - 1) + ' rows');
  Logger.log('Headers: ' + data[0].join(' | '));
  
  // Show first 5 data rows
  // Columns: ResponseID(0), SessionID(1), FormType(2), Timestamp(3), Q1(4), Q2(5)...Q10(13), RawJSON(14)
  for (let i = 1; i < Math.min(6, data.length); i++) {
    Logger.log('Row ' + i + ':');
    Logger.log('  SessionID: "' + data[i][1] + '"');
    Logger.log('  FormType: "' + data[i][2] + '"');
    Logger.log('  Q1: "' + data[i][4] + '"');
  }
  
  // List unique session IDs in FormResponses
  const sessionIds = new Set();
  for (let i = 1; i < data.length; i++) {
    sessionIds.add(String(data[i][1]).trim());
  }
  Logger.log('Unique Session IDs in FormResponses: ' + [...sessionIds].join(', '));
}

/**
 * Diagnostic: Check why a specific session's form isn't capturing responses
 * Run this with a session ID to diagnose issues
 */
function diagnoseSessionForm(sessionId) {
  // Default to test session if not provided
  sessionId = sessionId || 'GLOB2512024';
  
  Logger.log('=== DIAGNOSING SESSION: ' + sessionId + ' ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Find session and get form URL
  let feedbackUrl = null;
  let sessionFound = false;
  
  for (let i = 1; i < sessData.length; i++) {
    if (String(sessData[i][1]).trim() === sessionId) {
      sessionFound = true;
      feedbackUrl = sessData[i][15] || ''; // Column P (index 15)
      Logger.log('1. SESSION FOUND in row ' + (i + 1));
      Logger.log('   Feedback Form URL: ' + (feedbackUrl || '(empty)'));
      break;
    }
  }
  
  if (!sessionFound) {
    Logger.log('❌ Session not found in Sessions sheet!');
    return;
  }
  
  if (!feedbackUrl) {
    Logger.log('❌ No feedback form URL set for this session');
    return;
  }
  
  // Check the form
  Logger.log('\n2. CHECKING FORM...');
  let formId = null;
  try {
    const form = FormApp.openByUrl(feedbackUrl);
    formId = form.getId();
    const formTitle = form.getTitle();
    const formDesc = form.getDescription();
    
    Logger.log('   Form ID: ' + formId);
    Logger.log('   Form Title: ' + formTitle);
    Logger.log('   Form Description: ' + formDesc);
    
    // Check if description has session ID
    if (formDesc && formDesc.includes('TMS_SESSION:')) {
      const match = formDesc.match(/TMS_SESSION:([^|]+)/);
      const embeddedSessionId = match ? match[1].trim() : '(not found)';
      Logger.log('   ✓ Embedded Session ID: ' + embeddedSessionId);
      
      if (embeddedSessionId !== sessionId) {
        Logger.log('   ⚠️ WARNING: Embedded session ID does not match! Expected: ' + sessionId);
        Logger.log('   → This form will save responses under: ' + embeddedSessionId);
      }
    } else {
      Logger.log('   ❌ Form description is missing TMS_SESSION tag!');
      Logger.log('   → This means responses cannot be linked to the session');
    }
    
    // Check triggers (may fail due to permissions)
    Logger.log('\n3. CHECKING TRIGGERS...');
    try {
      const triggers = ScriptApp.getProjectTriggers();
      let formHasTrigger = false;
      
      triggers.forEach(t => {
        try {
          if (t.getTriggerSourceId() === formId) {
            formHasTrigger = true;
            Logger.log('   ✓ Trigger found: ' + t.getHandlerFunction());
          }
        } catch (e) {}
      });
      
      if (!formHasTrigger) {
        Logger.log('   ❌ NO TRIGGER found for this form!');
        Logger.log('   → Responses will NOT be captured automatically');
      }
      
      Logger.log('\n4. TRIGGER SUMMARY (Total: ' + triggers.length + '/20 max)');
      let formTriggerCount = 0;
      triggers.forEach(t => {
        const handler = t.getHandlerFunction();
        if (handler === 'onFeedbackFormSubmit' || handler === 'onSurveyFormSubmit') {
          formTriggerCount++;
        }
      });
      Logger.log('   Form triggers: ' + formTriggerCount);
      Logger.log('   Other triggers: ' + (triggers.length - formTriggerCount));
      
      if (triggers.length >= 20) {
        Logger.log('   ⚠️ WARNING: At trigger limit! New triggers cannot be created.');
      }
    } catch (triggerErr) {
      Logger.log('   ⚠️ Cannot check triggers (permission issue)');
      Logger.log('   Run: reauthorizeScript() to fix permissions');
    }
    
  } catch (e) {
    Logger.log('   ❌ Error accessing form: ' + e.message);
  }
  
  // Check FormResponses for this session
  Logger.log('\n5. CHECKING FORMRESPONSES SHEET...');
  const frSheet = ss.getSheetByName('FormResponses');
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    let responseCount = 0;
    for (let i = 1; i < frData.length; i++) {
      if (String(frData[i][1]).trim() === sessionId) {
        responseCount++;
      }
    }
    Logger.log('   Responses found for ' + sessionId + ': ' + responseCount);
  } else {
    Logger.log('   FormResponses sheet not found');
  }
  
  Logger.log('\n=== DIAGNOSIS COMPLETE ===');
}

/**
 * COMPREHENSIVE: Check ALL sessions with feedback forms
 * This will reveal which sessions have working triggers vs broken ones
 */
function diagnoseAllSessionForms() {
  Logger.log('=== DIAGNOSING ALL SESSION FORMS ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Get all triggers
  const triggers = ScriptApp.getProjectTriggers();
  const triggerFormIds = new Set();
  
  triggers.forEach(t => {
    try {
      const sourceId = t.getTriggerSourceId();
      if (sourceId) triggerFormIds.add(sourceId);
    } catch (e) {}
  });
  
  Logger.log('Total triggers in project: ' + triggers.length + '/20');
  Logger.log('Form triggers found: ' + triggerFormIds.size);
  Logger.log('\n');
  
  // Get all responses by session
  const frSheet = ss.getSheetByName('FormResponses');
  const responsesBySession = {};
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    for (let i = 1; i < frData.length; i++) {
      const sid = String(frData[i][1]).trim();
      responsesBySession[sid] = (responsesBySession[sid] || 0) + 1;
    }
  }
  
  // Check each session
  const working = [];
  const broken = [];
  const noForm = [];
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const feedbackUrl = String(sessData[i][15] || '').trim(); // Column P
    
    if (!feedbackUrl) {
      noForm.push(sessionId);
      continue;
    }
    
    try {
      const form = FormApp.openByUrl(feedbackUrl);
      const formId = form.getId();
      const desc = form.getDescription() || '';
      const hasTrigger = triggerFormIds.has(formId);
      const hasCorrectDesc = desc.includes('TMS_SESSION:' + sessionId);
      const responseCount = responsesBySession[sessionId] || 0;
      
      if (hasTrigger && hasCorrectDesc) {
        working.push({
          sessionId: sessionId,
          formId: formId,
          responses: responseCount,
          status: '✓ OK'
        });
      } else {
        broken.push({
          sessionId: sessionId,
          formId: formId,
          responses: responseCount,
          hasTrigger: hasTrigger,
          hasCorrectDesc: hasCorrectDesc,
          actualDesc: desc.substring(0, 50)
        });
      }
    } catch (e) {
      broken.push({
        sessionId: sessionId,
        formId: '(inaccessible)',
        responses: responsesBySession[sessionId] || 0,
        error: e.message.substring(0, 50)
      });
    }
  }
  
  // Report
  Logger.log('=== WORKING FORMS (' + working.length + ') ===');
  working.forEach(w => {
    Logger.log('  ' + w.sessionId + ' - ' + w.responses + ' responses');
  });
  
  Logger.log('\n=== BROKEN FORMS (' + broken.length + ') ===');
  broken.forEach(b => {
    if (b.error) {
      Logger.log('  ' + b.sessionId + ' - ERROR: ' + b.error);
    } else {
      Logger.log('  ' + b.sessionId + ' - Trigger: ' + (b.hasTrigger ? '✓' : '✗') + 
                 ', Desc: ' + (b.hasCorrectDesc ? '✓' : '✗') + 
                 ', Responses: ' + b.responses);
      if (!b.hasCorrectDesc) {
        Logger.log('    Actual desc: "' + b.actualDesc + '"');
      }
    }
  });
  
  Logger.log('\n=== NO FEEDBACK FORM (' + noForm.length + ') ===');
  if (noForm.length <= 10) {
    Logger.log('  ' + noForm.join(', '));
  } else {
    Logger.log('  ' + noForm.slice(0, 10).join(', ') + '... and ' + (noForm.length - 10) + ' more');
  }
  
  Logger.log('\n=== SUMMARY ===');
  Logger.log('Working: ' + working.length);
  Logger.log('Broken: ' + broken.length);
  Logger.log('No form: ' + noForm.length);
  
  return { working, broken, noForm };
}

/**
 * SIMPLE DIAGNOSTIC: Check forms WITHOUT needing trigger permissions
 * Run this if diagnoseAllSessionForms fails due to permissions
 */
function diagnoseAllSessionFormsSimple() {
  Logger.log('=== DIAGNOSING ALL SESSION FORMS (Simple) ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Get all responses by session from FormResponses
  const frSheet = ss.getSheetByName('FormResponses');
  const responsesBySession = {};
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    for (let i = 1; i < frData.length; i++) {
      const sid = String(frData[i][1]).trim();
      responsesBySession[sid] = (responsesBySession[sid] || 0) + 1;
    }
  }
  
  Logger.log('Sessions with responses in FormResponses: ' + Object.keys(responsesBySession).length);
  Logger.log('Total responses: ' + Object.values(responsesBySession).reduce((a, b) => a + b, 0));
  Logger.log('\n');
  
  // Check each session's form
  const results = {
    hasFormAndResponses: [],
    hasFormNoResponses: [],
    noForm: [],
    formAccessError: []
  };
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const sessionName = String(sessData[i][2] || '').trim();
    const feedbackUrl = String(sessData[i][15] || '').trim(); // Column P
    const responseCount = responsesBySession[sessionId] || 0;
    
    if (!feedbackUrl) {
      results.noForm.push(sessionId);
      continue;
    }
    
    // Try to access the form and check its description
    try {
      const form = FormApp.openByUrl(feedbackUrl);
      const desc = form.getDescription() || '';
      const hasCorrectDesc = desc.includes('TMS_SESSION:');
      const embeddedId = desc.match(/TMS_SESSION:([^|]+)/);
      const embeddedSessionId = embeddedId ? embeddedId[1].trim() : '';
      
      const info = {
        sessionId: sessionId,
        responses: responseCount,
        hasDesc: hasCorrectDesc,
        embeddedId: embeddedSessionId,
        descMismatch: hasCorrectDesc && embeddedSessionId !== sessionId
      };
      
      if (responseCount > 0) {
        results.hasFormAndResponses.push(info);
      } else {
        results.hasFormNoResponses.push(info);
      }
    } catch (e) {
      results.formAccessError.push({
        sessionId: sessionId,
        responses: responseCount,
        error: e.message.substring(0, 60)
      });
    }
  }
  
  // Report
  Logger.log('=== FORMS WITH RESPONSES (' + results.hasFormAndResponses.length + ') ===');
  results.hasFormAndResponses.forEach(r => {
    let status = r.responses + ' responses';
    if (!r.hasDesc) status += ' [NO DESC!]';
    if (r.descMismatch) status += ' [ID MISMATCH: ' + r.embeddedId + ']';
    Logger.log('  ✓ ' + r.sessionId + ' - ' + status);
  });
  
  Logger.log('\n=== FORMS WITHOUT RESPONSES (' + results.hasFormNoResponses.length + ') ===');
  results.hasFormNoResponses.forEach(r => {
    let issue = '';
    if (!r.hasDesc) issue = '← MISSING TMS_SESSION in description!';
    else if (r.descMismatch) issue = '← WRONG ID: ' + r.embeddedId;
    else issue = '← Has correct desc, maybe no submissions yet or trigger issue';
    Logger.log('  ✗ ' + r.sessionId + ' - 0 responses ' + issue);
  });
  
  Logger.log('\n=== FORM ACCESS ERRORS (' + results.formAccessError.length + ') ===');
  results.formAccessError.forEach(r => {
    const hasResponses = r.responses > 0 ? ' [BUT HAS ' + r.responses + ' RESPONSES!]' : '';
    Logger.log('  ✗ ' + r.sessionId + ' - ' + r.error + hasResponses);
  });
  
  Logger.log('\n=== NO FEEDBACK FORM (' + results.noForm.length + ') ===');
  if (results.noForm.length <= 5) {
    Logger.log('  ' + results.noForm.join(', '));
  } else {
    Logger.log('  ' + results.noForm.length + ' sessions without feedback forms');
  }
  
  Logger.log('\n=== SUMMARY ===');
  Logger.log('Has form + responses: ' + results.hasFormAndResponses.length);
  Logger.log('Has form, NO responses: ' + results.hasFormNoResponses.length);
  Logger.log('Form access error: ' + results.formAccessError.length);
  Logger.log('No form: ' + results.noForm.length);
  
  return results;
}

/**
 * COMPARE: Sessions with forms vs Sessions with responses
 * This bypasses form access entirely - just compares data
 */
function compareFormsVsResponses() {
  Logger.log('=== COMPARING FORMS VS RESPONSES ===\n');
  
  const ss = getSpreadsheet_();
  
  // Get sessions with feedback form URLs
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  const sessionsWithForms = {};
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const feedbackUrl = String(sessData[i][15] || '').trim(); // Column P (index 15)
    if (feedbackUrl) {
      sessionsWithForms[sessionId] = feedbackUrl;
    }
  }
  
  Logger.log('Sessions with feedback form URL: ' + Object.keys(sessionsWithForms).length);
  
  // Get sessions with responses
  const frSheet = ss.getSheetByName('FormResponses');
  const responsesBySession = {};
  
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    Logger.log('FormResponses headers: ' + frData[0].join(' | '));
    
    for (let i = 1; i < frData.length; i++) {
      const sessionId = String(frData[i][1]).trim();
      const formType = String(frData[i][2]).trim();
      
      if (formType.toLowerCase() === 'feedback') {
        if (!responsesBySession[sessionId]) {
          responsesBySession[sessionId] = { count: 0, lastTimestamp: null };
        }
        responsesBySession[sessionId].count++;
        responsesBySession[sessionId].lastTimestamp = frData[i][3];
      }
    }
  }
  
  Logger.log('Sessions with feedback responses: ' + Object.keys(responsesBySession).length);
  Logger.log('\n');
  
  // Compare
  const hasFormWithResponses = [];
  const hasFormNoResponses = [];
  const hasResponsesNoForm = [];
  
  // Check sessions with forms
  for (const sessionId in sessionsWithForms) {
    if (responsesBySession[sessionId]) {
      hasFormWithResponses.push({
        sessionId: sessionId,
        responses: responsesBySession[sessionId].count,
        lastResponse: responsesBySession[sessionId].lastTimestamp
      });
    } else {
      hasFormNoResponses.push(sessionId);
    }
  }
  
  // Check for responses without forms (shouldn't happen but good to check)
  for (const sessionId in responsesBySession) {
    if (!sessionsWithForms[sessionId]) {
      hasResponsesNoForm.push({
        sessionId: sessionId,
        responses: responsesBySession[sessionId].count
      });
    }
  }
  
  // Report
  Logger.log('=== WORKING: Has Form + Has Responses (' + hasFormWithResponses.length + ') ===');
  hasFormWithResponses.forEach(r => {
    Logger.log('  ✓ ' + r.sessionId + ' - ' + r.responses + ' responses (last: ' + r.lastResponse + ')');
  });
  
  Logger.log('\n=== PROBLEM: Has Form + NO Responses (' + hasFormNoResponses.length + ') ===');
  hasFormNoResponses.forEach(sessionId => {
    Logger.log('  ✗ ' + sessionId + ' - 0 responses');
  });
  
  if (hasResponsesNoForm.length > 0) {
    Logger.log('\n=== ORPHAN: Has Responses but NO Form URL (' + hasResponsesNoForm.length + ') ===');
    hasResponsesNoForm.forEach(r => {
      Logger.log('  ? ' + r.sessionId + ' - ' + r.responses + ' responses');
    });
  }
  
  Logger.log('\n=== ANALYSIS ===');
  if (hasFormNoResponses.length > 0 && hasFormWithResponses.length > 0) {
    Logger.log('Some forms work, some don\'t. Possible causes:');
    Logger.log('1. Missing trigger on specific forms');
    Logger.log('2. Wrong/missing TMS_SESSION in form description');
    Logger.log('3. Form was deleted or recreated');
    Logger.log('\nWorking session IDs: ' + hasFormWithResponses.map(r => r.sessionId).join(', '));
    Logger.log('Not working session IDs: ' + hasFormNoResponses.join(', '));
  }
  
  return {
    working: hasFormWithResponses,
    notWorking: hasFormNoResponses,
    orphan: hasResponsesNoForm
  };
}

/**
 * LIST ALL TRIGGERS: See what forms have triggers and which sessions they map to
 */
function listAllFormTriggers() {
  Logger.log('=== LISTING ALL FORM TRIGGERS ===\n');
  
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers: ' + triggers.length + '/20');
  
  const formTriggers = [];
  const otherTriggers = [];
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    const eventType = t.getEventType().toString();
    
    if (handler === 'onFeedbackFormSubmit' || handler === 'onSurveyFormSubmit') {
      try {
        const formId = t.getTriggerSourceId();
        formTriggers.push({
          handler: handler,
          formId: formId,
          triggerId: t.getUniqueId()
        });
      } catch (e) {
        formTriggers.push({
          handler: handler,
          formId: '(error getting ID)',
          triggerId: t.getUniqueId()
        });
      }
    } else {
      otherTriggers.push({
        handler: handler,
        eventType: eventType
      });
    }
  });
  
  Logger.log('\n=== FORM TRIGGERS (' + formTriggers.length + ') ===');
  
  // Try to get session ID from each form
  formTriggers.forEach((ft, index) => {
    let sessionInfo = '';
    try {
      const form = FormApp.openById(ft.formId);
      const desc = form.getDescription() || '';
      const title = form.getTitle() || '';
      const match = desc.match(/TMS_SESSION:([^|]+)/);
      const sessionId = match ? match[1].trim() : '(no session ID in description)';
      sessionInfo = ' → Session: ' + sessionId + ' | Title: ' + title.substring(0, 40);
    } catch (e) {
      sessionInfo = ' → (cannot access form: ' + e.message.substring(0, 30) + ')';
    }
    Logger.log('  ' + (index + 1) + '. ' + ft.handler + ' | Form: ' + ft.formId + sessionInfo);
  });
  
  Logger.log('\n=== OTHER TRIGGERS (' + otherTriggers.length + ') ===');
  otherTriggers.forEach(ot => {
    Logger.log('  - ' + ot.handler + ' (' + ot.eventType + ')');
  });
  
  // Now compare with sessions
  Logger.log('\n=== CHECKING SESSIONS WITHOUT TRIGGERS ===');
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Get form IDs that have triggers
  const triggeredFormIds = new Set(formTriggers.map(ft => ft.formId));
  
  // Get sessions with form URLs and check if they have triggers
  const sessionsNeedingTriggers = [];
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const feedbackUrl = String(sessData[i][15] || '').trim();
    
    if (!feedbackUrl) continue;
    
    // Extract form ID from URL - try different patterns
    let formIdFromUrl = null;
    
    // Pattern 1: /forms/d/FORM_ID/edit or /forms/d/FORM_ID/viewform
    const editMatch = feedbackUrl.match(/\/forms\/d\/([a-zA-Z0-9_-]+)\//);
    if (editMatch) {
      formIdFromUrl = editMatch[1];
    }
    
    // Pattern 2: /forms/d/e/PUBLISHED_ID/viewform (published URL - can't get form ID from this)
    const publishedMatch = feedbackUrl.match(/\/forms\/d\/e\/([a-zA-Z0-9_-]+)\//);
    if (publishedMatch && !formIdFromUrl) {
      // This is a published URL - we can't extract the real form ID from it
      // We need to check by trying to match session IDs from triggers
      formIdFromUrl = 'PUBLISHED:' + publishedMatch[1];
    }
    
    Logger.log('  ' + sessionId + ': URL type = ' + (formIdFromUrl ? formIdFromUrl.substring(0, 20) + '...' : 'unknown'));
  }
  
  return { formTriggers, otherTriggers };
}

/**
 * FIX ALL BROKEN FORMS: Add triggers to forms that don't have them
 * This is the main fix function - run this to repair all broken sessions
 */
function fixAllBrokenFormTriggers() {
  Logger.log('=== FIXING ALL BROKEN FORM TRIGGERS ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Get existing triggers and map them by session ID
  const triggers = ScriptApp.getProjectTriggers();
  const triggersBySessionId = {};
  const existingTriggerFormIds = new Set();
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    if (handler !== 'onFeedbackFormSubmit' && handler !== 'onSurveyFormSubmit') return;
    
    try {
      const formId = t.getTriggerSourceId();
      existingTriggerFormIds.add(formId);
      
      const form = FormApp.openById(formId);
      const desc = form.getDescription() || '';
      const match = desc.match(/TMS_SESSION:([^|]+)/);
      if (match) {
        const sessionId = match[1].trim();
        if (!triggersBySessionId[sessionId]) {
          triggersBySessionId[sessionId] = [];
        }
        triggersBySessionId[sessionId].push({
          formId: formId,
          handler: handler,
          publishedUrl: form.getPublishedUrl()
        });
      }
    } catch (e) {
      // Form not accessible
    }
  });
  
  Logger.log('Existing triggers: ' + triggers.length);
  Logger.log('Sessions with triggers: ' + Object.keys(triggersBySessionId).length);
  
  // Get sessions that have responses (to know what's working)
  const frSheet = ss.getSheetByName('FormResponses');
  const workingSessions = new Set();
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    for (let i = 1; i < frData.length; i++) {
      workingSessions.add(String(frData[i][1]).trim());
    }
  }
  
  Logger.log('Sessions with responses: ' + workingSessions.size);
  Logger.log('\n');
  
  let fixed = 0;
  let urlUpdated = 0;
  let alreadyWorking = 0;
  let needsRecreate = 0;
  let errors = 0;
  
  // Process each session with a feedback form URL
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const sessionName = String(sessData[i][2] || '').trim();
    const feedbackUrl = String(sessData[i][15] || '').trim(); // Column P
    
    if (!feedbackUrl) continue;
    
    // Skip if already working
    if (workingSessions.has(sessionId)) {
      Logger.log('✓ ' + sessionId + ' - Already working (has responses)');
      alreadyWorking++;
      continue;
    }
    
    // Check if we have a trigger for this session
    if (triggersBySessionId[sessionId]) {
      // We have a trigger! Update column P with the correct URL
      const triggerInfo = triggersBySessionId[sessionId][0]; // Use first one
      const correctUrl = triggerInfo.publishedUrl;
      
      if (correctUrl !== feedbackUrl) {
        sessSheet.getRange(i + 1, 16).setValue(correctUrl); // Column P
        Logger.log('🔧 ' + sessionId + ' - Updated form URL to match trigger');
        urlUpdated++;
      } else {
        Logger.log('? ' + sessionId + ' - Has trigger, URL matches, but no responses. Maybe no submissions yet?');
      }
      continue;
    }
    
    // No trigger exists - just need to fix the form description for polling
    // Try to access the form
    try {
      const form = FormApp.openByUrl(feedbackUrl);
      const formId = form.getId();
      
      // Update description (required for polling to identify the session)
      form.setDescription('TMS_SESSION:' + sessionId + '|TMS_TYPE:feedback');
      
      // NOTE: Trigger creation skipped - polling handles response collection
      
      Logger.log('🔧 ' + sessionId + ' - Fixed form description for polling');
      fixed++;
      
    } catch (e) {
      // Can't access form - mark for recreation
      Logger.log('⚠️ ' + sessionId + ' - Cannot access form, needs recreation. Error: ' + e.message.substring(0, 40));
      needsRecreate++;
    }
  }
  
  Logger.log('\n=== FIX COMPLETE ===');
  Logger.log('Already working: ' + alreadyWorking);
  Logger.log('URL updated (had trigger): ' + urlUpdated);
  Logger.log('Description fixed: ' + fixed);
  Logger.log('Needs form recreation: ' + needsRecreate);
  Logger.log('Errors: ' + errors);
  
  Logger.log('\n📋 Next Steps:');
  Logger.log('1. Run setupFormResponsePolling() to enable automatic polling');
  Logger.log('2. Run removeAllFormTriggers() to clean up old per-form triggers');
  
  if (needsRecreate > 0) {
    Logger.log('3. Run recreateFormsForBrokenSessions() to fix inaccessible forms');
  }
  
  return { fixed, urlUpdated, alreadyWorking, needsRecreate, errors };
}

/**
 * RECREATE FORMS: For sessions where the form is inaccessible
 */
function recreateFormsForBrokenSessions() {
  Logger.log('=== RECREATING FORMS FOR BROKEN SESSIONS ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Get sessions that have responses (working)
  const frSheet = ss.getSheetByName('FormResponses');
  const workingSessions = new Set();
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    for (let i = 1; i < frData.length; i++) {
      workingSessions.add(String(frData[i][1]).trim());
    }
  }
  
  // Get sessions that have working triggers
  const triggers = ScriptApp.getProjectTriggers();
  const sessionsWithTriggers = new Set();
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    if (handler !== 'onFeedbackFormSubmit') return;
    
    try {
      const formId = t.getTriggerSourceId();
      const form = FormApp.openById(formId);
      const desc = form.getDescription() || '';
      const match = desc.match(/TMS_SESSION:([^|]+)/);
      if (match) {
        sessionsWithTriggers.add(match[1].trim());
      }
    } catch (e) {}
  });
  
  let recreated = 0;
  let skipped = 0;
  let errors = 0;
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const sessionName = String(sessData[i][2] || sessionId).trim();
    const feedbackUrl = String(sessData[i][15] || '').trim();
    
    if (!feedbackUrl) continue;
    
    // Skip if working or has trigger
    if (workingSessions.has(sessionId) || sessionsWithTriggers.has(sessionId)) {
      skipped++;
      continue;
    }
    
    // Try to access form - if fails, recreate
    try {
      const form = FormApp.openByUrl(feedbackUrl);
      // Form is accessible, skip
      skipped++;
    } catch (e) {
      // Form inaccessible - recreate it
      Logger.log('Recreating form for: ' + sessionId);
      
      try {
        // Create new form
        const title = 'Post-Training Feedback: ' + sessionName;
        const form = FormApp.create(title);
        
        // Add questions
        form.addScaleItem()
          .setTitle('Overall, how satisfied are you with this training?')
          .setBounds(1, 5)
          .setLabels('Very Dissatisfied', 'Very Satisfied')
          .setRequired(true);
        
        form.addScaleItem()
          .setTitle('How would you rate the trainer/facilitator?')
          .setBounds(1, 5)
          .setLabels('Poor', 'Excellent');
        
        form.addScaleItem()
          .setTitle('How relevant was the content to your work?')
          .setBounds(1, 5)
          .setLabels('Not Relevant', 'Very Relevant');
        
        form.addParagraphTextItem()
          .setTitle('What did you learn from this training?');
        
        form.addParagraphTextItem()
          .setTitle('Any suggestions for improvement?');
        
        // Set form settings
        form.setRequireLogin(false);
        form.setCollectEmail(false);
        form.setLimitOneResponsePerUser(false);
        form.setDescription('TMS_SESSION:' + sessionId + '|TMS_TYPE:feedback');
        
        const newFormUrl = form.getPublishedUrl();
        const formId = form.getId();
        
        // Update session with new URL and form ID
        sessSheet.getRange(i + 1, 16).setValue(newFormUrl); // Column P - URL
        sessSheet.getRange(i + 1, 19).setValue(formId);     // Column S - Form ID
        
        // NOTE: Trigger creation skipped - polling handles response collection
        // Responses will be captured within 10 minutes by pollAllFormResponses()
        
        Logger.log('✓ ' + sessionId + ' - Form recreated: ' + newFormUrl);
        recreated++;
        
      } catch (createErr) {
        Logger.log('❌ ' + sessionId + ' - Error creating form: ' + createErr.message);
        errors++;
      }
    }
  }
  
  Logger.log('\n=== RECREATION COMPLETE ===');
  Logger.log('Recreated: ' + recreated);
  Logger.log('Skipped (working): ' + skipped);
  Logger.log('Errors: ' + errors);
  Logger.log('\nTotal triggers now: ' + ScriptApp.getProjectTriggers().length + '/20');
  
  return { recreated, skipped, errors };
}

/**
 * CLEANUP: Remove duplicate/orphaned triggers to free up trigger quota
 */
function cleanupOrphanedTriggers() {
  Logger.log('=== CLEANING UP ORPHANED TRIGGERS ===\n');
  
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers before cleanup: ' + triggers.length);
  
  // Get sessions and their expected form IDs
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Map session IDs to check which are valid
  const validSessionIds = new Set();
  for (let i = 1; i < sessData.length; i++) {
    validSessionIds.add(String(sessData[i][1]).trim());
  }
  
  let removed = 0;
  let kept = 0;
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    
    // Only check form triggers
    if (handler !== 'onFeedbackFormSubmit' && handler !== 'onSurveyFormSubmit') {
      kept++;
      return;
    }
    
    try {
      const formId = t.getTriggerSourceId();
      const form = FormApp.openById(formId);
      const desc = form.getDescription() || '';
      const match = desc.match(/TMS_SESSION:([^|]+)/);
      const sessionId = match ? match[1].trim() : null;
      
      // Check if this session exists and has this form linked
      if (!sessionId) {
        Logger.log('Removing trigger: Form has no session ID in description');
        ScriptApp.deleteTrigger(t);
        removed++;
      } else if (!validSessionIds.has(sessionId)) {
        Logger.log('Removing trigger: Session ' + sessionId + ' no longer exists');
        ScriptApp.deleteTrigger(t);
        removed++;
      } else {
        kept++;
      }
    } catch (e) {
      // Can't access form - might be deleted
      Logger.log('Removing trigger: Cannot access form - ' + e.message.substring(0, 30));
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  
  Logger.log('\n=== CLEANUP COMPLETE ===');
  Logger.log('Removed: ' + removed);
  Logger.log('Kept: ' + kept);
  Logger.log('Total triggers now: ' + ScriptApp.getProjectTriggers().length);
  
  return { removed, kept };
}

/**
 * CLEANUP DUPLICATES: Remove duplicate triggers for the same session
 * Keeps only ONE trigger per session (the one with most responses or first found)
 */
function cleanupDuplicateTriggers() {
  Logger.log('=== CLEANING UP DUPLICATE TRIGGERS ===\n');
  
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers before cleanup: ' + triggers.length);
  
  // Get responses count by form ID to determine which trigger to keep
  const ss = getSpreadsheet_();
  const frSheet = ss.getSheetByName('FormResponses');
  const responsesBySession = {};
  
  if (frSheet) {
    const frData = frSheet.getDataRange().getValues();
    for (let i = 1; i < frData.length; i++) {
      const sessionId = String(frData[i][1]).trim();
      responsesBySession[sessionId] = (responsesBySession[sessionId] || 0) + 1;
    }
  }
  
  // Group triggers by session ID
  const triggersBySession = {};
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    if (handler !== 'onFeedbackFormSubmit' && handler !== 'onSurveyFormSubmit') return;
    
    try {
      const formId = t.getTriggerSourceId();
      const form = FormApp.openById(formId);
      const desc = form.getDescription() || '';
      const match = desc.match(/TMS_SESSION:([^|]+)/);
      const formType = desc.includes('TMS_TYPE:feedback') ? 'feedback' : 
                       desc.includes('TMS_TYPE:survey') ? 'survey' : 'unknown';
      
      if (match) {
        const sessionId = match[1].trim();
        const key = sessionId + '|' + formType; // Group by session + type
        
        if (!triggersBySession[key]) {
          triggersBySession[key] = [];
        }
        
        triggersBySession[key].push({
          trigger: t,
          formId: formId,
          handler: handler,
          publishedUrl: form.getPublishedUrl()
        });
      }
    } catch (e) {
      // Can't access - will be handled by cleanupOrphanedTriggers
    }
  });
  
  // Report duplicates and remove extras
  let removed = 0;
  
  for (const key in triggersBySession) {
    const sessionTriggers = triggersBySession[key];
    const [sessionId, formType] = key.split('|');
    
    if (sessionTriggers.length > 1) {
      Logger.log('\n' + sessionId + ' (' + formType + '): ' + sessionTriggers.length + ' triggers (DUPLICATES!)');
      
      // Keep the first one, remove the rest
      // In the future, could be smarter (keep the one that matches column P URL)
      const keepTrigger = sessionTriggers[0];
      Logger.log('  Keeping: Form ' + keepTrigger.formId.substring(0, 20) + '...');
      
      for (let i = 1; i < sessionTriggers.length; i++) {
        const removeTrigger = sessionTriggers[i];
        Logger.log('  Removing: Form ' + removeTrigger.formId.substring(0, 20) + '...');
        ScriptApp.deleteTrigger(removeTrigger.trigger);
        removed++;
      }
    } else {
      Logger.log(sessionId + ' (' + formType + '): 1 trigger ✓');
    }
  }
  
  Logger.log('\n=== DUPLICATE CLEANUP COMPLETE ===');
  Logger.log('Duplicates removed: ' + removed);
  Logger.log('Total triggers now: ' + ScriptApp.getProjectTriggers().length);
  
  return { removed };
}

// ============================================================================
// POLLING-BASED FORM RESPONSE COLLECTION
// This replaces per-form triggers with a single polling mechanism
// Benefits: Unlimited forms, only 1 trigger used
// ============================================================================

/**
 * SETUP: Initialize the polling system for form responses
 * Run this ONCE to set up the time-based trigger
 */
function setupFormResponsePolling() {
  Logger.log('=== SETTING UP FORM RESPONSE POLLING ===\n');
  
  // Remove any existing polling triggers
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'pollAllFormResponses') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  
  if (removed > 0) {
    Logger.log('Removed ' + removed + ' existing polling trigger(s)');
  }
  
  // Create new polling trigger - runs every 10 minutes
  ScriptApp.newTrigger('pollAllFormResponses')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  Logger.log('✓ Created polling trigger (every 10 minutes)');
  Logger.log('\nThe system will now automatically check all forms for new responses.');
  Logger.log('No per-form triggers needed - unlimited forms supported!');
  
  // Initialize last poll timestamp
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('LAST_FORM_POLL')) {
    props.setProperty('LAST_FORM_POLL', new Date().toISOString());
    Logger.log('✓ Initialized last poll timestamp');
  }
  
  Logger.log('\n=== SETUP COMPLETE ===');
  Logger.log('Total triggers now: ' + ScriptApp.getProjectTriggers().length);
  
  return { success: true };
}

/**
 * REMOVE: Disable the polling system
 */
function removeFormResponsePolling() {
  Logger.log('=== REMOVING FORM RESPONSE POLLING ===\n');
  
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'pollAllFormResponses') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  
  Logger.log('Removed ' + removed + ' polling trigger(s)');
  Logger.log('Total triggers now: ' + ScriptApp.getProjectTriggers().length);
  
  return { removed };
}

/**
 * MAIN POLLING FUNCTION: Check all forms for new responses
 * This runs automatically every 10 minutes via time-based trigger
 */
function pollAllFormResponses() {
  const startTime = new Date();
  Logger.log('=== POLLING ALL FORMS FOR RESPONSES ===');
  Logger.log('Started at: ' + startTime.toISOString());
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Get or create FormResponses sheet
  let frSheet = ss.getSheetByName('FormResponses');
  if (!frSheet) {
    frSheet = ss.insertSheet('FormResponses');
    frSheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 
                       'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
    Logger.log('Created FormResponses sheet');
  }
  
  // Get existing response IDs to avoid duplicates
  const existingResponseIds = new Set();
  const frData = frSheet.getDataRange().getValues();
  for (let i = 1; i < frData.length; i++) {
    existingResponseIds.add(String(frData[i][0])); // ResponseID column
  }
  
  Logger.log('Existing responses: ' + existingResponseIds.size);
  
  // Collect all forms to check (feedback, survey, assessment)
  // Columns: P=15 (FeedbackURL), Q=16 (SurveyURL), R=17 (AssessmentURL)
  //          S=18 (FeedbackFormID), T=19 (SurveyFormID), U=20 (AssessmentFormID)
  // Note: Array indices are 0-based, so P=15, Q=16, R=17, S=18, T=19, U=20
  const formsToCheck = [];
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    
    // Get form IDs (new columns S, T, U) or fall back to extracting from URL
    const feedbackFormId = String(sessData[i][18] || '').trim(); // Column S (index 18)
    const surveyFormId = String(sessData[i][19] || '').trim();   // Column T (index 19)
    const assessmentFormId = String(sessData[i][20] || '').trim(); // Column U (index 20)
    
    // Also get URLs for fallback
    const feedbackUrl = String(sessData[i][15] || '').trim(); // Column P
    const surveyUrl = String(sessData[i][16] || '').trim();   // Column Q
    const assessmentUrl = String(sessData[i][17] || '').trim(); // Column R
    
    if (feedbackFormId || feedbackUrl) {
      formsToCheck.push({ 
        sessionId, 
        formId: feedbackFormId,
        url: feedbackUrl, 
        type: 'feedback', 
        row: i + 1 
      });
    }
    if (surveyFormId || surveyUrl) {
      formsToCheck.push({ 
        sessionId, 
        formId: surveyFormId,
        url: surveyUrl, 
        type: 'survey', 
        row: i + 1 
      });
    }
    if (assessmentFormId || assessmentUrl) {
      formsToCheck.push({ 
        sessionId, 
        formId: assessmentFormId,
        url: assessmentUrl, 
        type: 'assessment', 
        row: i + 1 
      });
    }
  }
  
  Logger.log('Forms to check: ' + formsToCheck.length);
  
  let totalNewResponses = 0;
  let formsProcessed = 0;
  let formsWithErrors = 0;
  
  // Process each form
  for (const formInfo of formsToCheck) {
    try {
      let form;
      
      // Try to open by form ID first (preferred method)
      if (formInfo.formId) {
        form = FormApp.openById(formInfo.formId);
      } else if (formInfo.url) {
        // Try to extract form ID from edit URL pattern
        const editMatch = formInfo.url.match(/\/forms\/d\/([a-zA-Z0-9_-]+)\//);
        if (editMatch) {
          form = FormApp.openById(editMatch[1]);
        } else {
          // Can't open - published URLs don't work
          throw new Error('No form ID available - only published URL stored');
        }
      } else {
        continue; // No form info available
      }
      
      const responses = form.getResponses();
      
      let newForThisForm = 0;
      
      for (const response of responses) {
        const responseId = response.getId();
        
        // Skip if already processed
        if (existingResponseIds.has(responseId)) {
          continue;
        }
        
        // Process new response
        const timestamp = response.getTimestamp();
        const itemResponses = response.getItemResponses();
        
        // Extract up to 10 question responses
        const answers = [];
        for (let q = 0; q < 10; q++) {
          if (itemResponses[q]) {
            answers.push(itemResponses[q].getResponse());
          } else {
            answers.push('');
          }
        }
        
        // Build raw JSON for reference
        const rawData = {};
        itemResponses.forEach((ir, idx) => {
          rawData['Q' + (idx + 1)] = {
            question: ir.getItem().getTitle(),
            answer: ir.getResponse()
          };
        });
        
        // Add to FormResponses
        frSheet.appendRow([
          responseId,
          formInfo.sessionId,
          formInfo.type,
          timestamp,
          ...answers,
          JSON.stringify(rawData)
        ]);
        
        existingResponseIds.add(responseId);
        newForThisForm++;
        totalNewResponses++;
      }
      
      if (newForThisForm > 0) {
        Logger.log('  ' + formInfo.sessionId + ' (' + formInfo.type + '): +' + newForThisForm + ' new responses');
        
        // Update average score for feedback forms
        if (formInfo.type === 'feedback') {
          updateSessionAvgScoreFromPolling_(formInfo.sessionId, frSheet, sessSheet);
        }
      }
      
      formsProcessed++;
      
    } catch (e) {
      Logger.log('  ❌ ' + formInfo.sessionId + ' (' + formInfo.type + '): ' + e.message.substring(0, 50));
      formsWithErrors++;
    }
  }
  
  // Update last poll timestamp
  PropertiesService.getScriptProperties().setProperty('LAST_FORM_POLL', startTime.toISOString());
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  Logger.log('\n=== POLLING COMPLETE ===');
  Logger.log('Forms processed: ' + formsProcessed);
  Logger.log('Forms with errors: ' + formsWithErrors);
  Logger.log('New responses imported: ' + totalNewResponses);
  Logger.log('Duration: ' + duration.toFixed(1) + ' seconds');
  
  if (formsWithErrors > 0) {
    Logger.log('\n⚠️ Some forms had errors. Run migrateFormIdsFromTriggers() to fix.');
  }
  
  return { 
    formsProcessed, 
    formsWithErrors, 
    totalNewResponses, 
    duration 
  };
}

/**
 * Helper: Update session average score after polling
 */
function updateSessionAvgScoreFromPolling_(sessionId, frSheet, sessSheet) {
  // Get all feedback responses for this session
  const frData = frSheet.getDataRange().getValues();
  const scores = [];
  
  for (let i = 1; i < frData.length; i++) {
    if (String(frData[i][1]).trim() === sessionId && 
        String(frData[i][2]).toLowerCase() === 'feedback') {
      const q1 = frData[i][4]; // Q1 is overall satisfaction (column E, index 4)
      if (q1 && !isNaN(Number(q1))) {
        scores.push(Number(q1));
      }
    }
  }
  
  if (scores.length === 0) return;
  
  const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
  
  // Find session row and update
  const sessData = sessSheet.getDataRange().getValues();
  for (let i = 1; i < sessData.length; i++) {
    if (String(sessData[i][1]).trim() === sessionId) {
      sessSheet.getRange(i + 1, 9).setValue(avgScore.toFixed(2)); // Column I
      break;
    }
  }
}

/**
 * MANUAL: Run polling immediately (for testing or manual sync)
 */
function runFormPollingNow() {
  Logger.log('Running manual form polling...\n');
  return pollAllFormResponses();
}

/**
 * MIGRATION: Extract form IDs from existing triggers and save to columns S, T, U
 * Run this ONCE to migrate existing forms to the new polling system
 */
function migrateFormIdsFromTriggers() {
  Logger.log('=== MIGRATING FORM IDS FROM TRIGGERS ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Ensure columns S, T, U exist (up to column 21)
  const currentLastCol = sessSheet.getLastColumn();
  if (currentLastCol < 21) {
    sessSheet.insertColumnsAfter(currentLastCol, 21 - currentLastCol);
    Logger.log('Expanded Sessions sheet to column U (21)');
    
    // Add headers if they don't exist
    sessSheet.getRange(1, 19).setValue('FeedbackFormID');
    sessSheet.getRange(1, 20).setValue('SurveyFormID');
    sessSheet.getRange(1, 21).setValue('AssessmentFormID');
    Logger.log('Added column headers for form IDs');
  }
  
  // Get form IDs from existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  const formIdsBySession = {}; // { sessionId: { feedback: formId, survey: formId } }
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    if (handler !== 'onFeedbackFormSubmit' && handler !== 'onSurveyFormSubmit') return;
    
    try {
      const formId = t.getTriggerSourceId();
      const form = FormApp.openById(formId);
      const desc = form.getDescription() || '';
      
      const sessionMatch = desc.match(/TMS_SESSION:([^|]+)/);
      const typeMatch = desc.match(/TMS_TYPE:([^|]+)/);
      
      if (sessionMatch) {
        const sessionId = sessionMatch[1].trim();
        const formType = typeMatch ? typeMatch[1].trim() : (handler === 'onFeedbackFormSubmit' ? 'feedback' : 'survey');
        
        if (!formIdsBySession[sessionId]) {
          formIdsBySession[sessionId] = {};
        }
        formIdsBySession[sessionId][formType] = formId;
        
        Logger.log('Found: ' + sessionId + ' (' + formType + ') → ' + formId);
      }
    } catch (e) {
      Logger.log('Error reading trigger: ' + e.message);
    }
  });
  
  Logger.log('\nFound ' + Object.keys(formIdsBySession).length + ' sessions with form IDs from triggers');
  
  // Update Sessions sheet with form IDs
  let updated = 0;
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    
    if (formIdsBySession[sessionId]) {
      const formIds = formIdsBySession[sessionId];
      
      if (formIds.feedback) {
        sessSheet.getRange(i + 1, 19).setValue(formIds.feedback); // Column S
        updated++;
      }
      if (formIds.survey) {
        sessSheet.getRange(i + 1, 20).setValue(formIds.survey); // Column T
        updated++;
      }
      if (formIds.assessment) {
        sessSheet.getRange(i + 1, 21).setValue(formIds.assessment); // Column U
        updated++;
      }
      
      Logger.log('Updated: ' + sessionId);
    }
  }
  
  Logger.log('\n=== MIGRATION COMPLETE ===');
  Logger.log('Form IDs saved: ' + updated);
  Logger.log('\nNext steps:');
  Logger.log('1. Run setupFormResponsePolling() to enable polling');
  Logger.log('2. Run removeAllFormTriggers() to remove old per-form triggers');
  Logger.log('3. Run runFormPollingNow() to test');
  
  return { updated, sessionsWithFormIds: Object.keys(formIdsBySession).length };
}

/**
 * MIGRATION: Search Google Drive for TMS forms and extract form IDs
 * Use this when triggers are already deleted
 */
function migrateFormIdsFromDrive() {
  Logger.log('=== SEARCHING DRIVE FOR TMS FORMS ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Ensure columns S, T, U exist (up to column 21)
  const currentLastCol = sessSheet.getLastColumn();
  if (currentLastCol < 21) {
    sessSheet.insertColumnsAfter(currentLastCol, 21 - currentLastCol);
    Logger.log('Expanded Sessions sheet to column U (21)');
    
    // Add headers
    sessSheet.getRange(1, 19).setValue('FeedbackFormID');
    sessSheet.getRange(1, 20).setValue('SurveyFormID');
    sessSheet.getRange(1, 21).setValue('AssessmentFormID');
    Logger.log('Added column headers');
  }
  
  // Get all session IDs we're looking for
  const sessionIds = new Set();
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    if (sessionId) sessionIds.add(sessionId);
  }
  
  Logger.log('Looking for forms for ' + sessionIds.size + ' sessions');
  
  // Search for Google Forms in Drive
  const formIdsBySession = {};
  
  // Search for forms with TMS in title or that are Google Forms
  const formFiles = DriveApp.getFilesByType(MimeType.GOOGLE_FORMS);
  let formsChecked = 0;
  
  while (formFiles.hasNext()) {
    const file = formFiles.next();
    formsChecked++;
    
    try {
      const form = FormApp.openById(file.getId());
      const desc = form.getDescription() || '';
      const title = form.getTitle() || '';
      
      // Check if this is a TMS form
      const sessionMatch = desc.match(/TMS_SESSION:([^|]+)/);
      if (sessionMatch) {
        const sessionId = sessionMatch[1].trim();
        const typeMatch = desc.match(/TMS_TYPE:([^|]+)/);
        const formType = typeMatch ? typeMatch[1].trim() : 'feedback';
        
        if (sessionIds.has(sessionId)) {
          if (!formIdsBySession[sessionId]) {
            formIdsBySession[sessionId] = {};
          }
          formIdsBySession[sessionId][formType] = file.getId();
          
          Logger.log('Found: ' + sessionId + ' (' + formType + ') → ' + file.getId() + ' | ' + title);
        }
      }
    } catch (e) {
      // Skip forms we can't access
    }
  }
  
  Logger.log('\nChecked ' + formsChecked + ' forms in Drive');
  Logger.log('Found ' + Object.keys(formIdsBySession).length + ' sessions with TMS forms');
  
  // Update Sessions sheet with form IDs
  let updated = 0;
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    
    if (formIdsBySession[sessionId]) {
      const formIds = formIdsBySession[sessionId];
      
      if (formIds.feedback) {
        sessSheet.getRange(i + 1, 19).setValue(formIds.feedback); // Column S
        updated++;
      }
      if (formIds.survey) {
        sessSheet.getRange(i + 1, 20).setValue(formIds.survey); // Column T
        updated++;
      }
      if (formIds.assessment) {
        sessSheet.getRange(i + 1, 21).setValue(formIds.assessment); // Column U
        updated++;
      }
      
      Logger.log('Updated session: ' + sessionId);
    }
  }
  
  // Report sessions still missing form IDs
  const missingFormIds = [];
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    const hasUrl = String(sessData[i][15] || '').trim(); // Column P
    
    if (hasUrl && !formIdsBySession[sessionId]) {
      missingFormIds.push(sessionId);
    }
  }
  
  Logger.log('\n=== MIGRATION COMPLETE ===');
  Logger.log('Form IDs saved: ' + updated);
  
  if (missingFormIds.length > 0) {
    Logger.log('\n⚠️ Sessions with URLs but no matching form found (' + missingFormIds.length + '):');
    Logger.log(missingFormIds.join(', '));
    Logger.log('\nThese forms may have been deleted or are inaccessible.');
    Logger.log('Run recreateFormsForBrokenSessions() to create new forms for them.');
  }
  
  Logger.log('\nNext: Run runFormPollingNow() to test polling');
  
  return { updated, found: Object.keys(formIdsBySession).length, missing: missingFormIds };
}

/**
 * CLEANUP: Remove ALL per-form triggers (after switching to polling)
 * Run this after setupFormResponsePolling() to clean up old triggers
 */
function removeAllFormTriggers() {
  Logger.log('=== REMOVING ALL PER-FORM TRIGGERS ===\n');
  
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  let kept = 0;
  
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    
    // Remove form-specific triggers
    if (handler === 'onFeedbackFormSubmit' || handler === 'onSurveyFormSubmit') {
      ScriptApp.deleteTrigger(t);
      removed++;
    } else {
      kept++;
    }
  });
  
  Logger.log('Removed: ' + removed + ' form triggers');
  Logger.log('Kept: ' + kept + ' other triggers');
  Logger.log('Total triggers now: ' + ScriptApp.getProjectTriggers().length);
  
  return { removed, kept };
}

/**
 * Fix a form that's missing trigger or description
 * Run this with a session ID to fix it
 */
function fixSessionFormTrigger(sessionId) {
  sessionId = sessionId || 'GLOB2512024';
  
  Logger.log('=== FIXING SESSION FORM: ' + sessionId + ' ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Find session and get form URL
  let feedbackUrl = null;
  let sessionRow = -1;
  
  for (let i = 1; i < sessData.length; i++) {
    if (String(sessData[i][1]).trim() === sessionId) {
      feedbackUrl = sessData[i][15] || ''; // Column P (index 15)
      sessionRow = i + 1;
      break;
    }
  }
  
  if (!feedbackUrl) {
    Logger.log('❌ No feedback form URL found for session: ' + sessionId);
    return;
  }
  
  try {
    const form = FormApp.openByUrl(feedbackUrl);
    const formId = form.getId();
    
    // Fix description if needed
    const currentDesc = form.getDescription() || '';
    if (!currentDesc.includes('TMS_SESSION:')) {
      form.setDescription('TMS_SESSION:' + sessionId + '|TMS_TYPE:feedback');
      Logger.log('✓ Fixed form description');
    } else {
      Logger.log('✓ Form description already correct');
    }
    
    // NOTE: Per-form triggers are no longer needed!
    // The polling system handles all response collection.
    Logger.log('✓ Form is ready for polling-based response collection');
    Logger.log('  Responses will be captured within 10 minutes by pollAllFormResponses()');
    
    Logger.log('\n=== FIX COMPLETE ===');
    Logger.log('Form should now capture responses for session: ' + sessionId);
    Logger.log('Make sure polling is enabled: run setupFormResponsePolling()');
    
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('\n⚠️ PERMISSION ISSUE DETECTED');
    Logger.log('The script cannot access this form. This usually means:');
    Logger.log('1. Form was created by a different Google account');
    Logger.log('2. Form sharing/ownership was changed');
    Logger.log('3. Script authorization was reset');
    Logger.log('\nSOLUTION: Run recreateSessionForm("' + sessionId + '") to create a new form');
  }
}

/**
 * Recreate a feedback form for a session (when old form has permission issues)
 * This will:
 * 1. Create a brand new feedback form
 * 2. Set up the trigger properly
 * 3. Update the session with new form URL
 * 
 * Note: The old form will still exist but won't be linked to TMS
 */
function recreateSessionForm(sessionId, formType) {
  sessionId = sessionId || 'GLOB2512024';
  formType = formType || 'feedback';
  
  Logger.log('=== RECREATING ' + formType.toUpperCase() + ' FORM FOR: ' + sessionId + ' ===\n');
  
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName('Sessions');
  const sessData = sessSheet.getDataRange().getValues();
  
  // Find session
  let sessionRow = -1;
  let sessionName = '';
  
  for (let i = 1; i < sessData.length; i++) {
    if (String(sessData[i][1]).trim() === sessionId) {
      sessionRow = i + 1;
      sessionName = sessData[i][2] || sessionId; // Column C = Session Name
      break;
    }
  }
  
  if (sessionRow === -1) {
    Logger.log('❌ Session not found: ' + sessionId);
    return;
  }
  
  try {
    let form;
    let title = '';
    
    if (formType === 'feedback') {
      title = 'Post-Training Feedback: ' + sessionName;
      form = FormApp.create(title);
      
      // Q1: Fixed satisfaction scale (1-5) - Critical for scoring
      form.addScaleItem()
        .setTitle('Overall, how satisfied are you with this training?')
        .setBounds(1, 5)
        .setLabels('Very Dissatisfied', 'Very Satisfied')
        .setRequired(true);
      
      form.addScaleItem()
        .setTitle('How would you rate the trainer/facilitator?')
        .setBounds(1, 5)
        .setLabels('Poor', 'Excellent');
      
      form.addScaleItem()
        .setTitle('How relevant was the content to your work?')
        .setBounds(1, 5)
        .setLabels('Not Relevant', 'Very Relevant');
      
      form.addParagraphTextItem()
        .setTitle('What did you learn from this training?');
      
      form.addParagraphTextItem()
        .setTitle('Any suggestions for improvement?');
        
    } else if (formType === 'survey') {
      title = 'Pre-Training Survey: ' + sessionName;
      form = FormApp.create(title);
      
      form.addParagraphTextItem()
        .setTitle('What do you hope to learn from this training?')
        .setRequired(true);
      
      form.addScaleItem()
        .setTitle('How familiar are you with this topic?')
        .setBounds(1, 5)
        .setLabels('Not Familiar', 'Very Familiar');
      
      form.addParagraphTextItem()
        .setTitle('Do you have any specific questions you would like addressed?');
    }
    
    // Set form settings
    form.setRequireLogin(false);
    form.setCollectEmail(false);
    form.setLimitOneResponsePerUser(false);
    form.setDescription('TMS_SESSION:' + sessionId + '|TMS_TYPE:' + formType);
    
    const formUrl = form.getPublishedUrl();
    const formId = form.getId();
    
    Logger.log('✓ Form created: ' + title);
    Logger.log('  URL: ' + formUrl);
    
    // Update session with new form URL and form ID
    // URL columns: P=16, Q=17, R=18 | ID columns: S=19, T=20, U=21
    const formUrlCol = formType === 'feedback' ? 16 : (formType === 'survey' ? 17 : 18);
    const formIdCol = formType === 'feedback' ? 19 : (formType === 'survey' ? 20 : 21);
    
    sessSheet.getRange(sessionRow, formUrlCol).setValue(formUrl);
    sessSheet.getRange(sessionRow, formIdCol).setValue(formId);
    Logger.log('✓ Updated session with form URL and ID');
    
    // NOTE: Per-form triggers are no longer needed - polling handles response collection
    Logger.log('✓ Form ready for polling-based response collection');
    
    // Move to forms folder if configured
    const formsFolderId = PropertiesService.getScriptProperties().getProperty('FORMS_FOLDER_ID');
    if (formsFolderId) {
      try {
        const formFile = DriveApp.getFileById(formId);
        const targetFolder = DriveApp.getFolderById(formsFolderId);
        const parents = formFile.getParents();
        targetFolder.addFile(formFile);
        while (parents.hasNext()) {
          parents.next().removeFile(formFile);
        }
        Logger.log('✓ Moved form to TMS Forms folder');
      } catch (moveErr) {
        Logger.log('⚠️ Could not move form to folder: ' + moveErr.message);
      }
    }
    
    Logger.log('\n=== RECREATION COMPLETE ===');
    Logger.log('New form URL: ' + formUrl);
    Logger.log('\nIMPORTANT: Share this new URL with participants.');
    Logger.log('Responses will be captured within 10 minutes by the polling system.');
    
    return formUrl;
    
  } catch (e) {
    Logger.log('❌ Error creating form: ' + e.message);
  }
}

/**
 * Get session form URLs and response counts
 */
function getSessionForms(sessionToken, sessionId) {
  return withSession_(sessionToken, (userData) => {
    const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
    const sessData = sessSheet.getDataRange().getValues();
    
    let feedbackUrl = null, surveyUrl = null, assessmentUrl = null;
    let feedbackEditUrl = null, surveyEditUrl = null, assessmentEditUrl = null;
    
    // Find session row (Column B = Session ID)
    // Sessions columns: A=ProgramID, B=SessionID, C=Name, D=Status, E=Type, F=Date, G=Duration, H=Location, I=Score, J=Provider, K=IsIndivDev, L=LastModifiedBy, M=LastModifiedOn, N=Entity, O=TrackQR, P=FeedbackFormURL, Q=SurveyFormURL, R=AssessmentFormURL
    for (let i = 1; i < sessData.length; i++) {
      if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
        feedbackUrl = sessData[i][15] || null;  // Column P (index 15)
        surveyUrl = sessData[i][16] || null;    // Column Q (index 16)
        assessmentUrl = sessData[i][17] || null; // Column R (index 17)
        break;
      }
    }
    
    // Helper function to get edit URL from published URL
    function getEditUrl(publishedUrl) {
      if (!publishedUrl) return null;
      try {
        // Extract form ID from URL and construct edit URL
        // Published URL format: https://docs.google.com/forms/d/e/XXXX/viewform
        const match = publishedUrl.match(/\/forms\/d\/e\/([^\/]+)/);
        if (match) {
          // We need the actual form ID, not the published ID
          // Open the form by its published URL to get the real form ID
          const form = FormApp.openByUrl(publishedUrl);
          const formId = form.getId();
          return 'https://docs.google.com/forms/d/' + formId + '/edit';
        }
      } catch (e) {
        console.error('Error getting edit URL:', e);
      }
      return null;
    }
    
    // Get edit URLs for each form
    feedbackEditUrl = getEditUrl(feedbackUrl);
    surveyEditUrl = getEditUrl(surveyUrl);
    assessmentEditUrl = getEditUrl(assessmentUrl);
    
    // Get response counts and avg score from FormResponses sheet (centralized)
    let feedbackResponses = 0, surveyResponses = 0, assessmentResponses = 0;
    let totalQ1Score = 0, q1Count = 0;
    
    const formResponsesSheet = getSheetSafe_('FormResponses');
    if (formResponsesSheet) {
      const frData = formResponsesSheet.getDataRange().getValues();
      // Headers: ResponseID(0), SessionID(1), FormType(2), Timestamp(3), Q1(4), Q2(5)...
      for (let i = 1; i < frData.length; i++) {
        if (String(frData[i][1]).trim() === String(sessionId).trim()) {
          const formType = String(frData[i][2]).trim().toLowerCase();
          if (formType === 'feedback') {
            feedbackResponses++;
            // Calculate avg from Q1 (column index 4)
            const q1 = parseFloat(frData[i][4]) || 0;
            if (q1 > 0) {
              totalQ1Score += q1;
              q1Count++;
            }
          }
          else if (formType === 'survey') surveyResponses++;
          else if (formType === 'assessment') assessmentResponses++;
        }
      }
    }
    
    const avgScore = q1Count > 0 ? (totalQ1Score / q1Count) : null;
    
    // Auto-save avgScore to Sessions sheet (Column I) if we have feedback responses
    if (avgScore !== null) {
      try {
        const sessSheet = getSheetSafe_(SHEET_NAMES.SESSIONS);
        if (sessSheet) {
          const sessData = sessSheet.getDataRange().getValues();
          for (let i = 1; i < sessData.length; i++) {
            if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
              sessSheet.getRange(i + 1, 9).setValue(avgScore.toFixed(1)); // Column I
              break;
            }
          }
        }
      } catch (saveErr) {
        console.log('Could not auto-save avgScore:', saveErr.message);
      }
    }
    
    return {
      success: true,
      feedbackUrl: feedbackUrl,
      surveyUrl: surveyUrl,
      assessmentUrl: assessmentUrl,
      feedbackEditUrl: feedbackEditUrl,
      surveyEditUrl: surveyEditUrl,
      assessmentEditUrl: assessmentEditUrl,
      feedbackResponses: feedbackResponses,
      surveyResponses: surveyResponses,
      assessmentResponses: assessmentResponses,
      avgScore: avgScore
    };
  });
}

/**
 * Get all responses for a session from FormResponses sheet
 */
function getSessionResponses(sessionToken, sessionId) {
  return withSession_(sessionToken, (userData) => {
    try {
      const ss = getSpreadsheet_();
      let formResponsesSheet = ss.getSheetByName('FormResponses');
      
      // Create sheet if it doesn't exist
      if (!formResponsesSheet) {
        formResponsesSheet = ss.insertSheet('FormResponses');
        formResponsesSheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
        return { success: true, responses: [] };
      }
      
      const frData = formResponsesSheet.getDataRange().getValues();
      const responses = [];
      
      // Headers: ResponseID(0), SessionID(1), FormType(2), Timestamp(3), Q1(4)-Q10(13), RawJSON(14)
      for (let i = 1; i < frData.length; i++) {
        if (String(frData[i][1]).trim() === String(sessionId).trim()) {
          // Get rawJSON from column 14, or generate from Q1-Q10 if missing
          let rawJSON = frData[i][14];
          if (!rawJSON) {
            // Generate rawJSON from Q1-Q10 columns
            const generatedData = {};
            for (let q = 0; q < 10; q++) {
              if (frData[i][4 + q]) {
                generatedData['Q' + (q + 1)] = frData[i][4 + q];
              }
            }
            rawJSON = JSON.stringify(generatedData);
          }
          
          responses.push({
            responseId: frData[i][0],
            sessionId: frData[i][1],
            formType: String(frData[i][2] || '').toLowerCase(),
            timestamp: frData[i][3],
            q1: frData[i][4],
            q2: frData[i][5],
            q3: frData[i][6],
            q4: frData[i][7],
            q5: frData[i][8],
            q6: frData[i][9],
            q7: frData[i][10],
            q8: frData[i][11],
            q9: frData[i][12],
            q10: frData[i][13],
            rawJSON: rawJSON,
            q1Score: String(frData[i][2] || '').toLowerCase() === 'feedback' ? (parseFloat(frData[i][4]) || null) : null
          });
        }
      }
      
      console.log('getSessionResponses: Found ' + responses.length + ' responses for session ' + sessionId);
      return { success: true, responses: responses };
    } catch (e) {
      console.error('getSessionResponses error:', e);
      return { success: false, message: e.message, responses: [] };
    }
  });
}

/**
 * Export survey results to PDF
 */
function exportSurveyResultPDF(sessionToken, responses, sessionId, formType) {
  return withSession_(sessionToken, (userData) => {
    try {
      if (!responses || responses.length === 0) {
        return { success: false, message: 'No responses to export' };
      }
      
      const formTypeNames = {
        feedback: 'Post-Training Feedback',
        survey: 'Pre-Training Survey',
        assessment: 'Assessment Quiz'
      };
      
      const typeLabel = formType ? formTypeNames[formType] || formType : 'All Types';
      const sessionLabel = sessionId || 'All Sessions';
      
      // Build HTML for PDF with company logo
      let html = `
        <html>
        <head>
          <title>Survey Results Report</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 40px; font-size: 11px; }
            .logo { text-align: center; margin-bottom: 15px; }
            .logo img { max-width: 120px; height: auto; }
            h1 { color: #166534; font-size: 20px; margin-bottom: 5px; text-align: center; }
            .meta { color: #6b7280; font-size: 10px; margin-bottom: 20px; text-align: center; }
            table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            th { background: #f3f4f6; padding: 8px; text-align: left; border-bottom: 2px solid #e5e7eb; font-size: 10px; }
            td { padding: 8px; border-bottom: 1px solid #f3f4f6; vertical-align: top; }
            tr:nth-child(even) { background: #fafafa; }
            .type-badge { padding: 2px 8px; border-radius: 10px; font-size: 9px; }
            .feedback { background: #fef3c7; color: #92400e; }
            .survey { background: #dbeafe; color: #1e40af; }
            .assessment { background: #fce7f3; color: #9d174d; }
          </style>
        </head>
        <body>
          <div class="logo"><img src="${COMPANY_LOGO}" alt="Acme Corp"></div>
          <h1>Survey Results Report</h1>
          <p class="meta">
            Session: ${sessionLabel} | Type: ${typeLabel} | Total Responses: ${responses.length} | Generated: ${new Date().toLocaleString()}
          </p>
          <table>
            <thead>
              <tr>
                <th>Timestamp</th>
                <th>Session ID</th>
                <th>Type</th>
                <th>Email</th>
                <th>Q1</th>
                <th>Q2</th>
                <th>Q3</th>
              </tr>
            </thead>
            <tbody>`;
      
      responses.forEach(r => {
        const ts = r.timestamp ? new Date(r.timestamp).toLocaleString() : '-';
        const typeClass = r.formType || 'feedback';
        const typeDisplay = formTypeNames[r.formType] || r.formType || '-';
        
        html += `
          <tr>
            <td>${ts}</td>
            <td>${r.sessionId || '-'}</td>
            <td><span class="type-badge ${typeClass}">${typeDisplay}</span></td>
            <td>${r.email || '-'}</td>
            <td>${r.q1 || '-'}</td>
            <td>${r.q2 || '-'}</td>
            <td>${r.q3 || '-'}</td>
          </tr>`;
      });
      
      html += `
            </tbody>
          </table>
        </body>
        </html>`;
      
      // Create PDF
      const blob = Utilities.newBlob(html, 'text/html', 'report.html');
      const pdfBlob = blob.getAs('application/pdf');
      const base64 = Utilities.base64Encode(pdfBlob.getBytes());
      
      const filename = `Survey_Report_${sessionId || 'All'}_${new Date().toISOString().split('T')[0]}.pdf`;
      
      return { success: true, base64: base64, filename: filename };
    } catch (e) {
      console.error('exportSurveyResultPDF error:', e);
      return { success: false, message: e.message };
    }
  });
}

/**
 * Create a new Google Form for a session
 */
function createSessionForm(sessionToken, sessionId, formType, sessionName) {
  return withSession_(sessionToken, (userData) => {
    try {
      let form;
      let title = '';
      
      if (formType === 'feedback') {
        title = `Post-Training Feedback: ${sessionName}`;
        form = FormApp.create(title);
        
        // Q1: Fixed satisfaction scale (1-5) - This is critical for scoring
        form.addScaleItem()
          .setTitle('Overall, how satisfied are you with this training?')
          .setBounds(1, 5)
          .setLabels('Very Dissatisfied', 'Very Satisfied')
          .setRequired(true);
        
        // Additional optional questions
        form.addScaleItem()
          .setTitle('How would you rate the trainer/facilitator?')
          .setBounds(1, 5)
          .setLabels('Poor', 'Excellent');
        
        form.addScaleItem()
          .setTitle('How relevant was the content to your work?')
          .setBounds(1, 5)
          .setLabels('Not Relevant', 'Very Relevant');
        
        form.addParagraphTextItem()
          .setTitle('What did you learn from this training?');
        
        form.addParagraphTextItem()
          .setTitle('Any suggestions for improvement?');
          
      } else if (formType === 'survey') {
        title = `Pre-Training Survey: ${sessionName}`;
        form = FormApp.create(title);
        
        form.addParagraphTextItem()
          .setTitle('What do you hope to learn from this training?')
          .setRequired(true);
        
        form.addScaleItem()
          .setTitle('How familiar are you with this topic?')
          .setBounds(1, 5)
          .setLabels('Not Familiar', 'Very Familiar');
        
        form.addParagraphTextItem()
          .setTitle('Do you have any specific questions you would like addressed?');
          
      } else if (formType === 'assessment') {
        title = `Assessment: ${sessionName}`;
        form = FormApp.create(title);
        
        // Note: Email is automatically captured from Google sign-in
        // No need for Name or Email fields
        
        // Placeholder questions - trainer should customize these
        form.addMultipleChoiceItem()
          .setTitle('Sample Question 1: [Edit this question]')
          .setChoiceValues(['Option A', 'Option B', 'Option C', 'Option D'])
          .setRequired(true);
        
        form.addMultipleChoiceItem()
          .setTitle('Sample Question 2: [Edit this question]')
          .setChoiceValues(['True', 'False'])
          .setRequired(true);
      }
      
      // Set form settings - no sign-in required (for workers without email)
      form.setRequireLogin(false);
      form.setCollectEmail(false);
      form.setLimitOneResponsePerUser(false);
      form.setDescription(`TMS_SESSION:${sessionId}|TMS_TYPE:${formType}`);
      
      const formUrl = form.getPublishedUrl();
      const editUrl = form.getEditUrl();
      const formId = form.getId();
      
      // Add the user who created the form as an editor (so they can modify questions)
      try {
        form.addEditor(userData.email);
        console.log('Added ' + userData.email + ' as form editor');
      } catch (editorError) {
        console.log('Could not add editor (might already have access): ' + editorError.message);
      }
      
      // Move form to designated folder (for sustainability/organization)
      const formsFolderId = getFormsFolderId_();
      let movedToFolder = false;
      if (formsFolderId) {
        try {
          const formFile = DriveApp.getFileById(formId);
          const targetFolder = DriveApp.getFolderById(formsFolderId);
          
          // Get all parent folders
          const parents = formFile.getParents();
          
          // Add to target folder first
          targetFolder.addFile(formFile);
          
          // Remove from all other parent folders (including root/My Drive)
          while (parents.hasNext()) {
            const parent = parents.next();
            parent.removeFile(formFile);
          }
          
          movedToFolder = true;
          console.log('Form moved to folder: ' + targetFolder.getName());
        } catch (moveError) {
          console.error('Could not move form to folder:', moveError.message);
          // Form was still created, just not moved
        }
      } else {
        console.log('No FORMS_FOLDER_ID configured, form stays in root Drive');
      }
      
      // Save URL and Form ID to session
      const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
      const sessData = sessSheet.getDataRange().getValues();
      
      // Column numbers (1-indexed for getRange): 
      // P=16 (FeedbackURL), Q=17 (SurveyURL), R=18 (AssessmentURL)
      // S=19 (FeedbackFormID), T=20 (SurveyFormID), U=21 (AssessmentFormID)
      const formUrlCol = formType === 'feedback' ? 16 : (formType === 'survey' ? 17 : 18);
      const formIdCol = formType === 'feedback' ? 19 : (formType === 'survey' ? 20 : 21);
      
      for (let i = 1; i < sessData.length; i++) {
        if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
          // Ensure we have enough columns in the sheet (up to column U = 21)
          const currentLastCol = sessSheet.getLastColumn();
          if (currentLastCol < formIdCol) {
            // Expand sheet to have enough columns
            sessSheet.insertColumnsAfter(currentLastCol, formIdCol - currentLastCol);
            console.log('Expanded Sessions sheet to column ' + formIdCol);
          }
          
          // Save published URL (for users to access)
          sessSheet.getRange(i + 1, formUrlCol).setValue(formUrl);
          console.log('Saved form URL to row ' + (i + 1) + ', column ' + formUrlCol + ': ' + formUrl);
          
          // Save form ID (for polling to access)
          sessSheet.getRange(i + 1, formIdCol).setValue(formId);
          console.log('Saved form ID to row ' + (i + 1) + ', column ' + formIdCol + ': ' + formId);
          
          break;
        }
      }
      
      // NOTE: Per-form triggers are no longer created here.
      // The polling system (pollAllFormResponses) handles response collection
      // for ALL forms every 10 minutes. This eliminates the 20-trigger limit.
      // To enable polling, run: setupFormResponsePolling()
      
      // Legacy trigger creation - disabled in favor of polling
      // Uncomment below if you need instant response capture for specific forms
      /*
      let triggerCreated = false;
      try {
        const triggerFunction = formType === 'feedback' ? 'onFeedbackFormSubmit' : 'onSurveyFormSubmit';
        ScriptApp.newTrigger(triggerFunction)
          .forForm(formId)
          .onFormSubmit()
          .create();
        triggerCreated = true;
        console.log('Trigger created successfully for ' + triggerFunction);
      } catch (triggerError) {
        console.error('Trigger creation failed:', triggerError.message);
      }
      */
      
      return {
        success: true,
        formUrl: formUrl,
        editUrl: editUrl,
        movedToFolder: movedToFolder,
        triggerCreated: false, // Polling handles response collection now
        pollingEnabled: true
      };
    } catch (e) {
      console.error('createSessionForm error:', e);
      return { success: false, message: e.message };
    }
  });
}

/**
 * Link an existing Google Form to a session
 */
function linkSessionForm(sessionToken, sessionId, formType, formUrl) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
      const sessData = sessSheet.getDataRange().getValues();
      
      for (let i = 1; i < sessData.length; i++) {
        if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
          const col = formType === 'feedback' ? 16 : (formType === 'survey' ? 17 : 18); // P, Q, R
          sessSheet.getRange(i + 1, col).setValue(formUrl);
          break;
        }
      }
      
      return { success: true };
    } catch (e) {
      console.error('linkSessionForm error:', e);
      return { success: false, message: e.message };
    }
  });
}

/**
 * Unlink a form from a session
 */
function unlinkSessionForm(sessionToken, sessionId, formType) {
  return withSession_(sessionToken, (userData) => {
    try {
      const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
      const sessData = sessSheet.getDataRange().getValues();
      
      for (let i = 1; i < sessData.length; i++) {
        if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
          const col = formType === 'feedback' ? 16 : (formType === 'survey' ? 17 : 18); // P, Q, R
          sessSheet.getRange(i + 1, col).setValue('');
          break;
        }
      }
      
      return { success: true };
    } catch (e) {
      console.error('unlinkSessionForm error:', e);
      return { success: false, message: e.message };
    }
  });
}

function sendFormToParticipants(sessionToken, sessionId, formType, recipientIds) {
  return { success: false, sent: 0, failed: 0, message: 'Email disabled in demo.' };
}

/**
 * Trigger function for feedback form submissions
 * Stores ALL responses in FormResponses (centralized)
 */
function onFeedbackFormSubmit(e) {
  try {
    const response = e.response;
    const form = FormApp.openById(e.source.getId());
    const description = form.getDescription() || '';
    
    // Extract session ID from form description (new format: TMS_SESSION:xxx|TMS_TYPE:feedback)
    let sessionId = null;
    const newFormatMatch = description.match(/TMS_SESSION:([^|]+)/);
    const oldFormatMatch = description.match(/Session ID:\s*(\S+)/);
    
    if (newFormatMatch) {
      sessionId = newFormatMatch[1].trim();
    } else if (oldFormatMatch) {
      sessionId = oldFormatMatch[1].trim();
    }
    
    if (!sessionId) {
      console.error('Session ID not found in form description:', description);
      return;
    }
    
    console.log('Processing feedback for session:', sessionId);
    
    // Get all item responses
    const itemResponses = response.getItemResponses();
    if (itemResponses.length === 0) {
      console.log('No item responses found');
      return;
    }
    
    const timestamp = new Date();
    const responseId = response.getId();
    
    // Extract all answers (up to 10 questions)
    const answers = [];
    const rawData = {};
    
    itemResponses.forEach((ir, index) => {
      const question = ir.getItem().getTitle();
      const answer = ir.getResponse();
      answers[index] = answer;
      rawData[question] = answer;
    });
    
    // Pad answers array to 10 elements
    while (answers.length < 10) {
      answers.push('');
    }
    
    // Get spreadsheet
    const ss = getSpreadsheet_();
    
    // Store ALL responses in FormResponses (centralized - single source of truth)
    let formResponsesSheet = ss.getSheetByName('FormResponses');
    if (!formResponsesSheet) {
      formResponsesSheet = ss.insertSheet('FormResponses');
      formResponsesSheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
    }
    
    formResponsesSheet.appendRow([
      responseId,
      sessionId,
      'feedback',
      timestamp,
      answers[0], answers[1], answers[2], answers[3], answers[4],
      answers[5], answers[6], answers[7], answers[8], answers[9],
      JSON.stringify(rawData)
    ]);
    
    // Update session's average score in Sessions sheet
    updateSessionAvgScore_(sessionId);
    
    console.log('Feedback response saved successfully for session:', sessionId);
    
  } catch (e) {
    console.error('onFeedbackFormSubmit error:', e);
  }
}

/**
 * Trigger function for pre-training survey and assessment form submissions
 * Stores responses in FormResponses (no avg score calculation)
 */
function onSurveyFormSubmit(e) {
  try {
    const response = e.response;
    const form = FormApp.openById(e.source.getId());
    const description = form.getDescription() || '';
    
    // Extract session ID and form type from form description
    let sessionId = null;
    let formType = 'survey';
    
    // Try new format first: TMS_SESSION:xxx|TMS_TYPE:survey
    const newSessionMatch = description.match(/TMS_SESSION:([^|]+)/);
    const newTypeMatch = description.match(/TMS_TYPE:(\S+)/);
    const oldSessionMatch = description.match(/Session ID:\s*(\S+)/);
    const oldTypeMatch = description.match(/Form Type:\s*(\S+)/);
    
    if (newSessionMatch) {
      sessionId = newSessionMatch[1].trim();
      formType = newTypeMatch ? newTypeMatch[1].trim() : 'survey';
    } else if (oldSessionMatch) {
      sessionId = oldSessionMatch[1].trim();
      formType = oldTypeMatch ? oldTypeMatch[1].trim() : 'survey';
    }
    
    if (!sessionId) {
      console.error('Session ID not found in form description:', description);
      return;
    }
    
    console.log('Processing survey/assessment for session:', sessionId, 'type:', formType);
    
    // Get all item responses
    const itemResponses = response.getItemResponses();
    if (itemResponses.length === 0) {
      console.log('No item responses found');
      return;
    }
    
    const timestamp = new Date();
    const responseId = response.getId();
    
    // Extract all answers (up to 10 questions)
    const answers = [];
    const rawData = {};
    
    itemResponses.forEach((ir, index) => {
      const question = ir.getItem().getTitle();
      const answer = ir.getResponse();
      answers[index] = answer;
      rawData[question] = answer;
    });
    
    // Pad answers array to 10 elements
    while (answers.length < 10) {
      answers.push('');
    }
    
    // Get spreadsheet
    const ss = getSpreadsheet_();
    
    // Store ALL responses in FormResponses (centralized)
    let formResponsesSheet = ss.getSheetByName('FormResponses');
    if (!formResponsesSheet) {
      formResponsesSheet = ss.insertSheet('FormResponses');
      formResponsesSheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
    }
    
    formResponsesSheet.appendRow([
      responseId,
      sessionId,
      formType,
      timestamp,
      answers[0], answers[1], answers[2], answers[3], answers[4],
      answers[5], answers[6], answers[7], answers[8], answers[9],
      JSON.stringify(rawData)
    ]);
    
    console.log('Survey response saved successfully for session:', sessionId);
    
  } catch (e) {
    console.error('onSurveyFormSubmit error:', e);
  }
}

/**
 * Update session's average satisfaction score based on FormResponses
 */
function updateSessionAvgScore_(sessionId) {
  try {
    const formResponsesSheet = getSheetSafe_('FormResponses');
    if (!formResponsesSheet) return;
    
    const frData = formResponsesSheet.getDataRange().getValues();
    // Headers: ResponseID(0), SessionID(1), FormType(2), Timestamp(3), Q1(4)...
    let totalScore = 0;
    let count = 0;
    
    for (let i = 1; i < frData.length; i++) {
      if (String(frData[i][1]).trim() === String(sessionId).trim() && 
          String(frData[i][2]).trim().toLowerCase() === 'feedback') {
        const score = parseFloat(frData[i][4]) || 0; // Q1 is column index 4
        if (score > 0) {
          totalScore += score;
          count++;
        }
      }
    }
    
    if (count === 0) return;
    
    const avgScore = totalScore / count;
    
    // Update Sessions sheet Column I (Average Satisfaction Score)
    const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
    const sessData = sessSheet.getDataRange().getValues();
    
    for (let i = 1; i < sessData.length; i++) {
      if (String(sessData[i][1]).trim() === String(sessionId).trim()) {
        sessSheet.getRange(i + 1, 9).setValue(avgScore.toFixed(1)); // Column I
        break;
      }
    }
  } catch (e) {
    console.error('updateSessionAvgScore_ error:', e);
  }
}

/**
 * Manually recalculate all session average scores (utility function)
 */
function recalculateAllAvgScores() {
  const formResponsesSheet = getSheetSafe_('FormResponses');
  if (!formResponsesSheet) {
    console.log('FormResponses sheet not found');
    return;
  }
  
  const frData = formResponsesSheet.getDataRange().getValues();
  const sessionScores = {};
  
  // Aggregate scores by session from feedback responses
  for (let i = 1; i < frData.length; i++) {
    const sessionId = String(frData[i][1]).trim();
    const formType = String(frData[i][2]).trim().toLowerCase();
    
    if (formType !== 'feedback') continue;
    
    const score = parseFloat(frData[i][4]) || 0; // Q1 is column index 4
    if (score <= 0) continue;
    
    if (!sessionScores[sessionId]) {
      sessionScores[sessionId] = { total: 0, count: 0 };
    }
    sessionScores[sessionId].total += score;
    sessionScores[sessionId].count++;
  }
  
  // Update each session
  const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
  const sessData = sessSheet.getDataRange().getValues();
  
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = String(sessData[i][1]).trim();
    if (sessionScores[sessionId]) {
      const avg = sessionScores[sessionId].total / sessionScores[sessionId].count;
      sessSheet.getRange(i + 1, 9).setValue(avg.toFixed(1));
    }
  }
  
  console.log('Recalculated scores for ' + Object.keys(sessionScores).length + ' sessions');
}

// ============================================
// FORM ARCHIVAL FUNCTIONS (90-day cleanup)
// ============================================

/**
 * Archive forms older than 90 days
 * Moves forms to an "Archive" subfolder instead of deleting
 * Run this manually or set up a weekly time-based trigger
 * 
 * To set up automatic archival:
 * 1. Go to Triggers (clock icon in Apps Script)
 * 2. Add Trigger > archiveOldForms > Time-driven > Weekly
 */
function archiveOldForms() {
  const formsFolderId = getFormsFolderId_();
  if (!formsFolderId) {
    console.log('FORMS_FOLDER_ID not configured in Script Properties');
    return { archived: 0, message: 'FORMS_FOLDER_ID not configured in Script Properties' };
  }
  
  try {
    const formsFolder = DriveApp.getFolderById(formsFolderId);
    
    // Create or get Archive subfolder
    let archiveFolder;
    const archiveFolders = formsFolder.getFoldersByName('Archive');
    if (archiveFolders.hasNext()) {
      archiveFolder = archiveFolders.next();
    } else {
      archiveFolder = formsFolder.createFolder('Archive');
      console.log('Created Archive subfolder');
    }
    
    // Get all sessions to check dates
    const sessSheet = getSheet_(SHEET_NAMES.SESSIONS);
    const sessData = sessSheet.getDataRange().getValues();
    
    // Build map of form URLs to session dates
    const formUrlDates = {};
    for (let i = 1; i < sessData.length; i++) {
      const sessionDate = sessData[i][5]; // Column F = Complete Date
      const feedbackUrl = sessData[i][16] || '';  // Column Q
      const surveyUrl = sessData[i][17] || '';    // Column R
      const assessmentUrl = sessData[i][18] || ''; // Column S
      
      if (sessionDate) {
        const dateObj = new Date(sessionDate);
        if (feedbackUrl) formUrlDates[feedbackUrl] = dateObj;
        if (surveyUrl) formUrlDates[surveyUrl] = dateObj;
        if (assessmentUrl) formUrlDates[assessmentUrl] = dateObj;
      }
    }
    
    // Calculate 90 days ago
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 90);
    
    let archivedCount = 0;
    const files = formsFolder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Skip if it's not a form
      if (file.getMimeType() !== 'application/vnd.google-apps.form') {
        continue;
      }
      
      // Check if this form is linked to an old session
      let formUrl = null;
      try {
        const form = FormApp.openById(file.getId());
        formUrl = form.getPublishedUrl();
      } catch (e) {
        // Can't open as form, use file creation date
      }
      
      let shouldArchive = false;
      
      if (formUrl && formUrlDates[formUrl]) {
        // Check if session date is older than 90 days
        shouldArchive = formUrlDates[formUrl] < cutoffDate;
      } else {
        // No linked session found, check file creation date
        shouldArchive = file.getDateCreated() < cutoffDate;
      }
      
      if (shouldArchive) {
        // Move to archive folder
        archiveFolder.addFile(file);
        formsFolder.removeFile(file);
        archivedCount++;
        console.log('Archived: ' + fileName);
      }
    }
    
    const message = `Archived ${archivedCount} form(s) older than 90 days`;
    console.log(message);
    return { archived: archivedCount, message: message };
    
  } catch (e) {
    console.error('archiveOldForms error:', e);
    return { archived: 0, message: e.message };
  }
}

/**
 * View archive statistics
 */
function getArchiveStats() {
  const formsFolderId = getFormsFolderId_();
  if (!formsFolderId) {
    return { active: 0, archived: 0 };
  }
  
  try {
    const formsFolder = DriveApp.getFolderById(formsFolderId);
    
    // Count active forms
    let activeCount = 0;
    const activeFiles = formsFolder.getFiles();
    while (activeFiles.hasNext()) {
      if (activeFiles.next().getMimeType() === 'application/vnd.google-apps.form') {
        activeCount++;
      }
    }
    
    // Count archived forms
    let archivedCount = 0;
    const archiveFolders = formsFolder.getFoldersByName('Archive');
    if (archiveFolders.hasNext()) {
      const archiveFolder = archiveFolders.next();
      const archivedFiles = archiveFolder.getFiles();
      while (archivedFiles.hasNext()) {
        if (archivedFiles.next().getMimeType() === 'application/vnd.google-apps.form') {
          archivedCount++;
        }
      }
    }
    
    return { active: activeCount, archived: archivedCount };
  } catch (e) {
    console.error('getArchiveStats error:', e);
    return { active: 0, archived: 0 };
  }
}

/**
 * Test function to verify form folder and trigger setup
 * Run this manually in Apps Script to check configuration
 */
function testFormConfiguration() {
  const results = {
    spreadsheetId: null,
    spreadsheetAccess: false,
    formsFolderId: null,
    formsFolderAccess: false,
    triggers: [],
    sheets: []
  };
  
  // Check spreadsheet
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    results.spreadsheetId = ssId;
    if (ssId) {
      const ss = SpreadsheetApp.openById(ssId);
      results.spreadsheetAccess = true;
      results.sheets = ss.getSheets().map(s => s.getName());
    }
  } catch (e) {
    results.spreadsheetError = e.message;
  }
  
  // Check forms folder
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('FORMS_FOLDER_ID');
    results.formsFolderId = folderId;
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      results.formsFolderAccess = true;
      results.formsFolderName = folder.getName();
    }
  } catch (e) {
    results.formsFolderError = e.message;
  }
  
  // Check triggers
  try {
    const triggers = ScriptApp.getProjectTriggers();
    results.triggers = triggers.map(t => ({
      function: t.getHandlerFunction(),
      type: t.getEventType().toString(),
      source: t.getTriggerSourceId()
    }));
  } catch (e) {
    results.triggersError = e.message;
  }
  
  console.log('Form Configuration Test Results:');
  console.log(JSON.stringify(results, null, 2));
  
  return results;
}

/**
 * Manually sync existing form responses to FormResponses sheet
 * Use this if triggers weren't working previously
 */
function syncExistingFormResponses() {
  const ss = getSpreadsheet_();
  const sessSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  const sessData = sessSheet.getDataRange().getValues();
  
  let formResponsesSheet = ss.getSheetByName('FormResponses');
  if (!formResponsesSheet) {
    formResponsesSheet = ss.insertSheet('FormResponses');
    formResponsesSheet.appendRow(['ResponseID', 'SessionID', 'FormType', 'Timestamp', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'RawJSON']);
  }
  
  // Get existing response IDs to avoid duplicates
  const existingIds = new Set();
  const frData = formResponsesSheet.getDataRange().getValues();
  for (let i = 1; i < frData.length; i++) {
    existingIds.add(String(frData[i][0]));
  }
  
  let synced = 0;
  
  // Loop through sessions and check for forms
  for (let i = 1; i < sessData.length; i++) {
    const sessionId = sessData[i][1];
    const feedbackUrl = sessData[i][16]; // Column Q
    const surveyUrl = sessData[i][17];   // Column R  
    const assessmentUrl = sessData[i][18]; // Column S
    
    const formsToSync = [
      { url: feedbackUrl, type: 'feedback' },
      { url: surveyUrl, type: 'survey' },
      { url: assessmentUrl, type: 'assessment' }
    ];
    
    for (const formInfo of formsToSync) {
      if (!formInfo.url) continue;
      
      try {
        // Extract form ID from URL
        const formIdMatch = formInfo.url.match(/\/forms\/d\/([a-zA-Z0-9_-]+)/);
        if (!formIdMatch) continue;
        
        const formId = formIdMatch[1];
        const form = FormApp.openById(formId);
        const responses = form.getResponses();
        
        for (const response of responses) {
          const responseId = response.getId();
          if (existingIds.has(responseId)) continue;
          
          const itemResponses = response.getItemResponses();
          const answers = [];
          const rawData = {};
          
          itemResponses.forEach((ir, index) => {
            const question = ir.getItem().getTitle();
            const answer = ir.getResponse();
            answers[index] = answer;
            rawData[question] = answer;
          });
          
          while (answers.length < 10) {
            answers.push('');
          }
          
          formResponsesSheet.appendRow([
            responseId,
            sessionId,
            formInfo.type,
            response.getTimestamp(),
            answers[0], answers[1], answers[2], answers[3], answers[4],
            answers[5], answers[6], answers[7], answers[8], answers[9],
            JSON.stringify(rawData)
          ]);
          
          existingIds.add(responseId);
          synced++;
        }
      } catch (e) {
        console.log('Error syncing form ' + formInfo.url + ':', e.message);
      }
    }
  }
  
  console.log('Synced ' + synced + ' responses');
  return { synced: synced };
}