const SPREADSHEET_ID = "1YIM_Zj_V8Ck0vwtEGEyxuQKSmA7r7KAsGNwwRjeqeJE"; // <<-- แทนที่ด้วย ID Google Sheet ของคุณ
const TEMPLATE_SHEET_NAME = "Data_Template"; // <<-- ชื่อชีตต้นฉบับที่คุณต้องสร้าง
const DAILY_SHEET_PREFIX = "Data_"; // คำนำหน้าสำหรับชีตข้อมูลรายวัน

// กำหนดดัชนีคอลัมน์ (Column Index) สำหรับการเข้าถึงข้อมูลโดยตรง
// A=0, B=1, C=2, ...
const COL_DATE = 0; // คอลัมน์ A: Date
const COL_MC = 1; // คอลัมน์ B: Machine
const COL_SPEC = 2; // คอลัมน์ C: Spec
const COL_PART_NO = 3; // คอลัมน์ D: Part No.
const COL_PROCESS = 4; // คอลัมน์ E: Process
const COL_OP = 5; // คอลัมน์ F: OP
const COL_MAN = 6; // คอลัมน์ G: Man
const COL_CAP_HR = 7; // คอลัมน์ H: Cap/Hr.
const COL_PMC_PLAN_DETERMINES = 8; // คอลัมน์ I: PMC Plan Determines
const COL_ACTUAL_PRODUCTION_HOURS = 9; // คอลัมน์ J: Actual Production Hours
const COL_PLAN_QUANTITY = 10; // คอลัมน์ K: Plan Quantity
const COL_ACTUAL_QUANTITY = 11; // คอลัมน์ L: Actual Quantity
const COL_BALANCE = 12; // คอลัมน์ M: Balance
const COL_REMARK = 13; // คอลัมน์ N: Remark


function doGet() {
  Logger.log("doGet function called.");
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Production Dashboard');
}

function doPost(e) {
  Logger.log("doPost function called with event: " + JSON.stringify(e));

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Spreadsheet opened: " + SPREADSHEET_ID);

    let dataToProcess;
    if (e && e.postData && e.postData.contents) {
      dataToProcess = JSON.parse(e.postData.contents);
      Logger.log("Received data for processing (from POST): " + JSON.stringify(dataToProcess));
    } else if (Array.isArray(e)) {
      dataToProcess = e;
      Logger.log("Received data for processing (from google.script.run): " + JSON.stringify(dataToProcess));
    } else if (typeof e === 'object' && e !== null) {
      dataToProcess = [e];
      Logger.log("Received single object for processing (from google.script.run): " + JSON.stringify(dataToProcess));
    } else {
      Logger.log("Error: No valid data received.");
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No valid data received" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (dataToProcess.length === 0) {
      Logger.log("No valid data rows to append.");
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "No valid data rows to append" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    dataToProcess.forEach(data => {
      // สำหรับแต่ละแถว ให้หาหรือสร้างชีทตามวันที่ในข้อมูลนั้นๆ
      const targetDateString = data.date; // วันที่จาก frontend จะเป็น yyyy-MM-dd หรือ yyyy/MM/dd
      const sheet = getOrCreateDailySheet(ss, targetDateString); // ส่งวันที่ไปให้ getOrCreateDailySheet

      if (!sheet) {
        Logger.log("Error: Daily sheet could not be created or found for date: " + targetDateString);
        // อาจจะข้ามแถวนี้ไป หรือส่ง error กลับไป
        return; // ข้ามแถวนี้ไป
      }
      Logger.log("Target sheet for doPost (for this row): " + sheet.getName());

      const rowData = [];
      // เปลี่ยนรูปแบบวันที่จาก yyyy-mm-dd หรือ yyyy/mm/dd เป็น yyyy/mm/dd สำหรับบันทึก
      rowData[COL_DATE] = data.date ? data.date.replace(/-/g, '/') : ''; 
      rowData[COL_MC] = data.mc;
      rowData[COL_SPEC] = data.spec;
      rowData[COL_PART_NO] = data.partNo;
      rowData[COL_PROCESS] = data.process;
      rowData[COL_OP] = data.op;
      rowData[COL_MAN] = parseFloat(data.man) || 0;
      rowData[COL_CAP_HR] = parseFloat(data.capHr) || 0;
      rowData[COL_PMC_PLAN_DETERMINES] = parseFloat(data.pmcPlanDetermines) || 0;
      rowData[COL_ACTUAL_PRODUCTION_HOURS] = parseFloat(data.actualProductionHours) || 0;
      rowData[COL_PLAN_QUANTITY] = parseFloat(data.planQuantity) || 0;
      rowData[COL_ACTUAL_QUANTITY] = parseFloat(data.actualQuantity) || 0;
      rowData[COL_BALANCE] = parseFloat(data.balance) || 0;
      rowData[COL_REMARK] = data.remark;
      
      Logger.log("Appending row to " + sheet.getName() + ": " + JSON.stringify(rowData));
      sheet.appendRow(rowData);
    });

    Logger.log("All data appended successfully.");

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data appended successfully" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error in doPost: " + error.message);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Server error: " + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Gets or creates a daily sheet based on the provided date string or current date.
 * Copies from TEMPLATE_SHEET_NAME if the daily sheet doesn't exist.
 * Moves the new sheet to the first position.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {string} [dateString=null] Optional: The date string in 'yyyy-MM-dd' or 'yyyy/MM/dd' format. If null, uses current date.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The daily sheet for the specified or current date.
*/
function getOrCreateDailySheet(ss, dateString = null) {
  let targetDate;
  if (dateString) {
    // Try parsing yyyy-MM-dd first (from input type="date")
    // Then try parsing yyyy/MM/dd (from contenteditable cell or if user types it)
    try {
      targetDate = Utilities.parseDate(dateString, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } catch (e) {
      try {
        targetDate = Utilities.parseDate(dateString, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } catch (e2) {
        Logger.log("Error parsing dateString in getOrCreateDailySheet: " + dateString + ". Error 1: " + e.message + ", Error 2: " + e2.message);
        throw new Error("Invalid date format provided: " + dateString);
      }
    }
  } else {
    targetDate = new Date();
  }
  
  // Format the date to yyyy/MM/dd for the sheet name
  const sheetName = DAILY_SHEET_PREFIX + Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  Logger.log("Attempting to get or create sheet: " + sheetName);

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet " + sheetName + " not found. Creating from template.");
    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) {
      throw new Error("Template sheet '" + TEMPLATE_SHEET_NAME + "' not found. Please create it.");
    }
    sheet = templateSheet.copyTo(ss);
    sheet.setName(sheetName);
    ss.setActiveSheet(sheet); // Activate the new sheet
    ss.moveActiveSheet(1); // Move to the first position (index 1)
    Logger.log("Sheet " + sheetName + " created and moved to first position.");
  } else {
    Logger.log("Sheet " + sheetName + " already exists.");
  }
  return sheet;
}

// Helper to get week number (ไม่เปลี่ยนแปลง)
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  return weekNo;
}

/**
 * Retrieves dashboard data from all relevant daily sheets based on filter options.
 * @param {object} filterOptions Options for filtering data (filterType, startDate, endDate, weekNumber, year).
 * @returns {object} An object containing detailed data, weekly summary, and monthly summary.
*/
function getDashboardData(filterOptions = {}) {
  Logger.log("getDashboardData function called with filterOptions: " + JSON.stringify(filterOptions));
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    let sheetsToProcess = [];

    // Identify daily sheets based on naming convention
    const dailySheets = allSheets.filter(sheet => sheet.getName().startsWith(DAILY_SHEET_PREFIX));
    Logger.log("Found " + dailySheets.length + " daily sheets.");

    if (filterOptions.filterType === 'date' && filterOptions.startDate && filterOptions.endDate) {
      // ใช้ Utilities.parseDate เพื่อให้แน่ใจว่าโซนเวลาสอดคล้องกันสำหรับวันที่จาก UI
      const startDate = Utilities.parseDate(filterOptions.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const endDate = Utilities.parseDate(filterOptions.endDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      endDate.setHours(23, 59, 59, 999); // Set to end of day

      Logger.log("Parsed startDate (script timezone): " + startDate);
      Logger.log("Parsed endDate (script timezone, end of day): " + endDate);

      sheetsToProcess = dailySheets.filter(sheet => {
        try {
          // ต้องปรับการ parse ชื่อชีทให้ตรงกับรูปแบบ yyyy/MM/dd
          const sheetDateStr = sheet.getName().substring(DAILY_SHEET_PREFIX.length); // e.g., "YYYY/MM/DD"
          const [year, month, day] = sheetDateStr.split('/').map(Number);
          // สร้าง Date object ในโซนเวลาของสคริปต์
          const sheetDate = new Date(year, month - 1, day); 
          Logger.log(`Comparing sheet: ${sheet.getName()} (${sheetDate}) with range: ${startDate} to ${endDate}`);
          return sheetDate >= startDate && sheetDate <= endDate;
        } catch (e) {
          Logger.log("Error parsing sheet name date for sheet: " + sheet.getName() + " - " + e.message);
          return false;
        }
      });
      Logger.log("Sheets filtered by date. Processing " + sheetsToProcess.length + " sheets.");

    } else if (filterOptions.filterType === 'week' && filterOptions.weekNumber && filterOptions.year) {
      const targetWeek = parseInt(filterOptions.weekNumber);
      const targetYear = parseInt(filterOptions.year);

      sheetsToProcess = dailySheets.filter(sheet => {
        try {
          // ต้องปรับการ parse ชื่อชีทให้ตรงกับรูปแบบ yyyy/MM/dd
          const sheetDateStr = sheet.getName().substring(DAILY_SHEET_PREFIX.length);
          const [year, month, day] = sheetDateStr.split('/').map(Number);
          const sheetDate = new Date(year, month - 1, day);
          const rowYear = sheetDate.getFullYear();
          const rowWeek = getWeekNumber(sheetDate);
          return rowYear === targetYear && rowWeek === targetWeek;
        } catch (e) {
          Logger.log("Error parsing sheet name date for week filter sheet: " + sheet.getName() + " - " + e.message);
          return false;
        }
      });
      Logger.log("Sheets filtered by week. Processing " + sheetsToProcess.length + " sheets.");

    } else {
      // Default: process only the current day's sheet if no filter or invalid filter
      const currentDailySheet = getOrCreateDailySheet(ss); // Ensure current day's sheet exists
      sheetsToProcess = [currentDailySheet];
      Logger.log("No specific filter. Processing current daily sheet: " + currentDailySheet.getName());
    }

    let combinedRawData = [];
    if (sheetsToProcess.length === 0) {
      Logger.log("No sheets to process after filtering.");
      return { status: "success", detailedData: [], weeklySummary: [], monthlySummary: [], dailySummary: [], yearlySummary: [] };
    }

    // Sort sheets by date to ensure consistent order in dashboard
    sheetsToProcess.sort((a, b) => {
      // ต้องปรับการ parse ชื่อชีทให้ตรงกับรูปแบบ yyyy/MM/dd
      const dateA = new Date(a.getName().substring(DAILY_SHEET_PREFIX.length).replace(/\//g, '-'));
      const dateB = new Date(b.getName().substring(DAILY_SHEET_PREFIX.length).replace(/\//g, '-'));
      return dateA.getTime() - dateB.getTime();
    });


    sheetsToProcess.forEach(sheet => {
      const range = sheet.getDataRange();
      const values = range.getValues();

      if (values.length > 1) { // Skip header only sheets
        values.slice(1).forEach((row, index) => { // Data starts from row 2 (index 1)
          let dateFromSheet = row[COL_DATE];
          let parsedDate = null;
          Logger.log(`Raw date from sheet (${sheet.getName()}, row ${index + 2}): ${dateFromSheet} (Type: ${typeof dateFromSheet})`);

          if (dateFromSheet instanceof Date) {
              parsedDate = dateFromSheet;
          } else if (typeof dateFromSheet === 'string') {
              // Try to parse YYYY-MM-DD
              if (dateFromSheet.match(/^\d{4}-\d{2}-\d{2}$/)) {
                  parsedDate = Utilities.parseDate(dateFromSheet, Session.getScriptTimeZone(), "yyyy-MM-dd");
              } 
              // Try to parse YYYY/MM/DD
              else if (dateFromSheet.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
                  const standardizedDateStr = dateFromSheet.replace(/\//g, '-');
                  parsedDate = Utilities.parseDate(standardizedDateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
              } 
              // Fallback for other string formats using new Date()
              else {
                  const tempDate = new Date(dateFromSheet);
                  if (!isNaN(tempDate.getTime())) { // Check if it's a valid date
                      const formattedDateStr = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
                      parsedDate = Utilities.parseDate(formattedDateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
                  } else {
                      Logger.log(`Warning: Could not parse date string '${dateFromSheet}' from sheet. Keeping original value.`);
                      parsedDate = dateFromSheet; // Keep original if parsing fails
                  }
              }
          } else if (dateFromSheet) {
              // If it's not a Date object or string, but not null/undefined, try new Date()
              const tempDate = new Date(dateFromSheet);
              if (!isNaN(tempDate.getTime())) {
                  const formattedDateStr = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
                  parsedDate = Utilities.parseDate(formattedDateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
              } else {
                  Logger.log(`Warning: Could not parse date value '${dateFromSheet}' from sheet. Keeping original value.`);
                  parsedDate = dateFromSheet;
              }
          }
          
          Logger.log(`Parsed date from sheet (script timezone): ${parsedDate}`);

          combinedRawData.push({
            sheetName: sheet.getName(), // Store sheet name
            originalRowIndex: index + 2, // Store 1-based index for sheet operations (header is row 1, data starts from row 2)
            Date: parsedDate, // ใช้ Date object ที่ถูกแปลงแล้ว
            MC: row[COL_MC],
            Spec: row[COL_SPEC],
            PartNo: row[COL_PART_NO],
            Process: row[COL_PROCESS],
            OP: row[COL_OP],
            Man: row[COL_MAN],
            CapHr: row[COL_CAP_HR],
            PMCPlanDetermines: row[COL_PMC_PLAN_DETERMINES],
            ActualProductionHours: row[COL_ACTUAL_PRODUCTION_HOURS],
            PlanQuantity: row[COL_PLAN_QUANTITY],
            ActualQuantity: row[COL_ACTUAL_QUANTITY],
            Balance: row[COL_BALANCE],
            Remark: row[COL_REMARK]
          });
        });
      }
    });
    Logger.log("Combined raw data from all relevant sheets: " + combinedRawData.length + " rows.");
    if (combinedRawData.length > 0) {
      Logger.log("Sample of combinedRawData (first row): " + JSON.stringify(combinedRawData[0]));
    }

    // --- Process and calculate detailed data for the new table ---
    const detailedData = combinedRawData.map(row => {
      const planQuantity = parseFloat(row.PlanQuantity) || 0;
      const actualQuantity = parseFloat(row.ActualQuantity) || 0;
      const pmcPlanDetermines = parseFloat(row.PMCPlanDetermines) || 0;
      const actualProductionHours = parseFloat(row.ActualProductionHours) || 0;
      const capHr = parseFloat(row.CapHr) || 0;

      const actualProductionPercent = (planQuantity > 0) ? (actualQuantity / planQuantity * 100) : 0;
      const planPercent = (planQuantity > 0) ? (actualQuantity / planQuantity * 100) : 0;
      const capPercent = (capHr > 0) ? (actualProductionHours / capHr * 100) : 0;
      const hrOvertime = actualProductionHours - pmcPlanDetermines;

      // Reverted the filtering for machine values with '/' for now
      let machineValue = row.MC || ''; 

      return {
        sheetName: String(row.sheetName || ''), // Ensure string
        originalRowIndex: row.originalRowIndex, 
        // เปลี่ยนรูปแบบวันที่สำหรับแสดงผลในตาราง Detailed Production Data
        date: row.Date ? Utilities.formatDate(row.Date, Session.getScriptTimeZone(), "yyyy/MM/dd") : '', 
        partNo: String(row.PartNo || ''), // Ensure string
        machine: String(machineValue || ''), // Ensure string
        spec: String(row.Spec || ''), // Ensure string
        process: String(row.Process || ''), // Ensure string
        op: String(row.OP || ''), // Ensure string
        man: parseFloat(row.Man) || 0,
        capHr: capHr,
        pmcPlanDeterminesPlan: planQuantity,
        pmcPlanDeterminesHr: pmcPlanDetermines,
        actualProductionActual: actualQuantity,
        actualProductionHr: actualProductionHours,
        balance: parseFloat(row.Balance) || 0,
        remark: String(row.Remark || ''), // Ensure string
        actualProductionPercent: actualProductionPercent.toFixed(1) + '%',
        planPercent: planPercent.toFixed(1) + '%',
        capPercent: capPercent.toFixed(1) + '%',
        hrOvertime: hrOvertime.toFixed(1)
      };
    });
    Logger.log("Detailed data processed.");
    if (detailedData.length > 0) {
      Logger.log("Sample of detailedData (first row): " + JSON.stringify(detailedData[0]));
    }

    // --- การสรุปข้อมูลรายวัน --- NEW
    const dailySummary = {};
    combinedRawData.forEach(row => {
      const date = row.Date; // ใช้ Date object ที่ถูกแปลงแล้ว
      // เปลี่ยนรูปแบบวันที่สำหรับ key ของ dailySummary
      const dateKey = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
      if (!dailySummary[dateKey]) {
        dailySummary[dateKey] = { date: dateKey, plan: 0, actual: 0, balance: 0 };
      }
      dailySummary[dateKey].plan += (parseFloat(row.PlanQuantity) || 0);
      dailySummary[dateKey].actual += (parseFloat(row.ActualQuantity) || 0);
      dailySummary[dateKey].balance += (parseFloat(row.ActualQuantity) || 0) - (parseFloat(row.PlanQuantity) || 0);
    });
    const sortedDailySummary = Object.values(dailySummary).sort((a, b) => b.date.localeCompare(a.date));
    Logger.log("Daily summary generated.");

    // --- การสรุปข้อมูลรายสัปดาห์ (ใช้ข้อมูลที่ถูกกรองแล้ว) ---
    const weeklySummary = {};
    combinedRawData.forEach(row => {
      const date = row.Date; // ใช้ Date object ที่ถูกแปลงแล้ว
      const year = date.getFullYear();
      const week = getWeekNumber(date);
      const weekKey = `${year}-W${String(week).padStart(2, '0')}`;

      if (!weeklySummary[weekKey]) {
        weeklySummary[weekKey] = { week: weekKey, plan: 0, actual: 0, balance: 0 }; // Added plan, actual, balance
      }
      weeklySummary[weekKey].plan += (parseFloat(row.PlanQuantity) || 0);
      weeklySummary[weekKey].actual += (parseFloat(row.ActualQuantity) || 0);
      weeklySummary[weekKey].balance += (parseFloat(row.ActualQuantity) || 0) - (parseFloat(row.PlanQuantity) || 0);
    });
    const sortedWeeklySummary = Object.values(weeklySummary).sort((a, b) => b.week.localeCompare(a.week));
    Logger.log("Weekly summary generated.");

    // --- การสรุปข้อมูลรายเดือน (ใช้ข้อมูลที่ถูกกรองแล้ว) ---
    const monthlySummary = {};
    combinedRawData.forEach(row => {
      const date = row.Date; // ใช้ Date object ที่ถูกแปลงแล้ว
      const year = date.getFullYear();
      const month = date.getMonth() + 1;
      const monthKey = `${year}-${String(month).padStart(2, '0')}`;

      if (!monthlySummary[monthKey]) {
        monthlySummary[monthKey] = { month: monthKey, plan: 0, actual: 0, balance: 0 }; // Added balance
      }
      monthlySummary[monthKey].plan += (parseFloat(row.PlanQuantity) || 0);
      monthlySummary[monthKey].actual += (parseFloat(row.ActualQuantity) || 0);
      monthlySummary[monthKey].balance += (parseFloat(row.ActualQuantity) || 0) - (parseFloat(row.PlanQuantity) || 0);
    });
    const sortedMonthlySummary = Object.values(monthlySummary).sort((a, b) => b.month.localeCompare(a.month));
    Logger.log("Monthly summary generated.");

    // --- การสรุปข้อมูลรายปี --- NEW
    const yearlySummary = {};
    combinedRawData.forEach(row => {
      const date = row.Date; // ใช้ Date object ที่ถูกแปลงแล้ว
      const yearKey = String(date.getFullYear());
      if (!yearlySummary[yearKey]) {
        yearlySummary[yearKey] = { year: yearKey, plan: 0, actual: 0, balance: 0 };
      }
      yearlySummary[yearKey].plan += (parseFloat(row.PlanQuantity) || 0);
      yearlySummary[yearKey].actual += (parseFloat(row.ActualQuantity) || 0);
      yearlySummary[yearKey].balance += (parseFloat(row.ActualQuantity) || 0) - (parseFloat(row.PlanQuantity) || 0);
    });
    const sortedYearlySummary = Object.values(yearlySummary).sort((a, b) => b.year.localeCompare(a.year));
    Logger.log("Yearly summary generated.");

    return {
      status: "success",
      detailedData: detailedData,
      weeklySummary: sortedWeeklySummary,
      monthlySummary: sortedMonthlySummary,
      dailySummary: sortedDailySummary, // NEW
      yearlySummary: sortedYearlySummary // NEW
    };

  } catch (e) {
    Logger.log("Error in getDashboardData: " + e.message);
    return { status: "error", message: "Server error in dashboard data: " + e.message };
  }
}

/**
 * Retrieves a list of unique machine names from all daily sheets.
 * @returns {string[]} An array of unique machine names.
 */
function getUniqueMachineNames() {
  Logger.log("getUniqueMachineNames function called.");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const uniqueMachines = new Set();

    const dailySheets = allSheets.filter(sheet => sheet.getName().startsWith(DAILY_SHEET_PREFIX));

    dailySheets.forEach(sheet => {
      const range = sheet.getDataRange();
      const values = range.getValues();

      if (values.length > 1) { // Skip header only sheets
        values.slice(1).forEach(row => {
          const machineName = String(row[COL_MC] || '').trim();
          if (machineName) {
            uniqueMachines.add(machineName);
          }
        });
      }
    });

    const sortedMachines = Array.from(uniqueMachines).sort();
    Logger.log("Unique machine names found: " + sortedMachines.length);
    return sortedMachines;

  } catch (e) {
    Logger.log("Error in getUniqueMachineNames: " + e.message);
    return []; // Return empty array on error
  }
}


/**
 * Updates a row in a specific Google Sheet.
 * @param {string} sheetName The name of the sheet to update.
 * @param {number} rowIndex The 1-based index of the row to update.
 * @param {object} updatedData The data to update the row with.
 * @returns {object} Status and message.
*/
function updateRowInSheet(sheetName, rowIndex, updatedData) {
  Logger.log("updateRowInSheet called for sheet: " + sheetName + ", row: " + rowIndex + " with data: " + JSON.stringify(updatedData));
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName); // Get sheet by name

    if (!sheet) {
      Logger.log("Error: Sheet '" + sheetName + "' not found in updateRowInSheet.");
      return { status: "error", message: "Sheet not found" };
    }

    // Prepare row values in the correct column order based on COL_ constants
    const rowValues = [];
    // เปลี่ยนรูปแบบวันที่จาก yyyy-mm-dd หรือ yyyy/mm/dd เป็น yyyy/mm/dd สำหรับบันทึก
    rowValues[COL_DATE] = updatedData.date ? updatedData.date.replace(/-/g, '/') : '';
    rowValues[COL_MC] = updatedData.mc;
    rowValues[COL_SPEC] = updatedData.spec;
    rowValues[COL_PART_NO] = updatedData.partNo;
    rowValues[COL_PROCESS] = updatedData.process;
    rowValues[COL_OP] = updatedData.op;
    rowValues[COL_MAN] = parseFloat(updatedData.man) || 0;
    rowValues[COL_CAP_HR] = parseFloat(updatedData.capHr) || 0;
    rowValues[COL_PMC_PLAN_DETERMINES] = parseFloat(updatedData.pmcPlanDetermines) || 0;
    rowValues[COL_ACTUAL_PRODUCTION_HOURS] = parseFloat(updatedData.actualProductionHours) || 0;
    rowValues[COL_PLAN_QUANTITY] = parseFloat(updatedData.planQuantity) || 0;
    rowValues[COL_ACTUAL_QUANTITY] = parseFloat(updatedData.actualQuantity) || 0;
    rowValues[COL_BALANCE] = parseFloat(updatedData.balance) || 0;
    rowValues[COL_REMARK] = updatedData.remark;

    const targetRange = sheet.getRange(rowIndex, 1, 1, rowValues.length);
    targetRange.setValues([rowValues]);

    Logger.log("Row updated successfully at index: " + rowIndex + " in sheet: " + sheetName);
    return { status: "success", message: "Row updated successfully" };

  } catch (e) {
    Logger.log("Error in updateRowInSheet: " + e.message);
    return { status: "error", message: "Server error: " + e.message };
  }
}

/**
 * Deletes a row from a specific Google Sheet.
 * @param {string} sheetName The name of the sheet to delete from.
 * @param {number} rowIndex The 1-based index of the row to delete.
 * @returns {object} Status and message.
*/
function deleteRowInSheet(sheetName, rowIndex) {
  Logger.log("deleteRowInSheet called for sheet: " + sheetName + ", row: " + rowIndex);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName); // Get sheet by name

    if (!sheet) {
      Logger.log("Error: Sheet '" + sheetName + "' not found in deleteRowInSheet.");
      return { status: "error", message: "Sheet not found" };
    }

    if (rowIndex > 1 && rowIndex <= sheet.getLastRow()) {
      sheet.deleteRow(rowIndex);
      Logger.log("Row deleted successfully at index: " + rowIndex + " from sheet: " + sheetName);
      return { status: "success", message: "Row deleted successfully" };
    } else {
      Logger.log("Invalid row index for deletion: " + rowIndex + " in sheet: " + sheetName);
      return { status: "error", message: "Invalid row index for deletion" };
    }

  } catch (e) {
    Logger.log("Error in deleteRowInSheet: " + e.message);
    return { status: "error", message: "Server error: " + e.message };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
