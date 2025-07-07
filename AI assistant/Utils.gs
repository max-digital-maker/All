// =========================================================================================
// ФАЙЛ: Utils.gs
// =========================================================================================

/**
 * Получает лист по имени. Если лист не существует, он будет создан с заданными заголовками.
 * @param {string} sheetName Имя листа.
 * @param {string[]} headers Массив заголовков для нового листа.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Объект листа.
 */

/**
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

*/
// --- Вспомогательная функция для записи общих логов в лог-лист Google Таблицы ---
function directLogToSheet(message) {
  try {
    if (!SPREADSHEET_ID) { 
      Logger.log('Error: SPREADSHEET_ID is not defined for directLogToSheet. Message: ' + message);
      return;
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('05_Логи_Ошибок');
    if (!sheet) {
      Logger.log('Error: Sheet "05_Логи_Ошибок" not found for directLogToSheet. Message: ' + message);
      return;
    }
    sheet.appendRow([new Date().toLocaleString('ru-RU'), 'Direct Log', message, '', 'SYSTEM_TRACE']);
  } catch (e) {
    Logger.log('FATAL ERROR: Could not log to sheet from directLogToSheet: ' + e.toString() + ' Original Message: ' + message);
  }
}

// --- Вспомогательная функция для записи ошибок в лог-лист Google Таблицы ---
function logErrorToSheet(errorType, errorMessage, context, userId) {
  try {
    if (!SPREADSHEET_ID) { 
      Logger.log('Error: SPREADSHEET_ID is not defined for logErrorToSheet. Error: ' + errorMessage);
      return;
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('05_Логи_Ошибок');
    if (!sheet) {
      Logger.log('Error: Sheet "05_Логи_Ошибок" not found for logErrorToSheet. Error: ' + errorMessage);
      return;
    }
    const now = new Date().toLocaleString('ru-RU');
    sheet.appendRow([now, errorType, errorMessage, context, userId]);
    Logger.log(`Logged error: ${errorType} - ${errorMessage}`);
  } catch (e) {
    Logger.log('FATAL ERROR: Could not log error to sheet from logErrorToSheet: ' + e.toString() + ' Original Error: ' + errorMessage);
  }
}

/**
// --- Вспомогательная функция для получения листа таблицы по имени ---
function getSheet(sheetName) {
  try {
    if (!SPREADSHEET_ID) { 
      throw new Error('SPREADSHEET_ID is not defined for getSheet.');
    }
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in spreadsheet ID: ${SPREADSHEET_ID}`);
    }
    return sheet;
  } catch (e) {
    Logger.log('Error getting sheet: ' + e.toString());
    logErrorToSheet('Sheet Access Error', e.toString(), `Attempting to get sheet: ${sheetName}`, 'SYSTEM');
    throw e; 
  }
}
*/

/**
 * Получает лист по имени. Если это системный лист (Логи, История) и его нет, создает его.
 * @param {string} sheetName Имя листа.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Объект листа.
 */
function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    if (sheetName === '05_Логи_Ошибок' || sheetName === CONVERSATION_HISTORY_SHEET_NAME) {
       directLogToSheet(`Sheet "${sheetName}" not found, creating it.`);
       return spreadsheet.insertSheet(sheetName);
    }
    throw new Error(`Sheet "${sheetName}" not found and cannot be auto-created.`);
  }
  return sheet;
}

// --- Вспомогательные функции для форматирования даты и времени ---
function formatDateForSheet(dateOrDateString) {
  if (typeof dateOrDateString === 'string' && dateOrDateString.match(/^\d{2}\.\d{2}\.\d{4}$/)) {
      return dateOrDateString;
  }
  if (dateOrDateString instanceof Date && !isNaN(dateOrDateString.getTime())) {
    return Utilities.formatDate(dateOrDateString, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }
  return ''; 
}

function formatTimeForSheet(dateOrTimeString) {
  if (dateOrTimeString instanceof Date && !isNaN(dateOrTimeString.getTime())) {
    return Utilities.formatDate(dateOrTimeString, Session.getScriptTimeZone(), 'HH:mm');
  }
  if (typeof dateOrTimeString === 'string' && dateOrTimeString.match(/^\d{2}:\d{2}$/)) {
      return dateOrTimeString;
  }
  return '';
}

// --- Вспомогательная функция для парсинга ДД.ММ.ГГГГ в объект Date ---
function parseDdMmYyyyToDate(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }
  const parts = dateString.split('.');
  if (parts.length !== 3 || !parts[0].match(/^\d{2}$/) || !parts[1].match(/^\d{2}$/) || !parts[2].match(/^\d{4}$/)) {
    return null; 
  }
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; 
  const year = parseInt(parts[2], 10);
  
  const dateObj = new Date(year, month, day);
  if (dateObj.getFullYear() !== year || dateObj.getMonth() !== month || dateObj.getDate() !== day) {
      return null; 
  }
  return dateObj;
}

// --- КРИТИЧЕСКИ ВАЖНАЯ ФУНКЦИЯ ДЛЯ ПАРСИНГА СОДЕРЖИМОГО ЯЧЕЕК ДАТЫ/ВРЕМЕНИ ---
function parseDateCellContent(cellValue) {
  if (cellValue instanceof Date) {
      return cellValue; 
  }
  if (typeof cellValue === 'string') {
      const parsedDate = parseDdMmYyyyToDate(cellValue);
      if (parsedDate) return parsedDate;
      const genericDate = new Date(cellValue);
      if (!isNaN(genericDate.getTime())) return genericDate;
  }
  return null;
}

// --- Вспомогательная функция для парсинга длительности ---
function parseDuration(durationString) {
  if (!durationString || typeof durationString !== 'string') return 0;

  const lowerDuration = durationString.toLowerCase();
  let totalMinutes = 0;

  const hoursMatch = lowerDuration.match(/(\d+)\s*(час|ч|h)/);
  if (hoursMatch) {
    totalMinutes += parseInt(hoursMatch[1], 10) * 60;
  }

  const minutesMatch = lowerDuration.match(/(\d+)\s*(минут|мин|m)/);
  if (minutesMatch) {
    totalMinutes += parseInt(minutesMatch[1], 10);
  }
  
  if (totalMinutes === 0 && !isNaN(parseInt(lowerDuration, 10))) {
      totalMinutes = parseInt(lowerDuration, 10);
  }

  return totalMinutes;
}

// --- Вспомогательная функция для парсинга диапазона дат (Версия 3.0, Финальная для MVP) ---
function parseDateScope(dateScope, referenceDate) {
  let startDate = null;
  let endDate = null;
  const now = referenceDate || new Date();
  const currentDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (!dateScope) {
    directLogToSheet("parseDateScope received empty input.");
    return { startDate: null, endDate: null };
  }

  const scope = dateScope.toLowerCase();
  const days = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  const dayMatch = scope.match(/воскресенье|понедельник|вторник|среда|четверг|пятница|суббота/);
  
  if (dayMatch) {
      const targetDayName = dayMatch[0];
      const targetDayIndex = days.indexOf(targetDayName);
      let dayOffset = targetDayIndex - currentDay.getDay();
      
      if (scope.includes('следующ') && dayOffset <= 0) { dayOffset += 7; } 
      else if (!scope.includes('следующ') && dayOffset < 0) { dayOffset += 7; }
      
      startDate = new Date(currentDay.getTime());
      startDate.setDate(startDate.getDate() + dayOffset);
      endDate = new Date(startDate.getTime());
      endDate.setHours(23, 59, 59, 999);
  } else {
      switch (scope) {
        case 'сегодня':
          startDate = new Date(currentDay.getTime());
          endDate = new Date(startDate);
          endDate.setHours(23, 59, 59, 999); 
          break;
        case 'завтра':
          startDate = new Date(currentDay.getTime());
          startDate.setDate(startDate.getDate() + 1);
          endDate = new Date(startDate);
          endDate.setHours(23, 59, 59, 999);
          break;
        case 'на этой неделе':
          const dayOfWeekThis = currentDay.getDay() === 0 ? 7 : currentDay.getDay();
          startDate = new Date(currentDay.getTime());
          startDate.setDate(startDate.getDate() - dayOfWeekThis + 1);
          endDate = new Date(startDate);
          endDate.setDate(endDate.getDate() + 6); 
          endDate.setHours(23, 59, 59, 999);
          break;
        // === ИСПРАВЛЕНИЕ И ДОПОЛНЕНИЕ ЗДЕСЬ ===
        case 'на следующей неделе':
          const startOfThisWeekForNext = new Date(currentDay.getTime());
          const dayOfWeekForNext = currentDay.getDay() === 0 ? 7 : currentDay.getDay();
          startOfThisWeekForNext.setDate(startOfThisWeekForNext.getDate() - dayOfWeekForNext + 1);
          startDate = new Date(startOfThisWeekForNext.getTime());
          startDate.setDate(startDate.getDate() + 7);
          endDate = new Date(startDate.getTime());
          endDate.setDate(endDate.getDate() + 6);
          endDate.setHours(23, 59, 59, 999);
          break;
        case 'прошедшая неделя':
        case 'на прошлой неделе':
           const startOfThisWeekForPrev = new Date(currentDay.getTime());
           const dayOfWeekForPrev = currentDay.getDay() === 0 ? 7 : currentDay.getDay();
           startOfThisWeekForPrev.setDate(startOfThisWeekForPrev.getDate() - dayOfWeekForPrev + 1);
           startDate = new Date(startOfThisWeekForPrev.getTime());
           startDate.setDate(startDate.getDate() - 7);
           endDate = new Date(startDate.getTime());
           endDate.setDate(endDate.getDate() + 6);
           endDate.setHours(23, 59, 59, 999);
           break;
        // ===================================
        default: 
          if (scope.match(/^\d{2}\.\d{2}\.\d{4}$/)) { 
            startDate = parseDdMmYyyyToDate(scope); 
            if (startDate) {
                endDate = new Date(startDate.getTime());
                endDate.setHours(23, 59, 59, 999); 
            }
          } else {
            directLogToSheet(`'${scope}' did not match any known date scopes in switch.`);
          }
          break;
      }
  }
  
  directLogToSheet(`Parsed date scope '${dateScope}': Start=${startDate}, End=${endDate}`);
  return { startDate, endDate };
}

// --- Вспомогательные функции для работы с CacheService ---
// Эти функции перенесены из DialogManager.gs в Utils.gs, т.к. они вспомогательные
function getContext(key) {
  try {
    // ИСПОЛЬЗУЕМ getUserCache() для пользовательского контекста
    const userCache = CacheService.getUserCache(); 
    const contextString = userCache.get(key);
    return contextString ? JSON.parse(contextString) : null;
  } catch(e) {
    logErrorToSheet('Cache Error', `Failed to get or parse context for key ${key}`, e.toString(), 'SYSTEM');
    return null;
  }
}

function saveContext(key, contextObject, expirationInSeconds) {
  try {
    // ИСПОЛЬЗУЕМ getUserCache() для пользовательского контекста
    const userCache = CacheService.getUserCache(); 
    userCache.put(key, JSON.stringify(contextObject), expirationInSeconds);
    directLogToSheet(`Context saved for key: ${key}`);
  } catch(e) {
    logErrorToSheet('Cache Error', `Failed to save context for key ${key}`, e.toString(), 'SYSTEM');
  }
}

function clearContext(key) {
  try {
    // ИСПОЛЬЗУЕМ getUserCache() для пользовательского контекста
    const userCache = CacheService.getUserCache(); 
    userCache.remove(key);
    directLogToSheet(`Context cleared for key: ${key}`);
  } catch(e) {
    logErrorToSheet('Cache Error', `Failed to clear context for key ${key}`, e.toString(), 'SYSTEM');
  }
}