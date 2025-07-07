// =========================================================================================
// ФАЙЛ: HistoryManager.gs (Создайте этот файл, если его нет)
// =========================================================================================

/**
 * Загружает историю чата из Google Sheet для заданного пользователя.
 * @param {string|number} userId - Уникальный идентификатор пользователя.
 * @returns {Array<Object>} Массив объектов в формате, понятном Gemini API.
 */
function loadHistory(userId) {
  try {
    // ИСПРАВЛЕНО: Используем getSheet вместо getOrCreateSheet
    const sheet = getSheet(CONVERSATION_HISTORY_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][0]) === String(userId)) {
        const historyJson = data[i][2];
        if (historyJson) return JSON.parse(historyJson);
      }
    }
  } catch (e) {
    // ИСПРАВЛЕНО: Логируем ошибку через стандартную функцию
    directLogToSheet(`History Parse Error for user ${userId}: ${e.message}`);
  }
  return []; 
}

/**
 * Добавляет новый обмен сообщениями в историю пользователя и сохраняет ее в Google Sheet.
 * @param {string|number} userId - Уникальный идентификатор пользователя.
 * @param {string} userMessage - Сообщение от пользователя.
 * @param {string} botResponse - Ответ бота.
 */
function updateHistory(userId, userMessage, botResponse) {
  let history = loadHistory(userId); 
  history.push({ "role": "user", "parts": [{ "text": userMessage }] });
  history.push({ "role": "model", "parts": [{ "text": botResponse }] });
  
  while (history.length > MAX_HISTORY_TURNS * 2) {
    history.shift(); 
  }
  
  clearHistory(userId, true);
  
  // ИСПРАВЛЕНО: Используем getSheet
  const sheet = getSheet(CONVERSATION_HISTORY_SHEET_NAME);
  sheet.appendRow([userId, new Date(), JSON.stringify(history)]);
}

/**
 * Полностью удаляет историю для пользователя из таблицы.
 * @param {string|number} userId - ID пользователя.
 * @param {boolean} silent - Если true, не выводить сообщение в лог.
 */
function clearHistory(userId, silent = false) {
  // ИСПРАВЛЕНО: Используем getSheet
  const sheet = getSheet(CONVERSATION_HISTORY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  if (!silent && rowsToDelete.length > 0) {
    directLogToSheet(`Cleared all history from sheet for user ${userId}`);
  }
}