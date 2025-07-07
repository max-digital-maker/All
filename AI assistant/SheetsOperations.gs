// =========================================================================================
// ФАЙЛ: SheetsOperations.gs (Версия 7.0 "ФИНАЛЬНЫЙ КОМПЛЕКСНЫЙ КОД")
// Полная замена. Учтены все исправления для Календаря, Заметок и Задач.
// =========================================================================================

// ===============================================================
// БЛОК: КАЛЕНДАРЬ
// ===============================================================

/**
 * Добавляет новое событие в календарь.
 */
function addEventToSheet(eventData, userId, chatId) {
  try {
    const sheet = getSheet('01_Календарь_События');
    const now = new Date();

    let eventDate = eventData.date;
    if (!eventDate || !String(eventDate).match(/^\d{2}\.\d{2}\.\d{4}$/)) {
        const parsedDates = parseDateScope(eventData.date_scope || 'сегодня', now);
        eventDate = parsedDates.startDate ? formatDateForSheet(parsedDates.startDate) : formatDateForSheet(now);
    }

    let eventTime = eventData.time || '09:00';
    let eventEndTime = eventData.end_time;

    if (!eventEndTime && eventData.duration) {
        const durationMinutes = parseDuration(eventData.duration);
        if (durationMinutes > 0) {
            const tempDate = parseDdMmYyyyToDate(eventDate); 
            if (tempDate) {
                const [hour, minute] = eventTime.split(':').map(Number);
                tempDate.setHours(hour, minute, 0, 0); 
                tempDate.setMinutes(tempDate.getMinutes() + durationMinutes); 
                eventEndTime = formatTimeForSheet(tempDate);
            }
        }
    }

    const newRow = [
      Utilities.getUuid(),
      eventData.event_name || 'Новое событие',
      eventDate,
      eventTime,
      eventEndTime || '',
      eventData.description || '',
      eventData.location || '',
      (eventData.participants || []).join(', '),
      'Запланировано',
      false, // Напоминание Отправлено
      now.toLocaleString('ru-RU'),
      userId,
      chatId,
      false // Удалено
    ];
    sheet.appendRow(newRow);
    sendMessageToTelegram(chatId, `Событие <b>"${eventData.event_name || 'Без названия'}"</b> записано на ${eventDate} с ${eventTime}!`);

  } catch (e) {
    directLogToSheet('Error in addEventToSheet: ' + e.toString() + e.stack);
    sendMessageToTelegram(chatId, 'Произошла ошибка при добавлении события.');
  }
}

/**
 * Ищет события в календаре по заданным фильтрам.
 */
function handleCalendarQuery(filters, userId, chatId) {
  try {
    const sheet = getSheet('01_Календарь_События');
    const data = sheet.getDataRange().getValues();
    const header = data[0]; 
    
    const NAME_COL = header.indexOf('Название События');
    const DATE_COL = header.indexOf('Дата');
    const TIME_COL = header.indexOf('Время');
    const END_TIME_COL = header.indexOf('Время_Окончания'); 
    const DESC_COL = header.indexOf('Описание');
    const LOCATION_COL = header.indexOf('Место');
    const PARTICIPANTS_COL = header.indexOf('Участники');
    const USER_ID_COL = header.indexOf('Пользователь_ID');
    const DELETED_COL = header.indexOf('Удалено');

    const filteredEvents = [];
    let { startDate, endDate } = parseDateScope(filters.date_scope, new Date());

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[DELETED_COL] === true || String(row[USER_ID_COL]) !== String(userId)) continue;
      
      let matches = true; 
      
      const eventDateObj = parseDateCellContent(row[DATE_COL]);
      if (eventDateObj && startDate && endDate) {
        if (eventDateObj < startDate || eventDateObj > endDate) matches = false;
      }
      
      if (matches && filters.participants && filters.participants.length > 0) {
        const eventParticipants = String(row[PARTICIPANTS_COL] || '').toLowerCase();
        const requestedParticipants = filters.participants.map(p => String(p).toLowerCase());
        if (!requestedParticipants.every(reqP => eventParticipants.includes(reqP))) matches = false;
      }

      if (matches && (filters.keywords || filters.event_name)) {
        const keyword = String(filters.keywords || filters.event_name).toLowerCase();
        if (!String(row[NAME_COL] || '').toLowerCase().includes(keyword)) matches = false;
      }

      if (matches) filteredEvents.push(row);
    }

    let responseText = 'По вашему запросу событий не найдено.';
    if (filteredEvents.length > 0) {
      responseText = 'Вот что я нашел:\n\n';
      filteredEvents.sort((a, b) => (parseDateCellContent(a[DATE_COL]) || 0) - (parseDateCellContent(b[DATE_COL]) || 0));
      filteredEvents.forEach(event => {
        responseText += `🗓️ <b>${event[NAME_COL]}</b>\n`;
        responseText += `  Дата: ${formatDateForSheet(parseDateCellContent(event[DATE_COL]))} с ${formatTimeForSheet(event[TIME_COL])}\n`;
        if (event[LOCATION_COL]) responseText += `  📍 Место: ${event[LOCATION_COL]}\n`;
        if (event[PARTICIPANTS_COL]) responseText += `  Участники: ${event[PARTICIPANTS_COL]}\n`;
        if (event[DESC_COL]) responseText += `  Описание: ${event[DESC_COL]}\n`;
        responseText += `---\n`;
      });
    }
    sendMessageToTelegram(chatId, responseText);
  } catch (e) {
    directLogToSheet('Error in handleCalendarQuery: ' + e.toString() + e.stack);
    sendMessageToTelegram(chatId, 'Произошла ошибка при поиске событий.');
  }
}

/**
 * Обновляет существующее событие.
 */
function updateEventInSheet(eventData, userId, chatId) {
  try {
    const sheet = getSheet('01_Календарь_События');
    const data = sheet.getDataRange().getValues();
    const header = data[0]; 
    const NAME_COL = header.indexOf('Название События');
    const DATE_COL = header.indexOf('Дата');
    const TIME_COL = header.indexOf('Время');
    const LOCATION_COL = header.indexOf('Место');
    const PARTICIPANTS_COL = header.indexOf('Участники');
    const DESC_COL = header.indexOf('Описание');
    const STATUS_COL = header.indexOf('Статус');
    const USER_ID_COL = header.indexOf('Пользователь_ID');
    const DELETED_COL = header.indexOf('Удалено');

    if (!eventData.date || !eventData.time) {
      sendMessageToTelegram(chatId, "Для обновления события мне нужна точная дата и время.");
      return;
    }

    let foundRowIndex = -1; 
    for (let i = 1; i < data.length; i++) { 
      const row = data[i];
      if (row[DELETED_COL] === true || String(row[USER_ID_COL]) !== String(userId)) continue;
      
      const eventDateInSheet = formatDateForSheet(parseDateCellContent(row[DATE_COL]));
      const eventTimeInSheet = formatTimeForSheet(row[TIME_COL]);
      
      let nameMatch = true;
      if (eventData.event_name) {
        nameMatch = String(row[NAME_COL] || '').toLowerCase().includes(String(eventData.event_name).toLowerCase());
      }
      
      if (eventDateInSheet === eventData.date && eventTimeInSheet === eventData.time && nameMatch) {
        foundRowIndex = i; 
        break; 
      }
    }

    if (foundRowIndex === -1) {
      sendMessageToTelegram(chatId, `Событие "${eventData.event_name || ''}" на ${eventData.date} в ${eventData.time} не найдено.`);
      return;
    }

    const fieldToUpdate = (eventData.field_to_update || '').toLowerCase();
    let colIndexToUpdate = -1;
    switch (fieldToUpdate) {
      case 'description': colIndexToUpdate = DESC_COL; break;
      case 'participants': colIndexToUpdate = PARTICIPANTS_COL; break;
      case 'location': colIndexToUpdate = LOCATION_COL; break;
      case 'time': colIndexToUpdate = TIME_COL; break;
      case 'date': colIndexToUpdate = DATE_COL; break;
      case 'status': colIndexToUpdate = STATUS_COL; break;
      default: sendMessageToTelegram(chatId, `Не могу обновить поле "${fieldToUpdate}".`); return;
    }

    if (colIndexToUpdate !== -1) {
      sheet.getRange(foundRowIndex + 1, colIndexToUpdate + 1).setValue(eventData.new_value);
      sendMessageToTelegram(chatId, `Событие "${data[foundRowIndex][NAME_COL]}" успешно обновлено.`);
    }
  } catch (e) {
    directLogToSheet('Error in updateEventInSheet: ' + e.toString() + e.stack);
    sendMessageToTelegram(chatId, 'Произошла ошибка при обновлении события.');
  }
}

/**
 * "Мягко" удаляет событие.
 */
function deleteEventFromSheet(eventData, userId, chatId) {
    try {
    const sheet = getSheet('01_Календарь_События');
    const data = sheet.getDataRange().getValues();
    const header = data[0]; 

    const NAME_COL = header.indexOf('Название События');
    const DATE_COL = header.indexOf('Дата');
    const TIME_COL = header.indexOf('Время');
    const USER_ID_COL = header.indexOf('Пользователь_ID');
    const DELETED_COL = header.indexOf('Удалено');

    if (!eventData.date || !eventData.time) { 
      sendMessageToTelegram(chatId, "Для удаления события мне нужно знать его точную дату и время.");
      return;
    }

    let foundRowIndex = -1; 
    for (let i = 1; i < data.length; i++) { 
      const row = data[i];
      if (row[DELETED_COL] === true || String(row[USER_ID_COL]) !== String(userId)) continue; 
      
      const eventNameInSheet = String(row[NAME_COL] || '').toLowerCase();
      const eventDateInSheet = formatDateForSheet(parseDateCellContent(row[DATE_COL]));
      const eventTimeInSheet = formatTimeForSheet(row[TIME_COL]);
      
      let nameMatch = true;
      if (eventData.event_name) {
        nameMatch = eventNameInSheet.includes(String(eventData.event_name).toLowerCase());
      }
      
      if (nameMatch && eventDateInSheet === eventData.date && eventTimeInSheet === eventData.time) {
        foundRowIndex = i; 
        break; 
      }
    }

    if (foundRowIndex === -1) {
      sendMessageToTelegram(chatId, `Событие на ${eventData.date} в ${eventData.time} не найдено.`);
      return;
    }

    sheet.getRange(foundRowIndex + 1, DELETED_COL + 1).setValue(true);
    sendMessageToTelegram(chatId, `Событие <b>"${data[foundRowIndex][NAME_COL]}"</b> перемещено в корзину.`);
    
  } catch (e) {
    directLogToSheet('Error soft-deleting event: ' + e.toString());
    sendMessageToTelegram(chatId, 'Произошла ошибка при удалении события.');
  }
}

// ===============================================================
// БЛОК: ЗАМЕТКИ
// ===============================================================

/**
 * Добавляет новую заметку.
 */
function addNoteToSheet(noteData, userId, chatId) {
  try {
    const sheet = getSheet('02_Личные_Заметки');
    const now = new Date();
    const newRow = [ Utilities.getUuid(), formatDateForSheet(now), formatTimeForSheet(now), noteData.content || 'Пустая заметка', noteData.category || 'Общее', userId, chatId, false ];
    sheet.appendRow(newRow);
    sendMessageToTelegram(chatId, `Заметка <b>"${(noteData.content || '').substring(0, 50)}..."</b> сохранена!`);
  } catch (e) {
    directLogToSheet('Error adding note: ' + e.toString());
    sendMessageToTelegram(chatId, 'Произошла ошибка при добавлении заметки.');
  }
}

/**
 * "Мягко" удаляет заметку.
 */
function deleteNoteFromSheet(noteData, userId, chatId) {
  try {
    const sheet = getSheet('02_Личные_Заметки');
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const CONTENT_COL = header.indexOf('Содержание Заметки');
    const USER_ID_COL = header.indexOf('Пользователь_ID');
    const DELETED_COL = header.indexOf('Удалено');

    const keywords = (noteData.keywords || '').toLowerCase();
    if (!keywords) {
      sendMessageToTelegram(chatId, "Чтобы удалить заметку, укажите ключевые слова из ее содержания.");
      return;
    }
    
    const foundNotes = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[DELETED_COL] === true || String(row[USER_ID_COL]) !== String(userId)) continue;
        if (String(row[CONTENT_COL] || '').toLowerCase().includes(keywords)) {
            foundNotes.push({ rowIndex: i + 1, content: row[CONTENT_COL] });
        }
    }

    if (foundNotes.length === 0) {
        sendMessageToTelegram(chatId, `Заметки, содержащей "${noteData.keywords}", не найдено.`);
    } else if (foundNotes.length === 1) {
        const noteToDelete = foundNotes[0];
        sheet.getRange(noteToDelete.rowIndex, DELETED_COL + 1).setValue(true);
        sendMessageToTelegram(chatId, `Заметка "<i>${noteToDelete.content.substring(0, 50)}...</i>" перемещена в корзину.`);
    } else {
        let responseText = `Найдено несколько заметок. Какую удалить?\n\n`;
        foundNotes.slice(0, 5).forEach((note, index) => { responseText += `${index + 1}. ${note.content.substring(0, 60)}...\n`; });
        responseText += `\nОтветьте, например: "удали первую" или "отмена".`;
        sendMessageToTelegram(chatId, responseText);
    }
  } catch (e) {
    directLogToSheet('Error soft-deleting note: ' + e.toString());
    sendMessageToTelegram(chatId, 'Произошла ошибка при удалении заметки.');
  }
}

// ===============================================================
// БЛОК: ЗАДАЧИ
// ===============================================================

/**
 * Добавляет задачу.
 */
function addTaskToSheet(taskData, userId, chatId) {
  try {
    const sheet = getSheet('03_Задачи');
    const now = new Date();
    let dueDate = '';
    if (taskData.due_date_scope) {
        const parsedDates = parseDateScope(taskData.due_date_scope, now);
        if(parsedDates.startDate) dueDate = formatDateForSheet(parsedDates.startDate);
    }
    const newRow = [
      Utilities.getUuid(),
      taskData.task_name || 'Новая задача',
      taskData.project || '',
      (taskData.tags || []).join(', '),
      'К выполнению',
      taskData.priority || 'Средний',
      dueDate,
      taskData.assignee || userId,
      now.toLocaleString('ru-RU'),
      '', 
      userId,
      chatId,
      false
    ];
    sheet.appendRow(newRow);
    sendMessageToTelegram(chatId, `✅ Задача "<b>${taskData.task_name}</b>" добавлена!`);
  } catch (e) {
    directLogToSheet('Error in addTaskToSheet: ' + e.toString() + e.stack);
    sendMessageToTelegram(chatId, 'Произошла ошибка при добавлении задачи.');
  }
}

/**
 * Ищет задачи.
 */
function handleTasksQuery(filters, userId, chatId) {
  try {
    const sheet = getSheet('03_Задачи');
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const NAME_COL = header.indexOf('Название Задачи');
    const PROJECT_COL = header.indexOf('Проект');
    const TAGS_COL = header.indexOf('Теги');
    const STATUS_COL = header.indexOf('Статус');
    const PRIORITY_COL = header.indexOf('Приоритет');
    const DUE_DATE_COL = header.indexOf('Срок');
    const ASSIGNEE_COL = header.indexOf('Ответственный');
    const DELETED_COL = header.indexOf('Удалено');

    const filteredTasks = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[DELETED_COL] === true) continue;

        let matches = true;
        if (filters.project && String(row[PROJECT_COL] || '').toLowerCase() !== filters.project.toLowerCase()) matches = false;
        if (filters.status && String(row[STATUS_COL] || '').toLowerCase() !== filters.status.toLowerCase()) matches = false;
        if (filters.priority && String(row[PRIORITY_COL] || '').toLowerCase() !== filters.priority.toLowerCase()) matches = false;
        if (filters.assignee && !String(row[ASSIGNEE_COL] || '').toLowerCase().includes(filters.assignee.toLowerCase())) matches = false;
        if (filters.tags && filters.tags.length > 0) {
            const rowTags = String(row[TAGS_COL] || '').toLowerCase().split(',').map(t => t.trim());
            if (!filters.tags.every(reqTag => rowTags.includes(reqTag.toLowerCase()))) matches = false;
        }
        
        if (matches) filteredTasks.push(row);
    }
    
    let responseText = 'По вашему запросу задач не найдено.';
    if (filteredTasks.length > 0) {
        responseText = 'Вот что я нашел:\n\n';
        filteredTasks.sort((a, b) => (parseDateCellContent(a[DUE_DATE_COL]) || new Date('2100-01-01')) - (parseDateCellContent(b[DUE_DATE_COL]) || new Date('2100-01-01')));
        filteredTasks.forEach(task => {
            responseText += `📝 <b>${task[NAME_COL]}</b> [${task[STATUS_COL]}]\n`;
            if(task[DUE_DATE_COL]) responseText += `   Срок: ${formatDateForSheet(parseDateCellContent(task[DUE_DATE_COL]))}\n`;
            if(task[PROJECT_COL]) responseText += `   Проект: ${task[PROJECT_COL]}\n`;
            if(task[ASSIGNEE_COL]) responseText += `   Ответственный: ${task[ASSIGNEE_COL]}\n`;
            if(task[TAGS_COL]) responseText += `   Теги: #${String(task[TAGS_COL]).replace(/, /g, ' #').replace(/,/g, ' #')}\n`;
            if(task[PRIORITY_COL]) responseText += `   Приоритет: ${task[PRIORITY_COL]}\n`;
            responseText += `---\n`;
        });
    }
    sendMessageToTelegram(chatId, responseText);
  } catch (e) {
      directLogToSheet('Error in handleTasksQuery: ' + e.toString() + e.stack);
      sendMessageToTelegram(chatId, 'Произошла ошибка при поиске задач.');
  }
}

/**
 * Обновляет статус задачи.
 */
function updateTaskStatusInSheet(taskData, userId, chatId) {
    try {
        const sheet = getSheet('03_Задачи');
        const data = sheet.getDataRange().getValues();
        const header = data[0];
        const NAME_COL = header.indexOf('Название Задачи');
        const STATUS_COL = header.indexOf('Статус');
        const DATE_COMPLETED_COL = header.indexOf('Дата_Завершения');
        const DELETED_COL = header.indexOf('Удалено');

        const keywords = (taskData.keywords || '').toLowerCase();
        if (!keywords) {
            sendMessageToTelegram(chatId, "Чтобы обновить задачу, укажите ключевые слова из ее названия.");
            return;
        }

        const foundTasks = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row[DELETED_COL] === true) continue;
            if (String(row[NAME_COL] || '').toLowerCase().includes(keywords)) {
                foundTasks.push({ rowIndex: i + 1, name: row[NAME_COL] });
            }
        }

        if (foundTasks.length === 0) {
            sendMessageToTelegram(chatId, `Задача, содержащая "${keywords}", не найдена.`);
        } else if (foundTasks.length === 1) {
            const taskToUpdate = foundTasks[0];
            const newStatus = taskData.new_status || 'Выполнено';
            sheet.getRange(taskToUpdate.rowIndex, STATUS_COL + 1).setValue(newStatus);
            if (newStatus === 'Выполнено' || newStatus === 'Отменено') {
                sheet.getRange(taskToUpdate.rowIndex, DATE_COMPLETED_COL + 1).setValue(new Date());
            }
            sendMessageToTelegram(chatId, `👍 Статус задачи "<b>${taskToUpdate.name}</b>" изменен на "${newStatus}"!`);
        } else {
            sendMessageToTelegram(chatId, "Найдено несколько задач. Пожалуйста, уточните, какую именно обновить.");
        }
    } catch (e) {
        directLogToSheet('Error in updateTaskStatusInSheet: ' + e.toString() + e.stack);
        sendMessageToTelegram(chatId, 'Произошла ошибка при обновлении статуса задачи.');
    }
}


/**
 * "Мягко" удаляет задачу.
 */
function deleteTaskFromSheet(taskData, userId, chatId) {
    try {
        const sheet = getSheet('03_Задачи');
        const data = sheet.getDataRange().getValues();
        const header = data[0];
        const NAME_COL = header.indexOf('Название Задачи');
        const DELETED_COL = header.indexOf('Удалено');

        const keywords = (taskData.keywords || '').toLowerCase();
        if (!keywords) {
            sendMessageToTelegram(chatId, "Чтобы удалить задачу, укажите ключевые слова из ее названия.");
            return;
        }
        
        const foundTasks = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row[DELETED_COL] === true) continue;
            if (String(row[NAME_COL] || '').toLowerCase().includes(keywords)) {
                foundTasks.push({ rowIndex: i + 1, name: row[NAME_COL] });
            }
        }

        if (foundTasks.length === 0) {
            sendMessageToTelegram(chatId, `Задача, содержащая "${keywords}", не найдена.`);
        } else if (foundTasks.length === 1) {
            const taskToDelete = foundTasks[0];
            sheet.getRange(taskToDelete.rowIndex, DELETED_COL + 1).setValue(true);
            sendMessageToTelegram(chatId, `🗑️ Задача "<b>${taskToDelete.name}</b>" перемещена в корзину.`);
        } else {
            sendMessageToTelegram(chatId, "Найдено несколько задач. Пожалуйста, уточните, какую именно удалить.");
        }
    } catch (e) {
        directLogToSheet('Error in deleteTaskFromSheet: ' + e.toString() + e.stack);
        sendMessageToTelegram(chatId, 'Произошла ошибка при удалении задачи.');
    }
}