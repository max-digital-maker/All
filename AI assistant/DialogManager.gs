// =========================================================================================
// ФАЙЛ: DialogManager.gs (ОТРЕФАКТОРЕННАЯ ВЕРСИЯ)
// =========================================================================================

const CLARIFY_CONTEXT_KEY_PREFIX = 'context_clarify_';

/**
 * Главная точка входа для обработки всех сообщений пользователя.
 */
function handleMessage(messageText, userId, chatId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  let ownerId = scriptProperties.getProperty('MAIN_USER_ID');
  if (!ownerId || ownerId === '0') {
      scriptProperties.setProperty('MAIN_USER_ID', String(userId));
      sendMessageToTelegram(chatId, "✨ Привет! Я твой новый личный ассистент. Теперь ты мой владелец. Я могу управлять календарем, заметками и просто общаться. С чего начнем?");
      return; 
  }
  if (String(userId) !== String(ownerId)) {
      sendMessageToTelegram(chatId, "Извините, у вас нет прав для выполнения этой операции.");
      return;
  }

  if (messageText.toLowerCase().trim() === '/reset') {
    clearContext(CLARIFY_CONTEXT_KEY_PREFIX + userId);
    clearHistory(userId);
    sendMessageToTelegram(chatId, "Контекст диалога сброшен. Начинаем заново!");
    return;
  }

  const clarifyContext = getContext(CLARIFY_CONTEXT_KEY_PREFIX + userId);
  const history = loadHistory(userId);
  
  let intent;
  if (clarifyContext) {
    intent = "COMMAND";
    directLogToSheet(`Forcing INTENT=COMMAND due to active clarification context for user ${userId}`);
  } else {
    intent = getIntent(messageText); // userId не нужен, промпт общий
  }
  
  directLogToSheet(`Intent determined: ${intent}`);
  
  if (intent === "COMMAND") {
    processCommandFlow(messageText, userId, chatId, clarifyContext, history);
  } else { 
    processChatFlow(messageText, userId, chatId, history);
  }
}

/**
 * Обрабатывает поток диалога, связанный с выполнением команд.
 */
function processCommandFlow(messageText, userId, chatId, clarifyContext, history) {
  // ИЗМЕНЕНО: Убрана вся логика сборки промпта. Теперь мы просто вызываем getCommandJson.
  const geminiResponse = getCommandJson(messageText, userId, clarifyContext, history);
  
  if (!geminiResponse || !geminiResponse.action) {
    logErrorToSheet('Command Flow Error', 'Gemini response is invalid or missing action.', JSON.stringify(geminiResponse), userId);
    sendMessageToTelegram(chatId, "Что-то пошло не так при обработке команды. Попробуем пообщаться?");
    processChatFlow(messageText, userId, chatId, history);
    return;
  }
  
  // ИЗМЕНЕНО: Добавлен кейс general_chat для более гибкого перехода от команды к чату
  switch (geminiResponse.action) {
    case 'execute_command':
      clearContext(CLARIFY_CONTEXT_KEY_PREFIX + userId);
      executeAction(geminiResponse.command_details, userId, chatId);
      // Не очищаем историю здесь, чтобы можно было задать следующий вопрос в контексте. 
      // Например: "Добавь встречу", "Ок, добавил", "А что у меня еще на сегодня?"
      break;

    case 'clarify':
      const contextToSave = {
        original_action: geminiResponse.context.original_action,
        extracted_entities: geminiResponse.context.extracted_entities,
        question_to_user: geminiResponse.message
      };
      saveContext(CLARIFY_CONTEXT_KEY_PREFIX + userId, contextToSave, CACHE_EXPIRATION_SECONDS);
      sendMessageToTelegram(chatId, geminiResponse.message);
      updateHistory(userId, messageText, geminiResponse.message);
      break;

    case 'cancel_operation':
      clearContext(CLARIFY_CONTEXT_KEY_PREFIX + userId);
      clearHistory(userId);
      sendMessageToTelegram(chatId, 'Хорошо, операция отменена.');
      break;
    
    case 'general_chat':
      // Этот кейс — защита, если Gemini решил, что это все-таки чат.
      // Перенаправляем на обработчик чата.
      directLogToSheet(`Switching from COMMAND to CHAT based on Gemini response for user ${userId}`);
      processChatFlow(messageText, userId, chatId, history);
      break;

    default:
      sendMessageToTelegram(chatId, 'Неизвестный тип ответа от ИИ. Попробуйте еще раз.');
      logErrorToSheet('Unknown Action', `Received unknown action: ${geminiResponse.action}`, JSON.stringify(geminiResponse), userId);
  }
}

/**
 * Обрабатывает поток диалога, связанный с общим чатом.
 */
function processChatFlow(messageText, userId, chatId, history) {
  clearContext(CLARIFY_CONTEXT_KEY_PREFIX + userId); 
  const chatResponseText = getChatResponse(messageText, history, userId); // Имя функции исправлено
  sendMessageToTelegram(chatId, chatResponseText);
  updateHistory(userId, messageText, chatResponseText);
}


/** Исполняет команды, полученные от Gemini.
 * ВЕРСИЯ С ПОДДЕРЖКОЙ ЗАДАЧ.
 */
function executeAction(commandDetails, userId, chatId) {
  directLogToSheet(`Executing action: ${commandDetails.action} for user: ${userId}`);
  
  // В зависимости от команды, вызываем соответствующую функцию из SheetsOperations.gs
  switch(commandDetails.action) {
    // --- КАЛЕНДАРЬ ---
    case 'add_event':
      addEventToSheet(commandDetails, userId, chatId);
      break;
    case 'query_calendar':
      handleCalendarQuery(commandDetails.filters, userId, chatId);
      break;
    case 'update_event':
      updateEventInSheet(commandDetails, userId, chatId);
      break;
    case 'delete_event':
      deleteEventFromSheet(commandDetails, userId, chatId);
      break;

    // --- ЗАМЕТКИ ---
    case 'add_note':
      addNoteToSheet(commandDetails, userId, chatId);
      break;
    case 'query_notes':
      // Предполагаем, что у вас есть функция handleNotesQuery. Если нет, ее нужно будет создать.
      // handleNotesQuery(commandDetails.filters, userId, chatId);
      sendMessageToTelegram(chatId, "Функция поиска по заметкам в разработке.");
      break;
    case 'delete_note':
      deleteNoteFromSheet(commandDetails, userId, chatId);
      break;

    // --- ЗАДАЧИ (НОВЫЙ БЛОК) ---
    case 'add_task':
      addTaskToSheet(commandDetails, userId, chatId);
      break;
    case 'query_tasks':
      handleTasksQuery(commandDetails.filters, userId, chatId);
      break;
    case 'update_task_status':
      updateTaskStatusInSheet(commandDetails, userId, chatId);
      break;
    case 'delete_task':
      deleteTaskFromSheet(commandDetails, userId, chatId);
      break;

    // --- ОБРАБОТКА НЕИЗВЕСТНОЙ КОМАНДЫ ---
    default:
      sendMessageToTelegram(chatId, "Я не знаю, как выполнить эту команду.");
      logErrorToSheet('Unknown Action', `Command not implemented: ${commandDetails.action}`, JSON.stringify(commandDetails), userId);
  }
}