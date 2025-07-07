// =========================================================================================
// ФАЙЛ: GeminiHandler.gs (ИСПРАВЛЕННАЯ ВЕРСИЯ БЕЗ GeminiApp)
// Все вызовы идут через UrlFetchApp, что не требует биллинга.
// =========================================================================================

/**
 * НОВАЯ УНИВЕРСАЛЬНАЯ ФУНКЦИЯ для вызова Gemini API через UrlFetchApp.
 * @param {string} model - Название модели (например, 'gemini-1.5-flash' или 'gemini-pro').
 * @param {string} prompt - Текст промпта для модели.
 * @param {object} config - Дополнительные параметры (temperature, response_mime_type).
 * @returns {string} - Текстовый ответ от API.
 */
function callGeminiApi(model, prompt, config = {}) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: config.temperature || 0.1,
      // Добавляем mime_type только если он указан
      ...(config.response_mime_type && { response_mime_type: config.response_mime_type })
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    directLogToSheet(`Sending to Gemini (${model}): ${prompt.substring(0, 500)}...`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      logErrorToSheet(`Gemini API Error (${model})`, `HTTP ${responseCode}: ${responseText}`, `Prompt: ${prompt.substring(0, 100)}...`);
      return null; // Возвращаем null в случае ошибки API
    }
    
    // В API generateContent ответ находится в response.candidates[0].content.parts[0].text
    const parsedResponse = JSON.parse(responseText);
    const textContent = parsedResponse.candidates[0].content.parts[0].text;
    
    directLogToSheet(`Received from Gemini (${model}): ${textContent.substring(0, 500)}...`);
    return textContent;

  } catch (e) {
    logErrorToSheet(`Gemini Call/Parse Error (${model})`, e.toString(), `Prompt: ${prompt.substring(0, 100)}...`);
    return null; // Возвращаем null в случае ошибки парсинга
  }
}

/**
 * 1. Определяет намерение пользователя (COMMAND или CHAT).
 * ИЗМЕНЕНО: Используем новую функцию-хелпер callGeminiApi.
 */
function getIntent(userMessage) {
  const prompt = PROMPT_INTENT_DETECTOR.replace('{{user_message}}', userMessage);
  const responseText = callGeminiApi('gemini-2.5-flash', prompt, { temperature: 0.0 });

  if (!responseText) {
    logErrorToSheet('Intent Detection Fallback', 'API call failed, defaulting to CHAT.', `Message: ${userMessage}`);
    return "CHAT";
  }

  const intent = responseText.trim().toUpperCase();
  directLogToSheet(`Intent Detector - User: "${userMessage}" -> AI: "${intent}"`);

  if (intent === "COMMAND" || intent === "CHAT") {
    return intent;
  }
  
  return "CHAT";
}

/**
 * 2. Обрабатывает команду и возвращает JSON.
 * ИЗМЕНЕНО: Используем новую функцию-хелпер callGeminiApi.
 */
function getCommandJson(userMessage, userId, clarifyContext, history) {
  const currentDateString = formatDateForSheet(new Date());
  let prompt;
  
  if (clarifyContext) {
    directLogToSheet(`User ${userId} is in Clarification Mode.`);
    prompt = PROMPT_CLARIFY
      .replace('{original_action}', clarifyContext.original_action)
      .replace('{extracted_entities}', JSON.stringify(clarifyContext.extracted_entities))
      .replace('{question_to_user}', clarifyContext.question_to_user)
      .replace('{user_message}', userMessage)
      .replace('{{current_date}}', currentDateString)
      .replace('{{command_structure}}', LLM_PROMPT_COMMANDS_STRUCTURE);
  } else {
    directLogToSheet(`User ${userId} is in Command Mode.`);
    const tomorrowDate = new Date();
    tomorrowDate.setDate(new Date().getDate() + 1);
    const tomorrowFormatted = formatDateForSheet(tomorrowDate);

    prompt = PROMPT_COMMAND_HANDLER
      .replace('{{current_date}}', currentDateString)
      .replace('{{tomorrow_date}}', tomorrowFormatted)
      .replace('{{command_structure}}', LLM_PROMPT_COMMANDS_STRUCTURE)
      .replace('{{history}}', formatHistoryForPrompt(history))
      .replace('{{user_message}}', userMessage);
  }

  // Вызываем API с требованием получить JSON
  const responseText = callGeminiApi('gemini-2.5-flash', prompt, { temperature: 0.1, response_mime_type: "application/json" });

  if (!responseText) {
    return { action: "clarify", message: "Произошла ошибка при обработке команды. Попробуйте еще раз." };
  }
  
  try {
    const cleanedText = responseText.replace(/```json\s*|\s*```/g, '').trim();
    const parsedObject = JSON.parse(cleanedText);
    directLogToSheet('Successfully parsed JSON object from Gemini: ' + JSON.stringify(parsedObject));
    return parsedObject;
  } catch(e) {
    logErrorToSheet('Gemini Command/Parse Error', e.toString(), `Raw response: ${responseText || 'No response'}. Message: ${userMessage}`, userId);
    return { action: "clarify", message: "Извините, ответ от ИИ не удалось распознать. Попробуйте переформулировать." };
  }
}

/**
 * 3. Генерирует ответ для свободного диалога.
 * ИЗМЕНЕНО: Используем новую функцию-хелпер callGeminiApi.
 */
function getChatResponse(userMessage, history, userId) {
  // Для чата можно собрать простой промпт с историей
  const historyString = formatHistoryForPrompt(history);
  const prompt = `Ты — дружелюбный и полезный ИИ-ассистент. Вот история нашего диалога:\n${historyString}\n\nПользователь: ${userMessage}\nАссистент:`;

  const responseText = callGeminiApi('gemini-2.5-flash', prompt, { temperature: 0.8 });

  if (!responseText) {
    return "Извините, у меня возникла техническая проблема. Попробуйте повторить ваш вопрос чуть позже.";
  }
  return responseText;
}


/**
 * Вспомогательная функция для форматирования истории для промпта (без изменений).
 */
function formatHistoryForPrompt(history) {
  if (!history || history.length === 0) {
    return "Нет истории.";
  }
  return history.map(turn => `${turn.role === 'user' ? 'Пользователь' : 'Ассистент'}: ${turn.parts[0].text}`).join('\n');
}