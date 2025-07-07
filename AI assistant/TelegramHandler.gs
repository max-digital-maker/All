// =========================================================================================
// ФАЙЛ: TelegramHandler.gs
// ПОЛНАЯ ЗАМЕНА
// =========================================================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
        logErrorToSheet('Webhook Input Error', 'Invalid POST data', JSON.stringify(e), 'SYSTEM');
        return; 
    }

    const contents = JSON.parse(e.postData.contents);
    const message = contents.message;

    if (!message) {
      directLogToSheet('Non-message update received, skipping.');
      return;
    }

    const chatId = message.chat.id;
    const userId = message.from.id;
    const userMessageText = message.text;

    if (message.voice) {
        sendMessageToTelegram(chatId, "Извините, сейчас я обрабатываю только текстовые сообщения.");
        return;
    }

    if (!userMessageText) {
      sendMessageToTelegram(chatId, "Извините, я пока умею работать только с текстовыми сообщениями.");
      return;
    }

    // ВЫЗЫВАЕМ НОВУЮ ТОЧКУ ВХОДА В DIALOGMANAGER
    handleMessage(userMessageText, userId, chatId);

  } catch (error) {
    const errorString = error.toString();
    const contextData = e.postData ? e.postData.contents : 'No e.postData available.';
    logErrorToSheet('doPost General Error', errorString, String(contextData), 'SYSTEM');
    
    try {
        const contents = e.postData ? JSON.parse(e.postData.contents) : null;
        if (contents && contents.message && contents.message.chat && contents.message.chat.id) {
            sendMessageToTelegram(contents.message.chat.id, "Произошла внутренняя ошибка. Попробуйте еще раз.");
        }
    } catch (eSend) {
        Logger.log('Failed to send error message back to user: ' + eSend.toString());
    }
  }
}

/**
 * Отправляет сообщение в Telegram с разбивкой на части и логикой повторных попыток.
 * @param {string|number} chatId ID чата.
 * @param {string} text Текст сообщения.
 */
function sendMessageToTelegram(chatId, text) {
  const MAX_LENGTH = 4096;
  if (!text) {
    directLogToSheet(`Attempted to send empty message to chat ID: ${chatId}`);
    return;
  }
  
  const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
  const RETRIES = 3; // Количество повторных попыток

  // Разбиваем сообщение на части, если оно слишком длинное
  for (let i = 0; i < text.length; i += MAX_LENGTH) {
    const chunk = text.substring(i, i + MAX_LENGTH);
    
    // Пытаемся отправить каждый кусок несколько раз
    for (let attempt = 1; attempt <= RETRIES; attempt++) {
      try {
        const payload = {
          method: 'post',
          payload: JSON.stringify({
            chat_id: String(chatId),
            text: chunk,
            parse_mode: 'HTML'
          }),
          contentType: 'application/json',
          muteHttpExceptions: false // Временно ставим false, чтобы ловить ошибку
        };
        UrlFetchApp.fetch(url, payload);
        // Если успешно, выходим из цикла попыток
        directLogToSheet(`Chunk sent successfully on attempt ${attempt}.`);
        break; 
      } catch (e) {
        directLogToSheet(`Attempt ${attempt} failed: ${e.toString()}`);
        if (attempt < RETRIES) {
          // Если это не последняя попытка, ждем перед следующей
          Utilities.sleep(1000 * attempt); // Увеличиваем задержку (1с, 2с)
        } else {
          // Если все попытки провалились, логируем фатальную ошибку
          logErrorToSheet('Telegram Send Fatal Error', `Failed to send chunk after ${RETRIES} attempts.`, e.toString(), chatId);
        }
      }
    }
  }
}

