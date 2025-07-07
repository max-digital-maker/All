// --- Функция для настройки Webhook Telegram ---
function setupTelegramWebhook() {
  const TELEGRAM_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');

  const webAppUrl = "https://script.google.com/macros/s/AKfycbzOumH4BjmgBiLApuqFbqNca5c7QFT_rLjd7WbSFtzTjm3ezXLpI_EjiTirNnoAF2prtA/exec"; // ВАЖНО: Убедитесь, что это URL вашего *текущего* развертывания!


  if (!TELEGRAM_BOT_TOKEN || !webAppUrl) {
    directLogToSheet('Error: TELEGRAM_BOT_TOKEN или URL веб-приложения отсутствует. Проверьте свойства сценария.');
    return;
  }

  const setWebhookUrl = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/setWebhook?url=${encodeURIComponent(webAppUrl)}`;

  try {
    const response = UrlFetchApp.fetch(setWebhookUrl);
    const result = JSON.parse(response.getContentText());

    if (result.ok) {
      directLogToSheet('Webhook set successfully: ' + JSON.stringify(result));
    } else {
      directLogToSheet('Failed to set webhook: ' + JSON.stringify(result));
      logErrorToSheet('Webhook Setup Error', result.description, 'Failed to set webhook via Apps Script', 'SYSTEM');
    }
  } catch (e) {
    directLogToSheet('Error setting webhook: ' + e.toString());
    logErrorToSheet('Webhook Setup API Error', e.toString(), 'Exception during UrlFetchApp.fetch for setWebhook', 'SYSTEM');
  }
}

// --- НЕОБЯЗАТЕЛЬНО: Функция для удаления Webhook (для отладки) ---
function deleteTelegramWebhook() {
  const TELEGRAM_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!TELEGRAM_BOT_TOKEN) {
    directLogToSheet('Error: TELEGRAM_BOT_TOKEN отсутствует. Проверьте свойства сценария.');
    return;
  }

  const deleteWebhookUrl = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/deleteWebhook`;

  try {
    const response = UrlFetchApp.fetch(deleteWebhookUrl);
    const result = JSON.parse(response.getContentText());

    if (result.ok) {
      directLogToSheet('Webhook deleted successfully: ' + JSON.stringify(result));
    } else {
      directLogToSheet('Failed to delete webhook: ' + JSON.stringify(result));
      logErrorToSheet('Webhook Delete Error', result.description, 'Failed to delete webhook via Apps Script', 'SYSTEM');
    }
  } catch (e) {
    directLogToSheet('Error deleting webhook: ' + e.toString());
    logErrorToSheet('Webhook Delete API Error', e.toString(), 'Exception during UrlFetchApp.fetch for deleteWebhook', 'SYSTEM');
  }
}