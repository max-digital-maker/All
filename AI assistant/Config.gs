// =========================================================================================
// ФАЙЛ: Config.gs (Версия 7.1 "ПОЛНЫЙ КОНТРАКТ")
// Финальная версия. Содержит все примеры и правила для всех модулей.
// =========================================================================================

const TELEGRAM_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const MAIN_USER_ID = PropertiesService.getScriptProperties().getProperty('MAIN_USER_ID');
const CACHE_EXPIRATION_SECONDS = parseInt(PropertiesService.getScriptProperties().getProperty('CACHE_EXPIRATION_SECONDS'), 10) || 600;
const CONVERSATION_HISTORY_SHEET_NAME = "ConversationHistory";
const MAX_HISTORY_TURNS = 10;
const PROMPT_INTENT_DETECTOR = `Analyze the user's message. Is it a command to manage a calendar, tasks or notes? If yes, respond with COMMAND. If it's a general question or greeting, respond with CHAT. User: "Добавь задачу" COMMAND. User: "Привет" CHAT. User: "{{user_message}}"`;

const PROMPT_COMMAND_HANDLER = `Ты — надежный ассистент. Твоя единственная задача — вернуть JSON, строго следуя инструкциям.

**ПРАВИЛО №1 (Даты):** Если пользователь говорит "завтра", "во вторник", передавай эту фразу КАК ЕСТЬ в "date_scope". Если дата явная ("10.07.2025"), используй "date".
**ПРАВИЛО №2 (Год):** Если дата без года ("10.07"), ты ДОЛЖЕН добавить текущий год. Текущая дата: {{current_date}}.
**ПРАВИЛО №3 (Сущности):** Если есть имена ("с Игорем"), места ("в кафе"), проекты ("в проект Переезд"), извлекай их в "participants", "location", "project".
**ПРАВИЛО №4 (Умный поиск):** При обновлении или удалении задачи, извлеки ключевые слова для поиска по названию в "keywords", а уточняющие детали (проект, статус) в объект "filters".

**СТРУКТУРА ОТВЕТА:**
1.  {"action": "execute_command", "command_details": { ... }}
2.  {"action": "clarify", "message": "...", "context": { ... }}
3.  {"action": "cancel_operation"}
4.  {"action": "general_chat"}

--- ПРИМЕРЫ ---
*   Пользователь: "Запиши мне встречу с Игорем завтра в 10:00" -> {"action": "execute_command", "command_details": {"action": "add_event", "event_name": "Встреча", "date_scope": "завтра", "time": "10:00", "participants": ["Игорь"]}}
*   Пользователь: "добавь задачу купить молоко в проект Дом с тегом покупки" -> {"action": "execute_command", "command_details": {"action": "add_task", "task_name": "купить молоко", "project": "Дом", "tags": ["покупки"]}}
*   Пользователь: "задача позвонить подрядчику до пятницы, приоритет высокий" -> {"action": "execute_command", "command_details": {"action": "add_task", "task_name": "позвонить подрядчику", "due_date_scope": "пятница", "priority": "Высокий"}}
*   Пользователь: "какие у меня задачи по проекту Переезд?" -> {"action": "execute_command", "command_details": {"action": "query_tasks", "filters": {"project": "Переезд"}}}
*   Пользователь: "отметь задачу 'согласовать смету' по проекту 'Упаковка' как выполненную" -> {"action": "execute_command", "command_details": {"action": "update_task_status", "keywords": "согласовать смету", "filters": {"project": "Упаковка"}, "new_status": "Выполнено"}}
*   Пользователь: "Удали встречу с Игорем" -> {"action": "clarify", "message": "Конечно, какую именно встречу с Игорем? Укажите, пожалуйста, дату и время.", "context": {"original_action": "delete_event", "extracted_entities": {"participants": ["Игорь"], "event_name": "Встреча"}}}

--- СТРУКТУРА КОМАНД ---
{{command_structure}}
--- ИСТОРИЯ ДИАЛОГА ---
{{history}}
--- ТЕКУЩИЙ ЗАПРОС ПОЛЬЗОВАТЕЛЯ ---
{{user_message}}
`;

const PROMPT_CLARIFY = `Ты — помощник, завершающий команду. Отвечай только в JSON.
Цель: {original_action}.
Уже известно: {extracted_entities}.
Задан вопрос: "{question_to_user}".
Ответ пользователя: "{user_message}".
Текущая дата: {{current_date}}.

**ПРАВИЛА:** Не вычисляй даты. "завтра" -> "date_scope":"завтра". "10.07" -> "date":"10.07.2025".

--- ПРИМЕРЫ ---
*   Цель: add_task, Известно: {"task_name": "Сделать отчет"}, Вопрос: "Какой срок у задачи?", Ответ: "до конца недели" -> {"action": "execute_command", "command_details": {"action": "add_task", "task_name": "Сделать отчет", "due_date_scope": "до конца недели"}}
*   Цель: delete_event, Известно: {"participants": ["Игорь"]}, Вопрос: "Укажите дату", Ответ: "во вторник" -> {"action": "clarify", "message": "Хорошо, встреча с Игорем во вторник. А во сколько?", "context": {"original_action": "delete_event", "extracted_entities": {"participants": ["Игорь"], "date_scope": "вторник", "event_name": "Встреча"}}}

--- СТРУКТУРА КОМАНД ---
{{command_structure}}
`;

const LLM_PROMPT_COMMANDS_STRUCTURE = `
// КАЛЕНДАРЬ
1.  add_event: "action":"add_event", "event_name":"...", "date":"...", "date_scope":"...", "time":"...", "location":"...", "participants":["..."]
2.  query_calendar: "action":"query_calendar", "filters":{...}
3.  update_event: "action":"update_event", "event_name":"...", "date":"...", "time":"...", "field_to_update":"...", "new_value":"..."
4.  delete_event: "action":"delete_event", "event_name":"...", "date":"...", "time":"..."
// ЗАМЕТКИ
5.  add_note: "action":"add_note", "content":"...", "category":"..."
6.  query_notes: "action":"query_notes", "filters":{ "category":"...", "keywords":"..."} // <-- ВОЗВРАЩЕНА НА МЕСТО
7.  delete_note: "action":"delete_note", "keywords":"..."
// ЗАДАЧИ
8.  add_task: "action":"add_task", "task_name":"...", "project":"...", "tags":["..."], "priority":"...", "due_date_scope":"...", "assignee":"..."
9.  query_tasks: "action":"query_tasks", "filters":{...}
10. update_task_status: "action":"update_task_status", "keywords":"...", "filters":{"project":"..."}, "new_status":"..."
11. delete_task: "action":"delete_task", "keywords":"...", "filters":{"project":"..."}
`;