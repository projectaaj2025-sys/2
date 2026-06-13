// ==================== КОНФИГУРАЦИЯ ====================
const SHEET_ID = 'ВАШ_ID_ГУГЛ_ТАБЛИЦЫ'; // <-- ЗАМЕНИТЕ НА ID ВАШЕЙ ТАБЛИЦЫ
const QUESTIONS_SHEET = 'Questions';
const RESULTS_SHEET = 'Results';
const SETTINGS_SHEET = 'Settings';

// ==================== CORS ====================
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getQuestions') {
    return getQuestions();
  } else if (action === 'getResults') {
    return getResults();
  } else if (action === 'getSettings') {
    return getSettings();
  }

  return jsonResponse({error: 'Unknown action'});
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'saveResult') {
    return saveResult(data);
  } else if (data.action === 'saveSettings') {
    return saveSettings(data);
  }

  return jsonResponse({success: true});
}

function doOptions(e) {
  return jsonResponse({});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== НАСТРОЙКИ ====================
function getSettings() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!sheet) {
      // Создаём лист настроек по умолчанию
      sheet = ss.insertSheet(SETTINGS_SHEET);
      sheet.appendRow(['Key', 'Value']);
      sheet.appendRow(['allowRetake', 'true']);
      sheet.appendRow(['testTitle', 'Тестирование']);
      sheet.appendRow(['timeMinutes', '30']);
      sheet.appendRow(['questionsCount', '30']);
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }

    return jsonResponse(settings);
  } catch (error) {
    return jsonResponse({allowRetake: 'true', error: error.toString()});
  }
}

function saveSettings(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(SETTINGS_SHEET);
      sheet.appendRow(['Key', 'Value']);
    }

    // Очищаем и перезаписываем
    sheet.clear();
    sheet.appendRow(['Key', 'Value']);

    for (const [key, value] of Object.entries(data.settings)) {
      sheet.appendRow([key, value]);
    }

    return jsonResponse({success: true});
  } catch (error) {
    return jsonResponse({error: error.toString()});
  }
}

// ==================== ЗАГРУЗКА ВОПРОСОВ ====================
function getQuestions() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(QUESTIONS_SHEET);
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return jsonResponse([]);
    }

    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const questions = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const question = {};

      headers.forEach((header, index) => {
        const value = row[index];
        question[header] = value !== undefined && value !== null ? String(value) : '';
      });

      if (question.question && question.question.trim() !== '') {
        questions.push(question);
      }
    }

    return jsonResponse(questions);
  } catch (error) {
    return jsonResponse({error: error.toString()});
  }
}

// ==================== СОХРАНЕНИЕ РЕЗУЛЬТАТА ====================
function saveResult(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(RESULTS_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(RESULTS_SHEET);
      sheet.appendRow(['Timestamp', 'Name', 'Score', 'Total', 'Percent', 'TimeSpent', 'Device', 'Answers', 'Questions']);
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.score || 0,
      data.total || 0,
      data.percent || 0,
      data.timeSpent || 0,
      data.device || 'unknown',
      JSON.stringify(data.answers || {}),
      JSON.stringify(data.questions || [])
    ]);

    return jsonResponse({success: true});
  } catch (error) {
    return jsonResponse({error: error.toString()});
  }
}

// ==================== ПОЛУЧЕНИЕ РЕЗУЛЬТАТОВ ====================
function getResults() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(RESULTS_SHEET);

    if (!sheet) {
      return jsonResponse([]);
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return jsonResponse([]);
    }

    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const result = {};

      headers.forEach((header, index) => {
        const value = row[index];
        if (value instanceof Date) {
          result[header] = value.toISOString();
        } else {
          result[header] = value !== undefined && value !== null ? String(value) : '';
        }
      });

      results.push(result);
    }

    return jsonResponse(results);
  } catch (error) {
    return jsonResponse({error: error.toString()});
  }
}