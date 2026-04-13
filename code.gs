// ==========================================
// НАСТРОЙКИ И ДАННЫЕ
// ==========================================

const DOQ_AUTH = {
  TOKEN_ALMATY: { email: "emirmed@mail.ru", password: "KeY5J2XLAu", city_id: 3 },
  TOKEN_ASTANA: { email: "emirmedastana@doq.kz", password: "KeY5J2XLAu", city_id: 1 },
  TOKEN_SHYMKENT: { email: "emirmedshymkent@doq.kz", password: "emir131313", city_id: 5 },
  TOKEN_ALASH: { email: "clinic@mail.ru", password: "R4#tV9!a", city_id: 3 }  
};

const TWO_GIS_AUTH = {
  login: "m.orazkhan@emirmed.kz",
  password: "11010010qQ!"
};

// Ключи Algolia
const ALGOLIA_APP_ID = "T0Z5OYKY4Q"; 
const ALGOLIA_API_KEY = "08a899e50b646efe7469e059ec0340b4"; 

const BRANCHES = {
  "проспект Серкебаева, 81": { doq_id: 1428, gis_id: "70000001104842090", auth: "TOKEN_ALASH" },
  "Калкаман микрорайон, 4/6": { doq_id: 1206, gis_id: "70000001094462195", auth: "TOKEN_ALMATY" },
  "улица Манаса, 53а": { doq_id: 154, gis_id: "70000001100799536", auth: "TOKEN_ALMATY" },
  "улица Нусупбекова, 26/1": { doq_id: 430, gis_id: "70000001044956649", auth: "TOKEN_ALMATY" },
  "улица Пограничная, 1/1": { doq_id: 1149, gis_id: "70000001089989162", auth: "TOKEN_ALMATY" },
  "улица Розыбакиева, 37в": { doq_id: 627, gis_id: "70000001062947400", auth: "TOKEN_ALMATY" },
  "проспект Турара Рыскулова, 143в": { doq_id: 1158, gis_id: "70000001089454660", auth: "TOKEN_ALMATY" },
  "проспект Серкебаева, 79": { doq_id: 1103, gis_id: "70000001089150778", auth: "TOKEN_ALMATY" },
  "улица Куйши Дина, 9": { doq_id: 973, gis_id: "70000001075209941", auth: "TOKEN_ASTANA" },
  "улица Сауран, 1": { doq_id: 1395, gis_id: "70000001106037513", auth: "TOKEN_ASTANA" },
  "18-й микрорайон, 44": { doq_id: 989, gis_id: "70000001080895732", auth: "TOKEN_SHYMKENT" },
  "улица Рашидова, 36/15": { doq_id: 990, gis_id: "70000001080895653", auth: "TOKEN_SHYMKENT" }
};

// ==========================================
// МЕНЮ ПРИ ОТКРЫТИИ ТАБЛИЦЫ
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Меню Отзывов')
    .addItem('1. Собрать отзывы (DOQ + 2GIS)', 'fetchReviews')
    .addItem('2. Нормализовать имена 2GIS', 'normalize2GisNames')
    .addItem('3. Сформировать общую сводную', 'generateSummary')
    .addItem('4. Сформировать сводные по филиалам', 'generateBranchSummaries')
    .addItem('5. Создать Дашборд с графиками 📊', 'generateDashboard')
    .addItem('6. Удалить отзывы за период', 'deleteReviewsByPeriod')
    .addItem('7. Сгенерировать Отчет (Google Doc) 📄', 'generateDocReport')
    .addToUi();
}

// ==========================================
// ЧАСТЬ 1: СБОР ОТЗЫВОВ С ФИЛЬТРАЦИЕЙ
// ==========================================

function fetchReviews() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Запрашиваем даты
  let startPrompt = ui.prompt('Настройка периода', 'Введите НАЧАЛЬНУЮ дату (ГГГГ-ММ-ДД):', ui.ButtonSet.OK_CANCEL);
  if (startPrompt.getSelectedButton() !== ui.Button.OK) return;
  
  let endPrompt = ui.prompt('Настройка периода', 'Введите КОНЕЧНУЮ дату (ГГГГ-ММ-ДД):', ui.ButtonSet.OK_CANCEL);
  if (endPrompt.getSelectedButton() !== ui.Button.OK) return;

  const START_DATE = new Date(`${startPrompt.getResponseText().trim()}T00:00:00+05:00`);
  const END_DATE = new Date(`${endPrompt.getResponseText().trim()}T23:59:59+05:00`);

  if (isNaN(START_DATE.getTime()) || isNaN(END_DATE.getTime()) || START_DATE > END_DATE) {
    ui.alert('Ошибка', 'Неверный формат дат или начальная дата больше конечной.', ui.ButtonSet.OK);
    return;
  }

  // 2. Выбор филиалов
  let branchKeys = Object.keys(BRANCHES);
  let branchListText = branchKeys.map((b, i) => `${i + 1}. ${b}`).join('\n');
  
  let branchPrompt = ui.prompt(
    'Выбор филиалов', 
    `Введите номера нужных филиалов через запятую (например: 1, 3, 4)\nИли введите 0 для выбора ВСЕХ филиалов:\n\n${branchListText}`, 
    ui.ButtonSet.OK_CANCEL
  );
  if (branchPrompt.getSelectedButton() !== ui.Button.OK) return;
  
  let selectedInput = branchPrompt.getResponseText().trim();
  let branchesToProcess = [];
  
  if (selectedInput === "0") {
    branchesToProcess = branchKeys;
  } else {
    let indexes = selectedInput.split(',').map(s => parseInt(s.trim()) - 1);
    indexes.forEach(idx => {
      if (idx >= 0 && idx < branchKeys.length) branchesToProcess.push(branchKeys[idx]);
    });
  }

  if (branchesToProcess.length === 0) {
    ui.alert('Ошибка', 'Филиалы не выбраны или введен неверный номер.', ui.ButtonSet.OK);
    return;
  }

  // 3. Запрос на очистку листа
  let clearResponse = ui.alert('Очистка старых данных', 'Удалить все существующие отзывы с листов выбранных филиалов перед сбором?', ui.ButtonSet.YES_NO_CANCEL);
  if (clearResponse === ui.Button.CANCEL || clearResponse === ui.Button.CLOSE) return;
  let clearSheetsBeforeRun = (clearResponse === ui.Button.YES);

  console.log("=== НАЧАЛО СБОРА ОТЗЫВОВ ===");
  
  let stats = { doqTotal: 0, gisTotal: 0, branchesProcessed: 0 };

  let doqTokens = {};
  for (let key in DOQ_AUTH) doqTokens[key] = getDoqToken(DOQ_AUTH[key].email, DOQ_AUTH[key].password);
  
  let gisToken;
  try {
    gisToken = get2GisToken();
    console.log("[AUTH] Токен 2GIS успешно получен.");
  } catch (e) {
    console.error("[AUTH] Ошибка получения токена 2GIS: " + e.message);
    ui.alert("Ошибка авторизации", "Не удалось получить токен 2GIS. Проверьте логин и пароль.", ui.ButtonSet.OK);
    return;
  }

  for (let branchName of branchesToProcess) {
    let branch = BRANCHES[branchName];
    stats.branchesProcessed++;
    console.log(`\n===========================================`);
    console.log(`--- Обработка филиала: ${branchName} ---`);
    console.log(`===========================================`);
    
    let branchData = []; // Сюда соберем DOQ и 2GIS
    
    // --- ПАРСИНГ DOQ ---
    let doqToken = doqTokens[branch.auth];
    let doqNextUrl = `https://api.doq.kz/api/v0/cabinet/${branch.doq_id}/feedbacks/?page_size=10&page=1`;
    let doqBranchCount = 0;
    let doqOldConsecutive = 0; // Счетчик старых отзывов подряд
    
    console.log(`[DOQ] Старт сбора отзывов...`);
    while (doqNextUrl) {
      console.log(`[DOQ] Запрашиваем страницу: ${doqNextUrl}`);
      let response = UrlFetchApp.fetch(doqNextUrl, { method: "GET", headers: { "Authorization": "Bearer " + doqToken }, muteHttpExceptions: true });
      let responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        try {
          let json = JSON.parse(response.getContentText());
          if (json.results && json.results.length > 0) {
            console.log(`[DOQ] Получено отзывов на странице: ${json.results.length}`);
            let stopFetching = false;
            for (let r of json.results) {
              let reviewDate = new Date(r.updated_at);
              if (reviewDate < START_DATE) { 
                doqOldConsecutive++;
                if (doqOldConsecutive >= 25) { 
                  console.log(`[DOQ] Найдено 25 старых отзыва подряд. Останавливаем сбор.`);
                  stopFetching = true; 
                  break; 
                }
              } else {
                doqOldConsecutive = 0; // Сбрасываем счетчик
                if (reviewDate <= END_DATE) {
                  // Создан | Рейтинг | кому | Текст | Сайт | Gemini_Formula
                  branchData.push([r.updated_at, r.score / 2, r.doctor_name, r.text, "DOQ", ""]);
                  doqBranchCount++;
                }
              }
            }
            doqNextUrl = stopFetching ? null : json.next;
          } else {
            console.log(`[DOQ] Массив results пуст. Окончание списка.`);
            doqNextUrl = null;
          }
        } catch (e) {
          console.error(`[DOQ] Ошибка парсинга JSON: ${e.message}`);
          doqNextUrl = null;
        }
      } else {
        console.error(`[DOQ] Ошибка API! HTTP Код: ${responseCode}, Ответ: ${response.getContentText()}`);
        doqNextUrl = null;
      }
    }
    console.log(`[DOQ] Завершено для ${branchName}. Собрано новых: ${doqBranchCount}`);
    stats.doqTotal += doqBranchCount;

    // --- ПАРСИНГ 2GIS ---
    let gisUrl = `https://api.account.2gis.com/api/1.0/presence/branch/${branch.gis_id}/reviews?pinRequestedFirst=true&limit=20`;
    let gisHasMore = true;
    let gisBranchCount = 0;
    let gisOldConsecutive = 0; // Счетчик старых отзывов подряд

    console.log(`[2GIS] Старт сбора отзывов...`);
    while (gisHasMore) {
      console.log(`[2GIS] Запрашиваем страницу...`);
      let response = UrlFetchApp.fetch(gisUrl, { method: "GET", headers: { "Authorization": "Bearer " + gisToken }, muteHttpExceptions: true });
      let responseCode = response.getResponseCode();

      if (responseCode === 200) {
        try {
          let json = JSON.parse(response.getContentText());
          if (json.result && json.result.items && json.result.items.length > 0) {
            console.log(`[2GIS] Получено отзывов на странице: ${json.result.items.length}`);
            let stopFetching = false;
            for (let r of json.result.items) {
              let reviewDate = new Date(r.dateCreated);
              if (reviewDate < START_DATE) {
                gisOldConsecutive++;
                if (gisOldConsecutive >= 3) { 
                  console.log(`[2GIS] Найдено 3 старых отзыва подряд. Останавливаем сбор.`);
                  stopFetching = true; 
                  break; 
                }
              } else {
                gisOldConsecutive = 0; // Сбрасываем счетчик
                if (reviewDate <= END_DATE) {
                  // Формулу вставим позже при записи на лист
                  branchData.push([r.dateCreated, r.rating, "", r.text, "2GIS", ""]);
                  gisBranchCount++;
                }
              }
            }
            if (stopFetching) {
              gisHasMore = false;
            } else {
              let lastDate = json.result.items[json.result.items.length - 1].dateCreated;
              gisUrl = `https://api.account.2gis.com/api/1.0/presence/branch/${branch.gis_id}/reviews?pinRequestedFirst=true&offsetDate=${encodeURIComponent(lastDate)}&limit=20`;
              console.log(`[2GIS] Переход к следующей странице. Офсет: ${lastDate}`);
            }
          } else {
            console.log(`[2GIS] Массив items пуст или отсутствует. Окончание списка.`);
            gisHasMore = false;
          }
        } catch (e) {
          console.error(`[2GIS] Ошибка парсинга JSON: ${e.message}`);
          console.error(`[2GIS] Текст ответа: ${response.getContentText().substring(0, 300)}...`);
          gisHasMore = false;
        }
      } else {
        console.error(`[2GIS] Ошибка API! HTTP Код: ${responseCode}`);
        console.error(`[2GIS] Тело ответа: ${response.getContentText()}`);
        gisHasMore = false;
      }
    }
    console.log(`[2GIS] Завершено для ${branchName}. Собрано новых: ${gisBranchCount}`);
    stats.gisTotal += gisBranchCount;

    // --- ЗАПИСЬ НА ЛИСТ ФИЛИАЛА ---
    if (branchData.length > 0) {
      // Ищем лист, если нет - создаем
      let sheet = ss.getSheetByName(branchName);
      if (!sheet) {
        sheet = ss.insertSheet(branchName);
        clearSheetsBeforeRun = true; // Если лист только что создан, его надо инициализировать
      }
      
      if (clearSheetsBeforeRun) {
        sheet.clear();
        sheet.appendRow(["Создан", "Рейтинг", "кому", "Текст", "Сайт", "Gemini_Formula"]);
        sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
      }

      let lastRow = sheet.getLastRow();
      if (lastRow === 0) {
        sheet.appendRow(["Создан", "Рейтинг", "кому", "Текст", "Сайт", "Gemini_Formula"]);
        lastRow = 1;
      }

      const prompt = `You are an AI assistant specialized in text extraction. You will be provided with a review text written in Russian or Kazakh. Your task is to extract the names of the employees mentioned in the text based on the strict rules below.
RULES:
1. Extract the name (First Name, Last Name, Patronymic if available) of the employee mentioned in the text.
2. Convert the extracted name to its base dictionary form.
3. Check for specific medical roles: If the text indicates a nurse, receptionist, or administrator, you MUST prepend their role to the name.
4. The role MUST ALWAYS be written in Russian: 'медсестра', 'регистратор', or 'администратор'. Example: 'медсестра Айым'. 
5. Do not include specialties for other types of staff.
6. ERROR HANDLING: If no name is found, output exactly 'ОШИБКА'. 
7. STRICT CONSTRAINT: Output ONLY the extracted data or 'ОШИБКА'.`;

      // Проставляем формулы для 2GIS
      for (let j = 0; j < branchData.length; j++) {
        let rowNumForFormula = lastRow + j + 1;
        if (branchData[j][4] === "2GIS") {
           // Колонка D (4) - это текст отзыва
           branchData[j][5] = `=AI("${prompt}"; D${rowNumForFormula})`;
        }
      }

      sheet.getRange(lastRow + 1, 1, branchData.length, 6).setValues(branchData);
      console.log(`[ЗАПИСЬ] Данные записаны на лист: ${branchName} (Добавлено строк: ${branchData.length})`);
    } else {
      console.log(`[ЗАПИСЬ] Нет новых данных для записи на лист: ${branchName}`);
    }
  }

  console.log("=== СБОР ЗАВЕРШЕН ===");
  ui.alert('📊 Отчет: Сбор отзывов', `Сбор завершен!\n\nОбработано филиалов: ${stats.branchesProcessed}\nНовых отзывов DOQ: ${stats.doqTotal}\nНовых отзывов 2GIS: ${stats.gisTotal}`, ui.ButtonSet.OK);
}

// ==========================================
// ЧАСТЬ 2: НОРМАЛИЗАЦИЯ ИМЕН (ALGOLIA -> DOQ)
// ==========================================
function normalize2GisNames() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let branchKeys = Object.keys(BRANCHES);
  
  // Добавил в статистику errorsCleared для наглядности
  let stats = { processed: 0, found: 0, medStaff: 0, notFound: 0, skipped: 0, errorsCleared: 0 };

  for (let branchName of branchKeys) {
    let sheet = ss.getSheetByName(branchName);
    if (!sheet) continue;

    let data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let rowNum = i + 1;
      let site = data[i][4]; // Сайт (DOQ или 2GIS)
      
      if (site !== "2GIS") continue;

      let komuCurrent = data[i][2]; 
      let geminiOutput = data[i][5]; 

      // --- ДОРАБОТКА 1: Если значение "ОШИБКА", удаляем его и идем дальше ---
      if (geminiOutput === "ОШИБКА") {
        sheet.getRange(rowNum, 6).clearContent();
        stats.errorsCleared++;
        continue;
      }

      // Убрали "ОШИБКА" из условия ниже, так как отработали его выше
      if (!komuCurrent && geminiOutput && geminiOutput !== "Loading..." && !geminiOutput.toString().startsWith("#")) {
        stats.processed++;
        
        let isMedStaff = /медсестра|регистратор|администратор/i.test(geminiOutput);
        
        if (isMedStaff) {
           sheet.getRange(rowNum, 3).setValue(geminiOutput);
           sheet.getRange(rowNum, 6).clearContent(); 
           stats.medStaff++;
           continue; 
        }

        let branchConfig = BRANCHES[branchName];
        let cityId = DOQ_AUTH[branchConfig.auth].city_id;
        let doqTargetId = branchConfig.doq_id;

        let algoliaUrl = `https://${ALGOLIA_APP_ID.toLowerCase()}-dsn.algolia.net/1/indexes/dev_doq?query=${encodeURIComponent(geminiOutput)}&filters=city%3D${cityId}`;
        let algoliaResponse = UrlFetchApp.fetch(algoliaUrl, { method: "GET", headers: { "X-Algolia-Application-Id": ALGOLIA_APP_ID, "X-Algolia-API-Key": ALGOLIA_API_KEY }, muteHttpExceptions: true });

        if (algoliaResponse.getResponseCode() === 200) {
          let hits = JSON.parse(algoliaResponse.getContentText()).hits;
          let foundFinalName = "";

          for (let hit of hits) {
            if (hit.category === "Doctors") {
              let doctorResp = UrlFetchApp.fetch(`https://api.doq.kz/api/v1/doctors/${hit.slug}/?expand=clinic_branches`, { muteHttpExceptions: true });
              if (doctorResp.getResponseCode() === 200) {
                let doctorInfo = JSON.parse(doctorResp.getContentText());
                if (doctorInfo.clinic_branches.some(b => b.id === doqTargetId)) {
                  foundFinalName = doctorInfo.name;
                  break;
                }
              }
            }
          }

          if (foundFinalName) {
            sheet.getRange(rowNum, 3).setValue(foundFinalName);
            // --- ДОРАБОТКА 3: Удаляем значение из столбца Gemini, если нашли совпадение ---
            sheet.getRange(rowNum, 6).clearContent();
            stats.found++;
          } else {
            // --- ДОРАБОТКА 2: Если не нашли совпадения — ничего не трогаем ---
            stats.notFound++;
          }
        }
      } else {
        stats.skipped++;
      }
    }
  }

  ui.alert('📊 Отчет: Нормализация 2GIS', `Обработано строк: ${stats.processed}\nНайдено в DOQ: ${stats.found}\nМедперсонал (без поиска): ${stats.medStaff}\nНе найдено: ${stats.notFound}\nУдалено «ОШИБКА»: ${stats.errorsCleared}`, ui.ButtonSet.OK);
}

// ==========================================
// ЧАСТЬ 3: СОЗДАНИЕ СВОДНОЙ ТАБЛИЦЫ
// ==========================================

function generateSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- ДОРАБОТКА 1: Запрашиваем период у пользователя ---
  let startPrompt = ui.prompt('Период сводки', 'Введите НАЧАЛЬНУЮ дату (ГГГГ-ММ-ДД)\nИли оставьте поле пустым для сводки за ВСЁ ВРЕМЯ:', ui.ButtonSet.OK_CANCEL);
  if (startPrompt.getSelectedButton() === ui.Button.CANCEL) return;
  
  let startDateStr = startPrompt.getResponseText().trim();
  let endDateStr = "";
  
  if (startDateStr) {
    let endPrompt = ui.prompt('Период сводки', 'Введите КОНЕЧНУЮ дату (ГГГГ-ММ-ДД):', ui.ButtonSet.OK_CANCEL);
    if (endPrompt.getSelectedButton() === ui.Button.CANCEL) return;
    endDateStr = endPrompt.getResponseText().trim();
  }

  const START_DATE = startDateStr ? new Date(`${startDateStr}T00:00:00+05:00`) : null;
  const END_DATE = endDateStr ? new Date(`${endDateStr}T23:59:59+05:00`) : null;

  if (START_DATE && END_DATE && (isNaN(START_DATE.getTime()) || isNaN(END_DATE.getTime()) || START_DATE > END_DATE)) {
    ui.alert('Ошибка', 'Неверный формат дат.', ui.ButtonSet.OK);
    return;
  }
  // --------------------------------------------------------

  let sheetSummary = ss.getSheetByName("Сводная") || ss.insertSheet("Сводная");
  sheetSummary.clear();

  const headers = [
    "Сотрудник", 
    "DOQ (+)", "DOQ (-)", "2GIS (+)", "2GIS (-)", 
    "Yandex (+)", "Yandex (-)", "Google (+)", "Google (-)", 
    "103.kz(+)", "103.kz(-)", "Idoctor(+)", "Idoctor(-)", 
    "emirmed(+)", "emirmed(-)", "ИТОГО (+)", "ИТОГО (-)"
  ];

  let employeeData = {};
  
  function initEmployee(name) {
    if (!employeeData[name]) {
      employeeData[name] = { doq_pos: 0, doq_neg: 0, gis_pos: 0, gis_neg: 0, yan_pos: 0, yan_neg: 0, goo_pos: 0, goo_neg: 0, kz103_pos: 0, kz103_neg: 0, idoc_pos: 0, idoc_neg: 0, emir_pos: 0, emir_neg: 0 };
    }
  }

  let branchKeys = Object.keys(BRANCHES);
  for (let branchName of branchKeys) {
    let sheet = ss.getSheetByName(branchName);
    if (!sheet) continue;
    
    let data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let dateRaw = data[i][0]; // Дата (Колонка A)
      
      // --- ДОРАБОТКА 2: Фильтруем по дате ---
      if (START_DATE && END_DATE) {
        let rowDate = new Date(dateRaw);
        if (isNaN(rowDate.getTime()) || rowDate < START_DATE || rowDate > END_DATE) {
          continue; // Пропускаем отзыв, если он не входит в выбранный период
        }
      }
      // ---------------------------------------

      let score = parseFloat(data[i][1]); // Оценка (Колонка B)
      let nameRaw = data[i][2];           // Имя (Колонка C)
      let site = data[i][4];              // Сайт (Колонка E)
      
      let name = (nameRaw && nameRaw.toString().trim() !== "") ? nameRaw.toString().trim() : "Без имени / Общие отзывы";
      initEmployee(name);

      if (site === "DOQ") {
        if (score >= 4) employeeData[name].doq_pos++; else employeeData[name].doq_neg++;
      } else if (site === "2GIS") {
        if (score >= 4) employeeData[name].gis_pos++; else employeeData[name].gis_neg++;
      }
      // Если у тебя появятся условия для Yandex, Google и т.д., добавляй их сюда
    }
  }

  let outputRows = [];
  let totalsRow = Array(headers.length).fill(0);
  totalsRow[0] = "ИТОГО"; 
  
  for (let empName in employeeData) {
    let emp = employeeData[empName];
    let totalPos = emp.doq_pos + emp.gis_pos + emp.yan_pos + emp.goo_pos + emp.kz103_pos + emp.idoc_pos + emp.emir_pos;
    let totalNeg = emp.doq_neg + emp.gis_neg + emp.yan_neg + emp.goo_neg + emp.kz103_neg + emp.idoc_neg + emp.emir_neg;
    
    if (totalPos === 0 && totalNeg === 0) continue;

    let row = [
      empName, emp.doq_pos, emp.doq_neg, emp.gis_pos, emp.gis_neg, emp.yan_pos, emp.yan_neg,
      emp.goo_pos, emp.goo_neg, emp.kz103_pos, emp.kz103_neg, emp.idoc_pos, emp.idoc_neg,
      emp.emir_pos, emp.emir_neg, totalPos, totalNeg
    ];
    outputRows.push(row);

    for (let j = 1; j < row.length; j++) totalsRow[j] += row[j];
  }

  outputRows.sort((a, b) => {
    if (a[0] === "Без имени / Общие отзывы") return 1;
    if (b[0] === "Без имени / Общие отзывы") return -1;
    return a[0].localeCompare(b[0]);
  });

  let finalData = [headers, totalsRow, ...outputRows];
  sheetSummary.getRange(1, 1, finalData.length, headers.length).setValues(finalData);
  sheetSummary.getRange(1, 1, 2, headers.length).setFontWeight("bold"); 
  sheetSummary.setFrozenRows(2);
  sheetSummary.autoResizeColumns(1, headers.length); 

  // Выводим информативное сообщение в конце
  let periodMsg = START_DATE ? `за период с ${START_DATE.toLocaleDateString()} по ${END_DATE.toLocaleDateString()}` : `за всё время`;
  ui.alert("Готово!", `Сводная таблица ${periodMsg} успешно сформирована.`, ui.ButtonSet.OK);
}


// ==========================================
// ЧАСТЬ 4: УДАЛЕНИЕ ОТЗЫВОВ ЗА ПЕРИОД
// ==========================================

function deleteReviewsByPeriod() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let startPrompt = ui.prompt('Удаление за период', 'Введите НАЧАЛЬНУЮ дату для удаления (ГГГГ-ММ-ДД):', ui.ButtonSet.OK_CANCEL);
  if (startPrompt.getSelectedButton() !== ui.Button.OK) return;
  
  let endPrompt = ui.prompt('Удаление за период', 'Введите КОНЕЧНУЮ дату для удаления (ГГГГ-ММ-ДД):', ui.ButtonSet.OK_CANCEL);
  if (endPrompt.getSelectedButton() !== ui.Button.OK) return;

  const START_DATE = new Date(`${startPrompt.getResponseText().trim()}T00:00:00+05:00`);
  const END_DATE = new Date(`${endPrompt.getResponseText().trim()}T23:59:59+05:00`);

  if (isNaN(START_DATE.getTime()) || isNaN(END_DATE.getTime()) || START_DATE > END_DATE) {
    ui.alert('Ошибка', 'Неверный формат дат.', ui.ButtonSet.OK);
    return;
  }

  let confirm = ui.alert('Подтверждение', `Удалить отзывы с ${START_DATE.toLocaleDateString()} по ${END_DATE.toLocaleDateString()} со ВСЕХ листов филиалов?`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  let branchKeys = Object.keys(BRANCHES);
  let deletedCount = 0;

  for (let branchName of branchKeys) {
    let sheet = ss.getSheetByName(branchName);
    if (!sheet) continue;

    let lastRow = sheet.getLastRow();
    if (lastRow < 2) continue; // Пропускаем пустые листы или где только заголовки

    // 1. Читаем весь столбец с датами одним махом (Мгновенно)
    let data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let deleteCountBlock = 0;
    let deleteStartRow = -1;

    // 2. Идем снизу вверх по полученному массиву
    for (let i = data.length - 1; i >= 0; i--) {
      let rowNum = i + 2; // Прибавляем 2, так как данные начинаются со 2-й строки
      let rowDate = new Date(data[i][0]);
      
      let shouldDelete = !isNaN(rowDate.getTime()) && rowDate >= START_DATE && rowDate <= END_DATE;

      if (shouldDelete) {
        // Собираем строки в один блок
        deleteStartRow = rowNum; 
        deleteCountBlock++;
        deletedCount++;
      } else {
        // Если встретили строку, которую не надо удалять, удаляем собранный до этого блок
        if (deleteCountBlock > 0) {
          sheet.deleteRows(deleteStartRow, deleteCountBlock);
          deleteCountBlock = 0; // Сбрасываем счетчик блока
        }
      }
    }
    
    // 3. Если самые первые строки подходили под условие, удаляем последний собранный блок
    if (deleteCountBlock > 0) {
      sheet.deleteRows(deleteStartRow, deleteCountBlock);
    }
  }

  ui.alert('Готово', `Удалено строк: ${deletedCount}`, ui.ButtonSet.OK);
}

function generateDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- НАСТРОЙКИ КОЛОНОК ---
  const COL_DATE = 0;   // A: Дата
  const COL_SCORE = 1;  // B: Оценка
  const COL_TEXT = 3;   // D: Текст отзыва (для облака слов)
  const COL_SITE = 4;   // E: Сайт (Источник)

  ss.toast("Сбор данных... Это может занять около минуты.", "Дашборд");

  let dashboardSheet = ss.getSheetByName("Дашборд");
  if (dashboardSheet) ss.deleteSheet(dashboardSheet);
  dashboardSheet = ss.insertSheet("Дашборд", 0);

  let dataSheet = ss.getSheetByName("Данные_Графиков");
  if (dataSheet) { dataSheet.clear(); } 
  else { dataSheet = ss.insertSheet("Данные_Графиков"); dataSheet.hideSheet(); }

  // --- ФУНКЦИЯ ОПРЕДЕЛЕНИЯ РЕГИОНА ПО НАЗВАНИЮ ---
  function getRegion(branchName) {
    let name = branchName.toLowerCase();
    if (name.includes("серкебаева, 81") || name.includes("калкаман") || name.includes("манаса") || 
        name.includes("нусупбекова") || name.includes("пограничная") || name.includes("розыбакиева") || 
        name.includes("рыскулова") || name.includes("серкебаева, 79")) return "Алматы";
    
    if (name.includes("куйши дина") || name.includes("сауран")) return "Астана";
    
    if (name.includes("18-й") || name.includes("18 й") || name.includes("18-мкр") || name.includes("рашидова")) return "Шымкент";
    
    return "Другое";
  }

  let branchKeys = Object.keys(BRANCHES);
  
  let monthlyStats = {};
  let branchStats = {};
  let regionStats = { "Алматы": {pos: 0, neg: 0}, "Астана": {pos: 0, neg: 0}, "Шымкент": {pos: 0, neg: 0}, "Другое": {pos: 0, neg: 0} };
  let ratingDistribution = { "1 звезда": 0, "2 звезды": 0, "3 звезды": 0, "4 звезды": 0, "5 звезд": 0 };
  
  let posWords = [];
  let negWords = [];

  // --- 1. СБОР ДАННЫХ ---
  for (let branchName of branchKeys) {
    let sheet = ss.getSheetByName(branchName);
    if (!sheet) continue;
    
    let region = getRegion(branchName);
    let lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    let data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    branchStats[branchName] = { pos: 0, neg: 0 };

    for (let row of data) {
      let dateRaw = row[COL_DATE];
      let score = parseFloat(row[COL_SCORE]);
      let text = row[COL_TEXT] ? row[COL_TEXT].toString().replace(/[^а-яА-ЯёЁa-zA-Z ]/g, " ").toLowerCase() : "";

      if (isNaN(score) || !dateRaw) continue;

      let d = new Date(dateRaw);
      if (isNaN(d.getTime())) continue;
      let monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;

      // Инициализация месяца (считаем сумму оценок для среднего рейтинга)
      if (!monthlyStats[monthKey]) monthlyStats[monthKey] = { pos: 0, neg: 0, totalScore: 0, count: 0 };
      
      monthlyStats[monthKey].totalScore += score;
      monthlyStats[monthKey].count++;

      // Распределение по звездам
      let roundedScore = Math.round(score);
      if (roundedScore >= 1 && roundedScore <= 5) {
        let starKey = roundedScore === 1 ? "1 звезда" : (roundedScore > 1 && roundedScore < 5) ? `${roundedScore} звезды` : "5 звезд";
        ratingDistribution[starKey]++;
      }

      let isPos = score >= 4;
      if (isPos) {
        monthlyStats[monthKey].pos++;
        branchStats[branchName].pos++;
        if(regionStats[region]) regionStats[region].pos++;
        if (text.length > 10) posWords.push(text);
      } else {
        monthlyStats[monthKey].neg++;
        branchStats[branchName].neg++;
        if(regionStats[region]) regionStats[region].neg++;
        if (text.length > 10) negWords.push(text);
      }
    }
  }

  // --- 2. ЗАПИСЬ ДАННЫХ ДЛЯ ГРАФИКОВ ---
  let writeRow = 1;

  // 2.1 Динамика (Отзывы + Средняя оценка)
  let monthsSorted = Object.keys(monthlyStats).sort();
  dataSheet.getRange(writeRow, 1, 1, 3).setValues([["Месяц", "Всего отзывов", "Средняя оценка"]]);
  let monthDataRows = monthsSorted.map(m => {
    let stat = monthlyStats[m];
    let avgScore = stat.count > 0 ? (stat.totalScore / stat.count).toFixed(2) : 0;
    return [m, stat.count, parseFloat(avgScore)];
  });
  if (monthDataRows.length > 0) dataSheet.getRange(writeRow + 1, 1, monthDataRows.length, 3).setValues(monthDataRows);
  let monthRange = dataSheet.getRange(writeRow, 1, monthDataRows.length + 1, 3);
  writeRow += monthDataRows.length + 2;

  // 2.2 Филиалы
  dataSheet.getRange(writeRow, 1, 1, 3).setValues([["Филиал", "Положительные", "Негативные"]]);
  let branchDataRows = Object.keys(branchStats).map(b => [b, branchStats[b].pos, branchStats[b].neg]);
  branchDataRows.sort((a, b) => (b[1] + b[2]) - (a[1] + a[2])); 
  if (branchDataRows.length > 0) dataSheet.getRange(writeRow + 1, 1, branchDataRows.length, 3).setValues(branchDataRows);
  let branchRange = dataSheet.getRange(writeRow, 1, branchDataRows.length + 1, 3);
  writeRow += branchDataRows.length + 2;

  // 2.3 Регионы (Лояльность)
  dataSheet.getRange(writeRow, 1, 1, 2).setValues([["Регион", "Лояльность (%)"]]);
  let regionKeys = ["Алматы", "Астана", "Шымкент"];
  let regionDataRows = regionKeys.map(r => {
    let stat = regionStats[r];
    let total = stat.pos + stat.neg;
    let loyalty = total > 0 ? Math.round((stat.pos / total) * 100) : 0;
    return [r, loyalty];
  });
  dataSheet.getRange(writeRow + 1, 1, regionDataRows.length, 2).setValues(regionDataRows);
  let regionRange = dataSheet.getRange(writeRow, 1, regionDataRows.length + 1, 2);
  writeRow += regionDataRows.length + 2;

  // 2.4 Распределение оценок
  dataSheet.getRange(writeRow, 1, 1, 2).setValues([["Оценка", "Количество"]]);
  let ratingRows = Object.keys(ratingDistribution).map(k => [k, ratingDistribution[k]]);
  dataSheet.getRange(writeRow + 1, 1, ratingRows.length, 2).setValues(ratingRows);
  let ratingRange = dataSheet.getRange(writeRow, 1, ratingRows.length + 1, 2);

  // --- 3. ПОСТРОЕНИЕ ГРАФИКОВ ---
  
  // 1. КОМБИНИРОВАННЫЙ: Динамика + Линия рейтинга
  let chartDynamics = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(monthRange)
    .setPosition(2, 2, 0, 0)
    .setOption('title', 'Динамика объёма отзывов и среднего рейтинга')
    .setOption('series', {
      0: {type: 'bars', targetAxisIndex: 0, color: '#4285F4'}, // Столбцы (Количество)
      1: {type: 'line', targetAxisIndex: 1, color: '#FBBC05', lineWidth: 3, pointSize: 5} // Линия (Рейтинг)
    })
    .setOption('vAxes', {
      0: {title: 'Кол-во отзывов'},
      1: {title: 'Средняя оценка', minValue: 1, maxValue: 5}
    })
    .setOption('width', 700)
    .setOption('height', 400)
    .build();
  dashboardSheet.insertChart(chartDynamics);

  // 2. ЛОЯЛЬНОСТЬ ПО РЕГИОНАМ
  let chartRegions = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(regionRange)
    .setPosition(2, 12, 0, 0)
    .setOption('title', 'Лояльность пациентов по регионам (%)')
    .setOption('colors', ['#34A853'])
    .setOption('vAxis', {minValue: 0, maxValue: 100})
    .setOption('width', 500)
    .setOption('height', 400)
    .build();
  dashboardSheet.insertChart(chartRegions);

  // 3. СРАВНЕНИЕ ФИЛИАЛОВ
  let chartBranches = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(branchRange)
    .setPosition(24, 2, 0, 0)
    .setOption('title', 'Сравнение филиалов по количеству отзывов')
    .setOption('isStacked', true)
    .setOption('colors', ['#34A853', '#EA4335'])
    .setOption('width', 700)
    .setOption('height', 600)
    .build();
  dashboardSheet.insertChart(chartBranches);

  // 4. РАСПРЕДЕЛЕНИЕ ОЦЕНОК (1-5 звезд)
  let chartRatings = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(ratingRange)
    .setPosition(24, 12, 0, 0)
    .setOption('title', 'Распределение оценок (1-5 звезд)')
    .setOption('colors', ['#9AA0A6']) // Серый цвет
    .setOption('width', 500)
    .setOption('height', 300)
    .build();
  dashboardSheet.insertChart(chartRatings);


  // --- 4. ОБЛАКО СЛОВ ---
  ss.toast("Графики построены. Генерируем облака слов...", "Дашборд");

  function getWordCloudBlob(textArray, color) {
    if (textArray.length === 0) return null;
    let text = textArray.slice(0, 300).join(" "); 
    
    let payload = {
      format: "png", width: 500, height: 300, fontFamily: "sans-serif",
      colors: [color], text: text, removeStopwords: true, language: "ru"
    };

    let options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

    try {
      let response = UrlFetchApp.fetch("https://quickchart.io/wordcloud", options);
      if (response.getResponseCode() === 200) return response.getBlob();
    } catch (e) { Logger.log("Ошибка генерации облака слов: " + e); }
    return null;
  }

  let posBlob = getWordCloudBlob(posWords, "#34A853");
  let negBlob = getWordCloudBlob(negWords, "#EA4335");

  // Размещение облаков слов сбоку от филиалов
  dashboardSheet.getRange("U23").setValue("Позитивное семантическое ядро").setFontWeight("bold");
  if (posBlob) dashboardSheet.insertImage(posBlob, 21, 24);

  dashboardSheet.getRange("U40").setValue("Негативное семантическое ядро").setFontWeight("bold");
  if (negBlob) dashboardSheet.insertImage(negBlob, 21, 41);

  dashboardSheet.getRange("A1:Z100").setBackground("#F8F9FA");
  dashboardSheet.setHiddenGridlines(true);

  ui.alert("✅ Дашборд готов!", "Перейдите на лист 'Дашборд', чтобы посмотреть графики.", ui.ButtonSet.OK);
}

// ==========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ==========================================

function getDoqToken(email, password) {
  let response = UrlFetchApp.fetch("https://api.doq.kz/api/v0/token/", { method: "POST", contentType: "application/json", payload: JSON.stringify({ email: email, password: password }), muteHttpExceptions: true });
  if (response.getResponseCode() === 200) return JSON.parse(response.getContentText()).user.access_token;
  throw new Error("Не удалось получить токен DOQ");
}

function get2GisToken() {
  let response = UrlFetchApp.fetch("https://api.account.2gis.com/api/1.0/users/auth", { method: "POST", contentType: "application/json", payload: JSON.stringify(TWO_GIS_AUTH), muteHttpExceptions: true });
  if (response.getResponseCode() === 200) return JSON.parse(response.getContentText()).result.access_token;
  throw new Error("Не удалось получить токен 2GIS");
}