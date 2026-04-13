function generateBranchSummaries() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Запрос периода ---
  let startPrompt = ui.prompt('Период сводки', 'Введите НАЧАЛЬНУЮ дату (ГГГГ-ММ-ДД)\nИли оставьте пустым за ВСЁ ВРЕМЯ:', ui.ButtonSet.OK_CANCEL);
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

  const SUMMARY_MARKER = "### СВОДНЫЙ ОТЧЕТ ###";
  let branchKeys = Object.keys(BRANCHES);

  // --- 2. Обработка каждого филиала ---
  for (let branchName of branchKeys) {
    let sheet = ss.getSheetByName(branchName);
    if (!sheet) continue;

    let lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // Читаем все данные листа (берем первые 6 колонок)
    let data = sheet.getRange(1, 1, lastRow, 6).getValues(); 

    // Ищем, где начинается старый сводный отчет, чтобы его удалить
    let summaryStartRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().startsWith(SUMMARY_MARKER)) {
        summaryStartRow = i + 1;
        break;
      }
    }

    // Определяем, где заканчиваются реальные отзывы
    let actualDataEnd = summaryStartRow > -1 ? summaryStartRow - 1 : lastRow;

    // Удаляем старый отчет, если он был
    if (summaryStartRow > -1) {
      let rowsToDelete = lastRow - summaryStartRow + 1;
      sheet.deleteRows(summaryStartRow, rowsToDelete);
    }

    let employeeData = {};
    let activeSites = new Set();
    let totalStats = { pos: 0, neg: 0, sites: {} };

    // --- 3. Подсчет статистики ---
    for (let i = 1; i < actualDataEnd; i++) {
      let dateRaw = data[i][0]; // Колонка A
      
      if (START_DATE && END_DATE) {
        let rowDate = new Date(dateRaw);
        if (isNaN(rowDate.getTime()) || rowDate < START_DATE || rowDate > END_DATE) continue;
      }

      let score = parseFloat(data[i][1]); // Колонка B
      if (isNaN(score)) continue;
      
      let nameRaw = data[i][2]; // Колонка C
      let site = data[i][4];    // Колонка E (Сайт)
      
      if (!site) continue;
      site = site.toString().trim();

      let name = (nameRaw && nameRaw.toString().trim() !== "") ? nameRaw.toString().trim() : "Без имени / Общие отзывы";

      // Сохраняем сайт, чтобы потом создать для него колонку
      activeSites.add(site);

      // Инициализация структур данных
      if (!employeeData[name]) employeeData[name] = { totalPos: 0, totalNeg: 0, sites: {} };
      if (!employeeData[name].sites[site]) employeeData[name].sites[site] = { pos: 0, neg: 0 };
      if (!totalStats.sites[site]) totalStats.sites[site] = { pos: 0, neg: 0 };

      // Распределение на позитивные (>= 4) и негативные
      if (score >= 4) {
        employeeData[name].sites[site].pos++;
        employeeData[name].totalPos++;
        totalStats.sites[site].pos++;
        totalStats.pos++;
      } else {
        employeeData[name].sites[site].neg++;
        employeeData[name].totalNeg++;
        totalStats.sites[site].neg++;
        totalStats.neg++;
      }
    }

    let sitesArray = Array.from(activeSites).sort();
    
    // Если за выбранный период отзывов нет вообще — пропускаем филиал
    if (sitesArray.length === 0) continue;

    // --- 4. Формирование таблицы (только нужные колонки) ---
    let headers = ["Сотрудник"];
    sitesArray.forEach(site => {
      headers.push(`${site} (+)`, `${site} (-)`);
    });
    headers.push("ИТОГО (+)", "ИТОГО (-)");

    let outputRows = [];
    
    // Строка ИТОГО
    let totalsRow = ["ИТОГО"];
    sitesArray.forEach(site => {
      totalsRow.push(totalStats.sites[site].pos, totalStats.sites[site].neg);
    });
    totalsRow.push(totalStats.pos, totalStats.neg);

    // Сортировка имен (общие отзывы в конец)
    let empNames = Object.keys(employeeData).sort((a, b) => {
      if (a === "Без имени / Общие отзывы") return 1;
      if (b === "Без имени / Общие отзывы") return -1;
      return a.localeCompare(b);
    });

    for (let name of empNames) {
      let emp = employeeData[name];
      let row = [name];
      sitesArray.forEach(site => {
        row.push(emp.sites[site] ? emp.sites[site].pos : 0);
        row.push(emp.sites[site] ? emp.sites[site].neg : 0);
      });
      row.push(emp.totalPos, emp.totalNeg);
      outputRows.push(row);
    }

    // Собираем всё вместе с маркером периода
    let periodLabel = START_DATE ? `${SUMMARY_MARKER} (с ${START_DATE.toLocaleDateString()} по ${END_DATE.toLocaleDateString()})` : SUMMARY_MARKER;
    let finalData = [[periodLabel, ...Array(headers.length - 1).fill("")], headers, totalsRow, ...outputRows];
    
    // --- 5. Запись на лист ---
    let writeStartRow = actualDataEnd + 2; // Отступаем одну пустую строку
    let targetRange = sheet.getRange(writeStartRow, 1, finalData.length, headers.length);
    
    targetRange.setValues(finalData);
    
    // Оформление
    sheet.getRange(writeStartRow, 1, 1, headers.length).setBackground("#d9d9d9").setFontWeight("bold"); // Маркер
    sheet.getRange(writeStartRow + 1, 1, 2, headers.length).setFontWeight("bold"); // Заголовки и Итого
    sheet.getRange(writeStartRow + 1, 1, finalData.length - 1, headers.length).setBorder(true, true, true, true, true, true); // Рамки
    sheet.autoResizeColumns(1, headers.length);
  }

  let periodMsg = START_DATE ? `за период с ${START_DATE.toLocaleDateString()} по ${END_DATE.toLocaleDateString()}` : `за всё время`;
  ui.alert("Готово!", `Сводные таблицы ${periodMsg} успешно созданы внизу данных на каждом листе филиала.`, ui.ButtonSet.OK);
}