function runLPReport() {
  const propertiesProp =
    PropertiesService.getScriptProperties().getProperties();
  const { PROPERTY_ID, SPREADSHEET_ID, SHEET_NAME } = propertiesProp;

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  const lastColumn = sheet.getLastColumn();

  const METRIC_NAMES = ["activeUsers"];
  const DIMENSION_NAMES = ["eventName", "landingPagePlusQueryString"];

  const START_DATE = formatDateToYYYYMMDD(sheet.getRange("B2").getValue());
  const END_DATE = formatDateToYYYYMMDD(sheet.getRange("C2").getValue());

  const mediaNameList = {
    Google: "google_cpc",
    "Yahoo!": "yahoo_cpc",
    Microsoft: "microsoft_cpc",
  };

  const metrics = createAnalyticsDataItems(
    AnalyticsData.newMetric,
    METRIC_NAMES
  );
  const dimensions = createAnalyticsDataItems(
    AnalyticsData.newDimension,
    DIMENSION_NAMES
  );

  const dateRange = AnalyticsData.newDateRange();
  dateRange.startDate = START_DATE;
  dateRange.endDate = END_DATE;

  // console.log("dateRange")
  // console.log(dateRange)

  const request = AnalyticsData.newRunReportRequest();
  request.dimensions = dimensions;
  request.metrics = metrics;
  request.dateRanges = dateRange;

  try {
    const report = AnalyticsData.Properties.runReport(
      request,
      "properties/" + PROPERTY_ID
    );

    if (!report.rows) {
      Logger.log("No rows returned.");
      return;
    }

    const eventData = processEventData(sheet, lastColumn);

    const eventDataTmp = new Map();

    const firstFilterCondition = [];

    for (let i = 2; i < lastColumn; i += 2) {
      eventDataTmp.set(`${sheet.getRange(5, i).getValue()}${i / 2}`, 0);
      eventDataTmp.set(`completionRate${i / 2}`, "");
      firstFilterCondition.push(`${sheet.getRange(5, i).getValue()}${i / 2}`);
    }

    // console.log([...eventDataTmp.entries()]);

    const filterConditionsObject = {};

    for (let i = 2; i < lastColumn; i += 2) {
      const filterConditions = [];

      for (let row = 5; row <= 7; row++) {
        const cellValue = sheet.getRange(row, i).getValue();
        if (cellValue !== "") {
          filterConditions.push(cellValue);
        }
      }

      filterConditionsObject[firstFilterCondition[i / 2 - 1]] =
        filterConditions;
    }

    const eventOrder = [
      "session_start",
      "mcv_cushion_cv_button_popup01_cb",
      "mcv_cushion_cv_button01_cb",
      "mcv_cushion_cv_button02_cb",
      "mcv_cushion_cv_button03_cb",
      "mcv_cushion_cv_button04_cb",
      "mcv_cushion_btn_cb",
      "mcv_input_start_cb",
      "mcv_input_zip_cb",
      "mcv_input_current_street_cb",
      "mcv_input_new_prefecture_cb",
      "mcv_next_button_cb",
      "mcv_input_name_cb",
      "mcv_input_mail_cb",
      "mcv_input_tel_cb",
      "mcv_input_cv_cb",
      "thanks",
    ];

    const obj = Object.fromEntries(eventDataTmp);

    function updateEventData(eventData, eventName, propertyName, row) {
      if (!eventData[eventName]) {
        return;
      }

      const metricValueSum = row.metricValues.reduce(
        (acc, curr) => acc + parseInt(curr.value),
        0
      );
      eventData[eventName][propertyName] += metricValueSum;
    }

    for (const row of report.rows) {
      const eventName = row.dimensionValues[0].value;
      const landingPagePath = row.dimensionValues[1].value;
      //            if(eventName === "mcv_input_start_cb"){
      // // console.log("eventName")
      // // console.log(eventName)
      // console.log("row")
      // console.log(row)
      // console.log("landingPagePath")
      // console.log(row.dimensionValues[1].value)
      // }

      if (
        !eventData[eventName] &&
        eventOrder.some((orderEvent) => eventName.includes(orderEvent))
      ) {
        eventData[eventName] = {
          eventName: eventName,
          ...obj,
        };
      }

      for (const propertyName of firstFilterCondition) {
        const conditions = filterConditionsObject[propertyName];
        if (conditions.length === 2) {
          if (
            conditions[1] === "全体" &&
            landingPagePath.includes(conditions[0])
          ) {
            updateEventData(eventData, eventName, propertyName, row);
          } else if (
            landingPagePath.includes(conditions[0]) &&
            landingPagePath.includes(mediaNameList[conditions[1]])
          ) {
            updateEventData(eventData, eventName, propertyName, row);
          }
          // continue;
        } else if (conditions.length === 3) {
          if (
            conditions[1] === "全体" &&
            landingPagePath.includes(conditions[0]) &&
            landingPagePath.includes(conditions[1])
          ) {
            updateEventData(eventData, eventName, propertyName, row);
          } else if (
            landingPagePath.includes(conditions[0]) &&
            landingPagePath.includes(conditions[1]) &&
            landingPagePath.includes(mediaNameList[conditions[2]])
          ) {
            updateEventData(eventData, eventName, propertyName, row);
          }
        }
      }
    }

    const startRow = 8;
    const startColumn = 1;
    const numRows = sheet.getMaxRows() - startRow + 1;
    const numColumns = sheet.getMaxColumns() - startColumn + 1;

    sheet.getRange(startRow, startColumn, numRows, numColumns).clear();

    let headerName = ["イベント名"];
    for (i = 0; i < (lastColumn - 1) / 2; i++) {
      headerName.push("ユーザー数", "完了率");
    }
    sheet.appendRow(headerName);

    for (const eventName of eventOrder) {
      if (eventData[eventName]) {
        const eventValues = Object.values(eventData[eventName]);
        sheet.appendRow(eventValues);
      }
    }

    const columns = [2, 4, 6, 8, 10];

    const startCompletionRateRow = 14;
    const endCompletionRateRow = 23;

    for (let colIndex = 0; colIndex < columns.length; colIndex++) {
      const col = columns[colIndex];

      const startCompletionRateRange = sheet.getRange(
        startCompletionRateRow,
        col
      );
      const startRowRange = sheet.getRange(startRow + 1, col);

      const targetCell = sheet.getRange(startCompletionRateRow, col + 1);

      const startCompletionRateA1 = startCompletionRateRange.getA1Notation();
      const startRowA1 = startRowRange.getA1Notation();

      const startCompletionRateFormula = `=${startCompletionRateA1}/${startRowA1}`;
      targetCell.setFormula(startCompletionRateFormula);

      for (
        let row = startCompletionRateRow;
        row <= endCompletionRateRow;
        row++
      ) {
        const currentRange = sheet.getRange(row, col);
        const nextRange = sheet.getRange(row + 1, col);

        const currentA1 = currentRange.getA1Notation();
        const nextA1 = nextRange.getA1Notation();

        const formula = `=${nextA1}/${currentA1}`;
        const targetRowCell = sheet.getRange(row + 1, col + 1);
        targetRowCell.setFormula(formula);
      }
    }

    // Logger.log('Report spreadsheet created: %s', spreadsheet.getUrl());
  } catch (e) {
    Logger.log(e);
  }
}

function formatDateToYYYYMMDD(date) {
  const formattedDate = new Date(date);
  const year = formattedDate.getFullYear();
  const month = (formattedDate.getMonth() + 1).toString().padStart(2, "0");
  const day = formattedDate.getDate().toString().padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function createAnalyticsDataItems(itemType, names) {
  return names.map((name) => {
    const item = itemType();
    item.name = name;
    return item;
  });
}

function processEventData(sheet, lastColumn) {
  const eventData = {};
  for (let i = 2; i < lastColumn; i += 2) {
    const key = sheet.getRange(5, i).getValue();
    eventData[key] = 0;
  }
  return eventData;
}
