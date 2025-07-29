function fillLunchByMonthHeader(monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Plan");
  const ideasSheet = ss.getSheetByName("Lunch ideas");

  const meals = ideasSheet.getRange("A2:A").getValues().flat().filter(x => x);

  const row1 = sheet.getRange("1:1").getValues()[0].map(v => (v || "").toString().trim().toLowerCase());
  const row2 = sheet.getRange("2:2").getValues()[0].map(v => (v || "").toString().trim().toLowerCase());

  const monthCol = row1.findIndex(v => v === monthName.toLowerCase());
  if (monthCol === -1) throw new Error(`Month '${monthName}' not found in row 1.`);

  // Identify table starting with Day, Date, Lunch, Dinner
  let startCol = -1;
  for (let offset = -3; offset <= 0; offset++) {
    const idx = monthCol + offset;
    if (idx < 0 || idx + 3 >= row2.length) continue;

    const headers = row2.slice(idx, idx + 4);
    if (
      headers[0] === "day" &&
      headers[1] === "date" &&
      headers[2] === "lunch" &&
      headers[3] === "dinner"
    ) {
      startCol = idx;
      break;
    }
  }

  if (startCol === -1) {
    throw new Error(`Could not locate table with headers [Day, Date, Lunch, Dinner] under month '${monthName}'.`);
  }

  const dateCol = startCol + 1;
  const lunchCol = startCol + 2;

  const maxRows = sheet.getLastRow() - 2;
  const dateValues = sheet.getRange(3, dateCol + 1, maxRows, 1).getValues().map(r => r[0]);

  // Find first row where date == 1 and continue until date is null/empty
  let startRow = -1;
  for (let i = 0; i < dateValues.length; i++) {
    if (dateValues[i] === 1) {
      startRow = i;
      break;
    }
  }
  if (startRow === -1) throw new Error(`Could not find row with Date = 1 for month '${monthName}'.`);

  let rowCount = 0;
  for (let i = startRow; i < dateValues.length; i++) {
    const val = dateValues[i];
    if (val === "" || val === null) break;
    rowCount++;
  }

  if (rowCount === 0) throw new Error(`No valid data rows found for '${monthName}'.`);

  const lunchRange = sheet.getRange(3 + startRow, lunchCol + 1, rowCount, 1);
  const output = [];
  const usedMeals = [];

  for (let i = 0; i < rowCount; i++) {
    const recent = usedMeals.slice(-7);
    const available = meals.filter(m => !recent.includes(m));
    const mealPool = available.length > 0 ? available : meals;
    const meal = mealPool[Math.floor(Math.random() * mealPool.length)];
    output.push([meal]);
    usedMeals.push(meal);
  }

  lunchRange.setValues(output);
}
