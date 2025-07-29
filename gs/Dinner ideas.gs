function fillDinnerByMonthHeader(monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Plan");
  const ideasSheet = ss.getSheetByName("Dinner ideas");

  // Load meal ideas from Dinner ideas sheet, column A, ignoring blanks
  const meals = ideasSheet.getRange("A2:A").getValues().flat().filter(x => x !== "");

  // Get row 1 (month headers) and row 2 (table headers)
  const row1 = sheet.getRange("1:1").getValues()[0];
  const row2 = sheet.getRange("2:2").getValues()[0];

  let dinnerCol = -1;

  // Find the 'Dinner' column where the month above matches monthName
  for (let col = 0; col < row2.length; col++) {
    const header = row2[col];
    const monthAbove = row1[col];

    if (
      typeof header === "string" &&
      header.trim().toLowerCase() === "dinner" &&
      typeof monthAbove === "string" &&
      monthAbove.trim().toLowerCase() === monthName.toLowerCase()
    ) {
      dinnerCol = col;
      break;
    }
  }

  if (dinnerCol === -1) {
    throw new Error(`Could not find 'Dinner' column under month '${monthName}'.`);
  }

  // The table columns are Day, Date, Lunch, Dinner, so Dinner is the 4th
  const startCol = dinnerCol - 3;
  if (startCol < 0) throw new Error("Not enough columns before 'Dinner' to find table.");

  // Validate expected headers in row 2
  const expectedHeaders = ["Day", "Date", "Lunch", "Dinner"];
  for (let i = 0; i < 4; i++) {
    const header = row2[startCol + i];
    if (
      typeof header !== "string" ||
      header.trim().toLowerCase() !== expectedHeaders[i].toLowerCase()
    ) {
      throw new Error(
        `Expected header "${expectedHeaders[i]}" in column ${startCol + i + 1}, but found "${header}".`
      );
    }
  }

  const dayCol = startCol;
  const dateCol = startCol + 1;

  // Determine number of data rows (max 31), stopping when both Day and Date are empty
  const maxPossibleRows = 50;
  const dayValues = sheet.getRange(3, dayCol + 1, maxPossibleRows, 1).getValues();
  const dateValues = sheet.getRange(3, dateCol + 1, maxPossibleRows, 1).getValues();

  let dataRowCount = 0;
  for (let i = 0; i < maxPossibleRows; i++) {
    const dayVal = dayValues[i][0];
    const dateVal = dateValues[i][0];
    if (
      (dayVal === "" || dayVal === null) &&
      (dateVal === "" || dateVal === null)
    ) {
      break;
    }
    dataRowCount++;
  }

  if (dataRowCount === 0) {
    throw new Error(`No data rows found for month '${monthName}'.`);
  }

  // Get the full table data (Day, Date, Lunch, Dinner)
  const dataRange = sheet.getRange(3, startCol + 1, dataRowCount, 4);
  const data = dataRange.getValues();

  // Fill dinner with random meals, skip Wednesdays with "Date Night", avoid repeat within 7 days
  const usedMeals = [];
  const output = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const day = row[0];
    const date = row[1];

    if (!day || !date) {
      // Leave row unchanged if day or date is missing
      output.push(row);
      continue;
    }

    if (typeof day === "string" && day.trim().toLowerCase() === "wednesday") {
      row[3] = "Date Night";
      usedMeals.push("Date Night");
    } else {
      const recentMeals = usedMeals.slice(-7);
      const available = meals.filter(m => !recentMeals.includes(m));
      const pool = available.length > 0 ? available : meals;
      const meal = pool[Math.floor(Math.random() * pool.length)];
      row[3] = meal;
      usedMeals.push(meal);
    }

    output.push(row);
  }

  // Write updated rows back to the sheet
  dataRange.setValues(output);
}
