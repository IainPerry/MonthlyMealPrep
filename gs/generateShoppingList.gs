function generateShoppingList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = ss.getSheetByName("Plan");
  const shoppingSheet = ss.getSheetByName("Shoppinglist");
  const lunchIdeas = ss.getSheetByName("Lunch ideas");
  const dinnerIdeas = ss.getSheetByName("Dinner ideas");

  // Get selected month and week starting date
  const month = shoppingSheet.getRange("B2").getValue().toString().trim();
  const weekStart = parseInt(shoppingSheet.getRange("B3").getValue(), 10);

  if (!month || isNaN(weekStart)) {
    throw new Error(`Please select both Month and Week Commencing date. Got Month: "${month}", Week Start: "${weekStart}"`);
  }

  // Find start column of the table by matching month in row 1 and header row 2
  const planRow1 = planSheet.getRange("1:1").getValues()[0];
  const planRow2 = planSheet.getRange("2:2").getValues()[0];
  let dinnerCol = -1;

  for (let col = 0; col < planRow2.length; col++) {
    if (
      planRow2[col].toString().trim().toLowerCase() === "dinner" &&
      planRow1[col].toString().trim().toLowerCase() === month.toLowerCase()
    ) {
      dinnerCol = col;
      break;
    }
  }

  if (dinnerCol === -1) {
    throw new Error(`Could not find the Dinner column for month '${month}'`);
  }

  const dayCol = dinnerCol - 3;
  const dateCol = dinnerCol - 2;
  const lunchCol = dinnerCol - 1;

  // Read all relevant data (up to 50 rows max)
  const maxRows = 50;
  const data = planSheet.getRange(3, dayCol + 1, maxRows, 4).getValues();

  // Extract meals for the given week (7 days starting from selected date)
  const weekMeals = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const day = row[0];
    const date = row[1];
    const lunch = row[2];
    const dinner = row[3];

    if (parseInt(date, 10) === weekStart || weekMeals.length > 0) {
      weekMeals.push({ lunch, dinner });
      if (weekMeals.length === 7) break;
    }
  }

  if (weekMeals.length === 0) {
    throw new Error(`No meals found for week starting on date '${weekStart}'`);
  }

  // Load meal-to-ingredients mapping
  function loadMealMap(sheet) {
    const values = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < values.length; i++) {
      const mealName = values[i][0];
      if (!mealName) continue;
      const normalized = mealName.toString().trim().toLowerCase();
      const ingredients = values[i].slice(2).filter(cell => cell && cell.toString().trim() !== "");
      map[normalized] = ingredients.map(ing => ing.toString().trim());
    }
    return map;
  }

  const lunchMap = loadMealMap(lunchIdeas);
  const dinnerMap = loadMealMap(dinnerIdeas);

  // Accumulate ingredients with count
  const ingredientCount = {};

  for (const { lunch, dinner } of weekMeals) {
    [lunch, dinner].forEach(meal => {
      const name = meal ? meal.toString().trim().toLowerCase() : "";
      const ingredients = lunchMap[name] || dinnerMap[name] || [];
      for (const ing of ingredients) {
        const norm = ing.toLowerCase();
        ingredientCount[norm] = (ingredientCount[norm] || 0) + 1;
      }
    });
  }

  // Convert back to display format (capitalize, count)
  const result = Object.entries(ingredientCount)
    .sort((a, b) => a[0].localeCompare(b[0])) // Sort alphabetically
    .map(([key, count]) => [key.charAt(0).toUpperCase() + key.slice(1), count]);

  // Clear previous data from A5:B downwards
  const lastRow = shoppingSheet.getLastRow();
  if (lastRow > 4) {
    shoppingSheet.getRange("A5:B" + lastRow).clearContent();
  }

  // Set headers
  shoppingSheet.getRange("A4").setValue("Ingredient");
  shoppingSheet.getRange("B4").setValue("Count");

  // Write new ingredient list
  if (result.length > 0) {
    shoppingSheet.getRange(5, 1, result.length, 2).setValues(result);
  }
}
