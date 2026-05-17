# Meal Management System – Google Sheets & Apps Script

# Google Apps Script Code

## Code.gs

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Meal Management")
    .addItem("Calculate Meal Rate", "calculateMealRate")
    .addItem("Generate Summary", "generateSummary")
    .addItem("Reset Daily Meals", "resetMeals")
    .addToUi();
}

function calculateMealRate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meal");

  const totalMeal = sheet.getRange("B2").getValue();
  const totalCost = sheet.getRange("B3").getValue();

  if (totalMeal == 0) {
    SpreadsheetApp.getUi().alert("Total meal cannot be zero");
    return;
  }

  const mealRate = totalCost / totalMeal;

  sheet.getRange("B4").setValue(mealRate);

  SpreadsheetApp.getUi().alert("Meal Rate Calculated Successfully");
}

function generateSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  const mealSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meal");

  const totalMeal = mealSheet.getRange("B2").getValue();
  const totalCost = mealSheet.getRange("B3").getValue();
  const mealRate = mealSheet.getRange("B4").getValue();

  sheet.getRange("B2").setValue(totalMeal);
  sheet.getRange("B3").setValue(totalCost);
  sheet.getRange("B4").setValue(mealRate);

  SpreadsheetApp.getUi().alert("Summary Generated");
}

function resetMeals() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DailyMeal");

  sheet.getRange("B2:AF20").clearContent();

  SpreadsheetApp.getUi().alert("Daily Meals Reset Successfully");
}

function calculateMemberCost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName("Members");
  const mealSheet = ss.getSheetByName("Meal");

  const mealRate = mealSheet.getRange("B4").getValue();

  const lastRow = memberSheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    const totalMeal = memberSheet.getRange(i, 3).getValue();
    const deposit = memberSheet.getRange(i, 4).getValue();

    const totalCost = totalMeal * mealRate;
    const balance = deposit - totalCost;

    memberSheet.getRange(i, 5).setValue(totalCost);
    memberSheet.getRange(i, 6).setValue(balance);
  }

  SpreadsheetApp.getUi().alert("Member Costs Calculated");
}

function addDeposit(name, amount) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Deposits");

  const lastRow = sheet.getLastRow() + 1;

  sheet.getRange(lastRow, 1).setValue(new Date());
  sheet.getRange(lastRow, 2).setValue(name);
  sheet.getRange(lastRow, 3).setValue(amount);
}

function addBazaarCost(item, amount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bazaar");

  const lastRow = sheet.getLastRow() + 1;

  sheet.getRange(lastRow, 1).setValue(new Date());
  sheet.getRange(lastRow, 2).setValue(item);
  sheet.getRange(lastRow, 3).setValue(amount);
}

function monthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName("Report");
  const memberSheet = ss.getSheetByName("Members");

  reportSheet.clear();

  reportSheet.getRange("A1").setValue("Monthly Meal Management Report");

  const data = memberSheet.getDataRange().getValues();

  reportSheet.getRange(3, 1, data.length, data[0].length).setValues(data);

  SpreadsheetApp.getUi().alert("Monthly Report Generated");
}
```
