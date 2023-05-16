# Save button

<figure><img src="../../.gitbook/assets/image (15).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../.gitbook/assets/image (17).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../.gitbook/assets/image (7).png" alt=""><figcaption></figcaption></figure>

add script

```javascript
function submitData() 
{
  /// Variables ------ Start
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSS = ss.getSheetByName("User Form"); //User Form
    var datasheet = ss.getSheetByName("Database"); //Database
  /// Variables ------ End

  //Input Values  ------ Start
    var values = [[
    formSS.getRange("B2").getValue(), // B2 ----old----- C7-----Customer ID
    formSS.getRange("B3").getValue(), // Customer ID = B3 ----old----- C9
    formSS.getRange("B4").getValue(), // Customer ID = B4 ----old----- C11
    formSS.getRange("B5").getValue(), // Customer ID = B5 ----old----- C13
    formSS.getRange("B6").getValue(), // Customer ID = B6 ----old----- C15
    formSS.getRange("B7").getValue()  // Customer ID = B7 ----old----- C17
    ]];
  //Input Values  ------ End

  //Clear the fields after submit  ------ Start
    datasheet.getRange(datasheet.getLastRow()+1, 1, 1, 6).setValues(values);
    formSS.getRange("B2").clear(); // Customer ID = B2 ----old----- C7
    formSS.getRange("B3").clear(); // Customer ID = B2 ----old----- C7
    formSS.getRange("B4").clear(); // Customer ID = B2 ----old----- C7
    formSS.getRange("B5").clear(); // Customer ID = B2 ----old----- C7
    formSS.getRange("B6").clear(); // Customer ID = B2 ----old----- C7
    formSS.getRange("B7").clear(); // Customer ID = B2 ----old----- C7
  //Clear the fields after submit  ------ End
  }
```
