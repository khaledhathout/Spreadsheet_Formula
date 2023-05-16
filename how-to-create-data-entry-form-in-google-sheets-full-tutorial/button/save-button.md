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
    formSS.getRange("B2").getValue(), // B2 ----old----- C7 -----Customer ID
    formSS.getRange("B3").getValue(), // B3 ----old----- C9 -----Date
    formSS.getRange("B4").getValue(), // B4 ----old----- C11 ----Name
    formSS.getRange("B5").getValue(), // B5 ----old----- C13 ----Phone
    formSS.getRange("B6").getValue(), // B6 ----old----- C15 ----Email
    formSS.getRange("B7").getValue()  // B7 ----old----- C17 ----Address
    ]];
  //Input Values  ------ End

  //Clear the fields after submit  ------ Start
    datasheet.getRange(datasheet.getLastRow()+1, 1, 1, 6).setValues(values);
    formSS.getRange("B2").clear(); // B2 ----old----- C7 -----Customer ID
    formSS.getRange("B3").clear(); // B3 ----old----- C9 -----Date
    formSS.getRange("B4").clear(); // B4 ----old----- C11 ----Name
    formSS.getRange("B5").clear(); // B5 ----old----- C13 ----Phone
    formSS.getRange("B6").clear(); // B6 ----old----- C15 ----Email
    formSS.getRange("B7").clear();  // B7 ----old----- C17 ----Address
  //Clear the fields after submit  ------ End
  }
```

<figure><img src="../../.gitbook/assets/image (10).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../.gitbook/assets/image (12).png" alt=""><figcaption><p>submitData</p></figcaption></figure>

<figure><img src="../../.gitbook/assets/image (3).png" alt=""><figcaption></figcaption></figure>
