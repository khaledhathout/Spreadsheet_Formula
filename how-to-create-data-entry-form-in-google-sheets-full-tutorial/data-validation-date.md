# Data Validation \[Date]

<figure><img src="../.gitbook/assets/image (12).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../.gitbook/assets/image (9).png" alt=""><figcaption><p>Is Valid Date</p></figcaption></figure>

```javascript
function IsValidDate() 
{

  // Variables ------ Start
    var spreadsheet = SpreadsheetApp.getActive();
  // Variables ------ End
    
  // Clear Old Format ------ Start
    spreadsheet.getRange('B3').activate();
    spreadsheet.getRange('B3').clearDataValidations();
  // Clear Old Format ------ End

  // Clear New Format ------ Start
    spreadsheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(true)
    .requireDate()
    .build());
    spreadsheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(true)
    .requireDate()
    .build());
  // Clear New Format ------ End

};
```
