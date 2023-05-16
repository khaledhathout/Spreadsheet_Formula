# Data Validation - RequireTextIsEmail

```javascript
function RequireTextIsEmail () {
  var spreadsheet = SpreadsheetApp.getActive();

  // Clear Old Format ------ Start
    spreadsheet.getRange('B6').activate();
    spreadsheet.getRange('B6').clearDataValidations();
  // Clear Old Format ------ End  
  
  // Clear New Format ------ Start
    spreadsheet.getRange('B6').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireTextIsEmail()
    .build());
  // Clear New Format ------ End

}
```
