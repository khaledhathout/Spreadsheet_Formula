# Data Validation - requireDate

<figure><img src="../.gitbook/assets/image (14).png" alt=""><figcaption></figcaption></figure>

```javascript
function RequireDate() {

  // Variables ------ Start
    var spreadsheet = SpreadsheetApp.getActive();
  // Variables ------ End
    
  // Clear Old Format ------ Start
    spreadsheet.getRange('B4').activate();
    spreadsheet.getRange('B4').clearDataValidations();
  // Clear Old Format ------ End

  // Clear New Format ------ Start
    spreadsheet.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireDate()
    .build());
  // Clear New Format ------ End

}
```

<figure><img src="../.gitbook/assets/image (12).png" alt=""><figcaption></figcaption></figure>
