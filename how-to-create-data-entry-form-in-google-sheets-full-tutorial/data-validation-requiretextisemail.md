# Data Validation - RequireTextIsEmail

<figure><img src="../.gitbook/assets/image (12) (2).png" alt=""><figcaption></figcaption></figure>

## Way 001

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

## Way 002

<figure><img src="../.gitbook/assets/image (7) (1).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../.gitbook/assets/image (15) (1).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../.gitbook/assets/image (13).png" alt=""><figcaption></figcaption></figure>
