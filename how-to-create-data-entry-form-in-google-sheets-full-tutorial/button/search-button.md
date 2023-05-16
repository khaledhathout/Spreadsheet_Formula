# Search button

```javascript
function searchStr() {
  
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var formSS    = ss.getSheetByName("User Form"); //Form Sheet
  
  var str       = formSS.getRange("C4").getValue();
  var values    = ss.getSheetByName("Database").getDataRange().getValues();
  for (var i = 0; i < values.length; i++) 
  {
    var row = values[i];
    if (row[SEARCH_COL_IDX] == str) 
    {
 
    
      
      formSS.getRange("C7").setValue(row[0]);
      formSS.getRange("C9").setValue(row[1]);
      formSS.getRange("C11").setValue(row[2]);
      formSS.getRange("C13").setValue(row[3]);
      formSS.getRange("C15").setValue(row[4]);
      formSS.getRange("C17").setValue(row[5]);
      
      
           
      return row[RETURN_COL_IDX];
      
    }
  }
}
```
