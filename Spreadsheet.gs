function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Excel Tools')
      .addItem('Convert URL to Hyperlinks', 'ConvertUrlToHyperlinks')
      .addToUi();
}

function validURL(str) {
  var pattern = new RegExp('^(https?:\\/\\/)?'+ // protocol
    '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|'+ // domain name
    '((\\d{1,3}\\.){3}\\d{1,3}))'+ // OR ip (v4) address
    '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*'+ // port and path
    '(\\?[;&a-z\\d%_.~+=-]*)?'+ // query string
    '(\\#[-a-z\\d_]*)?$','i'); // fragment locator
  return !!pattern.test(str);
}

function ConvertUrlToHyperlinks() {
  
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  var ranges =  selection.getActiveRangeList().getRanges();
  for (var i = 0; i < ranges.length; i++) {
    var columns=ranges[i].getNumColumns();
    var rows=ranges[i].getNumRows();
    for (var column=1; column <= columns; column++) {
      for (var row=1; row <= rows; row++) { 
        var cell=ranges[i].getCell(row,column);
          var cellValue = cell.getValue();
          if(validURL(cellValue)){
            var richValue = SpreadsheetApp.newRichTextValue().setText("Click Here").setLinkUrl(cellValue).build();
            cell.setRichTextValue(richValue);
          }
          //Logger.log(cellValue);
      }
    }
  }
}
