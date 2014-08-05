var excelbuilder = require('msexcel-builder');

var workbook = excelbuilder.createWorkbook('./', 'planner.xlsx');

var now = new Date("August 24, 2014");
var inc = new Date("August 24, 2014");

var end = 1000 * 60 * 60 * 24 * 300;

while ((inc - now) < end) {

  var carts = ['Colleges (28)', 'Espanol (24)', 'Elements (25)'];

  inc.setDate(inc.getDate() + 1);
  var dateString = inc.toDateString().replace(/\s+/g, '.');

  console.log(dateString);

  var len = 31,
    mins = 15;

  if (inc.getDay() > 0 && inc.getDay() < 6) {
    var sheet = workbook.createSheet(dateString, 9, len);
    for (var j = 0; j < carts.length; j++) {

      var col = (j * 3) + 1;
      sheet.set(col, 1, carts[j]);
      sheet.set(col + 1, 2, 'Teacher');
      sheet.set(col + 2, 2, 'Reason');

      var dataRow = 3;

      inc.setHours(8);
      inc.setMinutes(0);
      inc.setSeconds(0);

      for (var i = dataRow; i <= len; i++) {
        sheet.set(col, i, inc.toLocaleTimeString("en-US"));
        inc.setMinutes(inc.getMinutes() + 15);
      }
    }
  }
}

workbook.save(function(ok) {
  if (!ok) {
    workbook.cancel();
  } else {
    console.log('congratulations, your workbook created');
  }
});
