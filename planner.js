var excelbuilder = require('msexcel-builder-colorfix');
var moment = require('moment');

var workbook = excelbuilder.createWorkbook('./', 'planner.xlsx');

var now = moment("2015-08-24");//new Date("August 24, 2014");
var inc = moment("2015-08-24");

var end = 1000 * 60 * 60 * 24 * 300;


while (inc.diff(now) < end) {

  var carts = ['Cart B (27)', 'Cart D 1&2 (33)', 'Cart C (25)', 'Ipod touch', 'Ipads (Cart 1)', 'Cart E (27)', 'Ipads (Cart 2)', 'Cart F (26)'];
  var colors = ['e81e42', '6a308e', 'ea0fcd', 'eae23b', '4148fb', 'f0aa01', '128b4e', 'a59578'];
  var timePeriods = 31; // number of 15 minute time periods in one day
  var timeSlot = 15; // 15 minute time slots

  if (inc.day() > 0 && inc.day() < 6) {
    var dateString = inc.format("ddd.MMM.D.YY");
    var sheet = workbook.createSheet(dateString, carts.length * 3, timePeriods);

    console.log("saving for date " + dateString);

    for (var j = 0; j < carts.length; j++) {
      var dataRow = 3;
      var col = (j * 3) + 1;
      console.log(carts[j]);
      sheet.set(col, 1, carts[j]); // set cart name
      sheet.set(col + 1, 2, 'Teacher'); // column header
      sheet.set(col + 2, 2, 'Reason');

      // fill with a color
      sheet.fill(col, 1, {'type': 'solid', 'fgColor': colors[j]});
      sheet.fill(col, 2, {'type': 'solid', 'fgColor': colors[j]});

      inc.hours(8);
      inc.minutes(0);
      inc.seconds(0);

      for(var i = dataRow; i <= timePeriods; i++) {
        sheet.set(col, i, inc.format("hh:mm a"));
        sheet.fill(col, i, {'type': 'solid', 'fgColor': colors[j]});
        inc.minute(inc.minute() + timeSlot);
      } 
    }
  }
  inc.add(1, 'd');
}

workbook.save(function(err) {
  if (err) {
    console.log(err);
  } else {
    console.log('congratulations, your workbook created');
  }
});
