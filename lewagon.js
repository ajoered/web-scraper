const jquery    = require('jquery');
const Nightmare = require('nightmare');
      nightmare = Nightmare();
const Excel     = require('exceljs');
const vo        = require('vo');

var workbook = new Excel.Workbook();
workbook.creator = 'Alvaro';
workbook.created = new Date(2017, 10, 4);
workbook.modified = new Date();
workbook.properties.date1904 = true;
var sheet = workbook.addWorksheet("Le Wagon");
sheet.columns = [
    { header: 'Country', key:'country' },
    { header: 'City', key:'city' },
    { header: 'Language', key:'language' },
    { header: 'Batch #', key:'batchnumber' },
    { header: 'Number of Students', key:'students' },
    { header: 'Quarter', key:'quarter' },
    { header: 'Year', key:'year' },
    { header: 'Price', key:'price' },
    { header: 'Currency', key:'currency' },
    { header: 'Turnover in Local', key:'turnover_local' },
    { header: 'Turnover in Eur/USD', key:'turnover_usd' },
];

var run = function * () {
  var batches = [];
  for (var i = 0; i < 92; i++) {
    var batch = yield nightmare.goto(`https://www.lewagon.com/demoday/`+i)
      .wait('body')
      .evaluate(function() {
        var city = $("span.batch-name>i").text();
        var language = $("button#languageSelector>span:nth-child(1)>i").text();
        var batchnumber = $("h3.project-name>small").text();
        var students = $(".demoday-title").text();
        var date = $("span.batch-date").text()
        var year = $("span.batch-name>i").text();

        var item = {};

        item["city"] = city
        item["language"] = language
        item["batchnumber"] = batchnumber
        item["students"] = students.substring(0, 3)
        item["quarter"] = date.slice(0, -4);
        item["year"] = date.substring(date.length - 4);

        return item
      });
    batches.push(batch);
  }
  yield nightmare.end();
  return batches;
}

vo(run)(function(err, batches) {
  console.dir(batches);
  for (var i = 0; i < batches.length; i++) {
    sheet.addRow({
      city: batches[i].city,
      language: batches[i].language,
      batchnumber: batches[i].batchnumber,
      students: batches[i].students,
      quarter: batches[i].quarter,
      year: batches[i].year,
    })
  };
  workbook.xlsx.writeFile("./lewagon.xlsx").then(function() {
      console.log("xls file is written.");
  });
});
