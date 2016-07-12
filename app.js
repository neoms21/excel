var Excel = require('exceljs');
var _ = require('lodash');

var workbook = new Excel.Workbook();
var data = [];
workbook.xlsx.readFile("data.xlsx")
	.then(function() {

			var sheet = workbook.getWorksheet(3);
			var rowCount = sheet._rows.length;
			var colCount = sheet.columns.count;

			var headers = sheet.getRow(1).values;

			for (var i = 2; i <= rowCount; i++) {

				var values = sheet.getRow(i).values;
				var idx =0;
				var obj = {};
				_.each(values, function(v) {
				
					obj[headers[idx]]= v;
					idx++;	
				});

				data.push(obj);
			}

			console.log(data);
		});

	