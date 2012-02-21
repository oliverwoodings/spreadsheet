# Spreadsheet

  A CommonJS module for reading Google Spreadsheets.


## Install

  It's available on npm, so a simple `npm install spreadsheet` should be enough.


## Usage


	var Spreadsheet = require("spreadsheet");
	
	// Instantiate a spreadsheet using the key directly.
	var sheet = new Spreadsheet("mykey");
	
	// Or just let the module extract it from an URL.
	var sheet = Spreadsheet.fromURL("http://shared...")
	
	// Load the worksheets, callback will be called for each worksheet
	sheet.worksheets(function(err,ws){
		// Each worksheet allows you to go through each row.
		ws.eachRow(function(err,row,meta){
			// `row` is an object with all the fields of that row.
			// `meta` is an object like {index: 1, total: 2, id: "https://...", update: Date()}
		})
		
		// Or each cell.
		ws.eachCell(function(err,cell,meta){
			// `cell` is an object like {row: 1, col: 1, value: "Hello!"}
			// `meta` is an object like {index: 1, total: 2, id: "https://...", update: Date()}
		})
	})
	
	// You can also work with just one worksheet by page number
	sheet.worksheet(1,function(err,ws){
		// Do stuff...
	})
	// Or worksheet id.
	sheet.worksheet("od6",function(err,ws){
		// Do stuff...
	})
	
	
	// Load the worksheets to an array
	sheet.worksheetArray(function(err,spreadsheet, worksheets){
		// Do stuff... like
		//res.render('worksheets', {
		//	spreadsheet : spreadsheet,
		//	worksheets : worksheets
		//});
	});




	if(wrk_id){
		//import one sheet
		sheet.worksheet(wrk_id,function(err,ws){
			if(err) throw(err);
			importSheet_(ws, true);
		});
	}else{
		//import all sheets
		sheet.worksheets(function(err,ws){
			if(err)	throw(err);
			importSheet_(ws);
		});
	}
	
	function importSheet_(ws, onlyOne)
	{
		// Each worksheet allows you to go through each row.
		ws.eachRow(function(err,row,meta){
			// `row` is an object with all the fields of that row.
			// `meta` is an object like {index: 1, total: 2, id: "https://...", update: Date()}
			if(meta.index === meta.total){
				//the last row
			}
			
			if((onlyOne && meta.index === meta.total) || (ws.spreadsheet.sheetCount === ws.index && meta.index === meta.total)){
				console.log('done');
			}
		});
	}
	
## History

### added by baryon

* [Feature] Add worksheetArray method.
* [Feature] Add author, title, updated timestamp, sheetCount property for spreadsheet
* [Feature] Add title and updated property for worksheet
* [Bug] Fixes for empty cell, now return null when the cell is an empty object.
* [Feature] Add some test cases for added method and properties.

### 0.3.0

* [Feature] Changed to `open-uri` instead of `request`.
* [Feature] Updated xml2js to 0.2.0.

### 0.2.1

* [Bug] Fixes for NPM 0.3+

### 0.2.0

* [Feature] Access row and cells directly by their IDs or URL through `Worksheet#cell(id,fn)`, `Worksheet#row(id,fn)` or using `Spreadsheet.fromURL(meta.id)` (meta.id is the meta you get in an `Worksheet#eachRow`- or `Worksheet.eachCell`-callback)

### 0.1.2

* [Feature] Error message when no rows were found.
* [Bug] Fixed an issue with the NPM package. It couldn't find the library when installed through NPM.

### 0.1.1

* [Feature] Documentation

### 0.1.0

* Initial Google Spreadsheet implementation.

## License 

(The MIT License)

Copyright (c) 2011 Robert Sk&ouml;ld &lt;robert@publicclass.se&gt;

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
'Software'), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.