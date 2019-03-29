'use strict';
function saveFormulaDefinition() {
	Excel.run(function (context) {            
		
		// Get all worksheets
		var sheets = context.workbook.worksheets;
		sheets.load('items/name');

		// Get workspace definition
		var xml = Blockly.Xml.workspaceToDom(workspace);
		var ws = {id:getFormulaID(xml), name:getFormulaName(xml), fullXML:Blockly.Xml.domToText(xml)}
		
		var rngDefinitions

		return context.sync()
		.then(function () {
		
			// Check if XLBlocks exists
			if (sheetExists(sheets.items, 'XLBlocks')) {
				// if 'XLBlocks' exists then getUsedRange
				var sht = sheets.getItem('XLBlocks');
				rngDefinitions = sht.getUsedRange();
				rngDefinitions.load('values');
				sht.delete();
			}
		})
		.then(context.sync)
		.then(function () {
			if (typeof rngDefinitions === 'undefined') {
				var xlValues = [];
			} else {
				var xlValues = rngDefinitions.values
				xlValues.shift();
			}
			var ids = getCol(xlValues,0)
			var index = ids.findIndex(function(id){return id === this},ws.id);
			if ( index === -1) {
				xlValues.push([ws.id, ws.name, ws.fullXML]);
			} else {
				xlValues[index][1] = ws.name;
				xlValues[index][2] = ws.fullXML;
			}
			xlValues.unshift(['ID', 'Name', 'XML'])
			var sht = sheets.add('XLBlocks');
			var rng = sht.getRange('A1:C' + xlValues.length)
			rng.values = xlValues;
			var ids = getCol(xlValues, 0);
			console.log(ws.name);
		})
	})
	.catch(function (error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
	});
}

function sheetExists(sheets, name) {
	var result = false;
	for (var i = 0; i < sheets.length; i++) {
		if (sheets[i].name == name) {
			result = true;
			break;
		}
	}
	return result;
}

function getFormulaID(xml) {
	var block = xml.getElementsByTagName('block');
	for (var i = 0; i < block.length; i++) {
		if (block[i].getAttribute('type') == 'formula') {
			var id = block[i].getAttribute('id');
			break;
		}
	}
	return id;
}
function getFormulaName(xml) {
	var field = xml.getElementsByTagName('field');
	for (var i = 0; i < field.length; i++) {
		if (field[i].getAttribute('name') == 'formula_name') {
			var name = field[i].innerText;
			break;
		}
	}
	return name;
}

function existsInArray(arr, col, key) {
	var result = false;
	for (var i = 0; i < arr.length; i++) {
		if(arr[i][col] === key) {
			result = true;
			break;
		}
	}
	return result;
}

function getCol(matrix, col) {
	var column = [];
	for (var i = 0; i < matrix.length; i++) {
		column.push(matrix[i][col]);
	}
	return column
}
