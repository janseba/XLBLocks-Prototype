'use strict';
(function () {  
	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				if(!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
					console.log('Sorry. The add-in uses Excel.js APIs that are not available in your version of Office');
				}
				OfficeExtension.config.extendedErrorLoggin = true;
				$('#update-formula').click(updateFormula);
				$('#refresh-page').click(refreshPage);
				$('#edit').click(editFormula);
				//$('#refresh-ddl').click(testDDL);
				getXLBlockList(replaceFormulaDdl);		
			});
		});

	function updateFormula() {
		saveFormulaDefinition();
	}
	
	function editFormula(formulas) {
		getXLBlockList(initWorkspace);
	}

	function getXLBlockList(callback) {
		Excel.run(function (context) {
			          
			var sheets = context.workbook.worksheets;
			sheets.load('items/name');

			return context.sync()
			.then(function () {
				if (sheetExists(sheets.items, 'XLBlocks')) {
					var xlBlockSht = sheets.getItem('XLBlocks');
					var definitionsRng = xlBlockSht.getRange('A1:C5');
					definitionsRng.load('values');
					return definitionsRng
				}
				return null
			})
			.then(context.sync)
			.then(function (definitionsRng) {
				if (definitionsRng !== null) {
					var definitionValues = definitionsRng.values;
				} else {
					var definitionValues = null;
				}
				callback(definitionValues);
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

	function refreshPage() {
		location.reload();
	}

	function updateFormulas(formulaDefString) {
		Excel.run(function (context) {

			var formula = JSON.parse(formulaDefString);
			var sheet = context.workbook.worksheets.getActiveWorksheet();
			var range = sheet.getRange(formula.outputRange);
			range.formulas = formula.statements;

			return context.sync();

		})
		.catch(function (error) {
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		})
	}

	function getFormulaID(blockDefinition) {
		var parser = new DOMParser();
		var xmlDoc = parser.parseFromString(blockDefinition, 'text/xml');
		var block = xmlDoc.getElementsByTagName('block');
		for (var i = 0; i < block.length; i++) {
			if (block[i].getAttribute('type') == 'formula') {
				var id = block[i].getAttribute('id');
				break;
			}
		}
		return id;
	}

	function getFormulaName(blockDefinition) {
		var parser = new DOMParser();
		var xmlDoc = parser.parseFromString(blockDefinition, 'text/xml');
		var field = xmlDoc.getElementsByTagName('field');
		for (var i = 0; i < field.length; i++) {
			if (field[i].getAttribute('name') == 'formula_name') {
				var name = field[i].innerText;
				break;
			}
		}
		return name;
	}
	
	function hex_to_ascii(str1) {
		var hex = str1.toString();
		var str = '';
		for (var n = 0; n < hex.length; n +=2) {
			str += String.fromCharCode(parseInt(hex.substr(n, 2), 16));
		}
		return str;
	}

	function ascii_to_hex(str) {
		var arr1 = [];
		for (var n = 0; n < str.length; n ++) {
			var hex = Number(str.charCodeAt(n)).toString(16);
			arr1.push(hex);
		}
		return arr1.join('');
	}
})();
