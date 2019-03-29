'use strict';
(function () {

	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				if(!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
					console.log('Sorry. The add-in uses Excel.js APIs that are not available in your version of Office');
				}
				$('#update-formula').click(updateFormula);
				$('#refresh-page').click(refreshPage);
				$('#edit').click(editFormula);
				$('#refresh-ddl').click(refreshDdl);
				//$('#ddlFormulas').change(testOnChange);
				//updateFormulaList();			
			});
		});

		function updateFormula() {
			console.log('update formula');
			saveFormulaDefinition();
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

		function refreshDdl() {
			
			var select = document.getElementById('ddlFormulas');
			for (var i = 0; i < 4; i++) {
				var option = document.createElement('option');
				option.text = 'omschrijving ' + i;
				select.add(option);
			}
			var ddlDiv = document.getElementById('bjaTest');
			var child = ddlDiv.querySelector('.ms-Dropdown-title');
			if (child != null) {
				ddlDiv.removeChild(child);
			}
			child = ddlDiv.querySelector('.ms-Dropdown-items');
			if (child != null) {
				ddlDiv.removeChild(child);
			}
			child = ddlDiv.querySelector('.ms-Dropdown-truncator');
			if (child != null) {
				ddlDiv.removeChild(child);
			}
			var DropdownHTMLElement = document.getElementById('bjaTest');
			var Dropdown = new fabric['Dropdown'](DropdownHTMLElement);
		}

		function editFormula() {
			var ddlFormulas = document.getElementById('ddlFormulas');
			var formulaID = ddlFormulas.options(ddlFormulas.selectedIndex).value;
			var formulaText = ddlFormulas.options(ddlFormulas.selectedIndex).text;
			console.log(formulaID + ' ' + formulaText);
		}

		function refreshPage() {
			location.reload();
		}

		var dialog;
		function startEditor() {
			
			Office.context.ui.displayDialogAsync(
				'https://localhost:3000/editor.html',
				{height: 90, width: 90},
				function (asyncResult) {
					dialog = asyncResult.value;
					dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage)
				})
		}
		function processMessage(arg) {
			var messageFromDialog = JSON.parse(arg.message);
			switch(messageFromDialog.Type) {
				case 'formula':
					updateFormulas(messageFromDialog.MessageContent);
					var formula = JSON.parse(messageFromDialog.MessageContent);
					var formulaID = getFormulaID(formula.blockDefinition);
					formulaID = ascii_to_hex(formulaID);
					addName('_Block' + formulaID, formula.blockDefinition);
					console.log(formulaID);
					break;
				case 'blockDefinition':
					localStorage.setItem("BlocklyWorkspace", messageFromDialog.MessageContent);
					break;
			}
			dialog.close();
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
function addName(id,value) {
    Excel.run(function (context) {            
      var workbook = context.workbook;
      const existingName = workbook.names.getItemOrNullObject(id);
      existingName.load('name, formula');


      return context.sync()
          .then(
              function() {
              	if (existingName.isNullObject) {
              		workbook.names.add(id, value);
              	} else {
              		existingName.formula = value;
              	}
              }
          )
          .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
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

function updateFormulaList() {
	Excel.run(function (context) {            
		
		// code before sync
		var names = context.workbook.names;

		names.load('items/name, items/value');

		return context.sync()
		.then(function () {
		
			// code after sync
			var list = document.createElement('ul');
			var select = document.getElementById('ddlFormulas');
			for (var i in names.items) {
				var option = document.createElement('option');
				var formulaName = getFormulaName(names.items[i].value);
				var formulaID = getFormulaID(names.items[i].value);
				option.text = formulaName;
				option.value = formulaID;
				select.add(option);
			}

/*			var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
			for (var i = 0; i < DropdownHTMLElements.length; ++i) {
				var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
			}*/
		})
	})
	.catch(function (error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
	});
}
})();
