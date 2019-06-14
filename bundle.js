/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};

/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {

/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;

/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			exports: {},
/******/ 			id: moduleId,
/******/ 			loaded: false
/******/ 		};

/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);

/******/ 		// Flag the module as loaded
/******/ 		module.loaded = true;

/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}


/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;

/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;

/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";

/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ function(module, exports) {

	'use strict';

	(function () {
		Office.onReady().then(function () {
			$(document).ready(function () {
				if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
					console.log('Sorry. The add-in uses Excel.js APIs that are not available in your version of Office');
				}
				OfficeExtension.config.extendedErrorLoggin = true;
				$('#validateFormula').click(updateFormula);
				$('#cancel').click(cancel);
				$('#changeFormula').click(editFormula);
				$('#newFormula').click(newFormula);
				$('#pasteRange').click(pasteRange);
				getXLBlockList(replaceFormulaDdl);
			});
		});

		function pasteRange() {
			Excel.run(function (context) {
				var range = context.workbook.getSelectedRange();
				range.load('address');

				return context.sync().then(function () {
					copyToClipboard(range.address.slice(range.address.indexOf('!') + 1));
				}).then(context.sync).then(function () {});
			}).catch(function (error) {
				console.log("Error: " + error);
				if (error instanceof OfficeExtension.Error) {
					console.log("Debug info: " + JSON.stringify(error.debugInfo));
				}
			});
		}

		function updateFormula() {
			Excel.run(function (context) {

				// Get all worksheets
				var sheets = context.workbook.worksheets;
				sheets.load('items/name');

				var rngDefinitions;

				return context.sync().then(function () {

					// Check if XLBlocks exists
					if (sheetExists(sheets.items, 'XLBlocks')) {
						// if 'XLBlocks' exists then getUsedRange
						var sht = sheets.getItem('XLBlocks');
						rngDefinitions = sht.getUsedRange();
						rngDefinitions.load('values');
						sht.delete();
					}
				}).then(context.sync).then(function () {
					if (typeof rngDefinitions === 'undefined') {
						var xlValues = [];
					} else {
						var xlValues = rngDefinitions.values;
						xlValues.shift();
					}
					var ws = getWorkspace();
					var ids = getCol(xlValues, 0);
					var index = ids.indexOf(ws.id);
					if (index === -1) {
						xlValues.push([ws.id, ws.name, ws.fullXML]);
					} else {
						var xml = Blockly.Xml.textToDom(xlValues[index][2]);
						var existingOutputRange = getFormulaOutput(xml);
						xlValues[index][1] = ws.name;
						xlValues[index][2] = ws.fullXML;
					}
					xlValues.unshift(['ID', 'Name', 'XML']);
					var sht = sheets.add('XLBlocks');
					var rng = sht.getRange('A1:C' + xlValues.length);
					rng.values = xlValues;
					replaceFormulaDdl(xlValues);
					sht.visibility = Excel.SheetVisibility.hidden;

					// update formulas in Excel
					var code = getFormula();
					var activeSheet = context.workbook.worksheets.getActiveWorksheet();
					if (code.outputRange !== existingOutputRange && existingOutputRange !== undefined) {
						var oldRange = activeSheet.getRange(existingOutputRange);
						oldRange.clear('Contents');
					}
					var formulaRange = activeSheet.getRange(code.outputRange);
					for (var i = 0; i < code.statements.length; i++) {
						for (var j = 0; j < code.statements[i].length; j++) {
							code.statements[i][j] = code.statements[i][j].replace(/\|/g, ',');
						}
					}
					formulaRange.formulas = code.statements;
				}).then(context.sync).then(function () {
					clearBlockly();
					hideBlockly();
					toggleButton('validateFormula', true);
					toggleButton('newFormula', false);
					toggleButton('cancel', false);
					toggleButton('changeFormula', false);
				});
			}).catch(function (error) {
				console.log("Error: " + error);
				if (error instanceof OfficeExtension.Error) {
					console.log("Debug info: " + JSON.stringify(error.debugInfo));
				}
			});
		}

		function editFormula(formulas) {
			toggleButton('newFormula', true);
			toggleButton('validateFormula', false);
			toggleButton('cancel', false);
			showBlockly();
			getXLBlockList(initWorkspace);
		}

		function getXLBlockList(callback) {
			Excel.run(function (context) {

				var sheets = context.workbook.worksheets;
				sheets.load('items/name');

				return context.sync().then(function () {
					if (sheetExists(sheets.items, 'XLBlocks')) {
						var xlBlockSht = sheets.getItem('XLBlocks');
						var definitionsRng = xlBlockSht.getUsedRange();
						definitionsRng.load('values');
						return definitionsRng;
					}
					return null;
				}).then(context.sync).then(function (definitionsRng) {
					if (definitionsRng !== null) {
						var definitionValues = definitionsRng.values;
					} else {
						var definitionValues = null;
					}
					callback(definitionValues);
				});
			}).catch(function (error) {
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
			for (var n = 0; n < hex.length; n += 2) {
				str += String.fromCharCode(parseInt(hex.substr(n, 2), 16));
			}
			return str;
		}

		function ascii_to_hex(str) {
			var arr1 = [];
			for (var n = 0; n < str.length; n++) {
				var hex = Number(str.charCodeAt(n)).toString(16);
				arr1.push(hex);
			}
			return arr1.join('');
		}
	})();

/***/ }
/******/ ]);