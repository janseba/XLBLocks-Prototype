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
/***/ (function(module, exports) {

	'use strict';

	(function () {

		Office.onReady().then(function () {
			$(document).ready(function () {
				if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
					console.log('Sorry. The add-in uses Excel.js APIs that are not available in your version of Office');
				}
				$('#start-editor').click(startEditor);
				console.log('onReady');
			});
		});

		var dialog;
		function startEditor() {

			Office.context.ui.displayDialogAsync('https://localhost:3000/editor.html', { height: 90, width: 90 }, function (asyncResult) {
				dialog = asyncResult.value;
				dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
			});
		}
		function processMessage(arg) {
			var messageFromDialog = JSON.parse(arg.message);
			switch (messageFromDialog.Type) {
				case 'formula':
					updateFormulas(messageFromDialog.MessageContent);
					var formula = JSON.parse(messageFromDialog.MessageContent);
					var formulaID = getFormulaID(formula.blockDefinition);
					formulaID = ascii_to_hex(formulaID);
					addName(formulaID, formula.blockDefinition);
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
			}).catch(function (error) {
				console.log("Error: " + error);
				if (error instanceof OfficeExtension.Error) {
					console.log("Debug info: " + JSON.stringify(error.debugInfo));
				}
			});
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
		function addName(id, value) {
			Excel.run(function (context) {
				var workbook = context.workbook;

				workbook.names.add(id, value);

				return context.sync().then(function () {
					console.log('test');
				}).then(context.sync);
			}).catch(function (error) {
				console.log("Error: " + error);
				if (error instanceof OfficeExtension.Error) {
					console.log("Debug info: " + JSON.stringify(error.debugInfo));
				}
			});
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

/***/ })
/******/ ]);