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
			var messageFromDialog = arg.message;
			console.log(messageFromDialog);
			// localStorage.setItem("BlocklyWorkspace", messageFromDialog)
			// document.getElementById("message").innerHTML = messageFromDialog;
			dialog.close();
		}
	})();

/***/ })
/******/ ]);