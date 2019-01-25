'use strict';
(function () {

	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				if(!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
					console.log('Sorry. The add-in uses Excel.js APIs that are not available in your version of Office');
				}
				$('#start-editor').click(startEditor)
				console.log('onReady');
			});
		});

		var dialog = null;
		function startEditor() {
			localStorage.setItem("formulaID","Een voorbeeldtekst");
			console.log('start editor log');
			Office.context.ui.displayDialogAsync(
				'https://localhost:3000/editor.html',
				{height: 45, width: 55},
				function (result) {
					dialog = result.value;
				})
		}
})();