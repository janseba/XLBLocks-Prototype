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
					console.log(messageFromDialog.Type)
					break;
				case 'blockDefinition':
					localStorage.setItem("BlocklyWorkspace", messageFromDialog.MessageContent);
					break;
			}
			dialog.close();
		}
})();