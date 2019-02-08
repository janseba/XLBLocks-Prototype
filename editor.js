(function () {
	"use strict";
	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				// TODO1: Assign handler to the OK button.
				$('#save-formulas').click(sendStringToParentPage);
				delete window.prompt;
			});
		});
	// TODO2: Create the OK button handler
	function sendStringToParentPage() {
		var testString = localStorage.getItem("formulaID");
		console.log(testString);
		Office.context.ui.messageParent(testString);
	}
}());