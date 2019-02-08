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
		var xml = Blockly.Xml.workspaceToDom(workspace);
		var xml_text = Blockly.Xml.domToText(xml);
		console.log(testString);
		Office.context.ui.messageParent(xml_text);
	}
}());