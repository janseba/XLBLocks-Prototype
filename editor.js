(function () {
	"use strict";
	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				// TODO1: Assign handler to the OK button.
				$('#save-formulas').click(sendStringToParentPage);
				$('#refresh-page').click(refreshPage);
				delete window.prompt;
				var xml_text = localStorage.getItem("BlocklyWorkspace");
				console.log("receive: " +xml_text);
				//var xml = Blockly.Xml.textToDom(xml_text);
				Blockly.Xml.domToBlock(xml_text, workspace);
			});
		});
	// TODO2: Create the OK button handler
	function sendStringToParentPage() {
		var testBlock = workspace.getBlocksByType("definenamedranges", false);
		var xml = Blockly.Xml.blockToDom(testBlock[0]);
		var xml_text = Blockly.Xml.domToText(xml);
		console.log("send: " + xml_text);
		Office.context.ui.messageParent(xml_text);
	}

	function refreshPage() {
		location.reload();
	}
}());