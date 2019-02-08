(function () {
	"use strict";
	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				// TODO1: Assign handler to the OK button.
				$('#save-formulas').click(sendStringToParentPage);
				delete window.prompt;
				var blocklyWorkspace = localStorage.getItem("BlocklyWorkspace");
				var xml = Blockly.Xml.textToDom(blocklyWorkspace);
				Blockly.Xml.domToWorkspace(xml, workspace);
			});
		});
	// TODO2: Create the OK button handler
	function sendStringToParentPage() {
		var testBlock = workspace.getTopBlocks(true);
		var xml = Blockly.Xml.blockToDom(block);
		var xml_text = Blockly.Xml.domToPrettyText(xml);
		console.log(xml_text);
		Office.context.ui.messageParent(xml_text);
	}
}());