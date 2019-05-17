(function () {
	"use strict";
	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				// TODO1: Assign handler to the OK button.
				$('#save-formulas').click(sendStringToParentPage);
				$('#refresh-page').click(refreshPage);
				$('#show-code').click(showCode);
				delete window.prompt;
				//var xml_text = localStorage.getItem("BlocklyWorkspace");
				//var xml_text = "<xml xmlns=\"http://www.w3.org/1999/xhtml\"><block id=\"en:~K%%+@iqU2{oG6?Z/\" type=\"definenamedranges\" /></xml>"
				var xml_text = localStorage.getItem("BlocklyWorkspace");
				console.log("receive: " +xml_text);
				var xml = Blockly.Xml.textToDom(xml_text);
				Blockly.Xml.domToWorkspace(xml, workspace);
			});
		});
	// TODO2: Create the OK button handler
	function sendStringToParentPage() {
		//var testBlock = workspace.getBlocksByType("definenamedranges", false);
		//var xml = Blockly.Xml.blockToDom(testBlock[0]);
		var xml = Blockly.Xml.workspaceToDom(workspace);
		var xml_text = Blockly.Xml.domToText(xml);
		//xml_text = "<xml xmlns=\"http://www.w3.org/1999/xhtml\">" + xml_text;
		//xml_text = xml_text + "</xml>";
		console.log("send: " + xml_text);
		var messageToTaskPane = new Object();
		messageToTaskPane.Type = 'blockDefinition';
		messageToTaskPane.MessageContent = xml_text;
		Office.context.ui.messageParent(JSON.stringify(messageToTaskPane));
	}

	function refreshPage() {
		location.reload();
	}

	function showCode() {
		var formulaBlocks = workspace.getBlocksByType('formula', false);
		var code = Blockly.JavaScript.blockToCode(formulaBlocks[0]);
		code = JSON.parse(code);
		var workspaceXML = Blockly.Xml.workspaceToDom(workspace);
		var workspaceString = Blockly.Xml.domToText(workspaceXML);
		code.blockDefinition = workspaceString;
		code = JSON.stringify(code);
		var messageToTaskPane = new Object();
		messageToTaskPane.Type = 'formula';
		messageToTaskPane.MessageContent = code;
		Office.context.ui.messageParent(JSON.stringify(messageToTaskPane));
	}
}());