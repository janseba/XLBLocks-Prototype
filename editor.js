(function () {
	"use strict";
	Office.onReady()
		.then(function() {
			$(document).ready(function () {
				// TODO1: Assign handler to the OK button.
				$('#ok-button').click(sendStringToParentPage);
				console.log('Dit is een test');
			});
		});
	// TODO2: Create the OK button handler
	function sendStringToParentPage() {
		console.log('Er is geklikt');
		var testString = localStorage.getItem("formulaID");
		console.log(testString);
	}
}());