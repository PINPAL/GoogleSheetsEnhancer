// ACTIVATE ON:
//  - https://docs.google.com/spreadsheets/d/*
// Google Sheets - Individual File (Editing a Sheet)

/**
 * Function to format the comments in Google Sheets conditional formatting rules.
 * This function will search for a comment in the formula and then display it above the formula.
 * The comment is expected to be in the form of (NOT case or whitespace sensitive):    LET(comment, "your comment here", yourFunctionHere)
 * @param {Element} cf_viewPill The parent element that contains the conditional formatting rule (conditionalformat-view-pill).
 */
function formatComments(cf_viewPill) {
	// Ignore if the element is already formatted
	if (cf_viewPill.classList.contains("pinpal_conditionalFormattingComment")) {
		return;
	}

	// Find the element that contains the conditional formatting rule
	const formattingRuleElement = cf_viewPill.getElementsByClassName("waffle-conditionalformat-condition")[0];

	// Yeet Google Sheets automatic "Custom formula is"
	let originalText = formattingRuleElement.innerText;
	originalText = originalText.replace(/^Custom formula is =/, "");

	// Search for formulas that have a LET(comment, "your comment here", yourFunctionHere) formula (NOT whitespace senitive)
	let foundCommentInFormula = false;
	let commentText = originalText.replace(/^\s{0,}[Ll][E-e][T-t]\(comment\s{0,},\s{0,}"/, function (token) {
		foundCommentInFormula = true;
		return "";
	});

	// If we found a comment then remove the actual formula from the text.
	if (foundCommentInFormula) {
		// Remove everything after the end of the comment
		// Store the comment text and the original formula text
		[commentText, originalText] = commentText.split('"');

		// Clean up the formula text by removing the comma and any whitespace at the start
		// Replace it with an equals sign
		originalText = originalText.replace(/^\s{0,},\s{0,}/, "=");

		// Replace the innerText of the element with our new formula text only
		formattingRuleElement.innerText = originalText;

		// Create a span element to contain our comment
		let commentElement = document.createElement("span");
		commentElement.className = "pinpal_conditionalFormattingCommentSpan";
		commentElement.innerText = commentText;
		commentElement.style.maxWidth = "fit-content";
		commentElement.style.borderRadius = "4px";
		commentElement.style.backgroundColor = "rgba(51, 51, 51, 0.15)";
		commentElement.style.textWrap = "nowrap";
		commentElement.style.display = "block";
		commentElement.style.padding = "4px";
		commentElement.style.width = "calc(100% - 8px)";
		commentElement.style.overflow = "hidden";
		commentElement.style.textOverflow = "ellipsis";

		// Put the commentElement at the start of the element
		formattingRuleElement.prepend(commentElement);

		// Adjust the maxHeight of the element to fit the comment
		formattingRuleElement.style.maxHeight = "40px";

		// Add a class to our element so we can find it in future searches
		cf_viewPill.classList.add("pinpal_conditionalFormattingComment");
	}
}

/**
 * Function to perform the initial formatting of the comments in the conditional formatting rules.
 * @param {Element} cf_ruleList The parent element that contains all the conditional formatting rules (conditionalformat-rule-list).
 */
function initialFormatComments(cf_ruleList) {
	// Get all the view pill elements
	const cf_viewPill = cf_ruleList.getElementsByClassName("waffle-conditionalformat-view-pill");

	// Iterate over each view pill element and format the comments
	for (let i = 0; i < cf_viewPill.length; i++) {
		formatComments(cf_viewPill[i]);
	}
}

/**
 * Function to modify the edit page for a conditional formatting rule
 * @param {Element} cf_editPill The parent element that contains the edit page for a conditional formatting rule (conditionalformat-edit-pill).
 */
function modifyEditPage(cf_editPill) {
	const cf_toggleTabs = cf_editPill.getElementsByClassName("waffle-conditionalformat-toggle-tabs")[0];
	const cf_arg1Holder = cf_editPill.getElementsByClassName("waffle-conditionalformat-arg1-holder")[0];
	const cf_formulaInput = cf_arg1Holder.getElementsByTagName("input")[0];
	const cf_formulaType = cf_editPill.getElementsByClassName("waffle-conditionalformat-condition-type-select")[0];

	console.log(cf_formulaType.innerText);

	// Delete all custom formula elements to prevent duplication
	document
		.querySelectorAll(".pinpal-condtionalformat-fake-formula, .pinpal-condtionalformat-rule-name-container")
		.forEach((element) => {
			element.remove();
		});

	// Observer to watch for changes to the attributes of cf_formulaType element eg: aria-activedescendant
	// This is to detect when the formula type is changed
	const OBSERVER_cf_formulaType = new MutationObserver(function (mutations) {
		mutations.forEach(function (mutation) {
			// Modify the edit page
			modifyEditPage(cf_editPill);
		});
	});
	// Start observing the cf_formulaType element
	OBSERVER_cf_formulaType.observe(cf_formulaType, {
		attributes: true,
		attributeFilter: ["aria-activedescendant"],
	});

	// Check if the formatting rule is a custom formula
	if (cf_formulaType.querySelectorAll('*[aria-selected="true"')[0].innerText !== "Custom formula is") {
		console.log("Not a custom formula, skipping...");
		return;
	}

	console.log("Generating edit page for " + cf_formulaInput.value);

	// Get the original input value
	let originalValue = cf_formulaInput.value;
	let originalComment = "";
	// Cleanup any pre-existing "comments" from the formula
	let foundCommentInFormula = false;
	let cleanedOriginalVal = originalValue.replace(
		/^\s*\=\s*[Ll][E-e][T-t]\(comment\s*,\s*".*"\s*\,\s*/,
		function (token) {
			foundCommentInFormula = true;
			return "";
		}
	);
	// If we found a comment then
	if (foundCommentInFormula) {
		// Remove the final bracket
		cleanedOriginalVal = cleanedOriginalVal.replace(/\)\s*$/, "");
		// Add a new equals sign to the start of the formula
		cleanedOriginalVal = "=" + cleanedOriginalVal;
		// Extract the comment text
		originalComment = originalValue.replace(/^\s*\=\s*[Ll][E-e][T-t]\(comment\s*,\s*"/, "");
		originalComment = originalComment.split('"')[0];
	}

	// Adjust width of custom formula input
	cf_formulaInput.parentElement.style.width = "100%";

	// Create "Rule Name" input
	let ruleNameContainer = document.createElement("div");
	ruleNameContainer.className = "pinpal-condtionalformat-rule-name-container";
	ruleNameContainer.style.padding = "0 18px";
	let ruleNameHeader = document.createElement("div");
	ruleNameHeader.className = "waffle-conditionalformat-edit-pill-section-header";
	ruleNameHeader.innerText = "Rule Name";
	ruleNameContainer.appendChild(ruleNameHeader);
	let ruleNameLabel = document.createElement("div");
	ruleNameLabel.className = "waffle-conditionalformat-edit-pill-section-label";
	ruleNameLabel.innerText = "Optional name to identify this rule...";
	ruleNameContainer.appendChild(ruleNameLabel);
	let ruleNameInput = document.createElement("input");
	ruleNameInput.className = "jfk-textinput";
	ruleNameInput.id = "pinpal-condtionalformat-rule-name";
	ruleNameInput.style.width = "100%";
	ruleNameInput.style.marginTop = 0;
	ruleNameInput.placeholder = "Optional Name for this Rule";
	ruleNameInput.value = originalComment;
	// Handler for "Rule Name" input events
	["input", "keydown", "keypress", "keyup", "change", "blur"].forEach((event) => {
		ruleNameInput.addEventListener(event, updateFormulaInputOnEdit);
	});
	// Append the "Rule Name" input to the container
	ruleNameContainer.appendChild(ruleNameInput);

	// Insert the "Rule Name" input inside cf_editPill just below the cf_toggleTabs element
	cf_editPill.insertBefore(ruleNameContainer, cf_toggleTabs.nextSibling);

	// Create a "fake" input to mimic the original input for the formula
	let fakeFormulaInput = document.createElement("input");
	fakeFormulaInput.className = "pinpal-condtionalformat-fake-formula jfk-textinput";
	fakeFormulaInput.type = "text";
	fakeFormulaInput.id = "pinpal-condtionalformat-rule-formula";
	fakeFormulaInput.style.width = "100%";
	fakeFormulaInput.placeholder = cf_formulaInput.placeholder;
	// Set the fake input value to match our cleaned up original value
	fakeFormulaInput.value = cleanedOriginalVal;
	// Handler for "fake" formula input events
	["input", "keydown", "keypress", "keyup", "change", "blur"].forEach((event) => {
		fakeFormulaInput.addEventListener(event, updateFormulaInputOnEdit);
	});
	// Insert the fake input just after the original input
	cf_formulaInput.parentElement.appendChild(fakeFormulaInput);
	// Hide the original input
	cf_formulaInput.style.backgroundColor = "red !important";
	cf_formulaInput.style.display = "none";
}

// Function to update the formula input with the new formula
function updateFormulaInputOnEdit(event) {
	let cf_formulaInput = document
		.getElementsByClassName("waffle-conditionalformat-arg1-holder")[0]
		.getElementsByTagName("input")[0];

	// If the input is empty ignore the event
	if (event.target.value === "") return;

	// Get the rule name from our "Rule Name" input
	let ruleName = document.getElementById("pinpal-condtionalformat-rule-name").value;

	// Get the rule formula from our "fake" Formula input
	let ruleFormula = document.getElementById("pinpal-condtionalformat-rule-formula").value;
	// Cleanup any preceeding = from the formula
	ruleFormula = ruleFormula.replace(/^\s*=\s*/, "");

	// Add the rule name to the original formula input
	cf_formulaInput.value = `=LET(comment, "${ruleName}", ${ruleFormula})`;
	// Trigger the change event to update the formula
	cf_formulaInput.dispatchEvent(new Event("input"));
}

// ==========================================================================
// OBSERVERS
// ==========================================================================

// Observer to watch for any conditionalformat-view-pill being added to the page
let OBSERVER_cf_viewPill = new MutationObserver(function (mutations) {
	mutations.forEach(function (mutation) {
		// Check if any nodes were added
		if (mutation.addedNodes.length) {
			mutation.addedNodes.forEach(function (addedNode) {
				// Ensure addedNode is an element and has classList before using contains
				if (addedNode.nodeType === Node.ELEMENT_NODE && addedNode.classList) {
					// Match cf_viewPill
					if (addedNode.classList.contains("waffle-conditionalformat-view-pill")) {
						formatComments(addedNode);
					}
					// Match cf_editPill
					if (addedNode.classList.contains("waffle-conditionalformat-edit-pill")) {
						modifyEditPage(addedNode);
					}
				}
			});
		}
	});
});

// Observer to watch for the initial appearance of the sidebar-container element
let OBSERVER_sidebarContainer = new MutationObserver(function (mutations) {
	mutations.forEach(function (mutation) {
		mutation.addedNodes.forEach(function (addedNode) {
			// Check if the desired container is added
			if (addedNode.nodeType === Node.ELEMENT_NODE && addedNode.classList.contains("waffle-sidebar-container")) {
				console.log("Found .waffle-sidebar-container, starting to observe...");

				const cf_ruleList = addedNode.getElementsByClassName("waffle-conditionalformat-rule-list");

				// Check if the element exists
				if (cf_ruleList.length) {
					console.log("Found .waffle-conditionalformat-rule-list, starting to observe...");

					// Perform the initial formatting of the comments
					initialFormatComments(cf_ruleList[0]);

					// Start observing this container for changes
					OBSERVER_cf_viewPill.observe(cf_ruleList[0], {
						childList: true,
						subtree: true,
					});

					// Stop the initial observer
					OBSERVER_sidebarContainer.disconnect();
				}
			}
		});
	});
});
// Initially observe the body for the first appearance of .sidebar-container
// Keep observing until we find the sidebar with .conditionalformat-rule-list
OBSERVER_sidebarContainer.observe(document.body, {
	childList: true,
	subtree: true,
});
