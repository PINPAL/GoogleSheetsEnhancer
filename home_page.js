// ACTIVATE ON:
//  - https://docs.google.com/spreadsheets/u/*
// Google Sheets - Home (Open Files)

/**
 * Function to format the names of files that contain "MASTER COPY" in the name.
 * @param {Element} docsHomescreenListItem The parent element that contains the file name (docs-homescreen-list-item-title-value).
 */
function formatFileName(docsHomescreenListItem) {
	// Add a class to our element so we can find it in future searches
	docsHomescreenListItem.classList.add("pinpal_masterCopy");
	// Find the child element that contains the file details
	const docsHomescreenListItemTitle = docsHomescreenListItem.getElementsByClassName(
		"docs-homescreen-list-item-title"
	)[0];
	// Find the element within docsHomescreenListItemTitle that contains the file name
	const docsHomescreenListItemTitleValue = docsHomescreenListItemTitle.getElementsByClassName(
		"docs-homescreen-list-item-title-value"
	)[0];
	let fileName = docsHomescreenListItemTitleValue.innerText;
	// Check if the file name contains "MASTER COPY"
	if (fileName.includes("MASTER COPY")) {
		console.log("Formatting file name for " + fileName);
		// Clean up the file name by removing "MASTER COPY" and any type of bracket ([{ surrounding that text
		fileName = fileName.replace(/\s*[\(,\[,\{]\s*MASTER\s*COPY\s*[\),\],\}]\s*/i, "");
		// Replace the innerText of the element with our new file name
		docsHomescreenListItemTitleValue.innerText = fileName;
		// Create a spawn to contain the "MASTER COPY" text
		let masterCopySpan = document.createElement("span");
		masterCopySpan.className = "pinpal_masterCopySpan";
		masterCopySpan.innerText = "MASTER COPY";
		masterCopySpan.style.fontWeight = "bold";
		masterCopySpan.style.padding = "2px 4px";
		masterCopySpan.style.borderRadius = "4px";
		masterCopySpan.style.backgroundColor = "#5f638"; // Match Icon Color
		masterCopySpan.style.color = "#fff"; // Match Background Color
		masterCopySpan.style.marginRight = "4px";
		masterCopySpan.style.marginLeft = "4px";

		docsHomescreenListItemTitle.append(masterCopySpan);
	}
}

// ==========================================================================
// OBSERVERS
// ==========================================================================

// Observer to watch for any docs-homescreen-list-item being added to the page
let OBSERVER_docsHomescreenListItem = new MutationObserver(function (mutations) {
	mutations.forEach(function (mutation) {
		// Check if any nodes were added
		if (mutation.addedNodes.length) {
			mutation.addedNodes.forEach(function (addedNode) {
				// Ensure addedNode is an element and has is the class we want, and not already formatted
				if (
					addedNode.nodeType === Node.ELEMENT_NODE &&
					addedNode.classList.contains("docs-homescreen-list-item") &&
					!addedNode.classList.contains("pinpal_masterCopy")
				) {
					// Perform the formatting of the file name
					formatFileName(addedNode);
				}
			});
		}
	});
});
// Observe the body for any appearance of .docs-homescreen-list-item
OBSERVER_docsHomescreenListItem.observe(document.body, {
	childList: true,
	subtree: true,
});
