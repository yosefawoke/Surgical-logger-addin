'use strict';

// Alert JS 1: Does the script file even load and start executing?
alert("taskpane.js: EXECUTION START");
console.log("taskpane.js: EXECUTION START");

try {
    // Attempt to change the status div *immediately*
    // Note: This might run *before* the DOM is fully ready, but it's a test.
    const statusDiv = document.getElementById('status');
    if (statusDiv) {
        alert("taskpane.js: statusDiv found immediately.");
        statusDiv.textContent = "taskpane.js was loaded!";
        statusDiv.style.color = 'purple';
        console.log("taskpane.js: Changed statusDiv immediately.");
    } else {
        alert("taskpane.js: statusDiv NOT found immediately.");
        console.error("taskpane.js: statusDiv NOT found immediately.");
        // Try again after a short delay in case it's a timing issue
         setTimeout(() => {
            const delayedStatusDiv = document.getElementById('status');
            if (delayedStatusDiv) {
                 alert("taskpane.js: statusDiv found after delay.");
                 delayedStatusDiv.textContent = "taskpane.js loaded (delayed check)!";
                 delayedStatusDiv.style.color = 'orange';
            } else {
                alert("taskpane.js: statusDiv STILL not found after delay.");
            }
         }, 500);
    }
} catch (error) {
    alert("taskpane.js: IMMEDIATE ERROR: " + error.message);
    console.error("taskpane.js: IMMEDIATE ERROR:", error);
}

// Alert JS Last: Did the script reach the end?
alert("taskpane.js: EXECUTION END");
console.log("taskpane.js: EXECUTION END");

// NO Office.onReady for this specific test. 
