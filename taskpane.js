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

(function () {

    // Define headers globally for reuse
    const headers = ["Date", "Age", "Sex", "MRN", "Diagnosis", "Procedure", "Attendant(s)", "Role", "Surgery Type"];

    Office.onReady(function(info){
        console.log("Office.onReady fired. Host:", info.host); // Log: Check if Office is ready

        // Check that we are running in Excel
        if (info.host === Office.HostType.Excel) {
            console.log("Host is Excel. Proceeding with initialization."); // Log: Check host
             // Attempt to initialize the add-in immediately after Office is ready
             // Wrap the entire initialization in a try/catch
            try {
                 // 1. Test showStatus immediately
                 showStatus("Status: Office Ready.", false);
                 console.log("Updated status inside Office.onReady.");

                 // 2. Test finding button and attaching listener
                 const submitButton = document.getElementById('submit-button');
                 if (!submitButton) throw new Error("Submit button not found.");
                 // Use simplified logData for this test
                 submitButton.addEventListener('click', logDataTestVersion);
                 showStatus("Status: Listener attached.", false);
                 console.log("Submit button listener attached (test version).");

                 // 3. Test finding other elements (without populating yet)
                 const dateInput = document.getElementById('date');
                 if (!dateInput) throw new Error("Date input not found.");
                 const listDiv = document.getElementById('attendants-list');
                 if (!listDiv) throw new Error("Attendants list not found.");
                 showStatus("Status: Core elements found.", false);
                 console.log("Core form elements found.");

                 // --- Temporarily disable population/header check for this test ---
                 console.log("Skipping populateAttendants for test.");
                 // populateAttendants(); // Disabled

                 console.log("Skipping setCurrentDate for test.");
                 // setCurrentDate(); // Disabled

                 console.log("Skipping ensureHeaders for test.");
                 // ensureHeaders().catch(handleError); // Disabled
                 // --- End temporary disable ---

                  // Update status again if possible
                  showStatus("Status: Init checks passed (tasks skipped).", false);
                  console.log("Initialization code within try block finished (mostly skipped).");

            } catch (error) {
                 // Catch synchronous errors during initialization (e.g., element not found)
                 console.error("Initialization error (sync):", error);
                 // Use showStatus which has internal check for statusDiv
                 showStatus("Init Error: " + error.message, true);
                 // Log details via handleError too if possible
                 try { handleError(error); } catch(e) { console.error("Error calling handleError:", e); }
            }

        } else {
             console.warn("Host is not Excel:", info.host);
             showStatus("This add-in only works in Excel.", true);
        }
    });

    const attendants = [
        "Dr Menarg", "Dr Mengist", "Dr Mequanint", "Dr Misganaw", "Dr Amare",
        "Dr Melese", "Dr Samrawit", "Dr Mesenbet", "Dr Abel", "Dr Leaynadis",
        "Dr Solomon", "Dr Sintaye", "Dr Cheru", "Dr Fasil", "Dr Meron", "Dr Adane"
    ];

    // Function definitions remain the same, but calls are disabled above for testing
    function populateAttendants() {
        console.log("Executing populateAttendants...");
        const listDiv = document.getElementById('attendants-list');
        if (!listDiv) { throw new Error("Attendants list container not found."); }
        listDiv.innerHTML = '';
        attendants.forEach((name, index) => {
             const div = document.createElement('div');
             const checkbox = document.createElement('input');
             checkbox.type = 'checkbox';
             checkbox.id = `attendant-${index}`;
             checkbox.name = 'attendant';
             checkbox.value = name;
             checkbox.className = 'ms-Checkbox-input'; // Optional Fabric UI styling

             const label = document.createElement('label');
             label.htmlFor = `attendant-${index}`;
             label.className = 'ms-Checkbox-label'; // Optional Fabric UI styling
             label.appendChild(document.createTextNode(name));

             const checkboxContainer = document.createElement('div');
             checkboxContainer.className = 'ms-Checkbox';
             checkboxContainer.appendChild(checkbox);
             checkboxContainer.appendChild(label);

             div.appendChild(checkboxContainer); // Append styled container
             listDiv.appendChild(div);
        });
        console.log("Attendants populated successfully.");
    }

    function setCurrentDate() {
        console.log("Executing setCurrentDate...");
        const dateInput = document.getElementById('date');
        if (!dateInput) { throw new Error("Date input field not found."); }
        const today = new Date();
        const year = today.getFullYear();
        const month = (today.getMonth() + 1).toString().padStart(2, '0');
        const day = today.getDate().toString().padStart(2, '0');
        dateInput.value = `${year}-${month}-${day}`;
        console.log("Date set to:", dateInput.value);
    }

    async function ensureHeaders() {
        console.log("Executing ensureHeaders...");
        await Excel.run(async (context) => {
             const sheet = context.workbook.worksheets.getActiveWorksheet();
             const headerCheckRange = sheet.getRange("A1");
             headerCheckRange.load("values");
             await context.sync();
             let headersNeedWriting = true;
             if (headerCheckRange.values && headerCheckRange.values[0] && headerCheckRange.values[0][0] === headers[0]) {
                 headersNeedWriting = false;
                 console.log("Header check indicates headers likely exist (A1 matches).");
             } else {
                 console.log("Header check failed (A1 empty or doesn't match). Will write headers.");
             }
             if (headersNeedWriting) {
                 console.log("Writing headers...");
                 const headerRangeToWrite = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                 headerRangeToWrite.values = [headers];
                 headerRangeToWrite.format.font.bold = true;
                 headerRangeToWrite.format.autofitColumns();
                 await context.sync();
                 console.log("Headers written successfully.");
             }
        });
    }

    // Keep original logData function available but don't call it directly in this test
    async function logData() { /* ... original full logData code ... */ }

    // Simplified logData for button testing - just update status
    function logDataTestVersion() {
        clearStatus();
        console.log("Executing logData (Test version)...");
        try {
            showStatus("Log Data Clicked! (No Action)", false); // Update status without Excel interaction
            console.log("Log Data button click handled (test version).");
        } catch (error) {
            handleError(error); // Still use error handler just in case
        }
    }

    // Helper functions unchanged
    function handleError(error) {
         console.error("Error caught by handleError:", error);
         let errorMessage = "An unexpected error occurred.";
         if (error instanceof Error) { errorMessage = error.message; }
         else if (typeof error === 'string') { errorMessage = error; }
         if (error instanceof OfficeExtension.Error) {
             console.error("OfficeExtension Error Details:", JSON.stringify(error.debugInfo));
             errorMessage = `Office API Error: ${error.message} (Code: ${error.code})`;
         }
         showStatus(`Error: ${errorMessage}`, true);
    }

    function showStatus(message, isError) {
        console.log(`showStatus: ${message} (isError: ${isError})`);
        const statusDiv = document.getElementById('status');
        if (statusDiv) {
            statusDiv.textContent = message;
            statusDiv.style.color = isError ? 'red' : 'green';
        } else {
            console.error("Status div (#status) not found! Cannot display message:", message);
        }
    }

    function clearStatus() {
        const statusDiv = document.getElementById('status');
         if (statusDiv) { statusDiv.textContent = ''; }
    }

})(); // End of IIFE 
