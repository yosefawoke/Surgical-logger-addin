'use strict';

(function () {

    // Define headers globally for reuse (still needed for potential future use)
    const headers = ["Date", "Age", "Sex", "MRN", "Diagnosis", "Procedure", "Attendant(s)", "Role", "Surgery Type"];

    Office.onReady(function(info){
        // Alert 1: Test if Office.onReady fires
        alert("Office.onReady fired!");
        console.log("Office.onReady fired.");

        if (info.host === Office.HostType.Excel) {
            console.log("Host is Excel. Scheduling minimal init check...");

            // Introduce delay
            setTimeout(function() {
                // Alert 2: Test if setTimeout callback executes
                alert("setTimeout callback executed!");
                console.log("Minimal delayed init starting...");
                try {
                    // Alert 3: Check for statusDiv
                    const statusDiv = document.getElementById('status');
                    alert(statusDiv ? "statusDiv FOUND!" : "statusDiv NOT FOUND!");
                    console.log("Checked for statusDiv:", statusDiv ? "Found" : "Not Found");

                    // Alert 4: Check for submitButton
                    const submitButton = document.getElementById('submit-button');
                    alert(submitButton ? "submitButton FOUND!" : "submitButton NOT FOUND!");
                     console.log("Checked for submitButton:", submitButton ? "Found" : "Not Found");

                    // Only proceed if both crucial elements are found
                    if (!statusDiv) {
                        console.error("CRITICAL: Status div (#status) not found!");
                        alert("Status div missing! Cannot proceed."); // Alert as fallback
                        return; // Stop if status div missing
                    }
                     if (!submitButton) {
                         console.error("CRITICAL: Submit button (#submit-button) not found!");
                         alert("Submit button missing! Cannot proceed."); // Alert as fallback
                         return; // Stop if button missing
                    }

                    statusDiv.textContent = "Attempting to attach listener..."; // Test message 1
                    statusDiv.style.color = 'orange';
                    console.log("Set status to 'Attempting to attach listener...'");

                    // Attach the listener
                    submitButton.addEventListener('click', logData);
                    console.log("Submit button listener attached for test.");

                    // Report success of this minimal init
                    statusDiv.textContent = "Minimal Init OK (Alert Test). Click button.";
                     statusDiv.style.color = 'green';
                    console.log("Minimal init seems OK.");

                } catch (error) {
                     console.error("Minimal delayed initialization error:", error);
                     // Try alerting the error message
                     alert("Error during init: " + error.message);
                     handleError(error); // Display error in status div if possible
                }
            }, 150); // Slightly longer delay just in case

        } else {
             console.warn("Host is not Excel:", info.host);
             alert("Host is not Excel: " + info.host); // Alert if not Excel
             // Use showStatus which checks for statusDiv
             showStatus("This add-in only works in Excel.", true);
        }
    });

    // --- Attendants List (Keep definition but don't call populate) ---
    const attendants = [
        "Dr Menarg", "Dr Mengist", "Dr Mequanint", "Dr Misganaw", "Dr Amare",
        "Dr Melese", "Dr Samrawit", "Dr Mesenbet", "Dr Abel", "Dr Leaynadis",
        "Dr Solomon", "Dr Sintaye", "Dr Cheru", "Dr Fasil", "Dr Meron", "Dr Adane"
    ];

    // --- Populate Attendants Function (Keep definition but don't call) ---
    function populateAttendants() {
        console.log("Executing populateAttendants..."); // Should not see this log now
        const listDiv = document.getElementById('attendants-list');
        if (!listDiv) throw new Error("Attendants list container not found.");
        listDiv.innerHTML = '';
         attendants.forEach((name, index) => {
             const div = document.createElement('div');
             const checkbox = document.createElement('input');
             checkbox.type = 'checkbox';
             checkbox.id = `attendant-${index}`;
             checkbox.name = 'attendant';
             checkbox.value = name;
             checkbox.className = 'ms-Checkbox-input';

             const label = document.createElement('label');
             label.htmlFor = `attendant-${index}`;
             label.className = 'ms-Checkbox-label';
             label.appendChild(document.createTextNode(name));

             const checkboxContainer = document.createElement('div');
             checkboxContainer.className = 'ms-Checkbox';
             checkboxContainer.appendChild(checkbox);
             checkboxContainer.appendChild(label);

             div.appendChild(checkboxContainer);
             listDiv.appendChild(div);
         });
          console.log("Attendants populated successfully.");
     }


    // --- Set Current Date Function (Keep definition but don't call) ---
    function setCurrentDate() {
        console.log("Executing setCurrentDate..."); // Should not see this log now
        const dateInput = document.getElementById('date');
         if (!dateInput) throw new Error("Date input field not found.");
         const today = new Date();
         const year = today.getFullYear();
         const month = (today.getMonth() + 1).toString().padStart(2, '0');
         const day = today.getDate().toString().padStart(2, '0');
         dateInput.value = `${year}-${month}-${day}`;
         console.log("Date set to:", dateInput.value);
     }

    // --- Ensure Headers Function (Keep definition) ---
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
                 console.log("Header check indicates headers likely exist.");
            } else {
                 console.log("Header check failed. Will write headers.");
            }

            if (headersNeedWriting) {
                console.log("Writing headers...");
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                headerRange.format.autofitColumns();
                await context.sync();
                console.log("Headers written successfully.");
            }
        });
    }

    // --- Log Data Function (SIMPLIFIED FOR TESTING) ---
    async function logData() {
        // Alert 5: Test if button click handler runs
        alert("Log Data button clicked!");
        console.log("Log Data button clicked (Test Version).");
        try {
            // ONLY update the status div
             showStatus("Log Button Clicked! (Alert Test)", false); // Test Message 2
             console.log("Set status to 'Log Button Clicked!'");

        } catch (error) {
             console.error("Error in test logData:", error);
             alert("Error in logData: " + error.message);
             handleError(error);
        }
    }

    // --- Error Handling Function ---
    function handleError(error) {
         console.error("Error caught by handleError:", error);
         let errorMessage = "An unexpected error occurred.";
         if (error instanceof Error) { errorMessage = error.message; }
         else if (typeof error === 'string') { errorMessage = error; }

        if (error instanceof OfficeExtension.Error) {
            console.error("OfficeExtension Error Details: Code=" + error.code + ", Message=" + error.message + ", DebugInfo=" + JSON.stringify(error.debugInfo));
            errorMessage = `Office API Error: ${error.message} (Code: ${error.code})`;
        }
        // Use showStatus to display the error
        showStatus(`Error: ${errorMessage}`, true);
    }

    // --- Status Display Function ---
    function showStatus(message, isError) {
        console.log(`showStatus: ${message} (isError: ${isError})`);
        const statusDiv = document.getElementById('status');
        if (statusDiv) {
            statusDiv.textContent = message;
            statusDiv.style.color = isError ? 'red' : 'green'; // Green for success/test clicks now
        } else {
            console.error("Status div (#status) not found! Cannot display message:", message);
            // alert((isError ? "Error: " : "Status: ") + message);
        }
    }

    // --- Clear Status Function ---
    function clearStatus() { // Keep this, logData calls it
        const statusDiv = document.getElementById('status');
         if (statusDiv) { statusDiv.textContent = ''; }
    }

})(); // End of IIFE 
