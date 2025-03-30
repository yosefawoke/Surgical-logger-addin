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
        console.log("Office.onReady fired."); // Log: Check if Office is ready

        // Check that we are running in Excel
        if (info.host === Office.HostType.Excel) {
            console.log("Host is Excel. Proceeding with initialization."); // Log: Check host
             // Attempt to initialize the add-in immediately after Office is ready
             // Wrap the entire initialization in a try/catch
            try {
                 // Assign event listeners FIRST
                 const submitButton = document.getElementById('submit-button');
                 if (!submitButton) throw new Error("Submit button not found.");
                 submitButton.addEventListener('click', logData);
                 console.log("Submit button listener attached.");

                 const formElements = document.querySelectorAll('#log-form input, #log-form textarea, #log-form input[type=checkbox]');
                 if (formElements.length === 0) console.warn("No form elements found for input listeners."); // Log: Warn if no elements found
                 formElements.forEach(el => el.addEventListener('input', clearStatus)); // Clear status on any input
                 console.log("Input listeners attached to form elements.");

                 // Populate the form elements
                 console.log("Attempting to populate attendants...");
                 populateAttendants(); // This function now throws an error if its container isn't found
                 console.log("Attempting to set current date...");
                 setCurrentDate(); // This function now throws an error if its input isn't found

                 // Ensure headers are present in the sheet (async operation)
                  console.log("Attempting to ensure headers...");
                 ensureHeaders()
                    .then(() => {
                         console.log("EnsureHeaders completed successfully (or headers already existed).");
                         showStatus("Add-in loaded and sheet headers checked.", false); // Update status after async header check completes
                    })
                    .catch(error => {
                        // Catch errors specifically from ensureHeaders
                        console.error("Error during ensureHeaders:", error);
                        handleError(error); // Use the central handler
                    });

            } catch (error) {
                 // Catch synchronous errors during initialization (e.g., element not found)
                 console.error("Initialization error (sync):", error);
                 showStatus("Error initializing add-in: " + error.message, true); // Show sync init errors
            }

        } else {
             // Handle non-Excel hosts
             console.warn("Host is not Excel:", info.host);
             // Try setting status even if DOMContentLoaded isn't guaranteed here
             const statusDiv = document.getElementById('status');
             if (statusDiv) {
                 statusDiv.textContent = "This add-in only works in Excel.";
                 statusDiv.style.color = 'red';
             } else {
                // Fallback if DOM isn't ready - less likely now but good practice
                 document.addEventListener('DOMContentLoaded', function() {
                    showStatus("This add-in only works in Excel.", true);
                });
             }
        }
    });

    const attendants = [
        "Dr Menarg", "Dr Mengist", "Dr Mequanint", "Dr Misganaw", "Dr Amare",
        "Dr Melese", "Dr Samrawit", "Dr Mesenbet", "Dr Abel", "Dr Leaynadis",
        "Dr Solomon", "Dr Sintaye", "Dr Cheru", "Dr Fasil", "Dr Meron", "Dr Adane"
    ];

    function populateAttendants() {
        console.log("Executing populateAttendants..."); // Debug log
        const listDiv = document.getElementById('attendants-list');
        if (!listDiv) {
            console.error("Attendants list container (#attendants-list) not found in HTML."); // Specific log
            throw new Error("Attendants list container not found."); // Crucial error
        }
        listDiv.innerHTML = ''; // Clear existing

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

            // Structure for Fabric UI Checkbox styling (optional)
            const checkboxContainer = document.createElement('div');
            checkboxContainer.className = 'ms-Checkbox';
            checkboxContainer.appendChild(checkbox);
            checkboxContainer.appendChild(label);

            div.appendChild(checkboxContainer); // Append styled container
            listDiv.appendChild(div);
        });
         console.log("Attendants populated successfully."); // Debug log
    }

    function setCurrentDate() {
        console.log("Executing setCurrentDate..."); // Debug log
        const dateInput = document.getElementById('date');
         if (!dateInput) {
             console.error("Date input field (#date) not found in HTML."); // Specific log
            throw new Error("Date input field not found."); // Crucial error
        }
        const today = new Date();
        const year = today.getFullYear();
        const month = (today.getMonth() + 1).toString().padStart(2, '0');
        const day = today.getDate().toString().padStart(2, '0');
        dateInput.value = `${year}-${month}-${day}`;
        console.log("Date set to:", dateInput.value); // Debug log
    }

     // Function to ensure headers exist in the sheet
    async function ensureHeaders() {
        console.log("Executing ensureHeaders..."); // Debug log
        // Note: Errors here are caught by the .catch() in the calling block
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            // Reduce scope slightly - just get A1 initially to check for headers
            const headerCheckRange = sheet.getRange("A1");
            headerCheckRange.load("values");
            await context.sync();

            let headersNeedWriting = true; // Assume we need to write unless proven otherwise
            if (headerCheckRange.values && headerCheckRange.values[0] && headerCheckRange.values[0][0] === headers[0]) {
                // Basic check passed, let's assume headers are okay for now
                // A more robust check could load the whole expected header range and compare all values
                headersNeedWriting = false;
                 console.log("Header check indicates headers likely exist (A1 matches)."); // Debug log
            } else {
                 console.log("Header check failed (A1 empty or doesn't match). Will write headers."); // Debug log
            }

            if (headersNeedWriting) {
                console.log("Writing headers..."); // Debug log
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                headerRange.format.autofitColumns(); // Autofit after writing
                await context.sync();
                console.log("Headers written successfully."); // Debug log
            }
        });
    }


    async function logData() {
        clearStatus();
        console.log("Executing logData..."); // Debug log

        try {
            // Collect data from form - Wrap this in a try block too for immediate feedback
            let date, age, sexElement, mrn, diagnosis, procedure, roleElement, surgeryTypeElement, sex, role, surgeryType;
             try {
                date = document.getElementById('date').value;
                age = document.getElementById('age').value;
                sexElement = document.querySelector('input[name="sex"]:checked');
                mrn = document.getElementById('mrn').value;
                diagnosis = document.getElementById('diagnosis').value;
                procedure = document.getElementById('procedure').value;
                roleElement = document.querySelector('input[name="role"]:checked');
                surgeryTypeElement = document.querySelector('input[name="surgery-type"]:checked');

                 // Validate required fields more explicitly
                if (!date) { throw new Error("Date is missing (should be auto-filled).");} // Date should exist
                if (!age) { throw new Error("Age is required."); }
                if (!sexElement) { throw new Error("Sex selection is required."); }
                if (!mrn) { throw new Error("MRN is required."); }
                if (!diagnosis) { throw new Error("Diagnosis is required."); }
                if (!procedure) { throw new Error("Procedure is required."); }
                if (!roleElement) { throw new Error("Role selection is required."); }
                if (!surgeryTypeElement) { throw new Error("Surgery Type selection is required."); }

                sex = sexElement.value;
                role = roleElement.value;
                surgeryType = surgeryTypeElement.value;

             } catch (formError) {
                  console.error("Error collecting data from form:", formError);
                  // Rethrow specifically as a form validation error for clarity
                  throw new Error(`Form Error: ${formError.message}`);
             }


            // Get selected attendants
            const selectedAttendants = [];
            const attendantCheckboxes = document.querySelectorAll('#attendants-list input[type="checkbox"]:checked');
            attendantCheckboxes.forEach(checkbox => {
                selectedAttendants.push(checkbox.value);
            });
            const attendantsString = selectedAttendants.join(', ');

            // Prepare data row for Excel (order MUST match headers)
            const dataToLog = [
                date, age, sex, mrn, diagnosis, procedure,
                attendantsString, role, surgeryType
            ];

            console.log("Data collected for logging:", dataToLog); // Debug log

            // Write data to Excel
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // Find the first empty row *after* the header row (robust check)
                const headerRange = sheet.getRange("A1"); // Check A1 specifically
                headerRange.load("values");
                await context.sync();

                let firstEmptyRowIndex;
                // Use getUsedRange starting from A1 only if A1 looks like a header
                if (headerRange.values && headerRange.values[0] && headerRange.values[0][0] === headers[0]) {
                     const dataRange = sheet.getRange("A1").getUsedRange(true); // Get range including headers + data
                     dataRange.load("rowCount");
                     await context.sync();
                     firstEmptyRowIndex = dataRange.rowCount; // The row *after* the last used row
                     console.log(`Headers found. Used range rows: ${dataRange.rowCount}. Writing to row index: ${firstEmptyRowIndex}`);
                } else {
                     // Headers missing or sheet doesn't start with our header.
                     // This case should be rare now due to ensureHeaders, but handle defensively.
                     console.warn("Headers not detected correctly in A1 before logging. Attempting to write from row 1.");
                     // We might need to write headers again here if ensureHeaders failed silently,
                     // but let's assume ensureHeaders worked or will run again on reload.
                     // For now, just log starting from row 1 (index 1).
                     firstEmptyRowIndex = 1; // Data goes below assumed headers
                     // Consider re-running ensureHeaders or writing them here if this becomes a common issue.
                }

                // Get the range for the new data row
                const targetRange = sheet.getRangeByIndexes(firstEmptyRowIndex, 0, 1, dataToLog.length);
                targetRange.values = [dataToLog];
                console.log(`Writing data to range: ${targetRange.address}`); // Debug log

                // Autofit columns for better readability
                sheet.getUsedRange(true).getEntireColumn().format.autofitColumns();

                await context.sync();
                console.log("Data logged successfully to Excel."); // Debug log
                showStatus("Data logged successfully!", false);

                // Optional: Clear form after successful logging
                 // document.getElementById('log-form').reset();
                 // setCurrentDate(); // Reset date if form is reset
                 // populateAttendants(); // Re-populate to clear checkboxes if form is reset

            });

        } catch (error) {
            // Catch errors from form validation OR Excel.run
            handleError(error); // Use central error handler
        }
    }

    function handleError(error) {
         console.error("Error caught by handleError:", error); // Log the raw error
         let errorMessage = "An unexpected error occurred. Please check console logs if possible."; // Default message
         if (error instanceof Error) {
             errorMessage = error.message; // Get message from Error object
         } else if (typeof error === 'string') {
            errorMessage = error; // Use string directly if error was thrown as string
         }

         // Check specifically for OfficeExtension.Error for more details
        if (error instanceof OfficeExtension.Error) {
            console.error("OfficeExtension Error Details: Code=" + error.code + ", Message=" + error.message + ", DebugInfo=" + JSON.stringify(error.debugInfo));
            // Provide a more user-friendly message for common Office errors if possible
            errorMessage = `Office API Error: ${error.message} (Code: ${error.code})`;
        }
        // Always show the error in the status div
        showStatus(`Error: ${errorMessage}`, true);
    }


    function showStatus(message, isError) {
        console.log(`showStatus called: ${message} (isError: ${isError})`); // Log status updates
        const statusDiv = document.getElementById('status');
        if (statusDiv) {
            statusDiv.textContent = message;
            statusDiv.style.color = isError ? 'red' : 'green';
        } else {
            // This should ideally not happen if the HTML is correct
            console.error("Status div (#status) not found in HTML! Cannot display message:", message);
        }
    }

    function clearStatus() {
        const statusDiv = document.getElementById('status');
         if (statusDiv) {
             statusDiv.textContent = '';
         }
    }

})(); // End of IIFE 
