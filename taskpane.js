'use strict';

(function () {

    // Define headers globally for reuse
    const headers = ["Date", "Age", "Sex", "MRN", "Diagnosis", "Procedure", "Attendant(s)", "Role", "Surgery Type"];

    Office.onReady(function(info){
        // Check that we are running in Excel
        if (info.host === Office.HostType.Excel) {
            // Ensure the DOM is fully loaded before interacting with it
            document.addEventListener('DOMContentLoaded', function() {
                // Assign event listeners
                document.getElementById('submit-button').addEventListener('click', logData);
                const formElements = document.querySelectorAll('#log-form input, #log-form textarea, #log-form input[type=checkbox]');
                formElements.forEach(el => el.addEventListener('input', clearStatus)); // Clear status on any input

                // Populate the form elements
                try {
                    populateAttendants();
                    setCurrentDate();
                    showStatus("Add-in loaded successfully.", false); // Initial success message
                } catch (error) {
                    showStatus("Error initializing form: " + error.message, true);
                    console.error("Initialization error:", error);
                }

                 // Ensure headers are present in the sheet
                 ensureHeaders().catch(handleError); // Call async function

            });
        } else {
             document.addEventListener('DOMContentLoaded', function() {
                showStatus("This add-in only works in Excel.", true);
            });
        }
    });

    const attendants = [
        "Dr Menarg", "Dr Mengist", "Dr Mequanint", "Dr Misganaw", "Dr Amare",
        "Dr Melese", "Dr Samrawit", "Dr Mesenbet", "Dr Abel", "Dr Leaynadis",
        "Dr Solomon", "Dr Sintaye", "Dr Cheru", "Dr Fasil", "Dr Meron", "Dr Adane"
    ];

    function populateAttendants() {
        console.log("Populating attendants..."); // Debug log
        const listDiv = document.getElementById('attendants-list');
        if (!listDiv) {
            throw new Error("Attendants list container not found.");
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
         console.log("Attendants populated."); // Debug log
    }

    function setCurrentDate() {
        console.log("Setting current date..."); // Debug log
        const dateInput = document.getElementById('date');
         if (!dateInput) {
            throw new Error("Date input field not found.");
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
        console.log("Ensuring headers..."); // Debug log
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.load("values");
                await context.sync();

                // Basic check: Does the first cell match the first header?
                // A more robust check could compare all header values.
                if (!headerRange.values || !headerRange.values[0] || headerRange.values[0][0] !== headers[0]) {
                    console.log("Headers not found or incorrect, writing headers..."); // Debug log
                    headerRange.values = [headers];
                    headerRange.format.font.bold = true;
                    headerRange.format.autofitColumns();
                    await context.sync();
                    console.log("Headers written."); // Debug log
                } else {
                     console.log("Headers found."); // Debug log
                }
            });
        } catch (error) {
             console.error("Error ensuring headers:", error);
             showStatus("Error setting up sheet headers: " + error.message, true);
             // Rethrow or handle as appropriate for initialization sequence
             throw error;
        }
    }


    async function logData() {
        clearStatus();
        console.log("Log Data button clicked."); // Debug log

        try {
            // Collect data from form
            const date = document.getElementById('date').value;
            const age = document.getElementById('age').value;
            const sexElement = document.querySelector('input[name="sex"]:checked');
            const mrn = document.getElementById('mrn').value;
            const diagnosis = document.getElementById('diagnosis').value;
            const procedure = document.getElementById('procedure').value;
            const roleElement = document.querySelector('input[name="role"]:checked');
            const surgeryTypeElement = document.querySelector('input[name="surgery-type"]:checked');

            // Validate required fields more explicitly
            if (!age) { throw new Error("Age is required."); }
            if (!sexElement) { throw new Error("Sex selection is required."); } // Should not happen with default checked
            if (!mrn) { throw new Error("MRN is required."); }
            if (!diagnosis) { throw new Error("Diagnosis is required."); }
            if (!procedure) { throw new Error("Procedure is required."); }
            if (!roleElement) { throw new Error("Role selection is required."); } // Should not happen
            if (!surgeryTypeElement) { throw new Error("Surgery Type selection is required."); } // Should not happen


            const sex = sexElement.value;
            const role = roleElement.value;
            const surgeryType = surgeryTypeElement.value;

            // Get selected attendants
            const selectedAttendants = [];
            const attendantCheckboxes = document.querySelectorAll('#attendants-list input[type="checkbox"]:checked');
            attendantCheckboxes.forEach(checkbox => {
                selectedAttendants.push(checkbox.value);
            });
            const attendantsString = selectedAttendants.join(', '); // Combine names, empty string if none selected

            // Prepare data row for Excel (order MUST match headers)
            const dataToLog = [
                date, age, sex, mrn, diagnosis, procedure,
                attendantsString, role, surgeryType
            ];

            console.log("Data collected:", dataToLog); // Debug log

            // Write data to Excel
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // Find the first empty row *after* the header row (more robust)
                 // Check row 1 for headers first to ensure we don't overwrite them
                const headerRange = sheet.getRange("A1"); // Check A1 specifically
                headerRange.load("values");
                await context.sync();

                let firstEmptyRowIndex;
                 // Check if A1 looks like the start of our headers
                if (headerRange.values && headerRange.values[0][0] === headers[0]) {
                    // Headers likely exist, find the next row using used range below headers
                     const usedRange = sheet.getRange("A1").getUsedRange(true); // Range including headers and data
                     usedRange.load("rowCount");
                     await context.sync();
                     firstEmptyRowIndex = usedRange.rowCount; // The first row *after* the used range
                     console.log(`Headers found. Used range rows: ${usedRange.rowCount}. Writing to row index: ${firstEmptyRowIndex}`); // Debug log
                } else {
                     // Headers don't seem to be in A1, assume sheet is empty or has other data.
                     // Attempt to write headers first, then data. This might overwrite unrelated data if sheet isn't truly empty.
                     console.warn("Headers not detected in A1. Writing headers and data from the top."); // Debug log
                     const writeHeaderRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                     writeHeaderRange.values = [headers];
                     writeHeaderRange.format.font.bold = true;
                     firstEmptyRowIndex = 1; // Data goes below the newly written headers
                }


                // Get the range for the new data row
                const dataRange = sheet.getRangeByIndexes(firstEmptyRowIndex, 0, 1, dataToLog.length);
                dataRange.values = [dataToLog];
                console.log(`Writing data to range: ${dataRange.address}`); // Debug log

                // Autofit columns for better readability (consider doing this less often if performance is an issue)
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
            handleError(error); // Use central error handler
        }
    }

    function handleError(error) {
         console.error("Error:", error);
         let errorMessage = "An unexpected error occurred.";
         if (error instanceof Error) {
             errorMessage = error.message;
         } else if (typeof error === 'string') {
            errorMessage = error;
         }

        if (error instanceof OfficeExtension.Error) {
            console.error("OfficeExtension Error Details: " + JSON.stringify(error.debugInfo));
            errorMessage = `Office API Error: ${error.message} (Code: ${error.code})`;
        }
        showStatus(`Error: ${errorMessage}`, true);
    }


    function showStatus(message, isError) {
        const statusDiv = document.getElementById('status');
        if (statusDiv) {
            statusDiv.textContent = message;
            statusDiv.style.color = isError ? 'red' : 'green';
        } else {
            console.warn("Status div not found. Message:", message);
        }
    }

    function clearStatus() {
        const statusDiv = document.getElementById('status');
         if (statusDiv) {
             statusDiv.textContent = '';
         }
    }

})();
