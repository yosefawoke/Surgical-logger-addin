'use strict';

(function () {

    // Define headers globally for reuse
    const headers = ["Date", "Age", "Sex", "MRN", "Diagnosis", "Procedure", "Attendant(s)", "Role", "Surgery Type"];

    Office.onReady(function(info){
        console.log("Office.onReady fired. Host:", info.host);

        if (info.host === Office.HostType.Excel) {
            console.log("Host is Excel. Proceeding with initialization.");
            try {
                 // Assign event listeners FIRST
                 const submitButton = document.getElementById('submit-button');
                 if (!submitButton) throw new Error("Submit button not found.");
                 // Point to the *real* logData function now
                 submitButton.addEventListener('click', logData);
                 console.log("Submit button listener attached to logData.");

                 const formElements = document.querySelectorAll('#log-form input, #log-form textarea, #log-form input[type=checkbox]');
                 if (formElements.length === 0) console.warn("No form elements found for input listeners.");
                 formElements.forEach(el => el.addEventListener('input', clearStatus));
                 console.log("Input listeners attached.");

                 // Populate the form elements
                 console.log("Attempting to populate attendants...");
                 populateAttendants();
                 console.log("Attempting to set current date...");
                 setCurrentDate();

                 // Ensure headers are present in the sheet (async operation)
                  console.log("Attempting to ensure headers...");
                 ensureHeaders()
                    .then(() => {
                         console.log("EnsureHeaders completed.");
                         showStatus("Add-in ready.", false); // Simple ready message
                    })
                    .catch(error => {
                        console.error("Error during ensureHeaders:", error);
                        handleError(error);
                    });

                  console.log("Initialization sync code finished.");

            } catch (error) {
                 console.error("Initialization error (sync):", error);
                 showStatus("Init Error: " + error.message, true);
                 try { handleError(error); } catch(e) { console.error("Error calling handleError during init catch:", e); }
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

    async function logData() {
        clearStatus();
        console.log("Executing logData (Full Version)...");

        try {
            // Collect data from form
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

                 // Validate required fields
                if (!date) { throw new Error("Date is missing.");}
                if (!age) { throw new Error("Age is required."); }
                if (!sexElement) { throw new Error("Sex selection is required."); }
                if (!roleElement) { throw new Error("Role selection is required."); }
                if (!surgeryTypeElement) { throw new Error("Surgery Type selection is required."); }
                if (!mrn) { throw new Error("MRN is required."); }
                if (!diagnosis) { throw new Error("Diagnosis is required."); }
                if (!procedure) { throw new Error("Procedure is required."); }

                sex = sexElement.value;
                role = roleElement.value;
                surgeryType = surgeryTypeElement.value;

             } catch (formError) {
                  console.error("Error collecting data from form:", formError);
                  throw new Error(`Form Error: ${formError.message}`);
             }

            // Get selected attendants
            const selectedAttendants = [];
            const attendantCheckboxes = document.querySelectorAll('#attendants-list input[type="checkbox"]:checked');
            attendantCheckboxes.forEach(checkbox => {
                selectedAttendants.push(checkbox.value);
            });
            const attendantsString = selectedAttendants.join(', ');

            // Prepare data row for Excel
            const dataToLog = [
                date, age, sex, mrn, diagnosis, procedure,
                attendantsString, role, surgeryType
            ];
            console.log("Data collected for logging:", dataToLog);

            // Write data to Excel
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // Find the first empty row after the header row
                const headerRange = sheet.getRange("A1");
                headerRange.load("values");
                await context.sync();

                let firstEmptyRowIndex;
                if (headerRange.values && headerRange.values[0] && headerRange.values[0][0] === headers[0]) {
                     const dataRange = sheet.getRange("A1").getUsedRange(true);
                     dataRange.load("rowCount");
                     await context.sync();
                     firstEmptyRowIndex = dataRange.rowCount;
                     console.log(`Headers found. Writing to row index: ${firstEmptyRowIndex}`);
                } else {
                     console.warn("Headers not detected correctly. Re-checking/writing headers.");
                     await ensureHeaders(); // Re-run header check/write
                     firstEmptyRowIndex = 1; // Data goes below headers
                     console.log(`Headers possibly corrected. Writing data to row index: ${firstEmptyRowIndex}`);
                }

                const targetRange = sheet.getRangeByIndexes(firstEmptyRowIndex, 0, 1, dataToLog.length);
                targetRange.values = [dataToLog];
                console.log(`Writing data to range: ${targetRange.address}`);

                sheet.getUsedRange(true).getEntireColumn().format.autofitColumns();
                await context.sync();
                console.log("Data logged successfully to Excel.");
                showStatus("Data logged successfully!", false);

            });

        } catch (error) {
            handleError(error);
        }
    }

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
