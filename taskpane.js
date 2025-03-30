'use strict';

(function () {

    // Define headers globally for reuse
    const headers = ["Date", "Age", "Sex", "MRN", "Diagnosis", "Procedure", "Attendant(s)", "Role", "Surgery Type"];

    Office.onReady(function(info){
        console.log("Office.onReady fired.");

        // Check that we are running in Excel
        if (info.host === Office.HostType.Excel) {
            console.log("Host is Excel. Scheduling initialization with slight delay...");

            // Introduce a small delay to ensure DOM is fully ready after Office.onReady
            setTimeout(function() {
                console.log("Delayed initialization starting...");
                // Wrap the delayed initialization in its own try/catch
                try {
                    // --- Check for essential elements FIRST ---
                    console.log("Checking for core HTML elements (delayed)...");
                    const statusDiv = document.getElementById('status');
                    const submitButton = document.getElementById('submit-button');
                    const dateInput = document.getElementById('date');
                    const attendantsList = document.getElementById('attendants-list');

                    // Use the statusDiv immediately if found, otherwise log error
                    const showInitStatus = (msg, isErr) => {
                        console.log(`Init Status: ${msg} (isError: ${isErr})`);
                        if (statusDiv) {
                            statusDiv.textContent = msg;
                            statusDiv.style.color = isErr ? 'red' : 'orange'; // Use orange for init status
                        } else {
                             console.error("CRITICAL: Status div (#status) not found during delayed init!");
                        }
                    };

                     if (!statusDiv) throw new Error("Status div (#status) not found.");
                     showInitStatus("Checking elements...", false); // Show initial check message

                     if (!submitButton) throw new Error("Submit button (#submit-button) not found.");
                     if (!dateInput) throw new Error("Date input (#date) not found.");
                     if (!attendantsList) throw new Error("Attendants list container (#attendants-list) not found.");
                     showInitStatus("Core elements verified.", false);
                     // --- End Element Checks ---

                     // Assign event listeners
                     submitButton.addEventListener('click', logData);
                     console.log("Submit button listener attached.");

                     const formElements = document.querySelectorAll('#log-form input, #log-form textarea, #log-form input[type=checkbox]');
                     formElements.forEach(el => el.addEventListener('input', clearStatus));
                     console.log("Input listeners potentially attached (count: " + formElements.length + ").");

                     // Populate the form elements
                     showInitStatus("Populating attendants...", false);
                     populateAttendants();
                     showInitStatus("Setting date...", false);
                     setCurrentDate();

                     // *** Run Header Check on Init ***
                     console.log("EnsureHeaders() call starting on init...");
                     ensureHeaders()
                        .then(() => {
                            showStatus("Add-in initialized. Header check complete.", false); // Use final status color
                            console.log("Delayed initialization completed successfully (inc. header check).");
                        })
                        .catch(handleError); // Catch errors from ensureHeaders

                } catch (error) {
                     // Catch synchronous errors during DELAYED initialization
                     console.error("Delayed initialization error caught:", error);
                     handleError(error); // handleError uses showStatus internally
                }
            }, 100); // 100ms delay

        } else {
             // Handle non-Excel hosts immediately
             console.warn("Host is not Excel:", info.host);
             showStatus("This add-in only works in Excel.", true); // Use showStatus safely
        }
    });

    // --- Attendants List ---
    const attendants = [
        "Dr Menarg", "Dr Mengist", "Dr Mequanint", "Dr Misganaw", "Dr Amare",
        "Dr Melese", "Dr Samrawit", "Dr Mesenbet", "Dr Abel", "Dr Leaynadis",
        "Dr Solomon", "Dr Sintaye", "Dr Cheru", "Dr Fasil", "Dr Meron", "Dr Adane"
    ];

    // --- Populate Attendants Function ---
    function populateAttendants() {
        console.log("Executing populateAttendants...");
        const listDiv = document.getElementById('attendants-list');
        listDiv.innerHTML = ''; // Clear existing

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

    // --- Set Current Date Function ---
    function setCurrentDate() {
        console.log("Executing setCurrentDate...");
        const dateInput = document.getElementById('date');
        const today = new Date();
        const year = today.getFullYear();
        const month = (today.getMonth() + 1).toString().padStart(2, '0');
        const day = today.getDate().toString().padStart(2, '0');
        dateInput.value = `${year}-${month}-${day}`;
        console.log("Date set to:", dateInput.value);
    }

    // --- Ensure Headers Function ---
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
        // Errors are caught by the caller's .catch()
    }

    // --- Log Data Function ---
    async function logData() {
        clearStatus();
        console.log("Executing logData...");

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
                  throw new Error(`Form Error: ${formError.message}`);
             }

            // Get selected attendants
            const selectedAttendants = [];
            const attendantCheckboxes = document.querySelectorAll('#attendants-list input[type="checkbox"]:checked');
            attendantCheckboxes.forEach(checkbox => { selectedAttendants.push(checkbox.value); });
            const attendantsString = selectedAttendants.join(', ');

            // Prepare data row
            const dataToLog = [ date, age, sex, mrn, diagnosis, procedure, attendantsString, role, surgeryType ];
            console.log("Data collected for logging:", dataToLog);

            // *** Ensure header check runs before logging data ***
             console.log("Running ensureHeaders before logging data...");
             await ensureHeaders(); // Run header check here now
             console.log("EnsureHeaders completed before logging.");


            // Write data to Excel
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();

                // Find the first empty row after the header row
                const dataRange = sheet.getRange("A1").getUsedRange(true); // Assumes headers are now in A1
                dataRange.load("rowCount");
                await context.sync();
                const firstEmptyRowIndex = dataRange.rowCount; // Row *after* last used row
                console.log(`Used range rows: ${dataRange.rowCount}. Writing to row index: ${firstEmptyRowIndex}`);


                // Get the range for the new data row
                const targetRange = sheet.getRangeByIndexes(firstEmptyRowIndex, 0, 1, dataToLog.length);
                targetRange.values = [dataToLog];
                console.log(`Writing data to range: ${targetRange.address}`);

                // Autofit columns
                sheet.getUsedRange(true).getEntireColumn().format.autofitColumns();

                await context.sync();
                console.log("Data logged successfully to Excel.");
                showStatus("Data logged successfully!", false);

                // Optional: Clear form
                 // document.getElementById('log-form').reset();
                 // setCurrentDate();
                 // populateAttendants(); // Need to re-run this to clear checkboxes visually

            });

        } catch (error) {
            // Catch errors from form validation, ensureHeaders, or Excel.run
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
            statusDiv.style.color = isError ? 'red' : 'green';
        } else {
            console.error("Status div (#status) not found! Cannot display message:", message);
        }
    }

    // --- Clear Status Function ---
    function clearStatus() {
        const statusDiv = document.getElementById('status');
         if (statusDiv) { statusDiv.textContent = ''; }
    }

})(); // End of IIFE 
