'use strict';

(function () {

    // The initialize function must be run each time a new page is loaded.
    Office.onReady(function(info){
        if (info.host === Office.HostType.Excel) {
            // Assign event handlers and populate initial data only after Office is ready
            document.addEventListener('DOMContentLoaded', function() {
                populateAttendants();
                setCurrentDate();
                document.getElementById('submit-button').addEventListener('click', logData);
                // Clear status on input change
                const formElements = document.querySelectorAll('#log-form input, #log-form textarea');
                formElements.forEach(el => el.addEventListener('input', clearStatus));
            });
        }
    });

    const attendants = [
        "Dr Menarg", "Dr Mengist", "Dr Mequanint", "Dr Misganaw", "Dr Amare",
        "Dr Melese", "Dr Samrawit", "Dr Mesenbet", "Dr Abel", "Dr Leaynadis", // Corrected spelling
        "Dr Solomon", "Dr Sintaye", "Dr Cheru", "Dr Fasil", "Dr Meron", "Dr Adane"
    ];

    function populateAttendants() {
        const listDiv = document.getElementById('attendants-list');
        listDiv.innerHTML = ''; // Clear existing

        attendants.forEach((name, index) => {
            const div = document.createElement('div');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `attendant-${index}`;
            checkbox.name = 'attendant';
            checkbox.value = name;

            const label = document.createElement('label');
            label.htmlFor = `attendant-${index}`;
            label.appendChild(document.createTextNode(name));

            div.appendChild(checkbox);
            div.appendChild(label);
            listDiv.appendChild(div);
        });
    }

    function setCurrentDate() {
        const dateInput = document.getElementById('date');
        const today = new Date();
        // Format date as YYYY-MM-DD for the input type="date"
        const year = today.getFullYear();
        const month = (today.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-indexed
        const day = today.getDate().toString().padStart(2, '0');
        dateInput.value = `${year}-${month}-${day}`;
        // Optional: make it readonly via JS too, though HTML attribute is usually sufficient
        // dateInput.readOnly = true;
    }

    function logData() {
        clearStatus(); // Clear previous status

        // Collect data from form
        const date = document.getElementById('date').value;
        const age = document.getElementById('age').value;
        const sex = document.querySelector('input[name="sex"]:checked').value;
        const mrn = document.getElementById('mrn').value;
        const diagnosis = document.getElementById('diagnosis').value;
        const procedure = document.getElementById('procedure').value;

        // Get selected attendants
        const selectedAttendants = [];
        const attendantCheckboxes = document.querySelectorAll('#attendants-list input[type="checkbox"]:checked');
        attendantCheckboxes.forEach(checkbox => {
            selectedAttendants.push(checkbox.value);
        });
        const attendantsString = selectedAttendants.join(', '); // Combine names

        const role = document.querySelector('input[name="role"]:checked').value;
        const surgeryType = document.querySelector('input[name="surgery-type"]:checked').value;

        // Basic validation (optional, add more as needed)
        if (!age || !mrn || !diagnosis || !procedure) {
             showStatus("Please fill in all required fields.", true);
             return;
        }

        // Prepare data row for Excel
        // IMPORTANT: The order here MUST match the desired column order in Excel
        const dataToLog = [
            date,
            age,
            sex,
            mrn,
            diagnosis,
            procedure,
            attendantsString, // Comma-separated list
            role,
            surgeryType
        ];

        // Write data to Excel
        Excel.run(async function (context) {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Define headers (must match dataToLog order)
            const headers = ["Date", "Age", "Sex", "MRN", "Diagnosis", "Procedure", "Attendant(s)", "Role", "Surgery Type"];

            // Find the first completely empty row after the header or data
            // A simple approach: find used range and add below it.
            const usedRange = sheet.getUsedRange(true); // Or false if hidden rows/cols shouldn't count
            usedRange.load('rowCount');
            await context.sync();

            let firstEmptyRowIndex = 0; // Default to first row if sheet is empty

            // Check if the sheet is truly empty or just has a header
            const checkRange = sheet.getRange("A1");
            checkRange.load("values");
            await context.sync();

            if (usedRange.rowCount === 1 && checkRange.values[0][0] === "") {
                 // Sheet is effectively empty, usedRange might report 1 row if A1 was touched
                firstEmptyRowIndex = 0;
            } else {
                firstEmptyRowIndex = usedRange.rowCount;
            }


            // Write headers if the sheet is empty (first row)
            if (firstEmptyRowIndex === 0) {
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.values = [headers];
                headerRange.format.font.bold = true; // Make headers bold
                firstEmptyRowIndex = 1; // Data goes in the next row
            }

            // Get the range for the new data row
            const dataRange = sheet.getRangeByIndexes(firstEmptyRowIndex, 0, 1, dataToLog.length);
            dataRange.values = [dataToLog];

            // Autofit columns for better readability (optional)
            sheet.getUsedRange(true).getEntireColumn().format.autofitColumns();

            await context.sync();
            showStatus("Data logged successfully!", false);

            // Optional: Clear form fields after successful logging
            // document.getElementById('log-form').reset();
            // setCurrentDate(); // Reset date after clearing form
            // populateAttendants(); // Reset checkboxes

        }).catch(function (error) {
            console.error("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.error("Debug info: " + JSON.stringify(error.debugInfo));
            }
            showStatus("Error logging data: " + error.message, true);
        });
    }

    function showStatus(message, isError) {
        const statusDiv = document.getElementById('status');
        statusDiv.textContent = message;
        statusDiv.style.color = isError ? 'red' : 'green';
    }

    function clearStatus() {
        const statusDiv = document.getElementById('status');
        statusDiv.textContent = '';
    }

})(); 