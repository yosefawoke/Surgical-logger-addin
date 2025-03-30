'use strict';

(function () {

    // Alert JS 1: Does taskpane.js even start parsing?
    alert("taskpane.js: Script Start");
    console.log("taskpane.js: Script Start");

    Office.onReady(function(info){
        // Alert JS 2: Does Office.onReady fire within taskpane.js?
        alert("taskpane.js: Office.onReady fired! Host: " + info.host);
        console.log("taskpane.js: Office.onReady fired! Host:", info.host);

        try {
             const statusDiv = document.getElementById('status');
             if (statusDiv) {
                // Alert JS 3: Can we write to the status div?
                alert("taskpane.js: Attempting to write to statusDiv...");
                 statusDiv.textContent = "taskpane.js: Office Ready!";
                 statusDiv.style.color = 'blue';
                 console.log("taskpane.js: Wrote to statusDiv.");
             } else {
                 // Alert JS 4: Status div not found from taskpane.js?
                 alert("taskpane.js: statusDiv NOT found!");
                 console.error("taskpane.js: statusDiv NOT found!");
             }
        } catch (error) {
            // Alert JS 5: Error within Office.onReady?
            alert("taskpane.js: Error in Office.onReady: " + error.message);
            console.error("taskpane.js: Error in Office.onReady:", error);
        }
    });

    // Alert JS 6: Did taskpane.js parse to the end?
    alert("taskpane.js: Script End");
    console.log("taskpane.js: Script End");

})(); 
