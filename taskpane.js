/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("run").onclick = run;

        // Load saved URL from localStorage
        const savedUrl = localStorage.getItem("sharepointBaseUrl");
        if (savedUrl) {
            document.getElementById("sharepoint-url").value = savedUrl;
        }
    }
});

async function run() {
    try {
        await Word.run(async (context) => {
            const sharePointBaseInput = document.getElementById("sharepoint-url").value;

            // Basic validation
            if (!sharePointBaseInput) {
                setStatus("Please enter a SharePoint Base URL.", "error");
                return;
            }

            // Ensure trailing slash
            const sharePointBase = sharePointBaseInput.endsWith('/') ? sharePointBaseInput : sharePointBaseInput + '/';

            // Save to localStorage
            localStorage.setItem("sharepointBaseUrl", sharePointBase);

            setStatus("Scanning links...", "");

            const hyperlinks = context.document.body.getRange().hyperlinkInfos;
            context.load(hyperlinks);
            await context.sync();

            let updateCount = 0;

            for (let i = 0; i < hyperlinks.items.length; i++) {
                let currentAddress = hyperlinks.items[i].address;

                // Check if it's a local/relative link (not starting with http/https)
                // Note: Some local links might start with file://, checking for http guarantees we strip that too if desired, 
                // but the original logic was specific about !startsWith("http")
                if (currentAddress && !currentAddress.toLowerCase().startsWith("http")) {

                    // Logic from original request:
                    // Extract filename from path (handles both backslash and forward slash)
                    let fileName = currentAddress.split('\\').pop().split('/').pop();

                    // Clean filename and encode
                    let cleanFileName = encodeURIComponent(fileName).replace(/'/g, "%27");

                    let newAddress = sharePointBase + cleanFileName;
                    let range = hyperlinks.items[i].getRange();
                    range.hyperlink = newAddress;
                    updateCount++;
                }
            }

            await context.sync();

            if (updateCount > 0) {
                setStatus(`Successfully updated ${updateCount} links!`, "success");
            } else {
                setStatus("No matching links found to update.", "");
            }
        });
    } catch (error) {
        console.error(error);
        setStatus("Error: " + error.message, "error");
    }
}

function setStatus(message, type) {
    const statusElement = document.getElementById("status-message");
    statusElement.innerText = message;
    statusElement.className = "status-message " + type;
}
