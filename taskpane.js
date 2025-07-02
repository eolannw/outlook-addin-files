// --- CONFIGURATION ---
// IMPORTANT: These URLs are for the Power Automate flows.
const CONFIG = {
    // Flow for creating a NEW request (your existing flow)
    REQUEST_CREATE_URL: "https://prod-135.westus.logic.azure.com:443/workflows/075b978523814f56951805720dc2da6d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7fsONcoc2c82EmQGTmubH_9PUgrWGRZz833KgAThavg",
    // Flow for LOOKING UP existing requests
    REQUEST_LOOKUP_URL: "https://prod-139.westus.logic.azure.com:443/workflows/939c3e7c315b43b8b12300ea476dbbd2/triggers/manual/paths/invoke?api-version=2016-06-01",
    // Flow for UPDATING an existing request
    REQUEST_UPDATE_URL: "https://prod-188.westus.logic.azure.com:443/workflows/13af96bdb60f4199856014b64e9f3188/triggers/manual/paths/invoke?api-version=2016-06-01"
};

// Global state variables
let currentItem;
let currentUser;
let existingRequests = [];

// --- INITIALIZATION ---
// FIX: Use Office.onReady for modern, reliable initialization.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Hide all panels initially
        document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
        document.getElementById("loading").style.display = "block";

        try {
            currentItem = Office.context.mailbox.item;
            currentUser = Office.context.mailbox.userProfile;

            // Ensure currentItem is valid before proceeding
            if (!currentItem) {
                throw new Error("Cannot access email data. Please select an email.");
            }

            setupGlobalEventHandlers();
            loadEmailData();
            checkExistingRequests(); // This can now safely access currentItem
        } catch (error) {
            showError("Error initializing add-in: " + error.message);
            showPanel('request-form'); // Fallback to new request form
        }
    }
});

function setupGlobalEventHandlers() {
    // New Request Form
    document.getElementById("submit-btn").onclick = submitNewRequest;
    document.getElementById("reset-btn").onclick = resetForm;
    document.getElementById("requestType").onchange = toggleReportsRequestedField;

    // Request List Panel
    document.getElementById("update-selected-btn").onclick = () => showUpdateForm();
    document.getElementById("create-new-btn").onclick = () => showPanel('request-form');
    document.getElementById("refresh-list-btn").onclick = checkExistingRequests;

    // Update Form Panel
    document.getElementById("submit-update-btn").onclick = submitUpdate;
    document.getElementById("back-to-list-btn").onclick = () => showPanel('request-list-panel');
    document.getElementById("update-status").onchange = toggleReportUrlField;
}

// --- DATA LOADING AND CHECKING ---

function loadEmailData() {
    try {
        if (!currentItem) {
            showError("Cannot access email data.");
            return;
        }
        document.getElementById("subject").value = currentItem.subject || "(No subject)";
        const sender = currentItem.from;
        document.getElementById("senderName").value = sender ? sender.displayName : "(Unknown sender)";
        document.getElementById("senderEmail").value = sender ? sender.emailAddress : "(Unknown email)";
        document.getElementById("sentDate").value = currentItem.dateTimeCreated ? formatDate(currentItem.dateTimeCreated, true) : "(Unknown date)";
    } catch (error) {
        showError("Error loading email data: " + error.message);
    }
}

async function checkExistingRequests() {
    showLoading(true, "Checking for existing requests...");
    const conversationId = currentItem.conversationId;

    if (!conversationId) {
        showError("Could not get conversation ID. Showing new request form.");
        showPanel('request-form');
        return;
    }

    if (CONFIG.REQUEST_LOOKUP_URL.includes("PASTE_YOUR")) {
        showError("Request Lookup Flow URL is not configured.");
        showPanel('request-form');
        return;
    }

    try {
        const response = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ conversationId: conversationId })
        });

        if (!response.ok) throw new Error(`HTTP error ${response.status}`);

        existingRequests = await response.json();

        if (existingRequests && existingRequests.length > 0) {
            showRequestsPanel(existingRequests);
        } else {
            showPanel('request-form'); // No requests found, show new form
        }
    } catch (error) {
        console.error("Error checking for existing requests:", error);
        showError("Could not check for existing requests. Please try again or create a new one.");
        showPanel('request-form'); // Fallback to new form on error
    } finally {
        showLoading(false);
    }
}

// --- UI NAVIGATION AND PANEL MANAGEMENT ---

function showPanel(panelId) {
    document.getElementById("loading").style.display = 'none';
    document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
    const panel = document.getElementById(panelId);
    if (panel) {
        panel.style.display = 'block';
    }
    clearMessages();
}

function showRequestsPanel(requests) {
    const container = document.getElementById('request-list-container');
    container.innerHTML = ''; // Clear previous list

    requests.forEach(req => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'request-list-item';
        itemDiv.innerHTML = `
            <input type="radio" name="requestSelection" value="${req.Id}" id="req-${req.Id}">
            <label for="req-${req.Id}" class="request-list-item-details">
                <strong>${req.RequestType}</strong>
                <span class="status-badge status-${req.RequestStatus.toLowerCase().replace(' ', '-')}">${req.RequestStatus}</span>
                <br>
                <small>Created: ${formatDate(req.TrackedDate)} | Priority: ${req.Priority || 'N/A'}</small>
            </label>
        `;
        container.appendChild(itemDiv);
    });

    showPanel('request-list-panel');
}

function showUpdateForm() {
    const selectedId = document.querySelector('input[name="requestSelection"]:checked')?.value;
    if (!selectedId) {
        showError("Please select a request to update.");
        return;
    }

    const selectedRequest = existingRequests.find(r => r.Id == selectedId);
    if (!selectedRequest) {
        showError("Could not find the selected request.");
        return;
    }

    // Pre-populate the update form
    document.getElementById('update-status').value = selectedRequest.RequestStatus;
    document.getElementById('update-notes').value = selectedRequest.Notes || '';
    document.getElementById('report-url').value = selectedRequest.ReportLink ? selectedRequest.ReportLink.Url : '';
    toggleReportUrlField(); // Show/hide report URL field based on status

    showPanel('update-form-panel');
}

// --- FORM SUBMISSION LOGIC (CREATE & UPDATE) ---

async function submitNewRequest() {
    // Validation
    const requestType = document.getElementById("requestType").value;
    const status = document.getElementById("status").value;
    if (!requestType || !status) {
        showError("Request Type and Status are required.");
        return;
    }

    showLoading(true, "Submitting new request...");

    try {
        console.log("Getting email body...");
        const emailBody = await getBodyAsText();
        
        const payload = {
            subject: document.getElementById("subject").value,
            senderName: document.getElementById("senderName").value,
            senderEmail: document.getElementById("senderEmail").value,
            sentDate: currentItem.dateTimeCreated ? currentItem.dateTimeCreated.toISOString() : null,
            requestType: requestType,
            reportsRequested: parseInt(document.getElementById("reportsRequested").value, 10) || null,
            requestStatus: status,
            notes: document.getElementById("notes").value || "",
            priority: document.getElementById("priority").value,
            dueDate: document.getElementById("dueDate").value || null,
            trackedDate: new Date().toISOString(),
            trackedBy: currentUser ? currentUser.emailAddress : "Unknown User",
            conversationId: currentItem.conversationId || "",
            messageId: currentItem.internetMessageId || currentItem.itemId || "",
            emailBody: emailBody || ""
        };
        
        console.log("Submitting with payload:", payload);
        
        // Create a Promise that wraps XMLHttpRequest for better error handling
        const submitRequest = () => {
            return new Promise((resolve, reject) => {
                const xhr = new XMLHttpRequest();
                
                // Log detailed state changes to help diagnose the issue
                xhr.onreadystatechange = function() {
                    console.log(`XHR state changed: ${xhr.readyState}, status: ${xhr.status}`);
                };
                
                xhr.open("POST", CONFIG.REQUEST_CREATE_URL, true);
                xhr.setRequestHeader("Content-Type", "application/json");
                
                xhr.onload = function() {
                    if (xhr.status >= 200 && xhr.status < 300) {
                        resolve({
                            status: xhr.status,
                            statusText: xhr.statusText,
                            response: xhr.responseText
                        });
                    } else {
                        reject({
                            status: xhr.status,
                            statusText: xhr.statusText,
                            response: xhr.responseText
                        });
                    }
                };
                
                xhr.onerror = function() {
                    // This is critical - it will catch network errors
                    console.error("XHR Network Error:", xhr);
                    reject({
                        status: 0,
                        statusText: "Network Error - Could not connect to Power Automate",
                        response: "The request couldn't be completed. This could be due to CORS restrictions, network issues, or an expired Power Automate URL."
                    });
                };
                
                xhr.ontimeout = function() {
                    reject({
                        status: 0,
                        statusText: "Timeout Error",
                        response: "The request timed out. Please try again or check your network connection."
                    });
                };
                
                xhr.timeout = 10000; // 10 seconds timeout
                
                try {
                    xhr.send(JSON.stringify(payload));
                } catch (e) {
                    reject({
                        status: 0,
                        statusText: "Request Error",
                        response: "Error sending request: " + e.message
                    });
                }
            });
        };

        // First, try a test request to make sure we can make HTTP requests at all
        console.log("Making test HTTP request to httpbin.org...");
        try {
            const testResponse = await fetch("https://httpbin.org/get?test=1");
            console.log("Test request succeeded:", await testResponse.text());
        } catch (testError) {
            console.error("Test request failed:", testError);
            throw new Error("Network connection test failed. Please check your internet connection.");
        }
        
        // Now try the actual request
        console.log("Making actual request to Power Automate...");
        const result = await submitRequest();
        console.log("XHR Success:", result);
        
        // Try to parse the response as JSON
        let responseData;
        try {
            responseData = JSON.parse(result.response);
            console.log("Response data:", responseData);
        } catch (e) {
            console.log("Response is not valid JSON:", result.response);
        }
        
        // Show success and reset the form
        showSuccess("Request created successfully!");
        resetForm();
        setTimeout(checkExistingRequests, 1500);

    } catch (error) {
        console.error("Submit error:", error);
        let errorMessage = "Error submitting request: ";
        
        if (error.status === 0) {
            errorMessage += "Network error - could not connect to Power Automate. ";
            errorMessage += "This might be due to CORS restrictions or an expired URL.";
        } else if (error.status === 409) {
            errorMessage += "This email has already been tracked for this Request Type.";
        } else if (error.response) {
            errorMessage += error.response;
        } else {
            errorMessage += error.message || "Unknown error";
        }
        
        showError(errorMessage);
        // Keep the form visible with the entered data so the user can try again
        showPanel('request-form');
    }
}

async function submitUpdate() {
    const selectedId = document.querySelector('input[name="requestSelection"]:checked')?.value;
    const newStatus = document.getElementById('update-status').value;
    const reportUrl = document.getElementById('report-url').value;

    if (!newStatus) {
        showError("Please select a status.");
        return;
    }
    if (newStatus === 'Completed' && !reportUrl) {
        showError("A Report Link is required for 'Completed' status.");
        return;
    }

    showLoading(true, "Submitting update...");

    try {
        const payload = {
            requestId: parseInt(selectedId, 10),
            requestStatus: newStatus,
            notes: document.getElementById('update-notes').value || "",
            reportUrl: reportUrl || "",
            updatedBy: currentUser ? currentUser.emailAddress : "Unknown User"
        };

        const response = await fetch(CONFIG.REQUEST_UPDATE_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });

        if (!response.ok) throw new Error(`HTTP Error ${response.status}`);

        showSuccess("Request updated successfully!");
        setTimeout(checkExistingRequests, 1500); // Refresh list after success

    } catch (error) {
        showError(error.message);
        // FIX: Show the update form again on error so the user can retry.
        showPanel('update-form-panel');
    }
}

// --- HELPER FUNCTIONS ---

function getBodyAsText() {
    // FIX: Return a promise that can be rejected on failure.
    return new Promise((resolve, reject) => {
        currentItem.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                console.error("Failed to get email body:", result.error);
                // Reject the promise so the main catch block can handle it.
                reject(new Error("Could not retrieve email body."));
            }
        });
    });
}

function toggleReportsRequestedField() {
    const requestType = document.getElementById("requestType").value;
    const reportsGroup = document.getElementById("reports-requested-group");
    reportsGroup.style.display = (requestType === "Compliance Request") ? "block" : "none";
    if (requestType === "Compliance Request") {
        document.getElementById("reportsRequested").value = "1";
    }
}

function toggleReportUrlField() {
    const status = document.getElementById('update-status').value;
    const reportGroup = document.getElementById('report-url-group');
    reportGroup.style.display = (status === 'Completed') ? 'block' : 'none';
}

function resetForm() {
    document.getElementById("request-form").reset();
    document.getElementById("priority").value = "Medium";
    toggleReportsRequestedField();
    clearMessages();
}

function formatDate(dateString, includeTime = false) {
    if (!dateString) return "";
    try {
        const date = new Date(dateString);
        if (includeTime) return date.toLocaleString();
        return date.toLocaleDateString();
    } catch (e) {
        return dateString;
    }
}

function showLoading(show, message = "Loading...") {
    const loading = document.getElementById("loading");
    if (show) {
        loading.textContent = message;
        loading.style.display = 'block';
        document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
    } else {
        loading.style.display = 'none';
    }
}

function showError(message) {
    const errorElement = document.getElementById("error-message");
    errorElement.textContent = message;
    errorElement.style.display = "block";
    errorElement.style.color = "red";
    errorElement.style.padding = "10px";
    errorElement.style.border = "1px solid red";
    errorElement.style.borderRadius = "4px";
    errorElement.style.marginBottom = "15px";
    showLoading(false);
    
    // Log the error to console for debugging
    console.error("ERROR SHOWN TO USER:", message);
}

function showSuccess(message) {
    const successElement = document.getElementById("success-message");
    successElement.textContent = message;
    successElement.style.display = "block";
    setTimeout(clearMessages, 4000);
}

function clearMessages() {
    document.getElementById("error-message").style.display = "none";
    document.getElementById("success-message").style.display = "none";
}
