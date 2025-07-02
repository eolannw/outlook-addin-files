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
        // FIX: Using the correct lookup flow URL
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
            // FIX: The payload was missing the required 'assignedTo' field.
            // The schema also includes 'trackedBy', so we will send both for completeness.
            assignedTo: currentUser ? currentUser.emailAddress : "Unknown User",
            trackedBy: currentUser ? currentUser.emailAddress : "Unknown User",
            conversationId: currentItem.conversationId || "",
            messageId: currentItem.internetMessageId || currentItem.itemId || "",
            emailBody: emailBody || ""
        };
        
        console.log("Submitting with corrected payload:", payload);
        
        // REFACTOR: Using fetch API directly for cleaner code and better error handling.
        const response = await fetch(CONFIG.REQUEST_CREATE_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        console.log("Response received from Power Automate:", response);

        // Check for non-successful responses and provide detailed error info.
        if (!response.ok) {
            const errorBody = await response.text();
            console.error("Power Automate Error Body:", errorBody);
            // Throw a detailed error that will be displayed to the user.
            throw new Error(`Submission failed. Status: ${response.status}. Details: ${errorBody}`);
        }
        
        // FIX: Optimistically update the UI to avoid race conditions.
        // Add the new request to our local array.
        const newRequestData = {
            Id: "new", // Placeholder ID
            RequestType: payload.requestType,
            RequestStatus: payload.requestStatus,
            TrackedDate: payload.trackedDate,
            Priority: payload.priority
        };
        existingRequests.push(newRequestData);
        
        // Show the success message and immediately switch to the list view.
        showSuccess("Request created successfully!");
        showRequestsPanel(existingRequests);
        resetForm();

        // Refresh the list from the server after a delay to get the real data.
        setTimeout(checkExistingRequests, 2500);

    } catch (error) {
        console.error("Submit error details:", error);
        // Display the specific error message to the user.
        showError(error.message);
        // Keep the form visible so the user can try again without re-entering data.
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

        // FIX: Using the correct update flow URL
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
    // FIX: Do not clear messages here, as it hides the success toast.
    // clearMessages(); 
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
