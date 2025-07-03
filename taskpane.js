// --- CONFIGURATION ---
// The CONFIG object has been moved to config.js and will be loaded from there.

// Centralized DOM element IDs for easier management and to prevent typos.
const DOM = {
    loading: 'loading',
    errorMessage: 'error-message',
    successMessage: 'success-message',
    // Panels
    requestForm: 'request-form',
    requestListPanel: 'request-list-panel',
    updateFormPanel: 'update-form-panel',
    // New Request Form
    subject: 'subject',
    senderName: 'senderName',
    senderEmail: 'senderEmail',
    sentDate: 'sentDate',
    requestType: 'requestType',
    reportsRequestedGroup: 'reports-requested-group',
    reportsRequested: 'reportsRequested',
    status: 'status',
    notes: 'notes',
    priority: 'priority',
    dueDate: 'dueDate',
    submitBtn: 'submit-btn',
    resetBtn: 'reset-btn',
    // Request List Panel
    requestListContainer: 'request-list-container',
    updateSelectedBtn: 'update-selected-btn',
    createNewBtn: 'create-new-btn',
    refreshListBtn: 'refresh-list-btn',
    // Update Form Panel
    updateRequestType: 'update-request-type',
    updateStatus: 'update-status',
    updateNotes: 'update-notes',
    updatePriority: 'update-priority',
    reportUrlGroup: 'report-url-group',
    reportUrl: 'report-url',
    submitUpdateBtn: 'submit-update-btn',
    backToListBtn: 'back-to-list-btn'
};

// Global state variables
let currentItem;
let currentUser;
let existingRequests = [];

// --- INITIALIZATION ---
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        // Hide all panels initially
        document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
        document.getElementById(DOM.loading).style.display = "block";

        try {
            setupGlobalEventHandlers();
            
            currentUser = Office.context.mailbox.userProfile;
            currentItem = Office.context.mailbox.item;
            
            // FIX: The primary logic flow should be to check for requests first.
            // The result of that check will then determine whether to show the
            // existing requests list or a new, populated form.
            populateDropdowns();
            // Await inside try to catch async errors
            await checkExistingRequests();

        } catch (error) {
            console.error("Initialization error:", error);
            showError("Could not initialize the add-in. Please try again.");
            // FIX: Load data before showing the form as a fallback.
            loadEmailData();
            showPanel(DOM.requestForm);
        }
    }
});

function setupGlobalEventHandlers() {
    // New Request Form
    document.getElementById(DOM.submitBtn).onclick = submitNewRequest;
    document.getElementById(DOM.resetBtn).onclick = resetForm;
    document.getElementById(DOM.requestType).onchange = toggleReportsRequestedField;

    // Request List Panel
    document.getElementById(DOM.updateSelectedBtn).onclick = () => showUpdateForm();
    // Modified to load email data when creating a new request from the list
    document.getElementById(DOM.createNewBtn).onclick = () => {
        loadEmailData();
        showPanel(DOM.requestForm);
    };
    document.getElementById(DOM.refreshListBtn).onclick = checkExistingRequests;

    // Update Form Panel
    document.getElementById(DOM.submitUpdateBtn).onclick = submitUpdate;
    document.getElementById(DOM.backToListBtn).onclick = () => showRequestsPanel(existingRequests);
    document.getElementById(DOM.updateStatus).onchange = toggleReportUrlField;
}

// --- UI TOGGLES ---

function toggleReportUrlField() {
    const selectedRequest = getSelectedRequest();
    if (!selectedRequest) return;

    const requestType = selectedRequest.RequestType;
    const status = document.getElementById(DOM.updateStatus).value;
    const reportUrlGroup = document.getElementById(DOM.reportUrlGroup);
    const reportUrlInput = document.getElementById(DOM.reportUrl);

    // Show the Report Link field only when the status is 'Completed' AND request type is 'Compliance Request'
    if (status === 'Completed' && requestType === 'Compliance Request') {
        reportUrlGroup.style.display = 'block';
        reportUrlInput.setAttribute('required', 'true');
    } else {
        // Hide the field and remove the required attribute for all other cases
        reportUrlGroup.style.display = 'none';
        reportUrlInput.removeAttribute('required');
    }
}

function toggleReportsRequestedField() {
    const requestType = document.getElementById(DOM.requestType).value;
    const reportsGroup = document.getElementById(DOM.reportsRequestedGroup);
    const reportsInput = document.getElementById(DOM.reportsRequested);

    if (requestType === "Compliance Request") {
        reportsGroup.style.display = "block";
        reportsInput.value = 1;
        reportsInput.disabled = false;
    } else {
        reportsGroup.style.display = "none";
        reportsInput.value = ""; // Clear value when hidden
        reportsInput.disabled = true;
    }
}

// --- DATA LOADING AND CHECKING ---

function populateDropdowns() {
    const requestTypes = [
        "Compliance Request",
        "Contract Extension",
        "Contract Termination",
        "Oracle Modification",
        "Deal Reporting",
        "Process Improvement",
        "Data Request",
        "Other"
    ];
    const statuses = [
        "New",
        "In Progress",
        "On Hold",
        "Completed",
        "Cancelled"
    ];

    const requestTypeDropdown = document.getElementById(DOM.requestType);
    const statusDropdown = document.getElementById(DOM.status);
    const updateStatusDropdown = document.getElementById(DOM.updateStatus);

    // Always clear all options before repopulating to prevent duplicates
    while (requestTypeDropdown.firstChild) {
        requestTypeDropdown.removeChild(requestTypeDropdown.firstChild);
    }
    while (statusDropdown.firstChild) {
        statusDropdown.removeChild(statusDropdown.firstChild);
    }
    while (updateStatusDropdown.firstChild) {
        updateStatusDropdown.removeChild(updateStatusDropdown.firstChild);
    }

    // Add a default placeholder
    const requestTypePlaceholder = document.createElement("option");
    requestTypePlaceholder.value = "";
    requestTypePlaceholder.textContent = "Select Request Type...";
    requestTypeDropdown.appendChild(requestTypePlaceholder);

    const statusPlaceholder = document.createElement("option");
    statusPlaceholder.value = "";
    statusPlaceholder.textContent = "Select Status...";
    statusDropdown.appendChild(statusPlaceholder);

    requestTypes.forEach(type => {
        const option = document.createElement("option");
        option.value = type;
        option.textContent = type;
        requestTypeDropdown.appendChild(option);
    });

    statuses.forEach(status => {
        const option = document.createElement("option");
        option.value = status;
        option.textContent = status;
        
        // Add the options to both the new request form and the update form
        statusDropdown.appendChild(option.cloneNode(true));
        updateStatusDropdown.appendChild(option.cloneNode(true));
    });
}

function loadEmailData() {
    try {
        if (!currentItem) {
            showError("Cannot access email data.");
            return;
        }
        document.getElementById(DOM.subject).value = currentItem.subject || "(No subject)";
        const sender = currentItem.from;
        document.getElementById(DOM.senderName).value = sender ? sender.displayName : "(Unknown sender)";
        document.getElementById(DOM.senderEmail).value = sender ? sender.emailAddress : "(Unknown email)";
        document.getElementById(DOM.sentDate).value = currentItem.dateTimeCreated ? formatDate(currentItem.dateTimeCreated, true) : "(Unknown date)";
    } catch (error) {
        showError("Error loading email data: " + error.message);
    }
}

async function checkExistingRequests() {
    if (!currentItem) {
        showError("Cannot access email data.");
        showLoading(false);
        return;
    }
    const conversationId = currentItem.conversationId;
    console.log("CRITICAL_DEBUG: Conversation ID for this email item is:", conversationId);
    console.log("CRITICAL_DEBUG: Conversation ID for this email item is:", conversationId);

    try {
        // --- Step 1: Primary lookup by Conversation ID ---
        let lookupPayload = { conversationId: conversationId };
        const response = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lookupPayload)
        });

        if (!response.ok) throw new Error(`HTTP error ${response.status}`);
        
        const potentialRequests = await response.json();

        if (potentialRequests && potentialRequests.length > 0) {
            // Success! Found a match with the Conversation ID.
            console.log("Found existing requests by Conversation ID:", potentialRequests);
            existingRequests = potentialRequests;
            showRequestsPanel(existingRequests);
        } else {
            // No match found by Conversation ID. Show the new request form.
            console.log("No existing requests found for this conversation. Showing new request form.");
            loadEmailData();
            showPanel(DOM.requestForm);
        }
    } catch (error) {
        console.error("Error checking for existing requests:", error);
        showError("Could not check for existing requests. Please try again.");
        // Fallback to the new request form on any error during lookup.
        loadEmailData();
        showPanel(DOM.requestForm);
    } finally {
        showLoading(false);
    }
}

// --- UI NAVIGATION AND PANEL MANAGEMENT ---

function showPanel(panelId, clear=true) {
    document.getElementById(DOM.loading).style.display = 'none';
    document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
    const panel = document.getElementById(panelId);
    if (panel) {
        panel.style.display = 'block';
        
        // Call toggleReportsRequestedField when showing the request form
        // to ensure the Reports Requested field is properly initialized
        if (panelId === DOM.requestForm) {
            toggleReportsRequestedField();
        }
    }
    // Only clear messages if the 'clear' flag is true.
    // This prevents the success toast from being hidden prematurely.
    if (clear) {
        clearMessages();
    }
}

function showRequestsPanel(requests) {
    const container = document.getElementById(DOM.requestListContainer);
    container.innerHTML = ''; // Clear previous list
    
    // Clear any lingering error messages when showing the list
    clearMessages();
    
    console.log("ENTERING showRequestsPanel with:", requests);
    
    // Super defensive check
    if (!requests) {
        console.error("requests is null or undefined");
        requests = [];
    }
    
    // Force convert to array if not already
    if (!Array.isArray(requests)) {
        console.error("requests is not an array, converting:", requests);
        requests = [requests]; // Wrap in array
    }
    
    // Process the requests to ensure the SharePoint complex objects are properly handled
    const processedRequests = requests.map(req => {
        const processed = { ...req };
        
        // Process ALL fields that might be complex objects
        Object.keys(req).forEach(key => {
            if (req[key] && typeof req[key] === 'object' && req[key].Value !== undefined) {
                console.log(`Converting complex object in field ${key}:`, req[key]);
                processed[key] = req[key].Value;
            }
        });
        
        // Double-check the critical fields
        if (req.RequestStatus && typeof req.RequestStatus === 'object') {
            console.log("Processing RequestStatus:", req.RequestStatus);
            processed.RequestStatus = req.RequestStatus.Value || "New";
        }
        
        if (req.Priority && typeof req.Priority === 'object') {
            console.log("Processing Priority:", req.Priority);
            processed.Priority = req.Priority.Value || "Medium";
        }
        
        if (req.RequestType && typeof req.RequestType === 'object') {
            console.log("Processing RequestType:", req.RequestType);
            processed.RequestType = req.RequestType.Value || "Unknown";
        }
        
        return processed;
    });
    
    // Replace the original requests with the processed ones
    existingRequests = processedRequests;
    
    if (processedRequests.length > 0) {
        processedRequests.forEach((req, index) => {
            try {
                // Create the element
                const itemDiv = document.createElement('div');
                itemDiv.className = 'request-list-item';
                
                // Get the request ID
                const reqId = (req && (req.ID !== undefined || req.Id !== undefined)) ? 
                    (req.ID !== undefined ? req.ID : req.Id) : `unknown-${index}`;
                    
                const uniqueId = `req-${reqId}-${Math.random().toString(36).substring(2, 8)}`;
                
                // Ensure we have string values for display
                const requestTypeText = String(req.RequestType || "Unknown");
                const statusText = String(req.RequestStatus || "New");
                const statusClass = statusText.toLowerCase().replace(/\s+/g, '-');
                const priorityText = String(req.Priority || "Medium");
                
                // Format date safely
                const trackedDate = (req && req.TrackedDate) ? formatDate(req.TrackedDate) : 'Unknown Date';
                
                // Build the HTML with safe values and new status badges
                itemDiv.innerHTML = `
                    <input type="radio" name="requestSelection" value="${reqId}" id="${uniqueId}">
                    <label for="${uniqueId}" class="request-list-item-details">
                        <div>
                            <strong>${requestTypeText}</strong>
                        </div>
                        <div>
                            <span class="status-badge status-${statusClass}">${statusText}</span>
                        </div>
                        <div>
                            <small>Created: ${trackedDate} | Priority: ${priorityText}</small>
                        </div>
                    </label>
                `;
                container.appendChild(itemDiv);
            } catch (err) {
                console.error(`Critical error processing request #${index}:`, err, req);
                const errorDiv = document.createElement('div');
                errorDiv.className = 'request-list-item error';
                errorDiv.textContent = `Error displaying request #${index}: ${err.message}`;
                errorDiv.style.color = "#cc0000";
                container.appendChild(errorDiv);
            }
        });
        
        showPanel(DOM.requestListPanel, false);
    } else {
        loadEmailData();
        showPanel(DOM.requestForm);
    }
}

function showUpdateForm() {
    clearMessages(); // Clear any previous error messages
    
    const selectedRequest = getSelectedRequest();
    if (!selectedRequest) {
        showError("Please select a request to update.");
        return;
    }
    
    console.log("Found selected request:", selectedRequest);
    
    // Display the request type
    document.getElementById(DOM.updateRequestType).textContent = selectedRequest.RequestType || "Unknown";

    // FIX: Handle RequestStatus properly for the dropdown
    let statusValue = "";
    if (selectedRequest.RequestStatus) {
        if (typeof selectedRequest.RequestStatus === 'string') {
            statusValue = selectedRequest.RequestStatus;
        } else if (typeof selectedRequest.RequestStatus === 'object' && selectedRequest.RequestStatus.Value) {
            statusValue = selectedRequest.RequestStatus.Value;
        }
    }
    
    // Pre-populate the update form
    document.getElementById(DOM.updateStatus).value = statusValue;
    document.getElementById(DOM.updateNotes).value = selectedRequest.Notes || '';
    
    // Add this code to set the priority value
    let priorityValue = "Medium"; // Default value
    if (selectedRequest.Priority) {
        if (typeof selectedRequest.Priority === 'string') {
            priorityValue = selectedRequest.Priority;
        } else if (typeof selectedRequest.Priority === 'object' && selectedRequest.Priority.Value) {
            priorityValue = selectedRequest.Priority.Value;
        }
    }
    document.getElementById(DOM.updatePriority).value = priorityValue;
    
    // FIX: Handle ReportLink from SharePoint format
    let reportUrl = "";
    if (selectedRequest.ReportLink) {
        if (typeof selectedRequest.ReportLink === 'string') {
            reportUrl = selectedRequest.ReportLink;
        } else if (typeof selectedRequest.ReportLink === 'object' && selectedRequest.ReportLink.Url) {
            reportUrl = selectedRequest.ReportLink.Url;
        }
    }
    
    document.getElementById(DOM.reportUrl).value = reportUrl;
    toggleReportUrlField(); // Show/hide report URL field based on status

    showPanel(DOM.updateFormPanel);
}

// --- FORM SUBMISSION LOGIC (CREATE & UPDATE) ---

async function submitNewRequest() {
    // Validation
    const requestType = document.getElementById(DOM.requestType).value;
    const status = document.getElementById(DOM.status).value;
    if (!requestType || !status) {
        showError("Request Type and Status are required.");
        return;
    }

    showLoading(true, "Submitting new request...");

    try {
        const emailBody = await getBodyAsText();
        
        const payload = {
            subject: document.getElementById(DOM.subject).value,
            senderName: document.getElementById(DOM.senderName).value,
            senderEmail: document.getElementById(DOM.senderEmail).value,
            sentDate: currentItem.dateTimeCreated ? new Date(currentItem.dateTimeCreated).toISOString() : null,
            requestType: requestType,
            reportsRequested: parseInt(document.getElementById(DOM.reportsRequested).value, 10) || null,
            requestStatus: status,
            notes: document.getElementById(DOM.notes).value || "",
            priority: document.getElementById(DOM.priority).value,
            dueDate: document.getElementById(DOM.dueDate).value || null,
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
            Id: "new-" + Date.now(), // Placeholder ID
            RequestType: payload.requestType,
            RequestStatus: payload.requestStatus,
            TrackedDate: payload.trackedDate,
            Priority: payload.priority
        };
        existingRequests.push(newRequestData);
        
        // Show the success message and immediately switch to the list view.
        showSuccess("Request created successfully!");
        showRequestsPanel(existingRequests);
        
        // No need to reset form here, as we are leaving the form view.

        // Refresh the list from the server after a delay to get the real data.
        setTimeout(checkExistingRequests, 2500);

    } catch (error) {
        console.error("Submit error details:", error);
        // Display the specific error message to the user.
        showError(error.message);
        // Keep the form visible so the user can try again without re-entering data.
        showPanel(DOM.requestForm);
    }
}

async function submitUpdate() {
    const selectedRequest = getSelectedRequest();
    if (!selectedRequest) {
        showError("Please select a request to update.");
        return;
    }
    const selectedId = selectedRequest.ID || selectedRequest.Id;

    // Ensure selectedId is a valid number before proceeding
    if (!selectedId || isNaN(Number(selectedId))) {
        showError("Invalid request ID. Cannot update this request.");
        return;
    }

    const newStatus = document.getElementById(DOM.updateStatus).value;
    const reportUrl = document.getElementById(DOM.reportUrl).value;
    const requestType = selectedRequest.RequestType;
    const priority = document.getElementById(DOM.updatePriority).value;

    // VALIDATION: Enforce Report Link requirement before submitting.
    if (requestType === 'Compliance Request' && newStatus === 'Completed' && !reportUrl) {
        showError('A Report Link is required to mark a Compliance Request as Completed.');
        document.getElementById(DOM.reportUrl).focus(); // Focus the input for user convenience
        return; // Stop the submission
    }

    showLoading(true, "Submitting update...");

    try {
        const notesValue = document.getElementById(DOM.updateNotes).value.trim();
        const payload = {
            requestId: parseInt(selectedId, 10),
            requestStatus: newStatus,
            // Add priority to payload
            priority: priority,
            updatedBy: currentUser ? currentUser.emailAddress : "Unknown User"
        };

        // Only include the reportUrl if it has a value.
        if (reportUrl) {
            payload.reportUrl = reportUrl;
        }
        
        // If new notes are added, prepend them to the existing notes to create a log.
        if (notesValue) {
            const timestamp = new Date().toLocaleString();
            const user = currentUser ? currentUser.displayName : "Unknown User";
            const existingNotes = selectedRequest.Notes || "";
            
            const newNoteEntry = `--- Note added by ${user} on ${timestamp} ---\n${notesValue}`;
            
            // Combine the new entry with previous notes.
            payload.notes = newNoteEntry + (existingNotes ? `\n\n${existingNotes}` : "");
        }

        console.log("Update payload:", payload);

        // Using the correct update flow URL
        const response = await fetch(CONFIG.REQUEST_UPDATE_URL, { 
            method: 'POST', 
            headers: { 'Content-Type': 'application/json' }, 
            body: JSON.stringify(payload) 
        });

        console.log("Update response status:", response.status);

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Update error response:", errorText);
            throw new Error(`HTTP Error ${response.status}: ${errorText}`);
        }

        showSuccess("Request updated successfully!");
        // Immediately refresh the list to show the update.
        await checkExistingRequests();

    } catch (error) {
        console.error("Update submission error:", error);
        showError(error.message);
        // Show the update form again on error so the user can retry.
        showPanel(DOM.updateFormPanel);
    } finally {
        showLoading(false);
    }
}

// --- HELPER FUNCTIONS ---

/**
 * Finds the selected request from the radio buttons in the list.
 * @returns {object | null} The request object from existingRequests or null if not found.
 */
function getSelectedRequest() {
    const selectedRadio = document.querySelector('input[name="requestSelection"]:checked');
    if (!selectedRadio) {
        return null;
    }
    const selectedId = selectedRadio.value;

    // Find the request using a flexible comparison for string/number IDs.
    return existingRequests.find(r => {
        const rId = r.ID !== undefined ? r.ID : r.Id;
        return String(rId) === String(selectedId);
    }) || null;
}

function getBodyAsText() {
    // FIX: Return a promise that can be rejected on failure.
    return new Promise((resolve, reject) => {
        currentItem.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                console.error("Failed to get email body:", result.error);
                // Reject the promise so the main catch block can handle it.
                reject(new Error("Failed to get email body: " + result.error.message));
            }
        });
    });
}

function resetForm() {
    document.getElementById(DOM.requestForm).reset();
    // The default value for priority is now set in the HTML markup.
    // FIX: Reload email data to ensure form is correctly populated.
    loadEmailData();
    // Ensure the Reports Requested field is properly set after reset
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
    const loading = document.getElementById(DOM.loading);
    if (show) {
        loading.textContent = message;
        loading.style.display = 'block';
        document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
    } else {
        loading.style.display = 'none';
    }
}

function showError(message) {
    const errorElement = document.getElementById(DOM.errorMessage);
    errorElement.textContent = message;
    errorElement.style.display = "block";
    errorElement.style.color = "white";
    errorElement.style.padding = "10px";
    errorElement.style.border = "1px solid #d32f2f";
    errorElement.style.borderRadius = "4px";
    errorElement.style.marginBottom = "15px";
    errorElement.style.backgroundColor = "#d32f2f";
    showLoading(false);
    
    // Log the error to console for debugging
    console.error("ERROR SHOWN TO USER:", message);
    
    // Auto-dismiss error after 6 seconds
    setTimeout(clearMessages, 6000);
}

function showSuccess(message) {
    const successElement = document.getElementById(DOM.successMessage);
    successElement.textContent = message;
    successElement.style.display = "block";
    successElement.style.color = "white";
    successElement.style.backgroundColor = "#FF6B00"; // Stryker orange
    successElement.style.padding = "10px";
    successElement.style.border = "1px solid #FF6B00";
    successElement.style.borderRadius = "4px";
    successElement.style.marginBottom = "15px";
    
    setTimeout(clearMessages, 6000);
}

function clearMessages() {
    const errorElem = document.getElementById(DOM.errorMessage);
    const successElem = document.getElementById(DOM.successMessage);
    if (errorElem) errorElem.style.display = "none";
    if (successElem) successElem.style.display = "none";
}
