// --- CONFIGURATION ---
// The CONFIG object has been moved to config.js and will be loaded from there.

// Global state variables
let currentItem;
let currentUser;
let existingRequests = [];

// --- INITIALIZATION ---
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        // Hide all panels initially
        document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
        document.getElementById("loading").style.display = "block";

        try {
            setupGlobalEventHandlers();
            
            currentUser = Office.context.mailbox.userProfile;
            currentItem = Office.context.mailbox.item;
            
            // FIX: The primary logic flow should be to check for requests first.
            // The result of that check will then determine whether to show the
            // existing requests list or a new, populated form.
            populateDropdowns();
            await checkExistingRequests();

        } catch (error) {
            console.error("Initialization error:", error);
            showError("Could not initialize the add-in. Please try again.");
            // FIX: Load data before showing the form as a fallback.
            loadEmailData();
            showPanel('request-form');
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
    document.getElementById("back-to-list-btn").onclick = () => showRequestsPanel(existingRequests);
    document.getElementById("update-status").onchange = toggleReportUrlField;
}

// --- UI TOGGLES ---

function toggleReportUrlField() {
    const selectedRadio = document.querySelector('input[name="requestSelection"]:checked');
    if (!selectedRadio) return;

    const selectedId = selectedRadio.value;
    const selectedRequest = existingRequests.find(r => {
        const rId = r.ID !== undefined ? r.ID : r.Id;
        return String(rId) === String(selectedId);
    });

    if (!selectedRequest) return;

    const requestType = selectedRequest.RequestType;
    const status = document.getElementById('update-status').value;
    const reportUrlGroup = document.getElementById('report-url-group');
    const reportUrlInput = document.getElementById('report-url');

    // Show the Report Link field only when the status is 'Completed'.
    if (status === 'Completed') {
        reportUrlGroup.style.display = 'block';
        // It's only REQUIRED if it's a Compliance Request.
        if (requestType === 'Compliance Request') {
            reportUrlInput.setAttribute('required', 'true');
        } else {
            reportUrlInput.removeAttribute('required');
        }
    } else {
        // Hide the field and remove the required attribute for all other statuses.
        reportUrlGroup.style.display = 'none';
        reportUrlInput.removeAttribute('required');
    }
}

function toggleReportsRequestedField() {
    const requestType = document.getElementById("requestType").value;
    const reportsGroup = document.getElementById("reports-requested-group");
    if (requestType === "Compliance Request") {
        reportsGroup.style.display = "block";
    } else {
        reportsGroup.style.display = "none";
    }
}

// --- DATA LOADING AND CHECKING ---

function populateDropdowns() {
    const requestTypes = [
        "Compliance Request",
        "Contract Extension",
        "Contract Termination",
        "Deal Reporting",
        "Other"
    ];
    const statuses = [
        "New",
        "In Progress",
        "On Hold",
        "Completed",
        "Cancelled"
    ];

    const requestTypeDropdown = document.getElementById("requestType");
    const statusDropdown = document.getElementById("status");
    const updateStatusDropdown = document.getElementById("update-status");

    // Clear existing options and add a default placeholder
    requestTypeDropdown.innerHTML = '<option value="">Select Request Type...</option>';
    statusDropdown.innerHTML = '<option value="">Select Status...</option>';
    updateStatusDropdown.innerHTML = ''; // No placeholder for the update form

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
    
    console.log("Looking up requests for conversation ID:", conversationId);

    if (!conversationId) {
        showError("Could not get conversation ID. Showing new request form.");
        loadEmailData();
        showPanel('request-form');
        return;
    }

    try {
        console.log("Calling Power Automate with URL:", CONFIG.REQUEST_LOOKUP_URL);
        
        const response = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ conversationId: conversationId })
        });

        console.log("Response status:", response.status, response.statusText);

        if (!response.ok) {
            let errorText = "";
            try {
                errorText = await response.text();
            } catch (e) {
                errorText = "Could not read error details";
            }
            console.error("Error response body:", errorText);
            throw new Error(`HTTP error ${response.status}: ${errorText}`);
        }

        const responseText = await response.text();
        console.log("Raw response:", responseText);
        
        try {
            console.log("Response type:", typeof responseText);
            console.log("Response length:", responseText.length);
            console.log("First 100 chars:", responseText.substring(0, 100));
            
            const rawRequests = responseText ? JSON.parse(responseText) : [];
            
            // ENHANCED PARSING: Deeply process the SharePoint complex objects
            existingRequests = rawRequests.map(req => {
                const processed = { ...req };
                
                // Process ALL fields that might be complex objects
                Object.keys(req).forEach(key => {
                    if (req[key] && typeof req[key] === 'object' && req[key].Value !== undefined) {
                        console.log(`Converting complex object in field ${key}:`, req[key]);
                        processed[key] = req[key].Value;
                    }
                });
                
                // Double-check the critical fields
                if (processed.RequestStatus && typeof processed.RequestStatus === 'object') {
                    console.log("Force converting RequestStatus:", processed.RequestStatus);
                    processed.RequestStatus = processed.RequestStatus.Value || String(processed.RequestStatus);
                }
                
                if (processed.Priority && typeof processed.Priority === 'object') {
                    console.log("Force converting Priority:", processed.Priority);
                    processed.Priority = processed.Priority.Value || String(processed.Priority);
                }
                
                if (processed.RequestType && typeof processed.RequestType === 'object') {
                    console.log("Force converting RequestType:", processed.RequestType);
                    processed.RequestType = processed.RequestType.Value || String(processed.RequestType);
                }
                
                return processed;
            });
            
            console.log("Processed requests:", existingRequests);
            
            if (existingRequests && existingRequests.length > 0) {
                showRequestsPanel(existingRequests);
            } else {
                loadEmailData();
                showPanel('request-form');
            }
        } catch (parseError) {
            console.error("Failed to parse or process response:", parseError);
            throw parseError;
        }
    } catch (error) {
        console.error("Error checking for existing requests:", error);
        showError("Could not check for existing requests. Please try again.");
        loadEmailData();
        showPanel('request-form');
    } finally {
        showLoading(false);
    }
}

// --- UI NAVIGATION AND PANEL MANAGEMENT ---

function showPanel(panelId, clear=true) {
    document.getElementById("loading").style.display = 'none';
    document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
    const panel = document.getElementById(panelId);
    if (panel) {
        panel.style.display = 'block';
    }
    // FIX: Only clear messages if the 'clear' flag is true.
    // This prevents the success toast from being hidden prematurely.
    if (clear) {
        clearMessages();
    }
}

function showRequestsPanel(requests) {
    const container = document.getElementById('request-list-container');
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
        
        showPanel('request-list-panel', false);
    } else {
        loadEmailData();
        showPanel('request-form');
    }
}

function showUpdateForm() {
    clearMessages(); // Clear any previous error messages
    
    const selectedRadio = document.querySelector('input[name="requestSelection"]:checked');
    if (!selectedRadio) {
        showError("Please select a request to update.");
        return;
    }
    
    const selectedId = selectedRadio.value;
    console.log("Selected ID for update:", selectedId, "type:", typeof selectedId);
    
    // CRITICAL DEBUG: Dump all the IDs in the existingRequests array for comparison
    console.log("All request IDs in memory:", existingRequests.map(r => {
        const id = r.ID !== undefined ? r.ID : r.Id;
        return {id: id, type: typeof id};
    }));
    
    // FIX: Use a more flexible comparison that handles number vs string issues
    const selectedRequest = existingRequests.find(r => {
        // Get ID regardless of case (ID or Id)
        const rId = r.ID !== undefined ? r.ID : r.Id;
        
        // Debug output for troubleshooting
        console.log(`Comparing request ID ${rId} (${typeof rId}) with selected ${selectedId} (${typeof selectedId})`);
        
        // Convert both to strings for comparison (handles both numeric and string IDs)
        return String(rId) === String(selectedId);
    });
    
    console.log("Found selected request:", selectedRequest);
    
    if (!selectedRequest) {
        showError("Could not find the selected request.");
        return;
    }

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
    document.getElementById('update-status').value = statusValue;
    document.getElementById('update-notes').value = selectedRequest.Notes || '';
    
    // FIX: Handle ReportLink from SharePoint format
    let reportUrl = "";
    if (selectedRequest.ReportLink) {
        if (typeof selectedRequest.ReportLink === 'string') {
            reportUrl = selectedRequest.ReportLink;
        } else if (typeof selectedRequest.ReportLink === 'object' && selectedRequest.ReportLink.Url) {
            reportUrl = selectedRequest.ReportLink.Url;
        }
    }
    
    document.getElementById('report-url').value = reportUrl;
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
        showPanel('request-form');
    }
}

async function submitUpdate() {
    const selectedRadio = document.querySelector('input[name="requestSelection"]:checked');
    if (!selectedRadio) {
        showError("Please select a request to update.");
        return;
    }
    const selectedId = selectedRadio.value;

    const selectedRequest = existingRequests.find(r => {
        const rId = r.ID !== undefined ? r.ID : r.Id;
        return String(rId) === String(selectedId);
    });

    if (!selectedRequest) {
        showError("Could not find the selected request.");
        return;
    }

    const newStatus = document.getElementById('update-status').value;
    const reportUrl = document.getElementById('report-url').value;
    const requestType = selectedRequest.RequestType;

    // VALIDATION: Enforce Report Link requirement before submitting.
    if (requestType === 'Compliance Request' && newStatus === 'Completed' && !reportUrl) {
        showError('A Report Link is required to mark a Compliance Request as Completed.');
        document.getElementById('report-url').focus(); // Focus the input for user convenience
        return; // Stop the submission
    }

    showLoading(true, "Submitting update...");

    try {
        const notesValue = document.getElementById('update-notes').value.trim();
        const payload = {
            requestId: parseInt(selectedId, 10),
            requestStatus: newStatus,
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
        showPanel('update-form-panel');
    } finally {
        showLoading(false);
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

function resetForm() {
    document.getElementById("request-form").reset();
    document.getElementById("priority").value = "Medium";
    toggleReportsRequestedField();
    // FIX: Reload email data to ensure form is correctly populated.
    loadEmailData();
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
    const successElement = document.getElementById("success-message");
    successElement.textContent = message;
    successElement.style.display = "block";
    successElement.style.color = "white";
    successElement.style.backgroundColor = "#FF6B00"; // Stryker orange
    successElement.style.padding = "10px";
    successElement.style.border = "1px solid #FF6B00";
    successElement.style.borderRadius = "4px";
    successElement.style.marginBottom = "15px";
    
    setTimeout(clearMessages, 4000);
}

function clearMessages() {
    document.getElementById("error-message").style.display = "none";
    document.getElementById("success-message").style.display = "none";
}
