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
// Holds the current email item object from the Outlook context (Office.context.mailbox.item).
// Holds the current Outlook user profile information (Office.context.mailbox.userProfile)
let currentItem;
let currentUser;
// Holds the list of request objects associated with the current email, used for duplicate checks and UI updates.
let existingRequests = [];
// --- INITIALIZATION ---
/**
 * Office.onReady initialization sequence:
 * 1. Hide all panels and show loading indicator.
 * 2. Set up global event handlers for UI controls.
 * 3. Retrieve current user and email item from Outlook context.
 * 4. Populate dropdowns for request types and statuses.
 * 5. Check for existing requests for the current email.
 *    - If found, show the request list panel.
 *    - If not found, show the new request form.
 * 6. On error, show an error message and fallback to loading email data and showing the form.
 */
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        try {
            // FIX: Previously, the form was shown before checking for existing requests, which could cause duplicate entries or confusion.
            // Now, we check for existing requests first; if any are found, we show the request list, otherwise we show a new, populated form.
            setupGlobalEventHandlers();
            
            currentUser = Office.context.mailbox.userProfile;
            currentItem = Office.context.mailbox.item;
            
            populateDropdowns();
            // Await inside try to catch async errors
            await checkExistingRequests(true);

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
    // Modified to reset form when creating a new request from the list
    document.getElementById(DOM.createNewBtn).onclick = () => {
        resetForm();
        showPanel(DOM.requestForm);
    };
    document.getElementById(DOM.refreshListBtn).onclick = () => checkExistingRequests(true);

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

async function checkExistingRequests(forceSwitchPanel = true) {
    if (!currentItem) {
        showError("Cannot access email data.");
        showLoading(false);
        return;
    }
    const internetMessageId = currentItem.internetMessageId;
    console.log("CRITICAL_DEBUG: Internet Message ID for this email item is:", internetMessageId);
    console.log("CRITICAL_DEBUG: Internet Message ID is shared across mailboxes (unlike conversationId)");

    try {
        // --- Step 1: Primary lookup by Internet Message ID ---
        // CRITICAL FIX: Use capital 'I' for InternetMessageId and keep as string to match schema
        let lookupPayload = { InternetMessageId: String(internetMessageId || "") };
        const response = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lookupPayload)
        });

        if (!response.ok) throw new Error(`HTTP error ${response.status}`);
        
        const potentialRequests = await response.json();

        if (potentialRequests && potentialRequests.length > 0) {
            // Success! Found matches with the Internet Message ID.
            console.log("Found existing requests by Internet Message ID:", potentialRequests);
            existingRequests = potentialRequests;
            
            // Only proceed with UI updates if forceSwitchPanel is true
            if (forceSwitchPanel) {
                // Simply show the request list panel without the redundant info box
                showRequestsPanel(existingRequests, false);
            }
            return true; // Return true if requests were found
        } else {
            // No match found by Internet Message ID. Show the new request form.
            console.log("No existing requests found for this Internet Message ID. Showing new request form.");
            if (forceSwitchPanel) {
                loadEmailData();
                showPanel(DOM.requestForm);
            }
            return false; // Return false if no requests were found
        }
    } catch (error) {
        console.error("Error checking for existing requests:", error);
        if (forceSwitchPanel) {
            showError("Could not check for existing requests. Please try again.");
            // Fallback to the new request form on any error during lookup.
            loadEmailData();
            showPanel(DOM.requestForm);
        }
        return false; // Return false on error
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
    // This prevents messages from being hidden prematurely.
    if (clear) {
        clearMessages();
    }
    
    // If we're showing the request list panel, make sure the success message is visible
    // above the panel
    if (panelId === DOM.requestListPanel) {
        const successElement = document.getElementById(DOM.successMessage);
        if (successElement && successElement.style.display === "block") {
            successElement.style.zIndex = "1000";
            successElement.style.position = "relative";
        }
    }
}

function showRequestsPanel(requests, showWithMessages = false) {
    const container = document.getElementById(DOM.requestListContainer);
    container.innerHTML = ''; // Clear previous list
    
    // Only clear messages if we're not explicitly showing with messages
    if (!showWithMessages) {
        clearMessages();
    }
    
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
    
    // Removed "Create Another Request" button as requested
    
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
        
        if (req.Priority) {
            if (typeof req.Priority === 'object') {
                console.log("Processing Priority object:", req.Priority);
                processed.Priority = req.Priority.Value || "Medium";
            } else if (!isNaN(req.Priority)) {
                // Convert numeric priority to string equivalent
                console.log("Converting numeric Priority:", req.Priority);
                const priorityMap = {
                    "1": "High",
                    "2": "Medium", 
                    "3": "Low"
                };
                processed.Priority = priorityMap[String(req.Priority)] || "Medium";
            } else {
                // Already a string value, use directly
                processed.Priority = String(req.Priority);
            }
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
        // Add a "Refresh List" info message if there are placeholder IDs
        const hasPlaceholderIds = processedRequests.some(req => {
            const reqId = req.ID !== undefined ? req.ID : req.Id;
            return String(reqId).startsWith('new-');
        });
        
        if (hasPlaceholderIds) {
            const infoDiv = document.createElement('div');
            infoDiv.className = 'request-list-item';
            infoDiv.style.backgroundColor = '#fff3e0';
            infoDiv.style.borderColor = '#ff9800';
            infoDiv.style.padding = '12px';
            infoDiv.style.marginBottom = '15px';
            infoDiv.innerHTML = `
                <div><strong>⚠️ Some requests are still being processed</strong></div>
                <div style="margin-top: 5px;">
                    New requests take a moment to be fully registered. 
                    Click "Refresh List" to update with actual IDs before updating.
                </div>
                <div style="margin-top: 8px;">
                    <button id="special-refresh-btn" style="margin-top: 0;">Refresh List</button>
                </div>
            `;
            container.appendChild(infoDiv);
            
            // Add event handler for the special refresh button
            setTimeout(() => {
                const refreshBtn = document.getElementById("special-refresh-btn");
                if (refreshBtn) {
                    refreshBtn.onclick = () => checkExistingRequests(true);
                }
            }, 0);
        }
        
        processedRequests.forEach((req, index) => {
            try {
                // Create the element
                const itemDiv = document.createElement('div');
                itemDiv.className = 'request-list-item';
                
                // Get the request ID
                const reqId = (req && (req.ID !== undefined || req.Id !== undefined)) ? 
                    (req.ID !== undefined ? req.ID : req.Id) : `unknown-${index}`;
                
                // Check if this is a placeholder ID
                const isPlaceholder = String(reqId).startsWith('new');
                
                if (isPlaceholder) {
                    itemDiv.style.opacity = "0.7";
                    itemDiv.style.border = "1px dashed #ccc";
                }
                    
                const uniqueId = `req-${reqId}-${Math.random().toString(36).substring(2, 8)}`;
                
                // Ensure we have string values for display
                const requestTypeText = String(req.RequestType || "Unknown");
                const statusText = String(req.RequestStatus || "New");
                const statusClass = statusText.toLowerCase().replace(/\s+/g, '-');
                
                // Get priority text, ensuring it's one of Low, Medium, High
                let priorityText;
                if (req.Priority) {
                    // If it's a number (1, 2, 3) convert to text
                    if (!isNaN(req.Priority)) {
                        const priorityMap = {
                            "1": "High",
                            "2": "Medium", 
                            "3": "Low"
                        };
                        priorityText = priorityMap[req.Priority] || String(req.Priority);
                    } else {
                        // If it's already text, use it directly
                        priorityText = String(req.Priority);
                    }
                } else {
                    priorityText = "Medium"; // Default
                }
                
                // Format date safely
                const trackedDate = (req && req.TrackedDate) ? formatDate(req.TrackedDate) : 'Unknown Date';
                
                // Calculate TimeToResolution if status is Completed
                let timeToResolutionText = "";
                if (req.RequestStatus === "Completed" && req.TrackedDate && req.CompletionDate) {
                    timeToResolutionText = calculateTimeToResolution(req.TrackedDate, req.CompletionDate);
                }
                
                // Build the HTML with safe values and new status badges
                let innerHTML = `
                    <input type="radio" name="requestSelection" value="${reqId}" id="${uniqueId}" ${isPlaceholder ? 'disabled' : ''}>
                    <label for="${uniqueId}" class="request-list-item-details">
                        <div>
                            <strong>${requestTypeText}</strong>
                            ${isPlaceholder ? ' <span style="color:#ff9800;font-style:italic">(Processing...)</span>' : ''}
                        </div>
                        <div>
                            <span class="status-badge status-${statusClass}">${statusText}</span>
                        </div>
                        <div>
                            <small>Created: ${trackedDate} | Priority: ${priorityText}</small>
                        </div>
                        ${timeToResolutionText ? `<div><small>Time to Resolution: ${timeToResolutionText}</small></div>` : ''}

                `;
                
                if (isPlaceholder) {
                    innerHTML += `
                        <div style="margin-top:5px;">
                            <small style="color:#ff9800;"><i>This request is still being processed and cannot be updated yet. 
                            Please refresh the list in a moment.</i></small>
                        </div>
                    `;
                }
                
                innerHTML += `</label>`;
                itemDiv.innerHTML = innerHTML;
                
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
            // If already a string, check if it's a number string and convert if needed
            if (!isNaN(selectedRequest.Priority)) {
                const priorityMap = {
                    "1": "High",
                    "2": "Medium", 
                    "3": "Low"
                };
                priorityValue = priorityMap[selectedRequest.Priority] || selectedRequest.Priority;
            } else {
                priorityValue = selectedRequest.Priority;
            }
        } else if (typeof selectedRequest.Priority === 'object' && selectedRequest.Priority.Value) {
            priorityValue = selectedRequest.Priority.Value;
        } else if (!isNaN(selectedRequest.Priority)) {
            // If it's a numeric value, convert to string
            const priorityMap = {
                1: "High",
                2: "Medium", 
                3: "Low"
            };
            priorityValue = priorityMap[selectedRequest.Priority] || "Medium";
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
    
    // Check for duplicate (InternetMessageId, RequestType) combination
    const duplicateRequest = findDuplicateRequest(requestType);
    if (duplicateRequest) {
        // Format the date in a user-friendly way
        let trackedDate = "Unknown date";
        try {
            if (duplicateRequest.TrackedDate) {
                trackedDate = new Date(duplicateRequest.TrackedDate).toLocaleDateString();
            }
        } catch (e) {
            console.error("Error formatting tracked date:", e);
        }
        
        // Get status and priority for better guidance
        let statusText = duplicateRequest.RequestStatus || "Unknown";
        if (typeof statusText === 'object' && statusText.Value) {
            statusText = statusText.Value;
        }
        
        // Create a descriptive error message with guidance and highlight the duplicate request details
        const errorMessage = `
            <p><strong>⚠️ Duplicate Request</strong></p>
            <p>A request of type "<strong>${requestType}</strong>" already exists for this email.</p>
            <div style="background-color: #fff3e0; padding: 10px; border-left: 4px solid #ff9800; margin: 10px 0;">
                <p style="margin: 0;"><strong>Existing Request Details:</strong></p>
                <p style="margin: 5px 0;">• Type: ${requestType}</p>
                <p style="margin: 5px 0;">• Status: ${statusText}</p>
                <p style="margin: 5px 0;">• Created: ${trackedDate}</p>
            </div>
            <p>You cannot create multiple requests of the same type for one email.</p>
            <p>Please view and update the existing request instead.</p>
            <div style="margin-top: 10px;">
                <button id="view-duplicate-btn" style="background-color: #ff9800; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer;">
                    View Existing Request
                </button>
            </div>
        `;
        
        // Show the error but stay on the form
        showError(errorMessage);
        
        // Add a button event handler to go to the list when clicked
        setTimeout(() => {
            const viewButton = document.getElementById("view-duplicate-btn");
            if (viewButton) {
                viewButton.addEventListener("click", () => {
                    showRequestsPanel(existingRequests);
                    // Highlight the duplicate request after the panel is shown
                    setTimeout(() => highlightDuplicateRequest(requestType), 100);
                });
            }
        }, 50);
        
        return;
    }

    showLoading(true, "Submitting new request...");

    try {
        const emailBody = await getBodyAsText();
        
        // Get the selected priority text value directly - no mapping needed
        // According to PowerAutomate-Schema-Types.md, priority must be a string value: "Low", "Medium", or "High"
        const priorityValue = document.getElementById(DOM.priority).value;
        
        console.log("New request: Using priority directly:", priorityValue);
        
        const payload = {
            subject: document.getElementById(DOM.subject).value,
            senderName: document.getElementById(DOM.senderName).value,
            senderEmail: document.getElementById(DOM.senderEmail).value,
            sentDate: currentItem.dateTimeCreated ? new Date(currentItem.dateTimeCreated).toISOString() : null,
            requestType: requestType,
            reportsRequested: parseInt(document.getElementById(DOM.reportsRequested).value, 10) || null,
            requestStatus: status,
            notes: document.getElementById(DOM.notes).value || "",
            priority: priorityValue,
            dueDate: document.getElementById(DOM.dueDate).value || null,
            trackedDate: new Date().toISOString(),
            assignedTo: currentUser ? currentUser.emailAddress : "Unknown User",
            trackedBy: currentUser ? currentUser.emailAddress : "Unknown User",
            conversationId: currentItem.conversationId || "", // Keep for backward compatibility
            InternetMessageId: String(currentItem.internetMessageId || ""), // Capital 'I' and string type as per schema
            messageId: currentItem.internetMessageId || currentItem.itemId || "",
            emailBody: emailBody || ""
        };
        
        console.log("Submitting new request with payload:", payload);
        console.log("InternetMessageId value (capital I):", payload.InternetMessageId);
        console.log("DATA TYPE CHECK - priority:", typeof payload.priority, payload.priority);
        console.log("DATA TYPE CHECK - reportsRequested:", typeof payload.reportsRequested, payload.reportsRequested);
        
        // Use standard JSON.stringify without a replacer to preserve types
        const payloadJson = JSON.stringify(payload);
        
        console.log("Final JSON payload to send:", payloadJson);
        console.log("JSON parsed back:", JSON.parse(payloadJson));
        
        // REFACTOR: Using fetch API directly for cleaner code and better error handling.
        const response = await fetch(CONFIG.REQUEST_CREATE_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: payloadJson
        });

        console.log("Response received from Power Automate:", response);

        // Check for non-successful responses and provide detailed error info.
        if (!response.ok) {
            let errorMessage = `Submission failed. Status: ${response.status}.`;
            try {
                const errorBody = await response.text();
                console.error("Power Automate Error Body:", errorBody);                    // Try to parse the error to get more details
                    try {
                        const parsedError = JSON.parse(errorBody);
                        if (parsedError.error && parsedError.error.message) {
                            errorMessage += ` Error: ${parsedError.error.message}`;
                            
                            // Check for any type mismatch errors
                            if (parsedError.error.message.includes("Invalid type. Expected")) {
                                // This will match both String/Integer and Integer/String type mismatches
                                console.error(`Type mismatch detected in error: ${parsedError.error.message}`);
                                console.log("Original payload:", payload);
                                
                                // Dump all values with their types for debugging
                                console.log("PAYLOAD FIELD TYPES:");
                                Object.entries(payload).forEach(([key, value]) => {
                                    console.log(`Field ${key}: type=${typeof value}, value=${value}`);
                                });
                                
                                // Run the schema analysis
                                parseSchemaRequirements(parsedError.error.message);
                                
                                // More user-friendly error message
                                errorMessage = `Submission failed: Data type mismatch in the request. Please contact IT support.`;
                            }
                        } else {
                            errorMessage += ` Details: ${errorBody}`;
                        }
                    } catch (parseError) {
                        // If not JSON, just use the raw response
                    errorMessage += ` Details: ${errorBody}`;
                }
            } catch (e) {
                errorMessage += " Could not retrieve error details.";
            }
            
            // Throw a detailed error that will be displayed to the user.
            throw new Error(errorMessage);
        }
        
        // Create a timestamp-based ID that will make it unlikely to clash
        // with other requests created in the same session
        const placeholderId = `new-${Date.now()}-${Math.random().toString(36).substring(2, 8)}`;
        
        // Add the new request to our local array
        const newRequestData = {
            Id: placeholderId, // Placeholder ID with added randomness
            RequestType: payload.requestType,
            RequestStatus: payload.requestStatus,
            TrackedDate: payload.trackedDate,
            Priority: payload.priority,
            _isPlaceholder: true // Add a flag to easily identify placeholder records
        };
        existingRequests.push(newRequestData);
        
        // Show the success message with enhanced instructions about placeholder IDs
        const successMsg = `
            <p>Request created successfully!</p>
            <p style="font-size: 90%;">Your request is being processed. <b>You'll need to refresh the list</b> before you can update it.</p>
        `;
        showSuccess(successMsg);
        
        // Clear the form so it's blank if the user creates another request
        resetForm();
        
        // Show the requests panel with the newly created request
        // Pass true to keep the success message visible
        showRequestsPanel(existingRequests, true);
        
        // Set up a timer to automatically refresh the list after a few seconds
        // This helps get the real IDs without requiring user action
        setTimeout(async () => {
            try {
                // Get the updated requests without triggering UI changes
                let lookupPayload = { InternetMessageId: String(currentItem.internetMessageId || "") };
                const response = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(lookupPayload)
                });

                if (response.ok) {
                    const updatedRequests = await response.json();
                    if (updatedRequests && updatedRequests.length > 0) {
                        // Update our local array without changing the view
                        existingRequests = updatedRequests;
                        // Refresh the list view without clearing success message
                        showRequestsPanel(existingRequests, true);
                        console.log("Auto-refreshed requests list after creation");
                    }
                }
            } catch (refreshError) {
                console.error("Error in auto-refresh of request list:", refreshError);
                // Don't show an error to the user, as the submission was successful
            }
        }, 5000); // Wait 5 seconds before auto-refreshing

        // Also schedule another refresh for those slower systems
        setTimeout(async () => {
            try {
                // Check if we still have placeholder IDs
                const stillHasPlaceholders = existingRequests.some(req => {
                    const reqId = req.ID !== undefined ? req.ID : req.Id;
                    return String(reqId).startsWith('new-');
                });
                
                if (stillHasPlaceholders) {
                    // Do another refresh if we still have placeholders
                    let lookupPayload = { InternetMessageId: String(currentItem.internetMessageId || "") };
                    const response = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(lookupPayload)
                    });

                    if (response.ok) {
                        const updatedRequests = await response.json();
                        if (updatedRequests && updatedRequests.length > 0) {
                            existingRequests = updatedRequests;
                            showRequestsPanel(existingRequests, true);
                            console.log("Second auto-refresh of requests list completed");
                        }
                    }
                }
            } catch (refreshError) {
                console.error("Error in second auto-refresh:", refreshError);
            }
        }, 10000); // Try again after 10 seconds

    } catch (error) {
        console.error("Submit error details:", error);
        // Display the specific error message to the user
        showError(error.message);
        // Keep the form visible so the user can try again
        showPanel(DOM.requestForm, false); // Don't clear the error message
    } finally {
        showLoading(false);
    }
}

async function submitUpdate() {
    const selectedRequest = getSelectedRequest();
    if (!selectedRequest) {
        showError("Please select a request to update.");
        return;
    }
    const selectedId = selectedRequest.ID || selectedRequest.Id;

    // Check if this is a placeholder ID (starts with "new-")
    if (selectedId && String(selectedId).startsWith("new-")) {
        showError("This request was just created and is still being processed. Please wait a moment and refresh the list before updating.");
        return;
    }

    // Ensure selectedId is a valid number before proceeding
    if (!selectedId || isNaN(Number(selectedId))) {
        showError("Invalid request ID. Cannot update this request.");
        return;
    }
    
    // Convert selectedId to an actual number to ensure it's not sent as a string
    const numericId = Number(selectedId);

    // Verify we have access to the InternetMessageId before proceeding
    let hasInternetMessageId = false;
    if (selectedRequest.InternetMessageId) {
        hasInternetMessageId = true;
    } else if (currentItem && currentItem.internetMessageId) {
        hasInternetMessageId = true;
    }
    
    if (!hasInternetMessageId) {
        showError("Cannot update request: InternetMessageId is not available. Please refresh the list and try again.");
        return;
    }

    const newStatus = document.getElementById(DOM.updateStatus).value;
    const reportUrl = document.getElementById(DOM.reportUrl).value;
    const requestType = selectedRequest.RequestType;
    const priorityValue = document.getElementById(DOM.updatePriority).value;

    // Priority value from dropdown is already the correct text value (Low, Medium, High)
    // No mapping needed as we're already using the correct string values in the HTML
    const priority = priorityValue; // Use the value directly from the dropdown


    // VALIDATION: Enforce Report Link requirement before submitting.
    if (requestType === 'Compliance Request' && newStatus === 'Completed' && !reportUrl) {
        showError('A Report Link is required to mark a Compliance Request as Completed.');
        document.getElementById(DOM.reportUrl).focus(); // Focus the input for user convenience
        return; // Stop the submission
    }

    showLoading(true, "Submitting update...");

    try {
        const notesValue = document.getElementById(DOM.updateNotes).value.trim();
        
        // Get the InternetMessageId - try from the selected request first, then fallback to current email
        let internetMessageId = null;
        
        // First check if the selected request has an InternetMessageId
        if (selectedRequest.InternetMessageId) {
            internetMessageId = selectedRequest.InternetMessageId;
            console.log("Using InternetMessageId from selected request:", internetMessageId);
        } else if (currentItem && currentItem.internetMessageId) {
            internetMessageId = currentItem.internetMessageId;
            console.log("Using InternetMessageId from current email item:", internetMessageId);
        } else {
            console.warn("No InternetMessageId found in request or current item!");
        }
        
        // Power Automate expects priority as a string with values "Low", "Medium", "High"
        // No mapping needed - we'll use the original string values
        
        console.log("Using text priority value directly:", priority);
        
        const payload = {
            // Keep requestId as a number as Power Automate expects an integer
            requestId: parseInt(selectedId, 10),
            // Keep strings as strings
            requestStatus: newStatus,
            // Send priority directly as the string value (Low, Medium, High)
            priority: priority,
            updatedBy: currentUser ? currentUser.emailAddress : "Unknown User",
            // CRITICAL FIX: Use "InternetMessageId" with capital I to match the schema
            // InternetMessageId is stored as text in SharePoint list, keep as string
            InternetMessageId: String(internetMessageId || "")
        };
        
        // If the status is being changed to Completed, add the completion date
        if (newStatus === 'Completed') {
            // Add the current timestamp as the completion date
            payload.completionDate = new Date().toISOString();
        }

        // Add some debugging output
        console.log("Update payload with InternetMessageId:", payload);
        console.log("DATA TYPE CHECK - requestId:", typeof payload.requestId, payload.requestId);
        console.log("DATA TYPE CHECK - priority:", typeof payload.priority, payload.priority);
        console.log("DATA TYPE CHECK - requestStatus:", typeof payload.requestStatus, payload.requestStatus);
        console.log("DATA TYPE CHECK - InternetMessageId:", typeof payload.InternetMessageId, payload.InternetMessageId);

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

        // Custom function to manage data types for Power Automate
        // Based on the error message, we need to keep requestId as a number but priority might need to be a string
        const payloadJson = JSON.stringify(payload);
        
        console.log("Final JSON payload to send:", payloadJson);
        
        // Also log parsed JSON to verify types
        console.log("JSON parsed back:", JSON.parse(payloadJson));

        // CRITICAL FIX: Validate that the payload has the correct InternetMessageId property
        // and that it's being sent as a string as expected by SharePoint
        let parsedPayload = JSON.parse(payloadJson);
        if (!parsedPayload.InternetMessageId && parsedPayload.InternetMessageId !== "") {
            console.error("ERROR: InternetMessageId is missing from the payload!");
        } else if (typeof parsedPayload.InternetMessageId !== 'string') {
            console.error("ERROR: InternetMessageId should be a string but is:", typeof parsedPayload.InternetMessageId);
            // Force it to be a string
            parsedPayload.InternetMessageId = String(parsedPayload.InternetMessageId || "");
            payloadJson = JSON.stringify(parsedPayload);
        }
        
        // Using the correct update flow URL
        const response = await fetch(CONFIG.REQUEST_UPDATE_URL, { 
            method: 'POST', 
            headers: { 'Content-Type': 'application/json' }, 
            body: payloadJson 
        });

        console.log("Update response status:", response.status);

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Update error response:", errorText);
            
            // Try to parse error as JSON for better error details
            let detailedErrorMessage = `HTTP Error ${response.status}`;
            try {
                const errorJson = JSON.parse(errorText);
                if (errorJson.error && errorJson.error.message) {
                    detailedErrorMessage += `: ${errorJson.error.message}`;
                    
                    // Check specifically for missing InternetMessageId
                    if (errorJson.error.message.includes("missing required property 'body/InternetMessageId'")) {
                        detailedErrorMessage = "Update failed: The InternetMessageId is missing. Please refresh and try again.";
                    }
                    // Check for any type mismatch errors
                    else if (errorJson.error.message.includes("Invalid type. Expected")) {
                        // This will match both String/Integer and Integer/String type mismatches
                        console.error(`Type mismatch detected in error: ${errorJson.error.message}`);
                        console.log("Original payload:", payload);
                        
                        // Dump all values with their types for debugging
                        console.log("PAYLOAD FIELD TYPES:");
                        Object.entries(payload).forEach(([key, value]) => {
                            console.log(`Field ${key}: type=${typeof value}, value=${value}`);
                        });
                        
                        // Run the schema analysis
                        parseSchemaRequirements(errorJson.error.message);
                        
                        // More user-friendly error message
                        detailedErrorMessage = `Update failed: Data type mismatch in the request. Please contact IT support.`;
                    }
                } else {
                    detailedErrorMessage += `: ${errorText}`;
                }
            } catch (e) {
                // If not JSON, just use text
                detailedErrorMessage += `: ${errorText}`;
            }
            
            throw new Error(detailedErrorMessage);
        }

        showSuccess("Request updated successfully!");
        
        // Refresh the list to show the update, but stay on the request list view
        try {
            // Get the updated requests without triggering UI changes
            // CRITICAL FIX: Use correct property name and string type for InternetMessageId in payload
            let lookupPayload = { 
                InternetMessageId: String(internetMessageId || "")  // Use capital 'I' and string type to match schema
            };
            console.log("Looking up requests with InternetMessageId after update:", String(internetMessageId || ""));
            
            const refreshResponse = await fetch(CONFIG.REQUEST_LOOKUP_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(lookupPayload)
            });

            if (refreshResponse.ok) {
                const updatedRequests = await refreshResponse.json();
                if (updatedRequests && updatedRequests.length > 0) {
                    console.log("Found updated requests:", updatedRequests.length);
                    // Update our local array and refresh the list view
                    existingRequests = updatedRequests;
                    console.log("Showing requests panel with updated data");
                    showRequestsPanel(existingRequests, true);
                    
                    // CRITICAL FIX: Ensure the panel is visible
                    document.getElementById(DOM.requestListPanel).style.display = 'block';
                } else {
                    console.warn("No requests found after update. Showing existing list.");
                    // Still show the requests panel even if no requests found
                    showRequestsPanel(existingRequests, true);
                    
                    // CRITICAL FIX: Ensure the panel is visible
                    document.getElementById(DOM.requestListPanel).style.display = 'block';
                }
            } else {
                console.warn("Error response when refreshing requests:", refreshResponse.status);
                // Still show the requests panel even if refresh fails
                showRequestsPanel(existingRequests, true);
                
                // CRITICAL FIX: Ensure the panel is visible
                document.getElementById(DOM.requestListPanel).style.display = 'block';
            }
        } catch (refreshError) {
            console.error("Error refreshing request list after update:", refreshError);
            // Still show the requests panel even if refresh fails
            showRequestsPanel(existingRequests, true);
            
            // CRITICAL FIX: Ensure the panel is visible
            document.getElementById(DOM.requestListPanel).style.display = 'block';
        }
        
        // CRITICAL FIX: Force showing the request list panel in case the above code didn't work
        console.log("Ensuring request list panel is shown");
        showPanel(DOM.requestListPanel, false);

    } catch (error) {
        console.error("Update submission error:", error);
        showError(error.message);
        // Show the update form again on error so the user can retry
        showPanel(DOM.updateFormPanel, false); // Don't clear the error message
    } finally {
        showLoading(false);
    }
}

// --- HELPER FUNCTIONS ---

/**
 * Parse an error message from Power Automate to determine expected schema types.
 * This helps developers understand what data types are expected by Power Automate.
 * @param {string} errorMessage - The error message from Power Automate
 */
function parseSchemaRequirements(errorMessage) {
    console.log("SCHEMA ANALYSIS: Analyzing Power Automate schema requirements");
    
    // Extract type mismatch information
    const typeMismatches = [];
    const regex = /Invalid type\. Expected (\w+) but got (\w+)(?:\.| for '([^']+)')/g;
    let match;
    
    while ((match = regex.exec(errorMessage)) !== null) {
        const expected = match[1];
        const received = match[2];
        const field = match[3] ? match[3].replace('body/', '') : 'unknown field';
        
        typeMismatches.push({
            field,
            expected,
            received
        });
    }
    
    if (typeMismatches.length > 0) {
        console.log("SCHEMA ANALYSIS: Found the following type mismatches:");
        typeMismatches.forEach(mismatch => {
            console.log(`Field: ${mismatch.field}, Expected: ${mismatch.expected}, Received: ${mismatch.received}`);
        });
        
        console.log("SCHEMA ANALYSIS: Suggested field types:");
        typeMismatches.forEach(mismatch => {
            console.log(`${mismatch.field}: ${mismatch.expected}`);
        });
    } else {
        console.log("SCHEMA ANALYSIS: No clear type mismatches found in error message");
    }
}

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
    // Reset all form fields
    const form = document.getElementById(DOM.requestForm);
    if (form) form.reset();
    
    // Clear specific fields that might not be fully reset by the form reset method
    const requestTypeElem = document.getElementById(DOM.requestType);
    if (requestTypeElem) requestTypeElem.value = "";
    const notesElem = document.getElementById(DOM.notes);
    if (notesElem) notesElem.value = "";
    const dueDateElem = document.getElementById(DOM.dueDate);
    if (dueDateElem) dueDateElem.value = "";
    
    // Reset priority to default "Medium" if it exists, otherwise use the first available option
    const priorityDropdown = document.getElementById(DOM.priority);
    let foundMedium = false;
    if (priorityDropdown && priorityDropdown.options) {
        for (let i = 0; i < priorityDropdown.options.length; i++) {
            if (priorityDropdown.options[i].value === "Medium") {
                priorityDropdown.value = "Medium";
                foundMedium = true;
                break;
            }
        }
        if (!foundMedium && priorityDropdown.options.length > 0) {
            priorityDropdown.value = priorityDropdown.options[0].value;
        }
    }
    
    // Reset status to default "New" only if the option exists
    const statusElem = document.getElementById(DOM.status);
    if (statusElem) {
        const hasNewOption = Array.from(statusElem.options).some(opt => opt.value === "New");
        if (hasNewOption) statusElem.value = "New";
        else statusElem.value = ""; // fallback to blank if "New" is not present
    }
    const statusDropdown = document.getElementById(DOM.status);
    let foundNew = false;
    if (statusDropdown && statusDropdown.options) {
        for (let i = 0; i < statusDropdown.options.length; i++) {
            if (statusDropdown.options[i].value === "New") {
                statusDropdown.value = "New";
                foundNew = true;
                break;
            }
        }
        if (!foundNew && statusDropdown.options.length > 0) {
            statusDropdown.value = statusDropdown.options[0].value;
        }
        // No need to get the element again, we already have it in statusDropdown
        statusDropdown.value = "New";
    }
    // No need to get the element again, we already have it in statusDropdown
    if (statusDropdown) statusDropdown.value = "New";
    
    // Reload email metadata but not user inputs
    loadEmailData();
    
    // Ensure the Reports Requested field is properly set after reset
    toggleReportsRequestedField();
    
    // Clear any messages
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
        // Create a spinner + message for better visual feedback
        loading.innerHTML = `
            <div style="display: flex; align-items: center; justify-content: center; flex-direction: column;">
                <div class="spinner" style="
                    border: 4px solid #f3f3f3;
                    border-top: 4px solid #FF6B00;
                    border-radius: 50%;
                    width: 30px;
                    height: 30px;
                    animation: spin 1s linear infinite;
                    margin-bottom: 10px;
                "></div>
                <div style="font-weight: bold;">${message}</div>
            </div>
            <style>
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
            </style>
        `;
        loading.style.display = 'flex';
        loading.style.alignItems = 'center';
        loading.style.justifyContent = 'center';
        loading.style.padding = '20px';
        loading.style.backgroundColor = '#f8f8f8';
        loading.style.borderRadius = '4px';
        loading.style.boxShadow = '0 2px 4px rgba(0, 0, 0, 0.1)';
        
        document.querySelectorAll('.panel').forEach(p => p.style.display = 'none');
    } else {
        loading.style.display = 'none';
    }
}

function showError(message) {
    const errorElement = document.getElementById(DOM.errorMessage);
    
    // Determine if this is a duplicate request error
    const isDuplicateError = message.includes('Duplicate Request') || 
                            message.includes('already exists for this email');
    
    // Support HTML content in error messages
    errorElement.style.padding = "12px";
    
    // Use different styling for duplicate errors (orange warning) vs other errors (red)
    if (isDuplicateError) {
        errorElement.style.border = "1px solid #ff9800";
        errorElement.style.backgroundColor = "#ff9800";
    } else {
        errorElement.style.border = "1px solid #d32f2f";
        errorElement.style.backgroundColor = "#d32f2f";
    }
    
    errorElement.style.borderRadius = "4px";
    errorElement.style.marginBottom = "15px";
    errorElement.style.fontWeight = "bold";
    errorElement.style.boxShadow = "0 2px 4px rgba(0, 0, 0, 0.2)";
    showLoading(false);
    
    // Log the error to console for debugging
    console.error("ERROR SHOWN TO USER:", message);
    
    // Auto-dismiss error after 8 seconds for errors that aren't critical or duplicates
    if (!isDuplicateError) {
        setTimeout(clearMessages, 8000);
    }
}

function showSuccess(message) {
    const successElement = document.getElementById(DOM.successMessage);
    
    // Support HTML content in success messages
    if (message.includes('<p>') || message.includes('<div>')) {
        successElement.innerHTML = message; // Use innerHTML for HTML content
    } else {
        successElement.textContent = message; // Use textContent for plain text (safer)
    }
    
    successElement.style.display = "block";
    successElement.style.color = "white";
    successElement.style.backgroundColor = "#FF6B00"; // Stryker orange
    successElement.style.padding = "12px";
    successElement.style.border = "1px solid #FF6B00";
    successElement.style.borderRadius = "4px";
    successElement.style.marginBottom = "15px";
    successElement.style.fontWeight = "bold";
    successElement.style.boxShadow = "0 2px 4px rgba(0, 0, 0, 0.2)";
    
    // Auto-dismiss success message after 6 seconds
    setTimeout(clearMessages, 6000);
}

function clearMessages() {
    const errorElem = document.getElementById(DOM.errorMessage);
    const successElem = document.getElementById(DOM.successMessage);
    if (errorElem) errorElem.style.display = "none";
    if (successElem) successElem.style.display = "none";
}

// Checks if a request with the same InternetMessageId and RequestType already exists
function findDuplicateRequest(requestType) {
    if (!existingRequests || !Array.isArray(existingRequests) || existingRequests.length === 0) {
        return null; // No existing requests to check against
    }
    
    console.log("Checking for duplicate requests with type:", requestType);
    
    // Normalize the request type for case-insensitive comparison
    const normalizedType = requestType.trim().toLowerCase();
    
    // Find any request with the same RequestType
    const duplicate = existingRequests.find(req => {
        // Handle both string and object formats for RequestType
        let existingType = req.RequestType;
        if (existingType && typeof existingType === 'object' && existingType.Value) {
            existingType = existingType.Value;
        }
        
        // Normalize for case-insensitive comparison
        const normalizedExistingType = existingType ? existingType.trim().toLowerCase() : '';
        
        console.log(`Comparing request type '${normalizedExistingType}' with '${normalizedType}'`);
        return normalizedExistingType === normalizedType;
    });
    
    // Enhance debugging info
    if (duplicate) {
        console.log("Found duplicate request:", duplicate);
        
        // Log additional details about the duplicate
        const duplicateId = duplicate.ID || duplicate.Id || "unknown";
        const duplicateStatus = duplicate.RequestStatus || "unknown";
        console.log(`Duplicate request details: ID=${duplicateId}, Status=${duplicateStatus}`);
        
        // Check if it's a placeholder ID
        if (String(duplicateId).startsWith('new-')) {
            console.log("Note: This duplicate has a placeholder ID and is still processing");
        }
    } else {
        console.log("No duplicate request found for type:", requestType);
    }
    
    return duplicate; // Will return the duplicate object or null if no duplicate is found
}

/**
 * Highlights a request in the list that matches the given request type.
 * Used when showing the list after a duplicate warning to make it obvious
 * which item is the duplicate.
 * @param {string} requestType - The type of request to highlight
 */
function highlightDuplicateRequest(requestType) {
    // Normalize the request type
    const normalizedType = requestType.trim().toLowerCase();
    
    // Get all the request items in the list
    const requestItems = document.querySelectorAll('.request-list-item');
    
    // Loop through them to find and highlight the matching one
    requestItems.forEach(item => {
        // Check if this item contains the request type text
        const typeElement = item.querySelector('strong');
        if (typeElement && typeElement.textContent.trim().toLowerCase() === normalizedType) {
            // This is the duplicate, highlight it
            item.style.border = "2px solid #ff9800";
            item.style.backgroundColor = "#fff8e1";
            
            // Add a duplicate indicator
            const duplicateIndicator = document.createElement('div');
            duplicateIndicator.innerHTML = `
                <div style="background-color: #ff9800; color: white; padding: 3px 8px; 
                border-radius: 12px; display: inline-block; margin-top: 5px; font-size: 11px; font-weight: bold;">
                    Duplicate Found
                </div>
            `;
            
            // Add it after the strong element
            typeElement.parentNode.appendChild(duplicateIndicator);
            
            // Scroll to this item
            setTimeout(() => {
                item.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, 300);
            
            // Auto-select the radio button for this item
            const radioBtn = item.querySelector('input[type="radio"]');
            if (radioBtn) {
                radioBtn.checked = true;
            }
        }
    });
}

/**
 * Calculates time difference between two dates with precision to minutes
 * @param {string|Date} startDate - The start date/time
 * @param {string|Date} endDate - The end date/time
 * @returns {string} Formatted time difference (e.g., "22.5 hours" or "2.2 days")
 */
function calculateTimeToResolution(startDate, endDate) {
    if (!startDate || !endDate) return "";
    
    try {
        // Convert strings to Date objects if needed
        const start = startDate instanceof Date ? startDate : new Date(startDate);
        const end = endDate instanceof Date ? endDate : new Date(endDate);
        
        // Calculate difference in milliseconds
        const diffMs = end.getTime() - start.getTime();
        
        // Convert to hours with 1 decimal precision
        const diffHours = (diffMs / (1000 * 60 * 60)).toFixed(1);
        
        // Format appropriately based on duration
        if (diffHours >= 24) {
            const diffDays = (diffHours / 24).toFixed(1);
            return `${diffDays} days`;
        } else {
            return `${diffHours} hours`;
        }
    } catch (e) {
        console.error("Error calculating time to resolution:", e);
        return "";
    }
}
