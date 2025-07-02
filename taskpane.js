// Wait for Office to be ready before doing anything
Office.initialize = function (reason) {
    // This function is called when the add-in is ready to start
    document.getElementById("loading").style.display = "block";
    
    try {
        loadEmailData();
    } catch (error) {
        showError("Error initializing add-in: " + error.message);
    }
};

function loadEmailData() {
    try {
        // Make sure we have access to the mailbox and item
        if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
            showError("Cannot access email. Please make sure you're viewing an email message.");
            return;
        }
        
        var item = Office.context.mailbox.item;
        
        // Direct property access (works with Mailbox API 1.1+)
        document.getElementById("subject").value = item.subject || "(No subject)";
        
        // Get sender info - carefully check for null/undefined
        var senderName = "";
        var senderEmail = "";
        if (item.from) {
            senderName = item.from.displayName || "";
            senderEmail = item.from.emailAddress || "";
        }
        document.getElementById("senderName").value = senderName || "(Unknown sender)";
        document.getElementById("senderEmail").value = senderEmail || "(Unknown email)";
        
        // Get date - carefully check for null/undefined
        var sentDate = "";
        if (item.dateTimeCreated) {
            sentDate = formatDate(item.dateTimeCreated);
        }
        document.getElementById("sentDate").value = sentDate || "(Unknown date)";
        
        // Hide loading, show form
        document.getElementById("loading").style.display = "none";
        document.getElementById("request-form").style.display = "block";
        
    } catch (error) {
        showError("Error loading email data: " + error.message);
    }
}

function formatDate(date) {
    if (!date) return "";
    
    try {
        return date.toLocaleDateString() + " " + date.toLocaleTimeString();
    } catch (e) {
        return date.toString();
    }
}

function showError(message) {
    var errorElement = document.getElementById("error-message");
    errorElement.textContent = message;
    errorElement.style.display = "block";
    document.getElementById("loading").style.display = "none";
}

function showSuccess(message) {
    var successElement = document.getElementById("success-message");
    successElement.textContent = message;
    successElement.style.display = "block";
    setTimeout(function() {
        successElement.style.display = "none";
    }, 5000);
}

// Set up event handlers when the page has loaded
window.onload = function() {
    document.getElementById("submit-btn").onclick = submitToSharePoint;
    document.getElementById("reset-btn").onclick = resetForm;
    document.getElementById("requestType").onchange = toggleReportsRequestedField;
};

function toggleReportsRequestedField() {
    var requestType = document.getElementById("requestType").value;
    var reportsGroup = document.getElementById("reports-requested-group");
    var reportsInput = document.getElementById("reportsRequested");

    if (requestType === "Compliance Request") {
        reportsGroup.style.display = "block";
    } else {
        reportsGroup.style.display = "none";
        reportsInput.value = ""; // Clear the value when hidden
    }
}

async function submitToSharePoint() {
    try {
        var status = document.getElementById("status").value;
        var notes = document.getElementById("notes").value;
        var requestType = document.getElementById("requestType").value;
        var reportsRequested = document.getElementById("reportsRequested").value;
        
        // Get new field values
        var priority = document.getElementById("priority").value;
        var dueDate = document.getElementById("dueDate").value;

        if (!status) {
            showError("Please select a request status");
            return;
        }
        if (!requestType) {
            showError("Please select a request type");
            return;
        }

        if (requestType === "Compliance Request" && reportsRequested && isNaN(parseInt(reportsRequested, 10))) {
            showError("Please enter a valid number for Reports Requested.");
            return;
        }
        
        // Show loading state
        document.getElementById("submit-btn").disabled = true;
        document.getElementById("submit-btn").textContent = "Submitting...";
        
        // Get email data and form inputs
        var item = Office.context.mailbox.item;
        var user = Office.context.mailbox.userProfile;

        // Get email body
        item.body.getAsync("text", async function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var emailBody = result.value;

                // Prepare the payload for Power Automate
                var payload = {
                    subject: document.getElementById("subject").value,
                    senderName: document.getElementById("senderName").value,
                    senderEmail: document.getElementById("senderEmail").value,
                    sentDate: document.getElementById("sentDate").value,
                    requestType: requestType,
                    reportsRequested: reportsRequested ? parseInt(reportsRequested, 10) : null,
                    requestStatus: status,
                    notes: notes || "",
                    // Add new and automatic fields
                    priority: priority,
                    dueDate: dueDate || null,
                    assignedTo: user ? user.emailAddress : "",
                    // ---
                    trackedDate: new Date().toISOString().split('T')[0],
                    trackedBy: user ? user.emailAddress : "Unknown User",
                    conversationId: item.conversationId || "",
                    messageId: item.internetMessageId || item.itemId || "",
                    emailBody: emailBody || ""
                };
                
                // Log the payload for debugging
                console.log("Sending payload to Power Automate:", payload);
                
                // Send the data directly to Power Automate flow
                sendRequestToFlow(payload);
            } else {
                console.error("Failed to get email body:", result.error);
                showError("Error getting email body: " + result.error.message);
                // Reset button state
                document.getElementById("submit-btn").disabled = false;
                document.getElementById("submit-btn").textContent = "Submit to SharePoint";
            }
        });

    } catch (error) {
        console.error("Error in submitToSharePoint (outer):", error);
        showError("Error: " + error.message);
        
        // Reset button state
        document.getElementById("submit-btn").disabled = false;
        document.getElementById("submit-btn").textContent = "Submit to SharePoint";
    }
}

// This function sends the request directly to the flow
function sendRequestToFlow(payloadData) {
    // This URL must include the complete endpoint with the SAS token (sig parameter)
    // The Power Automate flow must have "Allow anonymous requests" set to "On"
    const flowUrl = "https://prod-135.westus.logic.azure.com:443/workflows/075b978523814f56951805720dc2da6d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7fsONcoc2c82EmQGTmubH_9PUgrWGRZz833KgAThavg";
    
    var xhr = new XMLHttpRequest();
    xhr.open("POST", flowUrl, true);
    xhr.setRequestHeader("Content-Type", "application/json");
    xhr.setRequestHeader("Accept", "application/json");
    // No authentication header needed when using anonymous access with SAS token

    // Set up handlers for success and error
    xhr.onreadystatechange = function() {
        if (xhr.readyState === 4) {
            console.log("Response status:", xhr.status);
            console.log("Response text:", xhr.responseText);
            
            if (xhr.status >= 200 && xhr.status < 300) {
                // Success
                showSuccess("Request tracked successfully in the SharePoint list!");
                resetForm();
            } else if (xhr.status === 409) {
                // Duplicate found
                showError("This email has already been tracked for this Request Type. Please visit the SharePoint site.");
            } else {
                // Other Error
                showError("Error submitting to SharePoint List: HTTP " + xhr.status + 
                            (xhr.responseText ? " - " + JSON.parse(xhr.responseText).message : ""));
            }
            
            // Reset button state
            document.getElementById("submit-btn").disabled = false;
            document.getElementById("submit-btn").textContent = "Submit to SharePoint";
        }
    };

    // Handle network errors
    xhr.onerror = function() {
        console.error("Network error occurred");
        showError("Network error occurred while submitting to the SharePoint list. Please check your connection.");
        document.getElementById("submit-btn").disabled = false;
        document.getElementById("submit-btn").textContent = "Submit to SharePoint";
    };
    
    // Send the request with JSON payload (no token required for anonymous access)
    xhr.send(JSON.stringify(payloadData));
}

function resetForm() {
    document.getElementById("status").value = "";
    document.getElementById("notes").value = "";
    document.getElementById("requestType").value = "";
    document.getElementById("reportsRequested").value = "";
    document.getElementById("reports-requested-group").style.display = "none";
    document.getElementById("error-message").style.display = "none";
    // Reset new fields
    document.getElementById("priority").value = "Medium";
    document.getElementById("dueDate").value = "";
}
