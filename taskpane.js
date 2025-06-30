// Track Request Add-in JavaScript
// This script handles the taskpane functionality for tracking email requests

// Configuration - Update these values for your environment
const CONFIG = {
    // Power Automate Flow HTTP trigger URL - Replace with your actual endpoint
    POWER_AUTOMATE_ENDPOINT: 'https://prod-135.westus.logic.azure.com:443/workflows/075b978523814f56951805720dc2da6d/triggers/manual/paths/invoke?api-version=2016-06-01',
    
    // SharePoint details for reference (used in data payload)
    SHAREPOINT_SITE: 'https://stryker.sharepoint.com/sites/FlexFinancial-Europe',
    SHAREPOINT_FOLDER: '/Shared Documents/Compliance/Compliance App/Requests/',
    EXCEL_FILE: 'TrackedRequests.xlsx',
    TABLE_NAME: 'TrackedRequests'
};

// Global variables
let currentEmailData = {};

// Initialize the add-in when Office is ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Track Request add-in loaded successfully');
        loadEmailMetadata();
        setupEventListeners();
    }
});

/**
 * Load email metadata from the current message
 */
function loadEmailMetadata() {
    try {
        Office.context.mailbox.item.loadPropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const item = result.value;
                
                // Extract email metadata
                currentEmailData = {
                    subject: item.subject || 'No Subject',
                    sender: getSenderEmail(item),
                    sentDate: formatDate(item.dateTimeCreated),
                    conversationId: item.conversationId,
                    itemId: item.itemId,
                    messageId: getMessageId(item)
                };
                
                // Populate form fields
                populateFormFields();
                
            } else {
                console.error('Error loading email properties:', result.error);
                showStatusMessage('Error loading email metadata', 'error');
            }
        });
    } catch (error) {
        console.error('Error in loadEmailMetadata:', error);
        showStatusMessage('Error accessing email data', 'error');
    }
}

/**
 * Extract sender email address
 */
function getSenderEmail(item) {
    try {
        if (item.sender && item.sender.emailAddress) {
            return `${item.sender.displayName || ''} <${item.sender.emailAddress}>`;
        } else if (item.from && item.from.emailAddress) {
            return `${item.from.displayName || ''} <${item.from.emailAddress}>`;
        }
        return 'Unknown Sender';
    } catch (error) {
        console.error('Error getting sender email:', error);
        return 'Unknown Sender';
    }
}

/**
 * Get message ID for tracking
 */
function getMessageId(item) {
    try {
        // Try to get Internet message ID if available
        if (item.internetMessageId) {
            return item.internetMessageId;
        }
        // Fallback to item ID
        return item.itemId || 'Unknown ID';
    } catch (error) {
        console.error('Error getting message ID:', error);
        return 'Unknown ID';
    }
}

/**
 * Format date for display
 */
function formatDate(date) {
    try {
        if (!date) return 'Unknown Date';
        const d = new Date(date);
        return d.toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    } catch (error) {
        console.error('Error formatting date:', error);
        return 'Unknown Date';
    }
}

/**
 * Populate form fields with email metadata
 */
function populateFormFields() {
    try {
        document.getElementById('subject').value = currentEmailData.subject;
        document.getElementById('sender').value = currentEmailData.sender;
        document.getElementById('sentDate').value = currentEmailData.sentDate;
    } catch (error) {
        console.error('Error populating form fields:', error);
    }
}

/**
 * Setup event listeners for form interactions
 */
function setupEventListeners() {
    // Form submission
    document.getElementById('trackRequestForm').addEventListener('submit', handleFormSubmit);
    
    // Reset button
    document.getElementById('resetBtn').addEventListener('click', resetForm);
    
    // Form validation on change
    document.getElementById('requestStatus').addEventListener('change', validateForm);
}

/**
 * Handle form submission
 */
async function handleFormSubmit(event) {
    event.preventDefault();
    
    if (!validateForm()) {
        showStatusMessage('Please fill in all required fields', 'error');
        return;
    }
    
    try {
        // Show loading indicator
        showLoading(true);
        
        // Prepare data payload
        const formData = prepareDataPayload();
        
        // Send to Power Automate
        await sendToPowerAutomate(formData);
        
        // Show success message
        showStatusMessage('✅ Request tracked successfully!', 'success');
        
        // Reset form after successful submission
        setTimeout(() => {
            resetForm();
        }, 2000);
        
    } catch (error) {
        console.error('Error submitting form:', error);
        showStatusMessage('❌ Error submitting request. Please try again.', 'error');
    } finally {
        showLoading(false);
    }
}

/**
 * Prepare data payload for Power Automate
 */
function prepareDataPayload() {
    const formData = new FormData(document.getElementById('trackRequestForm'));
    
    const payload = {
        // Email metadata
        subject: currentEmailData.subject,
        sender: currentEmailData.sender,
        sentDate: currentEmailData.sentDate,
        conversationId: currentEmailData.conversationId,
        messageId: currentEmailData.messageId,
        
        // User input
        requestStatus: formData.get('requestStatus'),
        notes: formData.get('notes') || '',
        
        // Additional metadata
        trackedDate: new Date().toISOString(),
        trackedBy: getUserEmail(),
        
        // SharePoint configuration (for Power Automate reference)
        targetFile: CONFIG.EXCEL_FILE,
        targetTable: CONFIG.TABLE_NAME,
        targetSite: CONFIG.SHAREPOINT_SITE,
        targetFolder: CONFIG.SHAREPOINT_FOLDER
    };
    
    console.log('Prepared payload:', payload);
    return payload;
}

/**
 * Get current user's email address
 */
function getUserEmail() {
    try {
        return Office.context.mailbox.userProfile.emailAddress || 'Unknown User';
    } catch (error) {
        console.error('Error getting user email:', error);
        return 'Unknown User';
    }
}

/**
 * Send data to Power Automate endpoint
 */
async function sendToPowerAutomate(data) {
    try {
        const response = await fetch(CONFIG.POWER_AUTOMATE_ENDPOINT, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(data)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        console.log('Power Automate response:', result);
        
        return result;
        
    } catch (error) {
        console.error('Error sending to Power Automate:', error);
        throw error;
    }
}

/**
 * Validate form inputs
 */
function validateForm() {
    const requestStatus = document.getElementById('requestStatus').value;
    
    if (!requestStatus) {
        return false;
    }
    
    return true;
}

/**
 * Reset form to initial state
 */
function resetForm() {
    // Clear user input fields only
    document.getElementById('requestStatus').value = '';
    document.getElementById('notes').value = '';
    
    // Re-populate email metadata (in case it was cleared)
    populateFormFields();
    
    // Clear status messages
    clearStatusMessage();
    
    console.log('Form reset');
}

/**
 * Show loading indicator
 */
function showLoading(show) {
    const loadingDiv = document.getElementById('loadingIndicator');
    const submitBtn = document.getElementById('submitBtn');
    
    if (show) {
        loadingDiv.style.display = 'block';
        submitBtn.disabled = true;
        submitBtn.textContent = 'Submitting...';
    } else {
        loadingDiv.style.display = 'none';
        submitBtn.disabled = false;
        submitBtn.textContent = '📤 Submit to SharePoint';
    }
}

/**
 * Show status message
 */
function showStatusMessage(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = `status-message status-${type}`;
    statusDiv.style.display = 'block';
    
    // Auto-hide success messages after 5 seconds
    if (type === 'success') {
        setTimeout(() => {
            clearStatusMessage();
        }, 5000);
    }
}

/**
 * Clear status message
 */
function clearStatusMessage() {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.style.display = 'none';
    statusDiv.textContent = '';
    statusDiv.className = 'status-message';
}

/**
 * Error handling for Office.js API calls
 */
function handleOfficeError(error) {
    console.error('Office.js error:', error);
    showStatusMessage('Error accessing Outlook data', 'error');
}

// Export functions for testing (if needed)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        CONFIG,
        formatDate,
        validateForm,
        prepareDataPayload
    };
}
