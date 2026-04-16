// Global state
let currentStep = 1;
const totalSteps = 4;
let ownersUploaded = false;
let usersUploaded = false;

// Initialize app with auto-reset - CLEARS DATA ON REFRESH
document.addEventListener('DOMContentLoaded', async function () {
    // Clear data on refresh (as per user request)
    try {
        await fetch('/api/reset', { method: 'POST' });
        console.log('System reset on reload');
    } catch (e) {
        console.error('Auto-reset failed', e);
    }

    showStep(1);
    // Reset state flags
    ownersUploaded = false;
    usersUploaded = false;
    // Show empty tables
    displayOwners([]);
    displayUsers([], [], []);
});

// Step Navigation with Validation logic
function validateAndShowStep(targetStep) {
    const currentStepNum = currentStep;

    // Moving Forward validation
    if (targetStep > currentStepNum) {
        // Validation for leaving Step 1: Must have Owners
        if (currentStepNum === 1) {
            if (!ownersUploaded) {
                showConfirmation('btn-next-step-1', '⚠️ You must upload an Owners file before proceeding!', 'warning');
                return;
            }
        }

        // Validation for leaving Step 2: Must have Users
        if (currentStepNum === 2) {
            if (!usersUploaded) {
                showConfirmation('btn-next-step-2', '⚠️ You must upload a Users file before proceeding!', 'warning');
                return;
            }
        }
    }

    showStep(targetStep);
}

// Basic showStep function (UI switching only)
function showStep(step) {
    currentStep = step;

    // Hide ALL step content first
    document.querySelectorAll('.step-content').forEach(content => {
        content.style.display = 'none';
        content.classList.remove('active');
    });

    // Show ONLY the current step
    const currentStepElement = document.getElementById(`step${step}`);
    if (currentStepElement) {
        currentStepElement.style.display = 'block';
        currentStepElement.classList.add('active');
    }

    // Update sidebar navigation
    document.querySelectorAll('.nav-item').forEach((item, index) => {
        item.classList.remove('active');
        // Only mark active if we are on that step
        if (index === step - 1) {
            item.classList.add('active');
        }
    });

    // Update page title and subtitle
    const titles = {
        1: { title: 'Setup & Configuration', subtitle: 'Configure AI and email settings' },
        2: { title: 'User Management', subtitle: 'Upload and manage users' },
        3: { title: 'Review Process', subtitle: 'AI-powered email response monitoring' },
        4: { title: 'Reports & Analysis', subtitle: 'View statistics and export reports' }
    };

    if (titles[step]) {
        document.getElementById('page-title').textContent = titles[step].title;
        document.getElementById('page-subtitle').textContent = titles[step].subtitle;
    }

    // Auto-load data for specific steps
    if (step === 4) {
        loadStats();
        loadUsers('active'); // Fetch only ACTIVE users for final report (hide deleted)
        loadChangeLogs(); // Ensure logs are populated
    }
}

function nextStep() {
    if (currentStep < totalSteps) {
        validateAndShowStep(currentStep + 1);
    }
}

function prevStep() {
    if (currentStep > 1) {
        showStep(currentStep - 1);
    }
}

// Utility Functions
function showAlert(message, type = 'info') {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type}`;
    alertDiv.innerHTML = `
        <span>${getAlertIcon(type)}</span>
        <span>${message}</span>
    `;

    const container = document.querySelector('.step-content.active');
    container.insertBefore(alertDiv, container.firstChild);

    setTimeout(() => alertDiv.remove(), 5000);
}

function getAlertIcon(type) {
    const icons = {
        success: '✅',
        error: '❌',
        warning: '⚠️',
        info: 'ℹ️'
    };
    return icons[type] || icons.info;
}

function showLoading(elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerHTML = '<div class="spinner"></div>';
    }
}

function hideLoading(elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerHTML = '';
    }
}

// Show confirmation message below button
function showConfirmation(buttonId, message, type = 'success') {
    // Remove any existing confirmation
    const existingConfirmation = document.getElementById(`${buttonId}-confirmation`);
    if (existingConfirmation) {
        existingConfirmation.remove();
    }

    // Create confirmation message
    const confirmationDiv = document.createElement('div');
    confirmationDiv.id = `${buttonId}-confirmation`;
    confirmationDiv.className = `confirmation-message confirmation-${type}`;

    const icon = type === 'success' ? '✓' : type === 'error' ? '✗' : type === 'warning' ? '⚠' : 'ℹ';
    confirmationDiv.innerHTML = `<strong>${icon}</strong> ${message}`;

    // Insert after the button
    const button = document.getElementById(buttonId);
    if (button) {
        button.parentNode.insertBefore(confirmationDiv, button.nextSibling);

        // Auto-remove after 5 seconds
        setTimeout(() => {
            if (confirmationDiv.parentNode) {
                confirmationDiv.remove();
            }
        }, 5000);
    }
}

// ==================== STEP 1: CONFIGURATION ====================

async function configureGemini() {
    const apiKey = document.getElementById('gemini-api-key').value;

    if (!apiKey) {
        showAlert('Please enter Gemini API key', 'error');
        return;
    }

    try {
        const response = await fetch('/api/config/gemini', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ api_key: apiKey })
        });

        const data = await response.json();

        if (data.success) {
            showConfirmation('btn-configure-gemini', data.message, 'success');
            document.getElementById('gemini-status').innerHTML =
                '<span class="badge badge-success">✓ Initialized</span>';
        } else {
            showConfirmation('btn-configure-gemini', data.message, 'error');
        }
    } catch (error) {
        showConfirmation('btn-configure-gemini', 'Error: ' + error.message, 'error');
    }
}

async function configureEmail() {
    const formData = {
        smtp_server: document.getElementById('smtp-server').value,
        smtp_port: document.getElementById('smtp-port').value,
        email: document.getElementById('from-email').value,
        password: document.getElementById('email-password').value,
        imap_server: document.getElementById('imap-server').value,
        imap_port: document.getElementById('imap-port').value
    };

    if (!formData.smtp_server || !formData.email || !formData.password) {
        showAlert('Please fill in all required email fields', 'error');
        return;
    }

    try {
        const response = await fetch('/api/config/email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(formData)
        });

        const data = await response.json();

        if (data.success) {
            showConfirmation('btn-configure-email', data.message, 'success');
        } else {
            showConfirmation('btn-configure-email', data.message, 'error');
        }
    } catch (error) {
        showConfirmation('btn-configure-email', 'Error: ' + error.message, 'error');
    }
}

async function uploadOwners() {
    const fileInput = document.getElementById('owners-file');
    const file = fileInput.files[0];

    if (!file) {
        showAlert('Please select a file', 'error');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        showLoading('owners-status');

        const response = await fetch('/api/upload/owners', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        hideLoading('owners-status');

        if (data.success) {
            showConfirmation('owners-file', data.message, 'success');
            ownersUploaded = true; // FIX: Update global validation flag
            displayOwners(data.owners);
        } else {
            showConfirmation('owners-file', data.message, 'error');
        }
    } catch (error) {
        hideLoading('owners-status');
        showAlert('Error uploading file: ' + error.message, 'error');
    }
}

function displayOwners(owners) {
    const container = document.getElementById('owners-preview');

    if (!owners || owners.length === 0) {
        container.innerHTML = '<div class="alert alert-info">No owners uploaded yet. Please upload the Owners Excel file to continue.</div>';
        return;
    }

    const html = `
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>User Name</th>
                        <th>Email</th>
                        <th>Department</th>
                        <th>CC</th>
                    </tr>
                </thead>
                <tbody>
                    ${owners.map(owner => `
                        <tr>
                            <td>${owner.user_name}</td>
                            <td>${owner.email}</td>
                            <td><span class="status-badge ${owner.owner_type === 'IT' ? 'status-info' : 'status-warning'}">${owner.owner_type}</span></td>
                            <td>${owner.cc_email || ''}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    container.innerHTML = html;
}

async function loadOwners() {
    try {
        const response = await fetch('/api/owners');
        const data = await response.json();

        if (data.success && data.owners.length > 0) {
            ownersUploaded = true; // Update state
            displayOwners(data.owners);
        } else {
            ownersUploaded = false; // No owners found
            displayOwners([]); // Display empty state
        }
    } catch (error) {
        console.error('Error loading owners:', error);
        ownersUploaded = false;
        displayOwners([]);
    }
}

// ==================== STEP 2: USER MANAGEMENT ====================

async function uploadUsers() {
    const fileInput = document.getElementById('users-file');
    const file = fileInput.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    document.getElementById('users-status').innerHTML = '<div class="loader"></div> Uploading...';

    try {
        const response = await fetch('/api/upload/users', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        hideLoading('users-status');

        if (data.success) {
            showConfirmation('users-file', data.message, 'success');
            usersUploaded = true;
            displayUsers(data.users, data.it_users, data.business_users);
        } else {
            showConfirmation('users-file', data.message, 'error');
            // Do not update usersUploaded if failed
        }
    } catch (error) {
        hideLoading('users-status');
        showConfirmation('users-file', 'Error uploading file: ' + error.message, 'error');
    }
}

function displayUsers(users, itUsers, businessUsers) {
    // IT Users Table Logic (Step 2)
    const itContainer = document.getElementById('it-users-table');
    const itCard = document.getElementById('card-it-users');

    if (!itUsers || itUsers.length === 0) {
        if (itContainer) itContainer.innerHTML = '';
        if (itCard) itCard.style.display = 'none';
    } else {
        const tableHtml = createUsersTable(itUsers);
        if (itContainer) itContainer.innerHTML = tableHtml;
        if (itCard) itCard.style.display = 'block';
    }

    // Business Users Table Logic (Step 2)
    const businessContainer = document.getElementById('business-users-table');
    const businessCard = document.getElementById('card-business-users');

    if (!businessUsers || businessUsers.length === 0) {
        if (businessContainer) businessContainer.innerHTML = '';
        if (businessCard) businessCard.style.display = 'none';
    } else {
        const tableHtml = createUsersTable(businessUsers);
        if (businessContainer) businessContainer.innerHTML = tableHtml;
        if (businessCard) businessCard.style.display = 'block';
    }

    // Final Unified Table Logic (Step 4)
    const finalContainer = document.getElementById('final-users-table');
    if (finalContainer) {
        if (users && users.length > 0) {
            finalContainer.innerHTML = createUsersTable(users);
        } else {
            finalContainer.innerHTML = '<p class="text-center">No users found.</p>';
        }
    }

    // Update Global State
    if (users && users.length > 0) {
        usersUploaded = true;
    }
}

function createUsersTable(users) {
    const html = `
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>User Name</th>
                        <th>Email</th>
                        <th>Roles</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${users.map(user => `
                        <tr>
                            <td>${user.user_name}</td>
                            <td>${user.email}</td>
                            <td>${user.roles}</td>
                            <td><span class="status-badge ${user.status === 'active' ? 'status-success' : 'status-error'}">
                                ${(user.status || 'ACTIVE').toUpperCase()}
                            </span></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    return html;
}

async function loadUsers(status = 'active') {
    try {
        console.log(`Loading users with status: ${status}`);
        const response = await fetch(`/api/users?status=${status}`);
        const data = await response.json();

        if (data.success && data.users && data.users.length > 0) {
            console.log(`Loaded ${data.users.length} users`);
            usersUploaded = true; // Update state
            displayUsers(data.users, data.it_users, data.business_users);
        } else {
            console.warn('No users found');
            // Only set to false if we are looking for active users and found none
            if (status === 'active') usersUploaded = false;
            displayUsers([], [], []);

            // Show visible notification if manually triggered (refresh)
            const btn = document.querySelector('button[onclick*="loadUsers"]');
            if (btn) {
                // Optional: visual cue
            }
        }
    } catch (error) {
        console.error('Error loading users:', error);
        alert('Error loading users: ' + error.message);
        usersUploaded = false;
        displayUsers([], [], []);
    }
}

// ==================== STEP 3: REVIEW PROCESS ====================

async function generateTickets() {
    try {
        showLoading('tickets-status');

        const response = await fetch('/api/tickets/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({}) // No payload needed now
        });

        const data = await response.json();

        hideLoading('tickets-status');

        if (data.success) {
            showConfirmation('btn-generate-tickets', data.message, 'success');
            displayTickets(data.results);
        } else {
            showConfirmation('btn-generate-tickets', data.message, 'error');
        }
    } catch (error) {
        hideLoading('tickets-status');
        showConfirmation('btn-generate-tickets', 'Error: ' + error.message, 'error');
    }
}

function displayTickets(results) {
    const container = document.getElementById('tickets-list');

    const html = `
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Ticket ID</th>
                        <th>Owner Email</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${results.map(result => `
                        <tr>
                            <td><strong>${result.ticket_id}</strong></td>
                            <td>${result.owner_email}</td>
                            <td><span class="status-badge ${result.success ? 'status-success' : 'status-error'}">
                                ${result.success ? '✓ Sent' : '✗ Failed'}
                            </span></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    container.innerHTML = html;
}

async function fetchResponses() {
    const statusDiv = 'responses-status';
    try {
        const statusEl = document.getElementById(statusDiv);

        // Start simulated progress
        showProgress(statusEl, 5, 'Connecting to email server...');

        let progress = 5;
        const interval = setInterval(() => {
            if (progress < 85) {
                progress += Math.floor(Math.random() * 10); // Random increment
                let msg = progress < 30 ? 'Connecting...' :
                    progress < 60 ? 'Searching Inbox...' :
                        'Filtering Responses...';
                showProgress(statusEl, progress, msg);
            }
        }, 1000);

        const response = await fetch('/api/responses/fetch', {
            method: 'POST'
        });

        clearInterval(interval);

        const data = await response.json();

        if (data.success) {
            showProgress(statusEl, 100, 'Complete!');
            setTimeout(() => {
                showConfirmation(statusDiv, data.message, 'success');
                loadResponses();
            }, 800);
        } else {
            showConfirmation(statusDiv, data.message, 'error');
        }
    } catch (error) {
        showConfirmation(statusDiv, 'Error fetching responses: ' + error.message, 'error');
    }
}

async function loadResponses() {
    try {
        const response = await fetch('/api/responses');
        const data = await response.json();

        if (data.success) {
            displayResponses(data.responses);
        }
    } catch (error) {
        console.error('Error loading responses:', error);
    }
}

function displayResponses(responses) {
    const container = document.getElementById('responses-list');

    if (!responses || responses.length === 0) {
        container.innerHTML = '<p>No responses received yet.</p>';
        return;
    }

    const html = `
        <div class="table-container">
            <table>
                <thead>
                <thead>
                    <tr>
                        <th>Ticket ID</th>
                        <th>From</th>
                        <th>Content</th>
                        <th>Attachment</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${responses.map(resp => `
                        <tr>
                            <td><strong>${resp.ticket_id}</strong></td>
                            <td>${resp.from_email}</td>
                            <td style="white-space: pre-wrap; max-width: 500px;">${resp.body}</td>
                            <td>${resp.has_attachment ? '📎 Yes' : 'No'}</td>
                            <td><span class="status-badge ${resp.processed ? 'status-success' : 'status-warning'}">
                                ${resp.processed ? 'Processed' : 'Pending'}
                            </span></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    container.innerHTML = html;
}

async function processResponses() {
    const statusDiv = 'process-status';
    try {
        // showLoading(statusDiv); // Replaced with Progress
        const statusEl = document.getElementById(statusDiv);

        // 1. Start Progress
        showProgress(statusEl, 15, 'Initializing AI...');

        // 2. Simulate progress movement while waiting
        let progress = 15;
        const interval = setInterval(() => {
            if (progress < 90) {
                progress += Math.floor(Math.random() * 5);
                showProgress(statusEl, progress, 'Analyzing emails with Gemini...');
            }
        }, 500);

        const response = await fetch('/api/responses/process', {
            method: 'POST'
        });

        const data = await response.json();

        clearInterval(interval); // Stop simulation

        if (data.success) {
            // 3. Complete Progress
            showProgress(statusEl, 100, 'Processing Complete!');
            setTimeout(() => {
                showConfirmation(statusDiv, data.message, 'success');
                loadChangeLogs(); // Refresh logs
                loadUsers();      // Refresh users table if changes happened
                loadResponses();  // Refresh responses table (to show 'Processed')
            }, 1000); // Wait 1s to show 100%

        } else {
            showConfirmation(statusDiv, data.message, 'error');
        }
    } catch (error) {
        showConfirmation(statusDiv, 'Error: ' + error.message, 'error');
    }
}

// Helper: Show Percentage Loader
function showProgress(container, percent, text) {
    if (typeof container === 'string') container = document.getElementById(container);
    if (!container) return;

    container.innerHTML = `
        <div class="progress-container">
            <div class="progress-bar progress-bar-animated" style="width: ${percent}%">
                ${percent}% - ${text}
            </div>
        </div>
    `;
}

async function loadChangeLogs() {
    try {
        const response = await fetch('/api/changelogs');
        const data = await response.json();

        if (data.success) {
            displayChangeLogs(data.logs);
        }
    } catch (error) {
        console.error('Error loading change logs:', error);
    }
}

function displayChangeLogs(logs) {
    const container = document.getElementById('changelogs-list');
    const containerFinal = document.getElementById('changelogs-list-final');

    if (!logs || logs.length === 0) {
        const emptyHtml = '<p>No changes recorded yet.</p>';
        if (container) container.innerHTML = emptyHtml;
        if (containerFinal) containerFinal.innerHTML = emptyHtml;
        return;
    }

    const html = `
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Action</th>
                        <th>User</th>
                        <th>Description</th>
                        <th>Ticket ID</th>
                        <th>Time</th>
                    </tr>
                </thead>
                <tbody>
                    ${logs.map(log => `
                        <tr>
                            <td>
                                <span class="status-badge ${getBadgeClass(log.action_type)}">
                                    ${getActionLabel(log.action_type)}
                                </span>
                            </td>
                            <td>${log.user_email}</td>
                            <td>
                                ${log.description}
                                ${log.old_value && log.new_value ? `<div style="font-size:0.85em; color:#666; margin-top:4px;">${log.old_value} ➝ ${log.new_value}</div>` : ''}
                            </td>
                            <td><small>${log.ticket_id}</small></td>
                            <td><small>${new Date(log.created_at).toLocaleString()}</small></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    if (container) container.innerHTML = html;
    if (containerFinal) containerFinal.innerHTML = html;
}

function getBadgeClass(actionType) {
    switch (actionType) {
        case 'delete': return 'status-error';
        case 'update_role': return 'status-warning';
        case 'unauthorized': return 'status-error';
        default: return 'status-info';
    }
}

function getActionLabel(actionType) {
    switch (actionType) {
        case 'delete': return '🗑️ Deleted';
        case 'update_role': return '✏️ Role Update';
        case 'unauthorized': return '⛔ Unauthorized';
        default: return 'ℹ️ Info';
    }
}

// ==================== STEP 4: REPORTS ====================

async function loadStats() {
    try {
        const response = await fetch('/api/stats');
        const data = await response.json();

        if (data.success) {
            displayStats(data.stats);
        }
    } catch (error) {
        console.error('Error loading stats:', error);
    }
}

async function resetSystem() {
    if (!confirm('⚠️ Are you sure you want to delete ALL data? This cannot be undone.')) {
        return;
    }

    try {
        const response = await fetch('/api/reset', { method: 'POST' });
        const data = await response.json();

        if (data.success) {
            alert('System reset successfully!');
            location.reload(); // Reload page to show empty state
        } else {
            alert('Error: ' + data.message);
        }
    } catch (error) {
        alert('Error resetting system: ' + error.message);
    }
}

async function clearProcessData() {
    if (!confirm('⚠️ Are you sure you want to clear PROCESS data (Tickets, Responses, Logs)?\n\nOwners, Users, and Settings will be KEPT.')) {
        return;
    }

    try {
        const response = await fetch('/api/reset_process', { method: 'POST' });
        const data = await response.json();

        if (data.success) {
            alert(data.message);
            // Reload specific parts without full page refresh
            document.getElementById('tickets-list').innerHTML = '';
            document.getElementById('responses-list').innerHTML = '';
            document.getElementById('changelogs-list').innerHTML = '';
            document.getElementById('tickets-status').innerHTML = '';
        } else {
            alert('Error: ' + data.message);
        }
    } catch (error) {
        alert('Error clearing data: ' + error.message);
    }
}

function displayStats(stats) {
    document.getElementById('stats-grid').innerHTML = `
        <div class="stat-card">
            <div class="stat-value">${stats.total_users}</div>
            <div class="stat-label">Total Users</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${stats.it_users}</div>
            <div class="stat-label">IT Users</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${stats.business_users}</div>
            <div class="stat-label">Business Users</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${stats.total_tickets}</div>
            <div class="stat-label">Tickets Generated</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${stats.total_responses}</div>
            <div class="stat-label">Responses Received</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${stats.total_changes}</div>
            <div class="stat-label">Total Changes</div>
        </div>
    `;
}

function exportUsersReport() {
    window.location.href = '/api/reports/users';
}

function exportChangeLogsReport() {
    window.location.href = '/api/reports/changelogs';
}

async function exportToDesktop() {
    try {
        const btn = document.getElementById('btn-export-desktop');
        const originalText = btn.innerHTML;
        btn.innerHTML = '⏳ Exporting...';
        btn.disabled = true;

        const response = await fetch('/api/export_desktop', { method: 'POST' });
        const data = await response.json();

        btn.innerHTML = originalText;
        btn.disabled = false;

        if (data.success) {
            alert('✅ Success! ' + data.message);
        } else {
            alert('❌ Error: ' + data.message);
        }
    } catch (error) {
        alert('❌ Error exporting: ' + error.message);
        if (document.getElementById('btn-export-desktop')) {
            document.getElementById('btn-export-desktop').disabled = false;
            document.getElementById('btn-export-desktop').innerHTML = '🖥️ Finish & Export Results';
        }
    }
}

// Auto-refresh for Step 3
let autoRefreshInterval;

function startAutoRefresh() {
    if (currentStep === 3) {
        autoRefreshInterval = setInterval(() => {
            loadResponses();
            loadChangeLogs();
        }, 30000); // Refresh every 30 seconds
    }
}

function stopAutoRefresh() {
    if (autoRefreshInterval) {
        clearInterval(autoRefreshInterval);
    }
}

// Update auto-refresh when changing steps
const originalShowStep = showStep;
showStep = function (step) {
    stopAutoRefresh();
    originalShowStep(step);

    if (step === 3) {
        startAutoRefresh();
    } else if (step === 4) {
        loadStats();
    }
};
