// --- UI State ---
let structure = [];

function renderStructure() {
    const list = document.getElementById('structure-list');
    list.innerHTML = '';
    structure.forEach((step, idx) => {
        const li = document.createElement('li');
        li.innerHTML = `
            <input type="text" value="${step.keyword}" placeholder="Slide Type" onchange="updateStep(${idx}, 'keyword', this.value)">
            <select onchange="updateStep(${idx}, 'action', this.value)">
                <option${step.action === 'Copy from GTM (as is)' ? ' selected' : ''}>Copy from GTM (as is)</option>
                <option${step.action === 'Merge: Template Layout + GTM Content' ? ' selected' : ''}>Merge: Template Layout + GTM Content</option>
            </select>
            <button type="button" onclick="removeStep(${idx})">🗑️</button>
        `;
        list.appendChild(li);
    });
}

window.updateStep = function(idx, field, value) {
    structure[idx][field] = value;
};

window.removeStep = function(idx) {
    structure.splice(idx, 1);
    renderStructure();
};

document.getElementById('add-step').onclick = function() {
    structure.push({ keyword: '', action: 'Copy from GTM (as is)' });
    renderStructure();
};
document.getElementById('clear-steps').onclick = function() {
    structure = [];
    renderStructure();
};

// Add a reset form function
window.resetForm = function() {
    document.getElementById('api_key').value = '';
    document.getElementById('template_files').value = '';
    document.getElementById('gtm_file').value = '';
    structure = [];
    renderStructure();
    document.getElementById('success-message').style.display = 'none';
    document.getElementById('error-message').style.display = 'none';
    console.log('[FRONTEND] Form reset completed');
};

// --- Form Submission ---
let isSubmitting = false; // Prevent double submissions

document.getElementById('upload-form').onsubmit = async function(e) {
    e.preventDefault();
    
    // Prevent double submissions
    if (isSubmitting) {
        console.log('[FRONTEND] Submission already in progress, ignoring duplicate request');
        return;
    }
    
    isSubmitting = true;
    const submitButton = document.querySelector('button[type="submit"]');
    const originalButtonText = submitButton.textContent;
    submitButton.disabled = true;
    
    // Add progress indicator
    let progressText = 'Processing';
    let dotCount = 0;
    submitButton.textContent = progressText;
    
    // Animate the button text to show progress
    const progressInterval = setInterval(() => {
        dotCount = (dotCount + 1) % 4;
        submitButton.textContent = progressText + '.'.repeat(dotCount);
    }, 500);
    
    // Show processing message
    document.getElementById('success-message').style.display = 'none';
    document.getElementById('error-message').style.display = 'none';
    
    // Add a processing status div
    let statusDiv = document.getElementById('processing-status');
    if (!statusDiv) {
        statusDiv = document.createElement('div');
        statusDiv.id = 'processing-status';
        statusDiv.style.cssText = 'background:#e5f3ff;border:1px solid #3b82f6;padding:10px;margin:10px 0;border-radius:5px;display:none;';
        document.querySelector('.container').appendChild(statusDiv);
    }
    statusDiv.innerHTML = `
        <div style="color:#1e40af;font-weight:bold;">🤖 AI Processing in Progress...</div>
        <div style="color:#64748b;font-size:14px;margin-top:5px;">
            • Analyzing presentation structure<br>
            • AI selecting best content matches<br>
            • Merging template layouts with GTM content<br>
            • This may take 1-3 minutes depending on content complexity
        </div>
    `;
    statusDiv.style.display = 'block';
    
    document.getElementById('process-log').style.display = 'none';
    document.getElementById('success-message').style.display = 'none';
    document.getElementById('error-message').style.display = 'none';

    const api_key = document.getElementById('api_key').value.trim();
    const template_files = document.getElementById('template_files').files;
    const gtm_file = document.getElementById('gtm_file').files[0];
    
    if (!api_key || !template_files.length || !gtm_file || structure.length === 0) {
        alert('Please fill all fields and add at least one step.');
        // Reset submission state
        isSubmitting = false;
        submitButton.disabled = false;
        submitButton.textContent = originalButtonText;
        return;
    }
    const formData = new FormData();
    formData.append('api_key', api_key);
    for (let i = 0; i < template_files.length; i++) {
        formData.append('template_files', template_files[i]);
    }
    formData.append('gtm_file', gtm_file);
    formData.append('structure', JSON.stringify(structure));

    console.log('[FRONTEND] Form data prepared:', {
        api_key_length: api_key.length,
        template_files_count: template_files.length,
        gtm_file_name: gtm_file.name,
        gtm_file_size: gtm_file.size,
        structure_steps: structure.length
    });

    try {
        console.log('[FRONTEND] Starting async assembly request');
        const startTime = Date.now();
        
        // Step 1: Start the assembly process
        console.log('[FRONTEND] Step 1: Starting assembly...');
        const startResp = await fetch('/start_assemble', {
            method: 'POST',
            body: formData
        });
        
        if (!startResp.ok) {
            const errorText = await startResp.text();
            throw new Error(`Failed to start assembly: ${startResp.status} ${errorText}`);
        }
        
        const startResult = await startResp.json();
        const taskId = startResult.task_id;
        console.log(`[FRONTEND] Assembly started with task ID: ${taskId}`);
        
        // Step 2: Poll for completion
        console.log('[FRONTEND] Step 2: Polling for completion...');
        let pollCount = 0;
        const maxPolls = 60; // 5 minutes max (5 second intervals)
        
        while (pollCount < maxPolls) {
            await new Promise(resolve => setTimeout(resolve, 5000)); // Wait 5 seconds
            pollCount++;
            
            const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
            console.log(`[FRONTEND] Poll ${pollCount}: Checking status after ${elapsed}s...`);
            
            try {
                const statusResp = await fetch(`/task_status/${taskId}`);
                if (!statusResp.ok) {
                    console.warn('[FRONTEND] Status check failed, retrying...');
                    continue;
                }
                
                const status = await statusResp.json();
                console.log(`[FRONTEND] Status: ${status.status}, Progress: ${status.progress}%, Message: ${status.message}`);
                
                // Update UI with progress
                const statusDiv = document.getElementById('processing-status');
                if (statusDiv) {
                    const minutes = Math.floor(elapsed / 60);
                    const seconds = Math.floor(elapsed % 60);
                    const timeStr = minutes > 0 ? `${minutes}m ${seconds}s` : `${seconds}s`;
                    
                    statusDiv.innerHTML = `
                        <div style="color:#1e40af;font-weight:bold;">🤖 AI Processing in Progress... (${timeStr})</div>
                        <div style="color:#64748b;font-size:14px;margin-top:5px;">
                            • Progress: ${status.progress || 0}%<br>
                            • Status: ${status.message || 'Processing...'}<br>
                            • <strong>This approach avoids timeout issues!</strong>
                        </div>
                        <div style="background:#f0f0f0;border-radius:10px;height:10px;margin:10px 0;">
                            <div style="background:#3b82f6;height:100%;border-radius:10px;width:${status.progress || 0}%;transition:width 0.3s;"></div>
                        </div>
                    `;
                }
                
                if (status.status === 'completed') {
                    console.log('[FRONTEND] Assembly completed! Downloading file...');
                    
                    // Step 3: Download the result
                    const downloadResp = await fetch(`/download/${taskId}`);
                    if (!downloadResp.ok) {
                        throw new Error(`Download failed: ${downloadResp.status}`);
                    }
                    
                    const blob = await downloadResp.blob();
                    console.log('[FRONTEND] File downloaded:', blob.size, 'bytes');
                    
                    // Trigger download
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'Dynamic_AI_Assembled_Deck.pptx';
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    window.URL.revokeObjectURL(url);
                    
                    document.getElementById('success-message').textContent = '✨ Your new regional presentation has been assembled!';
                    document.getElementById('success-message').style.display = 'block';
                    return;
                    
                } else if (status.status === 'failed') {
                    throw new Error(`Assembly failed: ${status.error || status.message}`);
                }
                
            } catch (pollError) {
                console.warn(`[FRONTEND] Poll error: ${pollError.message}`);
                // Continue polling unless it's a critical error
                if (pollCount >= 3) { // Allow a few poll failures
                    throw new Error(`Status polling failed: ${pollError.message}`);
                }
            }
        }
        
        throw new Error('Assembly timed out after 5 minutes');
        
    } catch (err) {
        console.error('[FRONTEND] Error during request:', err);
        console.error('[FRONTEND] Error name:', err.name);
        console.error('[FRONTEND] Error message:', err.message);
        console.error('[FRONTEND] Error stack:', err.stack);
        
        // Provide more specific error message based on error type
        let userMessage = 'Error: ' + err.message;
        if (err.name === 'TypeError' && err.message.includes('Failed to fetch')) {
            userMessage = 'Connection error: Please check your internet connection and try again.';
        } else if (err.message.includes('CSP')) {
            userMessage = 'Browser security error: Please try refreshing the page or using a different browser.';
        } else if (err.message.includes('timeout')) {
            userMessage = 'Request timeout: The server is taking too long to respond. Please try again.';
        }
        
        document.getElementById('error-message').textContent = userMessage;
        document.getElementById('error-message').style.display = 'block';
    } finally {
        // Always reset submission state and clean up progress indicators
        if (typeof progressInterval !== 'undefined') {
            clearInterval(progressInterval);
        }
        
        isSubmitting = false;
        submitButton.disabled = false;
        submitButton.textContent = originalButtonText;
        
        // Hide processing status
        const statusDiv = document.getElementById('processing-status');
        if (statusDiv) {
            statusDiv.style.display = 'none';
        }
        
        console.log('[FRONTEND] Submission state reset');
    }
};

renderStructure();
