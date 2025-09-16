/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { authService } from '../auth/AuthService';
import { UserProfile } from '../types/auth';
import { sellerHistoryService } from '../api/SellerHistoryService';
import { bedrockAgentClient, BedrockAgentResponse } from '../api/BedrockAgentClient';
import { lambdaAgentClient, LambdaAgentResponse } from '../api/LambdaAgentClient';
import { bedrockAgentCoreClient } from '../api/BedrockAgentCoreClient';

// Check if we're running in Office context or standalone browser
let isInOfficeContext = false;
let currentUser: UserProfile | null = null;

try {
  isInOfficeContext = typeof Office !== 'undefined' && 
    Office.context && 
    Office.context.mailbox && 
    Office.context.host !== undefined;
    
  console.log("🔧 Script loaded. Office context:", Office?.context || "undefined");
  console.log("🔍 Is in real Office context:", isInOfficeContext);
} catch (error) {
  console.warn("⚠️ Error checking Office context:", error);
  isInOfficeContext = false;
}

// Handle callback routing
function handleCallbackRouting(): boolean {
  const currentPath = window.location.pathname;
  const hash = window.location.hash;
  console.log("🌍 Current path:", currentPath);
  console.log("🔗 Current hash:", hash);
  
  // Check if this is an authentication callback by looking for id_token in URL hash
  if (hash.includes('id_token=')) {
    console.log("🎯 Processing authentication callback...");
    
    // Let oidc-client handle the callback
    authService.handleCallback().then((user) => {
      console.log("✅ Callback processed successfully:", user);
      
      // Clean up the URL by removing the hash
      if (window.history && typeof window.history.replaceState === 'function') {
        const cleanUrl = `${window.location.protocol}//${window.location.host}${window.location.pathname}`;
        window.history.replaceState(null, '', cleanUrl);
      }
      
      // Update UI to show authenticated state
      updateAuthUI();
      
      // Initialize API UI after successful authentication
      initializeApiUI();
    }).catch((error) => {
      console.error("❌ Callback processing failed:", error);
      showError(`Authentication failed: ${error.message}`);
      
      // Clean up URL even on error
      if (window.history && typeof window.history.replaceState === 'function') {
        const cleanUrl = `${window.location.protocol}//${window.location.host}${window.location.pathname}`;
        window.history.replaceState(null, '', cleanUrl);
      }
    });
    
    return true; // Indicate we're processing a callback
  }
  
  return false; // Not a callback URL
}

// Initialize authentication when the page loads
async function initializeAuth() {
  try {
    console.log("🔐 Initializing authentication...");
    await authService.initialize();
    
    // Set up token expiration event listeners
    setupTokenEventListeners();
    
    // Check if user is already authenticated
    const user = await authService.getUser();
    console.log("👤 Current user:", user);
    
    // Update UI based on authentication state
    updateAuthUI();
    
    // If authenticated, immediately display user identity
    if (user) {
      console.log("✅ User authenticated, displaying identity...");
      displayUserIdentity();
    }
    
    console.log("✅ Authentication initialized successfully");
  } catch (error) {
    console.error("❌ Authentication initialization failed:", error);
    showError("Authentication setup failed: " + error.message);
  }
}

// Set up event listeners for token-related events
function setupTokenEventListeners() {
  // Token expiration warning
  authService.on('accessTokenExpiring', () => {
    console.log('⏰ Token expiring...');
    showInfo('Your session will expire soon. If you experience issues, please sign in again.');
  });

  // Token expired - show clear re-authentication prompt in Office Add-ins
  authService.on('accessTokenExpired', () => {
    console.log('❌ Token expired');
    showError('Your session has expired. Please sign in again.');
    showReAuthenticationPrompt();
    updateAuthUI(); // This will show the login interface
  });
  
  // Token expired from renewal failure
  authService.on('tokenExpired', () => {
    console.log('❌ Token expired - renewal not possible');
    showError('Your session has expired. Please sign in again.');
    showReAuthenticationPrompt();
    updateAuthUI(); // This will show the login interface
  });

  // Token renewal successful (only in non-Office environments)
  authService.on('tokenRenewed', (user) => {
    console.log('✅ Token renewed successfully', user);
    showSuccess('Session renewed successfully!');
    updateAuthUI();
  });

  // Listen for successful authentication to update UI immediately
  authService.on('authSuccess', (user) => {
    console.log('🎉 Authentication successful, updating UI...', user);
    updateAuthUI();
  });

  // Token renewal failed - user needs to login again
  authService.on('loginRequired', () => {
    console.log('🔐 Login required - redirecting to login');
    showError('Your session has expired. Please sign in again.');
    // Clear any existing user data
    const existingUserInfo = document.getElementById("user-identity-section");
    if (existingUserInfo) {
      existingUserInfo.remove();
    }
    updateAuthUI();
  });

  // Silent renewal error
  authService.on('silentRenewError', (error) => {
    console.error('🔄 Silent renewal error:', error);
    showError('Session renewal failed. You may need to sign in again.');
  });
}

function updateAuthUI() {
  const isAuthenticated = authService.isAuthenticated();
  const authState = authService.getAuthState();
  console.log("🎨 Updating auth UI. Authenticated:", isAuthenticated);
  console.log("🔍 Auth state details:", {
    isAuthenticated: authState.isAuthenticated,
    hasUser: !!authState.user,
    userName: authState.user?.name,
    isLoading: authState.isLoading,
    error: authState.error
  });

  // Update header to show user info or login
  updateHeader(isAuthenticated);
  
  // Show/hide main functionality based on auth state
  const appBody = document.getElementById("app-body");
  const sideloadMsg = document.getElementById("sideload-msg");
  const authLoading = document.getElementById("auth-loading");
  const authSignin = document.getElementById("auth-signin");
  
  console.log("🔍 DOM elements found:", {
    appBody: !!appBody,
    sideloadMsg: !!sideloadMsg,
    authLoading: !!authLoading,
    authSignin: !!authSignin
  });
  
  // Hide loading state
  if (authLoading) authLoading.style.display = "none";
  
  if (isAuthenticated) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    
    // Display user identity in the panel
    console.log("🆔 Displaying user identity panel...");
    displayUserIdentity();
    
    // Setup debug buttons for authenticated users
    setupDebugButtons();
  } else {
    if (appBody) appBody.style.display = "none";
    if (sideloadMsg) sideloadMsg.style.display = "block";
    if (authSignin) authSignin.style.display = "block";
    
    // Show login button
    const loginBtn = document.getElementById("login-btn");
    if (loginBtn) {
      loginBtn.style.display = "block";
      loginBtn.onclick = () => {
        console.log("🔑 Login button clicked");
        authService.login().catch(error => {
          console.error("❌ Login failed:", error);
          showError("Login failed: " + error.message);
        });
      };
    }
    
    // Add debug clear auth button functionality
    const clearAuthBtn = document.getElementById("clear-auth-btn");
    if (clearAuthBtn) {
      clearAuthBtn.onclick = async () => {
        console.log("🗑️ Clear auth button clicked");
        try {
          await authService.signOut();
          
          // Also manually clear any remaining OIDC data
          Object.keys(localStorage).forEach(key => {
            if (key.startsWith('oidc.') || key.includes('auth') || key.includes('token')) {
              localStorage.removeItem(key);
              console.log('🗑️ Cleared:', key);
            }
          });
          
          showSuccess("Auth data cleared! You can now test fresh login.");
          
          // Refresh the page to reset state
          setTimeout(() => {
            window.location.reload();
          }, 1000);
          
        } catch (error) {
          console.error("❌ Error clearing auth:", error);
          showError("Error clearing auth data: " + error.message);
        }
      };
    }
    
    // Remove any existing user identity panel
    const existingUserInfo = document.getElementById("user-identity-section");
    if (existingUserInfo) {
      existingUserInfo.remove();
    }
  }
}

function updateHeader(isAuthenticated: boolean) {
  const headerContent = document.querySelector('.header-content');
  if (!headerContent) return;

  // Keep header simple - just show the title
  headerContent.innerHTML = `
  `;
}

async function handleLogin() {
  try {
    console.log("🚀 Login button clicked");
    await authService.login();
  } catch (error) {
    console.error("❌ Login failed:", error);
    showError("Login failed. Please try again.");
  }
}

// Setup logout button
const logoutBtn = document.getElementById('logout-btn');
if (logoutBtn) {
  logoutBtn.addEventListener('click', async () => {
    try {
      console.log('🚪 User requested logout');
      showInfo('Signing out...');
      
      // Use signOut instead of logout for better Office Add-in compatibility
      await authService.signOut();
      showSuccess('Successfully signed out!');
      updateAuthUI();
    } catch (error) {
      console.error('❌ Logout failed:', error);
      showError('Logout failed. Please try again.');
    }
  });
}

function showError(message: string) {
  const errorDiv = document.createElement("div");
  errorDiv.className = "error-message";
  errorDiv.style.cssText = `
    background-color: #fef2f2;
    border: 1px solid #fecaca;
    color: #dc2626;
    padding: 12px;
    border-radius: 6px;
    margin: 10px 20px;
    font-size: 14px;
  `;
  errorDiv.textContent = message;
  
  const container = document.getElementById("app-body") || document.body;
  container.insertBefore(errorDiv, container.firstChild);
  
  // Remove error after 5 seconds
  setTimeout(() => {
    if (errorDiv.parentNode) {
      errorDiv.parentNode.removeChild(errorDiv);
    }
  }, 5000);
}

// Handle both Office and standalone contexts
if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    console.log("📋 Office ready:", info);
    initializeApp();
  });
} else {
  // Fallback for non-Office environments
  document.addEventListener('DOMContentLoaded', () => {
    console.log("🌐 DOM ready (standalone)");
    initializeApp();
  });
}

async function initializeApp() {
  console.log("🚀 Initializing application...");
  
  // Handle callback routing first
  const isCallback = handleCallbackRouting();
  if (isCallback) {
    return; // Exit early if processing callback
  }
  
  // Initialize authentication
  await initializeAuth();
  
  // Initialize API UI
  initializeApiUI();
  
  // Don't automatically run - let user click the button
  // if (isInOfficeContext) {
  //   await run(); // Automatically analyze current email
  // }
  
  console.log("✅ Application initialized");
}

export async function run() {
  /**
   * Enhanced Outlook add-in with LLM email summarization
   */
  
  // Check authentication first
  if (!authService.isAuthenticated()) {
    showError("Please sign in to analyze emails");
    return;
  }

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  
  // Clear previous content
  insertAt.innerHTML = "";
  
  // Create email info container
  const emailInfoContainer = document.createElement("div");
  emailInfoContainer.style.cssText = `
    background-color: #f8f9fa;
    border: 1px solid #e9ecef;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 20px;
  `;
  
  // Display email title
  const titleElement = document.createElement("div");
  titleElement.innerHTML = `<strong>📧 Subject:</strong> ${item.subject || "No subject"}`;
  titleElement.style.cssText = `
    font-size: 16px;
    margin-bottom: 12px;
    word-wrap: break-word;
  `;
  emailInfoContainer.appendChild(titleElement);
  
  // Display sender information
  const senderElement = document.createElement("div");
  if (item.from && item.from.emailAddress) {
    senderElement.innerHTML = `<strong>👤 From:</strong> ${item.from.displayName || item.from.emailAddress} &lt;${item.from.emailAddress}&gt;`;
  } else {
    senderElement.innerHTML = `<strong>👤 From:</strong> Information not available`;
  }
  senderElement.style.cssText = `
    font-size: 14px;
    margin-bottom: 15px;
    color: #666;
    word-wrap: break-word;
  `;
  emailInfoContainer.appendChild(senderElement);
  
  // Display To recipients
  const toRecipientsElement = document.createElement("div");
  if (item.to && item.to.length > 0) {
    const toRecipients = item.to.map(recipient => 
      `${recipient.displayName || recipient.emailAddress} &lt;${recipient.emailAddress}&gt;`
    ).join('; ');
    toRecipientsElement.innerHTML = `<strong>📨 To:</strong> ${toRecipients}`;
  } else {
    toRecipientsElement.innerHTML = `<strong>📨 To:</strong> No recipients`;
  }
  toRecipientsElement.style.cssText = `
    font-size: 14px;
    margin-bottom: 15px;
    color: #666;
    word-wrap: break-word;
  `;
  emailInfoContainer.appendChild(toRecipientsElement);
  
  // Display CC recipients
  const ccRecipientsElement = document.createElement("div");
  if (item.cc && item.cc.length > 0) {
    const ccRecipients = item.cc.map(recipient => 
      `${recipient.displayName || recipient.emailAddress} &lt;${recipient.emailAddress}&gt;`
    ).join('; ');
    ccRecipientsElement.innerHTML = `<strong>📋 CC:</strong> ${ccRecipients}`;
    ccRecipientsElement.style.cssText = `
      font-size: 14px;
      margin-bottom: 15px;
      color: #666;
      word-wrap: break-word;
    `;
    emailInfoContainer.appendChild(ccRecipientsElement);
  }
  
  insertAt.appendChild(emailInfoContainer);

  // Get email body for display and summarization
  try {
    // Get the email body
    item.body.getAsync("text", async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emailBody = result.value;
        
        // Extract first 20 words
        const words = emailBody.trim().split(/\s+/);
        const first20Words = words.slice(0, 20).join(' ');
        const hasMore = words.length > 20;
        
        // Display first 20 words
        const previewElement = document.createElement("div");
        previewElement.innerHTML = `<strong>📄 First 20 words:</strong><br/>${first20Words}${hasMore ? '...' : ''}`;
        previewElement.style.cssText = `
          background-color: #fff;
          border: 1px solid #dee2e6;
          border-radius: 6px;
          padding: 15px;
          margin-bottom: 15px;
          font-style: italic;
          line-height: 1.4;
        `;
        insertAt.appendChild(previewElement);
        
        // Display total length info
        const lengthInfo = document.createElement("div");
        lengthInfo.innerHTML = `<strong>📊 Email Stats:</strong> ${words.length} words, ${emailBody.length} characters`;
        lengthInfo.style.cssText = `
          font-size: 14px;
          color: #666;
          margin-bottom: 20px;
        `;
        insertAt.appendChild(lengthInfo);
        

        
        // Automatically trigger email logging and analysis
        await logEmailActivity(emailBody, insertAt, item);
        
        // Add divider
        const divider = document.createElement("hr");
        divider.style.cssText = `
          margin: 20px 0;
          border: none;
          border-top: 1px solid #dee2e6;
        `;
        insertAt.appendChild(divider);
        
      } else {
        const errorMsg = document.createElement("div");
        errorMsg.textContent = "Could not access email body";
        errorMsg.style.cssText = `
          color: #dc3545;
          background-color: #f8d7da;
          border: 1px solid #f5c6cb;
          padding: 12px;
          border-radius: 6px;
          margin-bottom: 15px;
        `;
        insertAt.appendChild(errorMsg);
      }
    });
  } catch (error) {
    const errorMsg = document.createElement("div");
    errorMsg.textContent = `Error: ${error.message}`;
    errorMsg.style.cssText = `
      color: #dc3545;
      background-color: #f8d7da;
      border: 1px solid #f5c6cb;
      padding: 12px;
      border-radius: 6px;
      margin-bottom: 15px;
    `;
    insertAt.appendChild(errorMsg);
  }
}

// Standalone testing function (simulates Office context)
export async function runStandalone() {
  console.log("🚀 runStandalone() called!");
  console.log("🔧 Running standalone test mode");
  
  // Check authentication first
  if (!authService.isAuthenticated()) {
    showError("Please sign in to log email activity");
    return;
  }
  
  let insertAt = document.getElementById("item-subject");
  console.log("📍 Insert target element:", insertAt);
  
  if (!insertAt) {
    console.error("❌ Cannot find item-subject element!");
    return;
  }
  
  // Clear previous content
  insertAt.innerHTML = "";
  console.log("🧹 Cleared previous content");
  
  // Display standalone mode message
  const standaloneMessage = document.createElement("div");
  standaloneMessage.style.cssText = `
    background-color: #fff3cd;
    border: 1px solid #ffeaa7;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 20px;
    text-align: center;
  `;
  
  standaloneMessage.innerHTML = `
    <h3 style="margin: 0 0 15px 0; color: #856404;">🔧 Standalone Test Mode</h3>
    <p style="margin: 0; color: #856404; font-size: 14px;">
      This mode is for testing outside of Outlook.<br/>
      In a real Outlook environment, the add-in would read actual email content.
    </p>
  `;
  
  insertAt.appendChild(standaloneMessage);
  
  // Add a simple test button
  const testButton = document.createElement("button");
  testButton.textContent = "📝 Test Log & Analyze";
  testButton.className = "ms-Button ms-Button--primary";
  testButton.style.cssText = `
    margin-bottom: 20px;
    padding: 12px 24px;
    font-size: 14px;
    display: block;
    margin-left: auto;
    margin-right: auto;
  `;
  
  testButton.onclick = () => {
    showBanner("Standalone test mode - no email content to log", false);
    console.log("📝 Standalone test mode - no real email data available");
  };
  
  insertAt.appendChild(testButton);
}




async function displayUserIdentity() {
  const greetingContainer = document.getElementById("user-greeting-container");
  if (!greetingContainer) return;

  try {
    const user = await authService.getUser();
    
    if (!user) return;

    // Clear previous user info
    const existingUserInfo = document.getElementById("user-identity-section");
    if (existingUserInfo) {
      existingUserInfo.remove();
    }

    // Create simple user greeting section
    const userIdentitySection = document.createElement("div");
    userIdentitySection.id = "user-identity-section";
    userIdentitySection.className = "user-greeting-panel";
    
    userIdentitySection.innerHTML = `
      <div class="user-greeting">
        <h2>Hi ${user.name || user.email || 'User'}!</h2>
      </div>
    `;

    // Insert into the dedicated greeting container
    greetingContainer.appendChild(userIdentitySection);

  } catch (error) {
    console.error("❌ Error displaying user identity:", error);
    showError("Failed to load user identity: " + error.message);
  }
}



// Helper function to show success messages
function showSuccess(message: string) {
  const successDiv = document.createElement("div");
  successDiv.className = "success-message";
  successDiv.style.cssText = `
    background-color: #f0f9f4;
    border: 1px solid #86efac;
    color: #059669;
    padding: 12px;
    border-radius: 6px;
    margin: 10px 20px;
    font-size: 14px;
    animation: slideIn 0.3s ease-out;
  `;
  successDiv.textContent = message;
  
  const container = document.getElementById("app-body") || document.body;
  container.insertBefore(successDiv, container.firstChild);
  
  // Remove success message after 3 seconds
  setTimeout(() => {
    if (successDiv.parentNode) {
      successDiv.parentNode.removeChild(successDiv);
    }
  }, 3000);
}

// Helper function to show info messages
function showInfo(message: string) {
  const infoDiv = document.createElement("div");
  infoDiv.className = "info-message";
  infoDiv.style.cssText = `
    background-color: #f0f9ff;
    border: 1px solid #93c5fd;
    color: #1d4ed8;
    padding: 12px;
    border-radius: 6px;
    margin: 10px 20px;
    font-size: 14px;
    animation: slideIn 0.3s ease-out;
  `;
  infoDiv.textContent = message;
  
  const container = document.getElementById("app-body") || document.body;
  container.insertBefore(infoDiv, container.firstChild);
  
  // Remove info message after 4 seconds
  setTimeout(() => {
    if (infoDiv.parentNode) {
      infoDiv.parentNode.removeChild(infoDiv);
    }
  }, 4000);
}

// Show a prominent re-authentication prompt
function showReAuthenticationPrompt() {
  // Remove any existing prompt
  const existingPrompt = document.getElementById("reauth-prompt");
  if (existingPrompt) {
    existingPrompt.remove();
  }

  const promptDiv = document.createElement("div");
  promptDiv.id = "reauth-prompt";
  promptDiv.className = "reauth-prompt";
  promptDiv.style.cssText = `
    background: linear-gradient(135deg, #f59e0b, #d97706);
    border: 2px solid #b45309;
    color: white;
    padding: 20px;
    border-radius: 12px;
    margin: 15px 20px;
    font-size: 16px;
    text-align: center;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    animation: slideIn 0.3s ease-out, pulse 2s infinite;
  `;
  
  promptDiv.innerHTML = `
    <div style="margin-bottom: 15px;">
      <strong>🔐 Session Expired</strong>
    </div>
    <div style="margin-bottom: 20px; font-size: 14px; opacity: 0.9;">
      Your authentication session has expired and couldn't be renewed automatically.<br>
      Please sign in again to continue using the application.
    </div>
    <button id="reauth-btn" style="
      background: white;
      color: #d97706;
      border: none;
      padding: 12px 24px;
      border-radius: 8px;
      font-weight: bold;
      cursor: pointer;
      font-size: 14px;
      transition: transform 0.2s ease;
    " onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
      🔓 Sign In Again
    </button>
  `;
  
  const container = document.getElementById("app-body") || document.body;
  container.insertBefore(promptDiv, container.firstChild);
  
  // Add click handler for the re-authentication button
  const reAuthBtn = document.getElementById("reauth-btn");
  if (reAuthBtn) {
    reAuthBtn.onclick = async () => {
      try {
        console.log("🔓 Re-authentication button clicked");
        await authService.login();
      } catch (error) {
        console.error("❌ Re-authentication failed:", error);
        showError("Re-authentication failed. Please try again.");
      }
    };
  }
  
  // Auto-remove after 30 seconds if user doesn't act
  setTimeout(() => {
    if (promptDiv.parentNode) {
      promptDiv.parentNode.removeChild(promptDiv);
    }
  }, 30000);
}

// Global variable to store the current due date
let currentDueDate: string | null = null;

function initializeDueDateUI() {
  const dueDateInput = document.getElementById("due-date-input") as HTMLInputElement;
  const clearDueDateBtn = document.getElementById("clear-due-date-btn") as HTMLButtonElement;
  const dueDateDisplay = document.getElementById("due-date-display") as HTMLDivElement;

  if (!dueDateInput || !clearDueDateBtn || !dueDateDisplay) {
    console.warn("Due date UI elements not found");
    return;
  }

  // Function to format date to ISO 8601 format
  function formatToISO8601(dateTimeLocalValue: string): string {
    if (!dateTimeLocalValue) return "";
    
    // Create a Date object from the datetime-local input value
    // Note: datetime-local returns format: YYYY-MM-DDTHH:mm
    const date = new Date(dateTimeLocalValue);
    
    // Convert to ISO 8601 format with milliseconds and Z suffix
    return date.toISOString();
  }

  // Function to update the display and global variable
  function updateDueDateDisplay() {
    const inputValue = dueDateInput.value;
    if (inputValue) {
      currentDueDate = formatToISO8601(inputValue);
      dueDateDisplay.textContent = `📅 Due Date: ${currentDueDate}`;
      dueDateDisplay.style.color = "#28a745";
      dueDateDisplay.style.fontWeight = "600";
    } else {
      currentDueDate = null;
      dueDateDisplay.textContent = "ISO 8601 format will appear here...";
      dueDateDisplay.style.color = "#6c757d";
      dueDateDisplay.style.fontWeight = "normal";
    }
  }

  // Event listeners
  dueDateInput.addEventListener("change", updateDueDateDisplay);
  dueDateInput.addEventListener("input", updateDueDateDisplay);

  clearDueDateBtn.addEventListener("click", () => {
    dueDateInput.value = "";
    updateDueDateDisplay();
  });

  // Initialize display
  updateDueDateDisplay();
}

// Function to get current due date for use in email analysis
function getCurrentDueDate(): string | null {
  return currentDueDate;
}

// Function to show success/failure banners
function showBanner(message: string, isSuccess: boolean = true) {
  // Remove any existing banners
  const existingBanner = document.getElementById("activity-banner");
  if (existingBanner) {
    existingBanner.remove();
  }

  const banner = document.createElement("div");
  banner.id = "activity-banner";
  banner.innerHTML = `<strong>${isSuccess ? '✅' : '❌'}</strong> ${message}`;
  banner.style.cssText = `
    position: fixed;
    top: 10px;
    left: 50%;
    transform: translateX(-50%);
    z-index: 9999;
    padding: 12px 20px;
    border-radius: 6px;
    font-size: 14px;
    font-weight: 500;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    max-width: 90%;
    text-align: center;
    animation: slideInFromTop 0.3s ease-out;
    ${isSuccess ? 
      'background: #d4edda; color: #155724; border: 1px solid #c3e6cb;' : 
      'background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb;'}
  `;

  document.body.appendChild(banner);

  // Auto-remove after 5 seconds
  setTimeout(() => {
    if (banner.parentNode) {
      banner.style.animation = "slideOutToTop 0.3s ease-in";
      setTimeout(() => {
        if (banner.parentNode) {
          banner.parentNode.removeChild(banner);
        }
      }, 300);
    }
  }, 5000);
}

// Function to log email activity instead of AI analysis
async function logEmailActivity(emailContent: string, container: HTMLElement, emailItem: any) {
  if (!authService.isAuthenticated()) {
    showError("Please sign in to log email activity");
    showBanner("Authentication required for email logging", false);
    return;
  }

  // Show loading message
  let loadingMsg = document.createElement("p");
  loadingMsg.textContent = "📝 Logging email activity...";
  loadingMsg.id = "loading-msg";
  container.appendChild(loadingMsg);

  try {
    const user = await authService.getUser();
    const dueDate = getCurrentDueDate();
    
    // Prepare email data for logging
    const emailData = {
      timestamp: new Date().toISOString(),
      user: user?.name || "Unknown User",
      emailLength: emailContent.length,
      wordCount: emailContent.trim().split(/\s+/).length,
      subject: emailItem?.subject || "No Subject",
      hasAttachments: emailItem?.attachments?.length > 0 || false,
      dueDate: dueDate,
      status: "WORK IN PROGRESS",
      activityType: "email_logged"
    };

    // Log to console (in a real app, this would go to a logging service)
    console.log("📧 EMAIL ACTIVITY LOGGED:");
    console.log("🕐 Timestamp:", emailData.timestamp);
    console.log("👤 User:", emailData.user);
    console.log("📄 Subject:", emailData.subject);
    console.log("📊 Email Length:", emailData.emailLength, "characters");
    console.log("📊 Word Count:", emailData.wordCount, "words");
    console.log("📎 Has Attachments:", emailData.hasAttachments);
    console.log("🚧 Status:", emailData.status);
    if (emailData.dueDate) {
      console.log("📅 Due Date:", emailData.dueDate);
    }
    console.log("📝 Activity Type:", emailData.activityType);
    console.log("📄 Email Preview:", emailContent.substring(0, 100) + "...");

    // Create simple loading UI for Lambda call
    if (loadingMsg) {
      loadingMsg.innerHTML = `
        <div style="display: flex; align-items: center; gap: 10px; padding: 15px;">
          <div class="loading-spinner" style="
            width: 20px; 
            height: 20px; 
            border: 2px solid #e3e3e3; 
            border-top: 2px solid #007acc; 
            border-radius: 50%; 
            animation: spin 1s linear infinite;
          "></div>
          <span>🤖 Processing...</span>
        </div>
      `;
    }

    // Call agent for brief email summary
    let agentResponse = "";
    try {
      // Extract sender email for seller lookup
      const senderEmail = emailItem?.from?.emailAddress || "";
      
      const summaryPrompt = `Create email activity with brief summary:

Subject: ${emailData.subject}
${emailData.dueDate ? `Due Date: ${emailData.dueDate}` : ''}
${senderEmail ? `Sender Email: ${senderEmail}` : ''}

Email Content:
${emailContent}

Instructions:
- Set status to WORK IN PROGRESS by default
- FIRST: Call the seller lookup tool with contactType="EMAIL" and contactValue="${senderEmail}" to get the seller name
- Include the seller name in the activity description (e.g., "Email activity from [Seller Name]:")
- Provide only a brief summary of this email in 2-3 sentences
- If seller lookup fails, proceed without seller name`;

      console.log("🤖 Calling Bedrock Agent Core for brief summary...");
      agentResponse = await invokeAgent(summaryPrompt);
      console.log("✅ Bedrock Agent Core summary completed");
    } catch (error) {
      console.error("⚠️ Agent summary failed:", error);
      agentResponse = "Brief summary unavailable at the moment.";
    }

    // Remove loading message
    const loading = document.getElementById("loading-msg");
    if (loading) loading.remove();

    // Create main results container
    let mainContainer = document.createElement("div");
    mainContainer.style.cssText = `
      margin-top: 15px;
    `;

    // Create AI analysis section
    let aiContainer = document.createElement("div");
    aiContainer.style.cssText = `
      border: 1px solid #28a745;
      border-radius: 6px;
      padding: 15px;
      margin-bottom: 15px;
      background: #f8fff9;
      box-shadow: 0 1px 4px rgba(40,167,69,0.1);
    `;

    let aiTitle = document.createElement("h4");
    aiTitle.textContent = "📝 Brief Summary";
    aiTitle.style.cssText = `
      margin: 0 0 10px 0;
      color: #28a745;
      font-size: 16px;
    `;
    aiContainer.appendChild(aiTitle);

    let aiResponse = document.createElement("div");
    aiResponse.textContent = agentResponse;
    aiResponse.style.cssText = `
      color: #333;
      line-height: 1.5;
      font-size: 14px;
      white-space: pre-wrap;
    `;
    aiContainer.appendChild(aiResponse);

    mainContainer.appendChild(aiContainer);

    // Create activity log section
    let logContainer = document.createElement("div");
    logContainer.style.cssText = `
      border: 1px solid #17a2b8;
      border-radius: 6px;
      padding: 15px;
      background: #f1f9ff;
      box-shadow: 0 1px 4px rgba(23,162,184,0.1);
    `;

    let logTitle = document.createElement("h4");
    logTitle.textContent = "📝 Activity Logged";
    logTitle.style.cssText = `
      margin: 0 0 10px 0;
      color: #17a2b8;
      font-size: 16px;
    `;
    logContainer.appendChild(logTitle);

    let logDetails = document.createElement("div");
    logDetails.innerHTML = `
      <p style="margin: 5px 0; font-size: 13px;"><strong>Timestamp:</strong> ${emailData.timestamp}</p>
      <p style="margin: 5px 0; font-size: 13px;"><strong>User:</strong> ${emailData.user}</p>
      <p style="margin: 5px 0; font-size: 13px;"><strong>Status:</strong> <span style="color: #ffc107; font-weight: 600;">🚧 ${emailData.status}</span></p>
      <p style="margin: 5px 0; font-size: 13px;"><strong>Email Stats:</strong> ${emailData.wordCount} words, ${emailData.emailLength} characters</p>
      ${emailData.dueDate ? `<p style="margin: 5px 0; font-size: 13px;"><strong>Due Date:</strong> ${emailData.dueDate}</p>` : ''}
      <p style="margin: 5px 0; font-size: 13px;"><strong>Activity:</strong> Email activity created with brief summary</p>
    `;
    logDetails.style.cssText = `
      color: #333;
      line-height: 1.4;
    `;
    logContainer.appendChild(logDetails);

    mainContainer.appendChild(logContainer);
    container.appendChild(mainContainer);

    // Show success banner
    showBanner("Email activity created with summary!", true);

    console.log("✅ Email activity created with brief summary successfully");

  } catch (error) {
    // Remove loading message
    const loading = document.getElementById("loading-msg");
    if (loading) loading.remove();

    console.error("❌ Email activity logging failed:", error);
    console.error("❌ Error details:", error.message);
    
    showError(`Email activity logging failed: ${error.message}`);
    showBanner("Failed to log email activity", false);
  }
}

function initializeApiUI() {
  // Initialize due date functionality
  initializeDueDateUI();

  // Get Seller Metrics button
  const getSellerHistoryBtn = document.getElementById("get-seller-history-btn");
  if (getSellerHistoryBtn) {
    getSellerHistoryBtn.onclick = handleGetSellerHistory;
  }

  // Invoke Agent button
  const invokeAgentBtn = document.getElementById("invoke-agent-btn") as HTMLButtonElement;
  if (invokeAgentBtn) {
    invokeAgentBtn.onclick = handleInvokeAgent;
  }

  // Analyze Email button
  const analyzeEmailBtn = document.getElementById("analyze-email-btn") as HTMLButtonElement;
  if (analyzeEmailBtn) {
    analyzeEmailBtn.onclick = async () => {
      console.log("📧 Analyze Email button clicked!");
      console.log("🔍 isInOfficeContext:", isInOfficeContext);
      
      const buttonLabel = analyzeEmailBtn.querySelector('.ms-Button-label');
      const originalText = buttonLabel ? buttonLabel.textContent : '📧 Analyze Current Email';
      
      try {
        // Show loading state
        if (buttonLabel) buttonLabel.textContent = '⏳ Analyzing...';
        analyzeEmailBtn.disabled = true;
        
        // Check Office context dynamically at runtime
        const hasOfficeContext = typeof Office !== 'undefined' && 
          Office.context && 
          Office.context.mailbox && 
          Office.context.mailbox.item;
        
        console.log("🔍 Dynamic Office check:");
        console.log("  - Office available:", typeof Office !== 'undefined');
        console.log("  - Office.context:", !!Office?.context);
        console.log("  - Office.context.mailbox:", !!Office?.context?.mailbox);
        console.log("  - Office.context.mailbox.item:", !!Office?.context?.mailbox?.item);
        console.log("  - hasOfficeContext:", hasOfficeContext);
        
        if (hasOfficeContext) {
          console.log("🏢 Running in Office context - reading real email");
          await run();
        } else {
          console.log("🌐 Running in standalone mode - no email data available");
          await runStandalone();
        }
        
        console.log("✅ Email analysis completed");
        
      } catch (error) {
        console.error("❌ Email analysis failed:", error);
        showError(`Email analysis failed: ${error.message}`);
      } finally {
        // Reset button in both success and error cases
        if (buttonLabel) buttonLabel.textContent = originalText;
        analyzeEmailBtn.disabled = false;
      }
    };
  }

  // Show appropriate status message with dynamic check
  const emailStatus = document.getElementById("email-status");
  if (emailStatus) {
    // Check Office context dynamically for status message
    const hasOfficeAtInit = typeof Office !== 'undefined' && 
      Office.context && 
      Office.context.mailbox;
    
    if (hasOfficeAtInit) {
      emailStatus.innerHTML = "✅ Connected to Outlook - Click button to analyze the current email";
    } else {
      emailStatus.innerHTML = "⚠️ Office context checking... Click button to analyze (will auto-detect context)";
    }
  }
}

async function handleInvokeAgent() {
  try {
    if (!authService.isAuthenticated()) {
      showError("Please sign in first before invoking the agent.");
      return;
    }

    const agentInput = document.getElementById("agent-input") as HTMLTextAreaElement;
    const invokeAgentBtn = document.getElementById("invoke-agent-btn") as HTMLButtonElement;
    const agentResults = document.getElementById("agent-results");
    const agentResponse = document.getElementById("agent-response");

    if (!agentInput || !agentResults || !agentResponse) {
      showError("Required UI elements not found.");
      return;
    }

    const inputText = agentInput.value.trim();
    if (!inputText) {
      showError("Please enter a message for the agent.");
      return;
    }

    // Show enhanced loading state
    const buttonLabel = invokeAgentBtn.querySelector('.ms-Button-label');
    
    if (buttonLabel) buttonLabel.textContent = '⏳ Invoking Bedrock Agent...';
    invokeAgentBtn.disabled = true;

    // Show simple loading in results area
    agentResults.style.display = "block";
    agentResponse.innerHTML = `
      <div style="display: flex; align-items: center; gap: 10px; padding: 20px;">
        <div class="loading-spinner" style="
          width: 24px; 
          height: 24px; 
          border: 3px solid #e3e3e3; 
          border-top: 3px solid #007acc; 
          border-radius: 50%; 
          animation: spin 1s linear infinite;
        "></div>
        <span>🤖 Processing...</span>
      </div>
    `;

    console.log("🤖 Invoking agent with input:", inputText);
    showInfo("Invoking Bedrock Agent Core...");

    try {
      // Call the agent invocation function
      const response = await invokeAgent(inputText);
      
      console.log("✅ Agent invocation successful:", response);
      showSuccess("Bedrock Agent Core invocation completed successfully!");
      
      // Display results
      agentResponse.innerHTML = `
        <div style="white-space: pre-wrap; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.5; padding: 15px; background: #f8fff9; border: 1px solid #28a745; border-radius: 6px; margin-top: 10px;">
          ${response.replace(/\n/g, '<br>')}
        </div>
      `;

    } catch (error) {
      console.error("❌ Agent invocation failed:", error);
      const errorMessage = (error as Error).message;
      
      showError(`Bedrock Agent Core invocation failed: ${errorMessage}`);
      
      // Show error in results area
      agentResponse.innerHTML = `
        <div style="padding: 15px; background: #fff5f5; border: 1px solid #f56565; border-radius: 6px; margin-top: 10px; color: #c53030;">
          <strong>❌ Error:</strong> ${errorMessage}
        </div>
      `;
    }

  } catch (error) {
    console.error("❌ Unexpected error in handleInvokeAgent:", error);
    showError(`Unexpected error: ${(error as Error).message}`);
  } finally {
    // Reset button state
    const invokeAgentBtn = document.getElementById("invoke-agent-btn") as HTMLButtonElement;
    const buttonLabel = invokeAgentBtn?.querySelector('.ms-Button-label');
    
    if (buttonLabel) buttonLabel.textContent = '🤖 Invoke AI Agent';
    if (invokeAgentBtn) invokeAgentBtn.disabled = false;
  }
}

async function invokeAgent(inputText: string): Promise<string> {
  try {
    console.log("🔧 Calling Bedrock Agent Core...");
    
    const response = await bedrockAgentCoreClient.invoke(inputText);
    console.log("✅ Bedrock Agent Core response received:", response);
    
    return response.response || "Agent response received but no content.";

  } catch (error) {
    console.error("❌ Error calling Bedrock Agent Core:", error);
    throw new Error(`Bedrock Agent Core call failed: ${(error as Error).message}`);
  }
}

async function handleGetSellerHistory() {
  try {
    if (!authService.isAuthenticated()) {
      showError("Please sign in first before calling the API.");
      return;
    }

    const marketplaceIdInput = document.getElementById("marketplace-id") as HTMLInputElement;

    const params = {
      ...(marketplaceIdInput?.value && { marketplaceId: marketplaceIdInput.value })
    };

    console.log("📞 Calling seller metrics API...");
    showApiLoading(true);

    const result = await sellerHistoryService.getSellerHistory(params);
    
    console.log("✅ API call successful:", result);
    showSuccess("Seller metrics retrieved successfully!");
    displayApiResults(result);

  } catch (error) {
    console.error("❌ API call failed:", error);
    const errorMessage = (error as Error).message;
    
    // Provide specific guidance for authentication errors
    if (errorMessage.includes('session has expired') || 
        errorMessage.includes('Authentication failed')) {
      showError(`${errorMessage} Please click "Sign In" to continue.`);
      updateAuthUI(); // Refresh the UI to show sign-in options
    } else {
      showError(`API call failed: ${errorMessage}`);
    }
    
    hideApiResults();
  } finally {
    showApiLoading(false);
  }
}

function displayApiResults(data: any) {
  const resultsDiv = document.getElementById("api-results");
  const outputPre = document.getElementById("api-output");
  const metricsDiv = document.getElementById("results-metrics-display");
  
  if (resultsDiv && outputPre && metricsDiv) {
    // Store raw JSON
    outputPre.textContent = JSON.stringify(data, null, 2);
    
    // Generate metrics display
    generateKeyMetrics(data, metricsDiv);
    
    // Show results
    resultsDiv.style.display = "block";
    
    // Setup tab functionality
    setupTabNavigation();
  }
}

function generateKeyMetrics(data: any, metricsDiv: HTMLElement) {
  const metrics = extractSellerMetrics(data);
  
  var html = '';
  html += '<div style="';
  html += 'display: -webkit-box; display: -ms-flexbox; display: flex;';
  html += '-webkit-box-wrap: wrap; -ms-flex-wrap: wrap; flex-wrap: wrap;';
  html += 'margin: -2px;';
  html += 'padding: 8px;';
  html += '">';
  
  for (var i = 0; i < metrics.length; i++) {
    var metric = metrics[i];
    
    var cardStyle = '';
    cardStyle += 'background: #f8f9fa;';
    cardStyle += 'border: 1px solid #e9ecef;';
    cardStyle += 'border-radius: 4px;';
    cardStyle += 'padding: 8px 6px;';
    cardStyle += 'text-align: left;';
    cardStyle += 'margin: 2px;';
    cardStyle += 'min-width: 90px;';
    cardStyle += 'max-width: none;';
    // More flexible width calculation
    cardStyle += 'width: -webkit-calc(50% - 4px);';
    cardStyle += 'width: -moz-calc(50% - 4px);';
    cardStyle += 'width: calc(50% - 4px);';
    cardStyle += '-webkit-box-sizing: border-box;';
    cardStyle += '-moz-box-sizing: border-box;';
    cardStyle += 'box-sizing: border-box;';
    cardStyle += 'cursor: pointer;';
    cardStyle += '-webkit-transition: all 0.2s ease;';
    cardStyle += '-moz-transition: all 0.2s ease;';
    cardStyle += '-o-transition: all 0.2s ease;';
    cardStyle += 'transition: all 0.2s ease;';
    
    var hoverEvents = '';
    hoverEvents += 'onmouseover="this.style.background=\'#e9ecef\'; ';
    hoverEvents += 'if (this.style.transform !== undefined) { this.style.transform=\'translateY(-1px)\'; } ';
    hoverEvents += 'this.style.boxShadow=\'0 2px 4px rgba(0,0,0,0.1)\'" ';
    hoverEvents += 'onmouseout="this.style.background=\'#f8f9fa\'; ';
    hoverEvents += 'if (this.style.transform !== undefined) { this.style.transform=\'translateY(0)\'; } ';
    hoverEvents += 'this.style.boxShadow=\'none\'" ';
    
    // Compact single line with abbreviations for better fit
    var shortLabel = metric.label
      .replace('Total Revenue', 'Revenue')
      .replace('Units Sold', 'Units')
      .replace('Buy Box Win Rate', 'Buy Box')
      .replace('Active Products', 'Products')
      .replace('Total Activities', 'Activities')
      .replace('Avg Selling Price', 'Avg Price');
    
    var singleLineContent = metric.icon + ' ' + shortLabel + ' ' + metric.value;
    if (metric.trend) {
      var trendArrow = '→';
      if (metric.trend.direction === 'up') {
        trendArrow = '↗';
      } else if (metric.trend.direction === 'down') {
        trendArrow = '↘';
      }
      singleLineContent += ' ' + trendArrow + metric.trend.text;
    }
    
    html += '<div style="' + cardStyle + '" ' + hoverEvents + '>';
    html += '<div style="';
    html += 'font-size: 11px;';
    html += 'color: #495057;';
    html += 'font-weight: 500;';
    html += 'line-height: 1.3;';
    html += 'word-wrap: break-word;';
    html += 'overflow-wrap: break-word;';
    html += '">' + singleLineContent + '</div>';
    html += '</div>';
  }
  
  html += '</div>';
  metricsDiv.innerHTML = html;
}

function extractSellerMetrics(data: any): MetricItem[] {
  const metrics: MetricItem[] = [];
  
  // Focus on the most important business metrics from seller_info
  const sellerInfo = data.seller_info || {};
  
  // 1. Total GMS (Revenue)
  if (sellerInfo.s_total_gms) {
    var trend = undefined;
    if (sellerInfo.s_total_gms_momp) {
      var isNegative = sellerInfo.s_total_gms_momp.indexOf('-') !== -1;
      trend = {
        direction: isNegative ? 'down' : 'up',
        text: sellerInfo.s_total_gms_momp + ' MoM'
      };
    }
    
    metrics.push({
      icon: '💰',
      label: 'Total Revenue',
      value: sellerInfo.s_total_gms,
      category: 'finance',
      trend: trend
    });
  }
  
  // 2. Total Ordered Units
  if (sellerInfo.total_ordered_units) {
    var trend = undefined;
    if (sellerInfo.total_ordered_units_momp) {
      var isNegative = sellerInfo.total_ordered_units_momp.indexOf('-') !== -1;
      trend = {
        direction: isNegative ? 'down' : 'up',
        text: sellerInfo.total_ordered_units_momp + ' MoM'
      };
    }
    
    metrics.push({
      icon: '📦',
      label: 'Units Sold',
      value: sellerInfo.total_ordered_units,
      category: 'sales',
      trend: trend
    });
  }
  
  // 3. Buy Box Win Rate
  if (sellerInfo.buy_box_win_rate) {
    const winRate = (parseFloat(sellerInfo.buy_box_win_rate) * 100).toFixed(1);
    metrics.push({
      icon: '🎯',
      label: 'Buy Box Win Rate',
      value: winRate + '%',
      category: 'performance'
    });
  }
  
  // 4. FBA Buyable ASINs
  if (sellerInfo.fba_buyable_asin_count_3p) {
    metrics.push({
      icon: '📋',
      label: 'Buyable Asins',
      value: sellerInfo.fba_buyable_asin_count_3p,
      category: 'inventory'
    });
  }
  
  // 5. IPI Score
  if (sellerInfo.ipi_scr) {
    metrics.push({
      icon: '📊',
      label: 'IPI Score',
      value: sellerInfo.ipi_scr,
      category: 'performance'
    });
  }
  
  // 6. Total Activities
  if (sellerInfo.activity_total_count) {
    metrics.push({
      icon: '📞',
      label: 'Total Activities',
      value: sellerInfo.activity_total_count,
      category: 'engagement'
    });
  }
  
  // 7. Marketplace
  if (sellerInfo.home_marketplace_id) {
    metrics.push({
      icon: '🌍',
      label: 'Marketplace',
      value: sellerInfo.home_marketplace_id,
      category: 'general'
    });
  }
  
  // 8. Brand Category
  if (sellerInfo.merchant_primary_pg_desc) {
    metrics.push({
      icon: '🏷️',
      label: 'Category',
      value: sellerInfo.merchant_primary_pg_desc,
      category: 'general'
    });
  }
  
  // 9. Average Selling Price
  if (sellerInfo.net_ordered_asp) {
    const asp = parseFloat(sellerInfo.net_ordered_asp);
    metrics.push({
      icon: '💴',
      label: 'Avg Selling Price',
      value: '¥' + asp.toLocaleString(),
      category: 'finance'
    });
  }
  
  // 10. FBA Adoption Status
  if (sellerInfo.fba_adoption_status) {
    metrics.push({
      icon: '🚚',
      label: 'FBA Status',
      value: sellerInfo.fba_adoption_status === 'Y' ? 'Active' : 'Inactive',
      category: 'general'
    });
  }
  
  // Return only first 10 metrics for compatibility
  var result = [];
  for (var i = 0; i < Math.min(metrics.length, 10); i++) {
    result.push(metrics[i]);
  }
  return result;
}

// Helper interfaces and functions
interface MetricItem {
  icon: string;
  label: string;
  value: string;
  category: string;
  trend?: {
    direction: 'up' | 'down' | 'neutral';
    text: string;
  };
}

function getMetricColor(index: number) {
  const colors = [
    { gradient: 'linear-gradient(135deg, #667eea, #764ba2)', shadow: 'rgba(102, 126, 234, 0.3)' },
    { gradient: 'linear-gradient(135deg, #f093fb, #f5576c)', shadow: 'rgba(240, 147, 251, 0.3)' },
    { gradient: 'linear-gradient(135deg, #4facfe, #00f2fe)', shadow: 'rgba(79, 172, 254, 0.3)' },
    { gradient: 'linear-gradient(135deg, #43e97b, #38f9d7)', shadow: 'rgba(67, 233, 123, 0.3)' },
    { gradient: 'linear-gradient(135deg, #fa709a, #fee140)', shadow: 'rgba(250, 112, 154, 0.3)' },
    { gradient: 'linear-gradient(135deg, #a8edea, #fed6e3)', shadow: 'rgba(168, 237, 234, 0.3)' },
    { gradient: 'linear-gradient(135deg, #ff9a9e, #fecfef)', shadow: 'rgba(255, 154, 158, 0.3)' },
    { gradient: 'linear-gradient(135deg, #a1c4fd, #c2e9fb)', shadow: 'rgba(161, 196, 253, 0.3)' },
    { gradient: 'linear-gradient(135deg, #ffecd2, #fcb69f)', shadow: 'rgba(255, 236, 210, 0.3)' },
    { gradient: 'linear-gradient(135deg, #e0c3fc, #9bb5ff)', shadow: 'rgba(224, 195, 252, 0.3)' }
  ];
  
  return colors[index % colors.length];
}

function formatCurrency(value: number): string {
  if (value >= 1000000) {
    return `$${(value / 1000000).toFixed(1)}M`;
  } else if (value >= 1000) {
    return `$${(value / 1000).toFixed(1)}K`;
  } else {
    return `$${value.toFixed(2)}`;
  }
}

function formatNumberValue(value: number, key: string): string {
  if (/revenue|sales|amount|price|cost|total|value|earning|dollar/i.test(key)) {
    return formatCurrency(value);
  } else if (/percent|rate|ratio/i.test(key)) {
    return `${value.toFixed(1)}%`;
  } else if (Number.isInteger(value)) {
    return value.toLocaleString();
  } else {
    return value.toFixed(2);
  }
}

function formatFieldName(key: string): string {
  return key.split(/[_-]/).map(word => 
    word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
  ).join(' ');
}

function getCategoryForField(key: string): string {
  if (/revenue|sales|amount|price|cost|total|value|earning/i.test(key)) return 'Revenue';
  if (/count|quantity|number|volume/i.test(key)) return 'Volume';
  if (/date|time|created|updated/i.test(key)) return 'Timeline';
  if (/status|state|type|category/i.test(key)) return 'Status';
  if (/id|identifier|key/i.test(key)) return 'Identity';
  return 'General';
}

function getNumberIcon(key: string): string {
  if (/revenue|sales|amount|price|cost|total|value|earning/i.test(key)) return '💰';
  if (/count|quantity|number|volume/i.test(key)) return '🔢';
  if (/percent|rate|ratio/i.test(key)) return '📊';
  return '📈';
}

function getStringIcon(key: string): string {
  if (/name|title/i.test(key)) return '🏷️';
  if (/status|state/i.test(key)) return '🔘';
  if (/id|identifier/i.test(key)) return '🆔';
  if (/email|mail/i.test(key)) return '📧';
  if (/phone|tel/i.test(key)) return '📞';
  return '📝';
}

function setupTabNavigation() {
  var metricsTab = document.getElementById('metrics-tab');
  var rawTab = document.getElementById('raw-tab');
  var metricsContent = document.getElementById('metrics-content');
  var rawContent = document.getElementById('raw-content');
  
  if (metricsTab && rawTab && metricsContent && rawContent) {
    metricsTab.onclick = function() {
      // Reset all tabs
      metricsTab.style.background = 'white';
      metricsTab.style.color = '#0078d4';
      metricsTab.style.borderBottom = '2px solid #0078d4';
      rawTab.style.background = '#f8f9fa';
      rawTab.style.color = '#6c757d';
      rawTab.style.borderBottom = '2px solid transparent';
      
      // Show/hide content
      metricsContent.style.display = 'block';
      rawContent.style.display = 'none';
    };
    
    rawTab.onclick = function() {
      // Reset all tabs
      metricsTab.style.background = '#f8f9fa';
      metricsTab.style.color = '#6c757d';
      metricsTab.style.borderBottom = '2px solid transparent';
      rawTab.style.background = 'white';
      rawTab.style.color = '#0078d4';
      rawTab.style.borderBottom = '2px solid #0078d4';
      
      // Show/hide content
      metricsContent.style.display = 'none';
      rawContent.style.display = 'block';
    };
  }
}

function hideApiResults() {
  const resultsDiv = document.getElementById("api-results");
  if (resultsDiv) {
    resultsDiv.style.display = "none";
  }
}

function showApiLoading(isLoading: boolean) {
  const getSellerHistoryBtn = document.getElementById("get-seller-history-btn");
  
  if (getSellerHistoryBtn) {
    const label = getSellerHistoryBtn.querySelector('.ms-Button-label');
    if (label) {
      label.textContent = isLoading ? "⏳ Loading..." : "📈 Get Seller Metrics";
    }
    (getSellerHistoryBtn as HTMLButtonElement).disabled = isLoading;
  }
}

// Setup debug buttons for authenticated users
function setupDebugButtons() {
  // Clear auth button (when logged in)
  const clearAuthLoggedInBtn = document.getElementById("clear-auth-logged-in-btn");
  if (clearAuthLoggedInBtn) {
    clearAuthLoggedInBtn.onclick = async () => {
      console.log("🗑️ Clear auth (logged in) button clicked");
      try {
        await authService.signOut();
        
        // Also manually clear any remaining OIDC data
        Object.keys(localStorage).forEach(key => {
          if (key.startsWith('oidc.') || key.includes('auth') || key.includes('token')) {
            localStorage.removeItem(key);
            console.log('🗑️ Cleared:', key);
          }
        });
        
        showSuccess("Auth data cleared! Refreshing to test fresh login...");
        
        // Refresh the page to reset state
        setTimeout(() => {
          window.location.reload();
        }, 1000);
        
      } catch (error) {
        console.error("❌ Error clearing auth:", error);
        showError("Error clearing auth data: " + error.message);
      }
    };
  }
  
  // Regular sign out button
  const signOutBtn = document.getElementById("sign-out-btn");
  if (signOutBtn) {
    signOutBtn.onclick = async () => {
      console.log("🚪 Sign out button clicked");
      try {
        await authService.signOut();
        showSuccess("Signed out successfully!");
        updateAuthUI();
      } catch (error) {
        console.error("❌ Error signing out:", error);
        showError("Error signing out: " + error.message);
      }
    };
  }
}
