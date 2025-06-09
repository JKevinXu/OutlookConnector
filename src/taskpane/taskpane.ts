/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { authService } from '../auth/AuthService';
import { UserProfile } from '../types/auth';
import { sellerHistoryService } from '../api/SellerHistoryService';

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
  console.log("🎨 Updating auth UI. Authenticated:", isAuthenticated);

  // Update header to show user info or login
  updateHeader(isAuthenticated);
  
  // Show/hide main functionality based on auth state
  const appBody = document.getElementById("app-body");
  const sideloadMsg = document.getElementById("sideload-msg");
  
  if (isAuthenticated) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    
    // Display user identity in the panel
    console.log("🆔 Displaying user identity panel...");
    displayUserIdentity();
  } else {
    if (appBody) appBody.style.display = "none";
    if (sideloadMsg) {
      sideloadMsg.style.display = "block";
      
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
  
  // Display email info
  let subjectLabel = document.createElement("b");
  subjectLabel.textContent = "📧 Subject: ";
  insertAt.appendChild(subjectLabel);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject || "No subject"));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createElement("br"));

  // Get email body for summarization
  try {
    // Get the email body
    item.body.getAsync("text", async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emailBody = result.value;
        
        // Display first 100 characters preview
        let previewLabel = document.createElement("b");
        previewLabel.textContent = "📄 Email Preview (first 100 chars): ";
        insertAt.appendChild(previewLabel);
        insertAt.appendChild(document.createElement("br"));
        
        let previewText = document.createElement("p");
        previewText.textContent = emailBody.substring(0, 100) + (emailBody.length > 100 ? "..." : "");
        previewText.style.backgroundColor = "#f8f9fa";
        previewText.style.padding = "10px";
        previewText.style.borderRadius = "5px";
        previewText.style.fontStyle = "italic";
        insertAt.appendChild(previewText);
        
        // Display original length
        let lengthInfo = document.createElement("p");
        lengthInfo.innerHTML = `<strong>📊 Total email length:</strong> ${emailBody.length} characters`;
        insertAt.appendChild(lengthInfo);
        
        // Add summarize button
        let summarizeBtn = document.createElement("button");
        summarizeBtn.textContent = "🤖 Summarize with AI";
        summarizeBtn.className = "ms-Button ms-Button--primary";
        summarizeBtn.onclick = () => summarizeEmail(emailBody, insertAt);
        insertAt.appendChild(summarizeBtn);
        
        // Add divider
        insertAt.appendChild(document.createElement("hr"));
        
      } else {
        let errorMsg = document.createElement("p");
        errorMsg.textContent = "Could not access email body";
        errorMsg.style.color = "red";
        insertAt.appendChild(errorMsg);
      }
    });
  } catch (error) {
    let errorMsg = document.createElement("p");
    errorMsg.textContent = `Error: ${error.message}`;
    errorMsg.style.color = "red";
    insertAt.appendChild(errorMsg);
  }
}

// Standalone testing function (simulates Office context)
export async function runStandalone() {
  console.log("🚀 runStandalone() called!");
  console.log("🔧 Running standalone test mode");
  
  // Check authentication first
  if (!authService.isAuthenticated()) {
    showError("Please sign in to analyze emails");
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
  
  // Use the actual email content directly
  const mockEmailSubject = "I want to onboard FBA";
  const mockEmailBody = `Subject: I want to onboard FBA

Dear Amazon FBA Support Team,

I hope this email finds you well. My name is Marcus Chen, and I'm the founder of TechGear Solutions, an e-commerce business that has been successfully selling electronics and tech accessories through various online platforms for the past three years.

I'm reaching out because I'm very interested in transitioning to Amazon's Fulfillment by Amazon (FBA) program to scale our operations and provide better customer service. After researching extensively, I believe FBA is the perfect solution for our growing business needs.

Current Business Overview:
Our company currently generates approximately $50,000 in monthly revenue selling items like wireless chargers, phone cases, laptop accessories, and smart home devices. We maintain inventory in a 2,000 sq ft warehouse in Phoenix, Arizona, and currently fulfill orders ourselves through multiple sales channels including our Shopify store, eBay, and other marketplaces.

Why We Want FBA:
The primary drivers for our FBA interest include accessing Amazon Prime customers, leveraging Amazon's world-class logistics network, reducing our fulfillment workload, and improving delivery speeds to customers nationwide. We're particularly excited about the potential for increased sales velocity through Prime eligibility and Amazon's trusted fulfillment reputation.

Product Portfolio:
We're looking to start with our top 15 SKUs, which represent about 80% of our current sales volume. These products range from $12 to $89 in retail price, with healthy profit margins that can accommodate FBA fees. Our products are primarily sourced from vetted suppliers in Taiwan and South Korea, with established quality control processes already in place.

Current Challenges:
We're currently struggling with shipping costs for individual orders, especially to customers on the East Coast. Our current 3-5 day shipping times are hurting our competitiveness, and we're spending too much time on fulfillment activities rather than focusing on product development and marketing growth strategies.

Questions and Next Steps:
I have several specific questions about the onboarding process. First, what's the typical timeline for FBA approval and first shipment acceptance? Second, can you provide guidance on optimal inventory planning for new FBA sellers? Third, what are the most common mistakes new FBA sellers make that we should avoid?

Additionally, I'd like to understand the requirements for product photography, listing optimization, and any compliance considerations for electronics products in the FBA program.

Investment Readiness:
We have $75,000 allocated specifically for initial FBA inventory investment and are prepared to commit to a long-term partnership with Amazon. Our goal is to reach $200,000 in monthly FBA sales within the first year.

I would greatly appreciate the opportunity to speak with someone from your team to discuss our FBA onboarding process. I'm available for a call any weekday between 9 AM and 5 PM PST.

Thank you for your time and consideration. I look forward to hearing from you soon and beginning our FBA journey.

Best regards,

Marcus Chen
Founder & CEO, TechGear Solutions
Email: marcus.chen@techgearsolutions.com
Phone: (602) 555-7892
Business Address: 1247 Industrial Blvd, Phoenix, AZ 85034`;
  
  // Display email info
  let subjectLabel = document.createElement("b");
  subjectLabel.textContent = "📧 Subject: ";
  insertAt.appendChild(subjectLabel);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(mockEmailSubject));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createElement("br"));

  // Display first 100 characters preview
  let previewLabel = document.createElement("b");
  previewLabel.textContent = "📄 Email Preview (first 100 chars): ";
  insertAt.appendChild(previewLabel);
  insertAt.appendChild(document.createElement("br"));
  
  let previewText = document.createElement("p");
  previewText.textContent = mockEmailBody.substring(0, 100) + (mockEmailBody.length > 100 ? "..." : "");
  previewText.style.backgroundColor = "#f8f9fa";
  previewText.style.padding = "15px";
  previewText.style.borderRadius = "8px";
  previewText.style.fontStyle = "italic";
  previewText.style.border = "1px solid #e1e5e9";
  previewText.style.marginTop = "10px";
  insertAt.appendChild(previewText);

  // Display original length
  let lengthInfo = document.createElement("p");
  lengthInfo.innerHTML = `<strong>📊 Total email length:</strong> ${mockEmailBody.length} characters`;
  insertAt.appendChild(lengthInfo);
  
  // Add summarize button
  let summarizeBtn = document.createElement("button");
  summarizeBtn.textContent = "🤖 Summarize with AI";
  summarizeBtn.className = "ms-Button ms-Button--primary";
  summarizeBtn.style.marginTop = "15px";
  summarizeBtn.onclick = () => summarizeEmail(mockEmailBody, insertAt);
  insertAt.appendChild(summarizeBtn);
  
  // Add divider
  insertAt.appendChild(document.createElement("hr"));
}

async function summarizeEmail(emailContent: string, container: HTMLElement) {
  // Check authentication before processing
  if (!authService.isAuthenticated()) {
    showError("Please sign in to use AI summarization");
    return;
  }

  // Show loading state
  let loadingMsg = document.createElement("p");
  loadingMsg.textContent = "🔄 Summarizing email with AI...";
  loadingMsg.id = "loading-msg";
  container.appendChild(loadingMsg);

  try {
    // Get user info for personalized experience
    const user = await authService.getUser();
    const userContext = user ? `\nAnalyzing for user: ${user.name} (${user.email})` : "";
    
    // Call OpenAI API (you'll need to add your API key)
    const summary = await callLLMApi(emailContent + userContext);
    
    // Remove loading message
    const loading = document.getElementById("loading-msg");
    if (loading) loading.remove();
    
    // Display summary
    let summaryContainer = document.createElement("div");
    summaryContainer.style.border = "1px solid #ccc";
    summaryContainer.style.padding = "10px";
    summaryContainer.style.marginTop = "10px";
    summaryContainer.style.backgroundColor = "#f9f9f9";
    
    let summaryTitle = document.createElement("h3");
    summaryTitle.textContent = "📝 AI Summary";
    summaryContainer.appendChild(summaryTitle);
    
    let summaryText = document.createElement("p");
    summaryText.textContent = summary;
    summaryContainer.appendChild(summaryText);
    
    container.appendChild(summaryContainer);
    
  } catch (error) {
    // Remove loading message
    const loading = document.getElementById("loading-msg");
    if (loading) loading.remove();
    
    let errorMsg = document.createElement("p");
    errorMsg.textContent = `❌ Error summarizing email: ${error.message}`;
    errorMsg.style.color = "red";
    container.appendChild(errorMsg);
  }
}

async function callLLMApi(emailContent: string): Promise<string> {
  // Example using OpenAI API - you can replace with any LLM service
  const API_KEY = "your-openai-api-key-here"; // You'll need to add this
  const API_URL = "https://api.openai.com/v1/chat/completions";
  
  // For demo purposes, return a mock summary if no API key
  if (!API_KEY || API_KEY === "your-openai-api-key-here") {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    const user = await authService.getUser();
    const userInfo = user ? `\n\n*Personalized for ${user.name}*` : "";
    
    return `🤖 **AI Summary** (Demo Mode):

**Email Type:** FBA Onboarding Inquiry

**Business Profile:**
• Company: TechGear Solutions (Marcus Chen, Founder & CEO)
• Current Revenue: $50,000/month selling electronics & tech accessories
• Experience: 3 years in e-commerce across multiple platforms
• Location: Phoenix, AZ (2,000 sq ft warehouse)

**FBA Interest & Goals:**
• Want to access Amazon Prime customers
• Improve delivery speeds nationwide
• Reduce fulfillment workload to focus on growth
• Target: $200,000/month within first year

**Product Portfolio:**
• Starting with top 15 SKUs (80% of current sales)
• Price range: $12-$89 with healthy margins
• Electronics: chargers, cases, laptop accessories, smart home devices
• Suppliers: Taiwan & South Korea with quality control

**Key Questions:**
• FBA approval timeline and first shipment process
• Inventory planning guidance for new sellers
• Common mistakes to avoid
• Product photography & listing requirements
• Electronics compliance considerations

**Investment Ready:**
• $75,000 allocated for initial FBA inventory
• Available for calls: weekdays 9 AM-5 PM PST
• Serious about long-term Amazon partnership

**Next Steps:** Schedule consultation call to discuss onboarding process

*To enable real AI analysis: Add your OpenAI/Claude/other LLM API key to the code*${userInfo}`;
  }
  
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${API_KEY}`
    },
    body: JSON.stringify({
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: "You are a helpful assistant that summarizes emails concisely. Focus on key points, action items, and important information."
        },
        {
          role: "user",
          content: `Please summarize this email:\n\n${emailContent}`
        }
      ],
      max_tokens: 200,
      temperature: 0.3
    })
  });
  
  if (!response.ok) {
    throw new Error(`API request failed: ${response.status}`);
  }
  
  const data = await response.json();
  return data.choices[0].message.content;
}

// Alternative LLM APIs you can use:

// Claude API example:
/*
async function callClaudeApi(emailContent: string): Promise<string> {
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': 'your-claude-api-key',
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-3-sonnet-20240229',
      max_tokens: 200,
      messages: [{
        role: 'user',
        content: `Summarize this email: ${emailContent}`
      }]
    })
  });
  
  const data = await response.json();
  return data.content[0].text;
}
*/

// Azure OpenAI example:
/*
async function callAzureOpenAI(emailContent: string): Promise<string> {
  const response = await fetch('https://your-resource.openai.azure.com/openai/deployments/your-model/chat/completions?api-version=2023-12-01-preview', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': 'your-azure-api-key'
    },
    body: JSON.stringify({
      messages: [
        { role: 'system', content: 'Summarize emails concisely.' },
        { role: 'user', content: `Summarize: ${emailContent}` }
      ],
      max_tokens: 200
    })
  });
  
  const data = await response.json();
  return data.choices[0].message.content;
}
*/

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

// Helper function to copy text to clipboard
function copyToClipboard(text: string) {
  navigator.clipboard.writeText(text).then(() => {
    showSuccess("Copied to clipboard!");
  }).catch(() => {
    showError("Failed to copy to clipboard");
  });
}

// Helper function to decode JWT token
function decodeJWT(token: string): any {
  try {
    // Split the JWT into parts
    const parts = token.split('.');
    if (parts.length !== 3) {
      throw new Error('Invalid JWT format');
    }

    // Decode the payload (second part)
    const payload = parts[1];
    // Add padding if needed for base64 decoding
    const paddedPayload = payload + '='.repeat((4 - payload.length % 4) % 4);
    const decodedPayload = atob(paddedPayload.replace(/-/g, '+').replace(/_/g, '/'));
    
    return JSON.parse(decodedPayload);
  } catch (error) {
    console.error('Failed to decode JWT:', error);
    return null;
  }
}

// Helper function to format timestamp
function formatTimestamp(timestamp: number): string {
  return new Date(timestamp * 1000).toLocaleString();
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

function initializeApiUI() {
  // Get Seller Metrics button
  const getSellerHistoryBtn = document.getElementById("get-seller-history-btn");
  if (getSellerHistoryBtn) {
    getSellerHistoryBtn.onclick = handleGetSellerHistory;
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
