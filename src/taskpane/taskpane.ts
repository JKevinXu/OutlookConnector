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
  
  if (currentPath === '/taskpane/callback' || hash.includes('id_token=')) {
    console.log("🎯 Processing authentication callback...");
    
    // Let oidc-client handle the callback
    authService.handleCallback().then((user) => {
      console.log("✅ Callback processed successfully:", user);
      
      // Redirect to main taskpane
      const baseUrl = `${window.location.protocol}//${window.location.host}`;
      const targetUrl = `${baseUrl}/taskpane.html`;
      console.log("🔄 Redirecting to:", targetUrl);
      
      // Use replace to avoid adding to history
      window.location.replace(targetUrl);
    }).catch((error) => {
      console.error("❌ Callback processing failed:", error);
      showError(`Authentication failed: ${error.message}`);
      
      // Still redirect to main page on error
      const baseUrl = `${window.location.protocol}//${window.location.host}`;
      const targetUrl = `${baseUrl}/taskpane.html`;
      setTimeout(() => {
        window.location.replace(targetUrl);
      }, 2000);
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

  if (isAuthenticated && currentUser) {
    headerContent.innerHTML = `
      <span class="header-icon">🛒</span>
      <h1 class="ms-font-xl">Seller Email Assistant</h1>
      <div class="user-info">
        <span class="user-avatar">👤</span>
        <span class="user-name">${currentUser.name}</span>
        <button id="logout-btn" class="ms-Button ms-Button--default logout-btn">Sign Out</button>
      </div>
    `;

    // Add logout button handler
    const logoutBtn = document.getElementById("logout-btn");
    if (logoutBtn) {
      logoutBtn.onclick = handleLogout;
    }
  } else {
    headerContent.innerHTML = `
      <span class="header-icon">🛒</span>
      <h1 class="ms-font-xl">Seller Email Assistant</h1>
    `;
  }
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

async function handleLogout() {
  try {
    console.log("🚪 Logout button clicked");
    await authService.logout();
  } catch (error) {
    console.error("❌ Logout failed:", error);
    showError("Logout failed. Please try again.");
  }
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
  const resultsContainer = document.getElementById("results-container");
  if (!resultsContainer) return;

  try {
    const user = await authService.getUser();
    const accessToken = await authService.getAccessToken();
    const idToken = await authService.getIdToken();
    
    if (!user) return;

    // Decode JWT claims if ID token is available
    let jwtClaims = null;
    if (idToken) {
      jwtClaims = decodeJWT(idToken);
    }

    // Clear previous user info
    const existingUserInfo = document.getElementById("user-identity-section");
    if (existingUserInfo) {
      existingUserInfo.remove();
    }

    // Create user identity section
    const userIdentitySection = document.createElement("div");
    userIdentitySection.id = "user-identity-section";
    userIdentitySection.className = "user-identity-panel";
    
    userIdentitySection.innerHTML = `
      <div class="user-identity-header">
        <h3>👤 User Identity Information</h3>
        <span class="identity-status">✅ Authenticated</span>
      </div>
      
      <div class="identity-grid">
        <div class="identity-item">
          <span class="identity-label">📧 Email:</span>
          <span class="identity-value">${user.email || 'Not provided'}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">👤 Name:</span>
          <span class="identity-value">${user.name}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">🆔 Subject ID:</span>
          <span class="identity-value">${user.sub}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">✅ Email Verified:</span>
          <span class="identity-value">${user.email_verified ? '✅ Yes' : '❌ No'}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">🏢 Organization:</span>
          <span class="identity-value">${user.org || 'Not provided'}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">🎭 Roles:</span>
          <span class="identity-value">${user.roles && user.roles.length > 0 ? user.roles.join(', ') : 'No roles assigned'}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">⏰ Issued At:</span>
          <span class="identity-value">${new Date(user.iat * 1000).toLocaleString()}</span>
        </div>
        
        <div class="identity-item">
          <span class="identity-label">⏳ Expires At:</span>
          <span class="identity-value">${new Date(user.exp * 1000).toLocaleString()}</span>
        </div>
      </div>
      
      ${jwtClaims ? `
      <div class="jwt-claims-section">
        <h4>🔍 Raw JWT Claims</h4>
        <div class="jwt-claims-container">
          <div class="jwt-claim-item">
            <span class="jwt-label">Audience (aud):</span>
            <span class="jwt-value">${jwtClaims.aud || 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Subject (sub):</span>
            <span class="jwt-value">${jwtClaims.sub || 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Issuer (iss):</span>
            <span class="jwt-value">${jwtClaims.iss || 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Not Before (nbf):</span>
            <span class="jwt-value">${jwtClaims.nbf ? formatTimestamp(jwtClaims.nbf) : 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Issued At (iat):</span>
            <span class="jwt-value">${jwtClaims.iat ? formatTimestamp(jwtClaims.iat) : 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Expires At (exp):</span>
            <span class="jwt-value">${jwtClaims.exp ? formatTimestamp(jwtClaims.exp) : 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Auth Time:</span>
            <span class="jwt-value">${jwtClaims.auth_time ? formatTimestamp(jwtClaims.auth_time) : 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">Nonce:</span>
            <span class="jwt-value">${jwtClaims.nonce || 'N/A'}</span>
          </div>
          <div class="jwt-claim-item">
            <span class="jwt-label">JWT ID (jti):</span>
            <span class="jwt-value">${jwtClaims.jti || 'N/A'}</span>
          </div>
          ${jwtClaims['https://aws.amazon.com/tags'] ? `
          <div class="jwt-claim-item">
            <span class="jwt-label">AWS Tags:</span>
            <span class="jwt-value">${JSON.stringify(jwtClaims['https://aws.amazon.com/tags'], null, 2)}</span>
          </div>
          ` : ''}
          <div class="jwt-claim-item">
            <span class="jwt-label">Token Purpose:</span>
            <span class="jwt-value">${jwtClaims.federate_token_purpose || 'N/A'}</span>
          </div>
        </div>
        
        <div class="jwt-raw-section">
          <h5>📋 Complete JWT Payload</h5>
          <textarea class="jwt-raw-textarea" readonly>${JSON.stringify(jwtClaims, null, 2)}</textarea>
          <button class="token-btn" onclick="copyToClipboard('${JSON.stringify(jwtClaims, null, 2).replace(/'/g, "\\'")}')">📋 Copy Claims</button>
        </div>
      </div>
      ` : ''}
      
      <div class="token-section">
        <h4>🔑 Token Information</h4>
        <div class="token-item">
          <span class="token-label">Access Token:</span>
          <span class="token-value">${accessToken ? '✅ Present' : '❌ Not available'}</span>
          ${accessToken ? `<button class="token-btn" onclick="copyToClipboard('${accessToken}')">📋 Copy</button>` : ''}
        </div>
        <div class="token-item">
          <span class="token-label">ID Token:</span>
          <span class="token-value">${idToken ? '✅ Present' : '❌ Not available'}</span>
          ${idToken ? `<button class="token-btn" onclick="copyToClipboard('${idToken}')">📋 Copy</button>` : ''}
        </div>
      </div>
      
      <div class="auth-actions">
        <button id="refresh-identity" class="ms-Button ms-Button--default">🔄 Refresh Identity</button>
        <button id="renew-token" class="ms-Button ms-Button--default">🔑 Renew Token</button>
      </div>
    `;

    // Insert at the beginning of results container
    resultsContainer.insertBefore(userIdentitySection, resultsContainer.firstChild);

    // Add event handlers
    const refreshBtn = document.getElementById("refresh-identity");
    if (refreshBtn) {
      refreshBtn.onclick = () => {
        displayUserIdentity();
      };
    }

    const renewBtn = document.getElementById("renew-token");
    if (renewBtn) {
      renewBtn.onclick = async () => {
        try {
          await authService.renewToken();
          displayUserIdentity();
          showSuccess("Token renewed successfully!");
        } catch (error) {
          showError("Failed to renew token: " + error.message);
        }
      };
    }

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

function initializeApiUI() {
  // Get Seller History button
  const getSellerHistoryBtn = document.getElementById("get-seller-history-btn");
  if (getSellerHistoryBtn) {
    getSellerHistoryBtn.onclick = handleGetSellerHistory;
  }

  // Get Current User History button
  const getCurrentUserHistoryBtn = document.getElementById("get-current-user-history-btn");
  if (getCurrentUserHistoryBtn) {
    getCurrentUserHistoryBtn.onclick = handleGetCurrentUserHistory;
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

    console.log("📞 Calling seller history API...");
    showApiLoading(true);

    const result = await sellerHistoryService.getSellerHistory(params);
    
    console.log("✅ API call successful:", result);
    displayApiResults(result);

  } catch (error) {
    console.error("❌ API call failed:", error);
    showError(`API call failed: ${(error as Error).message}`);
    hideApiResults();
  } finally {
    showApiLoading(false);
  }
}

async function handleGetCurrentUserHistory() {
  try {
    if (!authService.isAuthenticated()) {
      showError("Please sign in first before calling the API.");
      return;
    }

    const marketplaceIdInput = document.getElementById("marketplace-id") as HTMLInputElement;
    const marketplaceId = marketplaceIdInput?.value || undefined;

    console.log("📞 Calling seller history API for current user...");
    showApiLoading(true);

    const result = await sellerHistoryService.getCurrentUserSellerHistory(marketplaceId);
    
    console.log("✅ API call successful:", result);
    displayApiResults(result);

  } catch (error) {
    console.error("❌ API call failed:", error);
    showError(`API call failed: ${(error as Error).message}`);
    hideApiResults();
  } finally {
    showApiLoading(false);
  }
}

function displayApiResults(data: any) {
  const resultsDiv = document.getElementById("api-results");
  const outputPre = document.getElementById("api-output");
  
  if (resultsDiv && outputPre) {
    outputPre.textContent = JSON.stringify(data, null, 2);
    resultsDiv.style.display = "block";
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
  const getCurrentUserHistoryBtn = document.getElementById("get-current-user-history-btn");
  
  if (getSellerHistoryBtn) {
    const label = getSellerHistoryBtn.querySelector('.ms-Button-label');
    if (label) {
      label.textContent = isLoading ? "⏳ Loading..." : "📈 Get Seller History";
    }
    (getSellerHistoryBtn as HTMLButtonElement).disabled = isLoading;
  }
  
  if (getCurrentUserHistoryBtn) {
    const label = getCurrentUserHistoryBtn.querySelector('.ms-Button-label');
    if (label) {
      label.textContent = isLoading ? "⏳ Loading..." : "👤 Get My History";
    }
    (getCurrentUserHistoryBtn as HTMLButtonElement).disabled = isLoading;
  }
}
