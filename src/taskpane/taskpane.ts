/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Check if we're running in Office context or standalone browser
let isInOfficeContext = false;

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

if (isInOfficeContext) {
  try {
    Office.onReady((info) => {
      console.log("📧 Office.onReady fired", info);
      if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        
        const button = document.getElementById("run");
        console.log("🔘 Button element:", button);
        if (button) {
          button.onclick = run;
          console.log("✅ Button click handler attached (Office mode)");
        } else {
          console.error("❌ Button not found!");
        }
      }
    });
  } catch (error) {
    console.error("❌ Error in Office.onReady:", error);
  }
} else {
  // Standalone browser mode for testing
  console.log("🌐 Setting up standalone browser mode");
  try {
    document.addEventListener('DOMContentLoaded', () => {
      console.log("📄 DOM Content Loaded");
      console.log("🔧 Running in standalone browser mode for testing");
      
      try {
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        const button = document.getElementById("run");
        
        console.log("Elements found:", { sideloadMsg, appBody, button });
        
        if (sideloadMsg) sideloadMsg.style.display = "none";
        if (appBody) appBody.style.display = "flex";
        
        if (button) {
          button.onclick = (e) => {
            console.log("🔘 Button clicked!", e);
            try {
              runStandalone();
            } catch (error) {
              console.error("❌ Error in runStandalone:", error);
            }
          };
          console.log("✅ Button click handler attached (standalone mode)");
        } else {
          console.error("❌ Button not found!");
        }
      } catch (error) {
        console.error("❌ Error setting up DOM elements:", error);
      }
    });
  } catch (error) {
    console.error("❌ Error setting up DOMContentLoaded listener:", error);
  }
}

export async function run() {
  /**
   * Enhanced Outlook add-in with LLM email summarization
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  
  // Clear previous content
  insertAt.innerHTML = "";
  
  // Display email info
  let subjectLabel = document.createElement("b");
  subjectLabel.textContent = "Subject: ";
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
        
        // Display original length
        let lengthInfo = document.createElement("p");
        lengthInfo.innerHTML = `<strong>Original email length:</strong> ${emailBody.length} characters`;
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
  
  let insertAt = document.getElementById("item-subject");
  console.log("📍 Insert target element:", insertAt);
  
  if (!insertAt) {
    console.error("❌ Cannot find item-subject element!");
    return;
  }
  
  // Clear previous content
  insertAt.innerHTML = "";
  console.log("🧹 Cleared previous content");
  
  // Simulate email data
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

  // Display original length
  let lengthInfo = document.createElement("p");
  lengthInfo.innerHTML = `<strong>📊 Original email length:</strong> ${mockEmailBody.length} characters`;
  insertAt.appendChild(lengthInfo);
  
  // Add summarize button
  let summarizeBtn = document.createElement("button");
  summarizeBtn.textContent = "🤖 Summarize with AI";
  summarizeBtn.className = "ms-Button ms-Button--primary";
  summarizeBtn.style.marginTop = "10px";
  summarizeBtn.onclick = () => summarizeEmail(mockEmailBody, insertAt);
  insertAt.appendChild(summarizeBtn);
  
  // Add divider
  insertAt.appendChild(document.createElement("hr"));
  
  // Show demo instructions
  let demoInfo = document.createElement("div");
  demoInfo.className = "demo-info";
  demoInfo.innerHTML = `
    <h4 class="demo-title">🔧 Demo Mode Active</h4>
    <p class="demo-description">This is a standalone browser test. In a real Outlook add-in:</p>
    <div class="demo-list">
      <div class="demo-item">
        <span class="demo-bullet">•</span>
        <span class="demo-text">📧 Email subject and body would come from the selected message</span>
      </div>
      <div class="demo-item">
        <span class="demo-bullet">•</span>
        <span class="demo-text">🔄 This would run inside Outlook's task pane</span>
      </div>
      <div class="demo-item">
        <span class="demo-bullet">•</span>
        <span class="demo-text">🤖 AI summarization would work with real email content</span>
      </div>
    </div>
  `;
  insertAt.appendChild(demoInfo);
}

async function summarizeEmail(emailContent: string, container: HTMLElement) {
  // Show loading state
  let loadingMsg = document.createElement("p");
  loadingMsg.textContent = "🔄 Summarizing email with AI...";
  loadingMsg.id = "loading-msg";
  container.appendChild(loadingMsg);

  try {
    // Call OpenAI API (you'll need to add your API key)
    const summary = await callLLMApi(emailContent);
    
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

*To enable real AI analysis: Add your OpenAI/Claude/other LLM API key to the code*`;
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
