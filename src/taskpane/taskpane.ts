/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Check if we're running in Office context or standalone browser
const isInOfficeContext = typeof Office !== 'undefined' && 
  Office.context && 
  Office.context.mailbox && 
  Office.context.host;

console.log("🔧 Script loaded. Office context:", Office?.context || "undefined");
console.log("🔍 Is in real Office context:", isInOfficeContext);

if (isInOfficeContext) {
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
} else {
  // Standalone browser mode for testing
  console.log("🌐 Setting up standalone browser mode");
  document.addEventListener('DOMContentLoaded', () => {
    console.log("📄 DOM Content Loaded");
    console.log("🔧 Running in standalone browser mode for testing");
    
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    const button = document.getElementById("run");
    
    console.log("Elements found:", { sideloadMsg, appBody, button });
    
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    
    if (button) {
      button.onclick = (e) => {
        console.log("🔘 Button clicked!", e);
        runStandalone();
      };
      console.log("✅ Button click handler attached (standalone mode)");
    } else {
      console.error("❌ Button not found!");
    }
  });
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
  const mockEmailSubject = "Demo: Meeting Regarding Q1 Budget Planning";
  const mockEmailBody = `Hi Team,

I hope this email finds you well. I wanted to schedule a meeting to discuss our Q1 budget planning initiatives.

Key Discussion Points:
1. Review of last quarter's performance metrics
2. Budget allocation for new projects
3. Resource planning for the next quarter
4. Timeline for implementation

Please let me know your availability for next week. The meeting should take approximately 2 hours.

Action Items:
- Prepare Q4 financial reports
- Review project proposals
- Gather team feedback on resource needs

Looking forward to our discussion.

Best regards,
John Smith
Project Manager`;
  
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

**Main Topic:** Budget planning meeting for Q1

**Key Points:**
• Meeting request for Q1 budget planning discussion
• 2-hour meeting needed for next week
• Focus on performance review and resource allocation

**Action Items:**
• Prepare Q4 financial reports
• Review project proposals  
• Gather team feedback on resource needs

**Participants:** Team members need to confirm availability

*To enable real AI summarization: Add your OpenAI/Claude/other LLM API key to the code*`;
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
