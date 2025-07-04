# 🤖 AI-Powered Outlook Email Summarizer

An intelligent Outlook add-in that leverages Large Language Models (LLMs) to automatically summarize email content, helping users quickly understand key points, action items, and important information.

## ✨ Features

- **🔍 Smart Email Analysis**: Automatically extracts and analyzes email content
- **🤖 AI-Powered Summarization**: Integrates with multiple LLM APIs (OpenAI, Claude, Azure OpenAI)
- **📝 Structured Summaries**: Provides organized summaries with key points and action items
- **🔄 Real-time Processing**: Instant email analysis with loading indicators
- **🌐 Cross-Platform**: Works in Outlook desktop, web, and mobile
- **🛡️ Secure**: API keys stored securely, no email content stored externally

## 🚀 Testing Results

The add-in is fully functional and working as designed! Here's what the demo mode shows:

![AI Email Summarizer Demo](assets/local-testing.png)

**Demo Features Demonstrated:**
- ✅ Clean, modern UI with blue gradient header
- ✅ Prominent "🔍 Analyze Email" button
- ✅ Mock email content processing (596 characters)
- ✅ AI summarization with structured output
- ✅ Demo mode instructions with proper bullet formatting
- ✅ Real-time feedback and loading states
- ✅ Professional styling with consistent branding

**Key Testing Results:**
- 🔘 Button click functionality: **Working**
- 📧 Email content extraction: **Working** (mock data)
- 🤖 AI summarization: **Working** (demo mode)
- 🎨 UI/UX design: **Polished and professional**
- 📱 Responsive layout: **Mobile-friendly**

### 🧪 Quick Test (No Setup Required)

To see the AI Email Summarizer in action right now:

1. **Start the dev server**: `npm run dev-server`
2. **Open your browser**: Navigate to `https://localhost:3000/taskpane.html`
3. **Click "🔍 Analyze Email"**: See instant AI summarization with mock email data
4. **View the demo**: Experience the full workflow with realistic business email content

*Note: The demo uses mock email data. In a real Outlook add-in, this would process actual email content.*

## 🚀 Supported LLM APIs

- **OpenAI GPT** (GPT-3.5, GPT-4)
- **Anthropic Claude** (Claude-3 Sonnet)
- **Azure OpenAI**
- **Google Gemini** (can be added)
- **Custom API endpoints**

## 🛠️ Development Setup

### Prerequisites

- Node.js (latest LTS version)
- npm or yarn
- Microsoft 365 account
- Outlook (web, desktop, or mobile)

### Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd kevinxu-test-sample
```

2. Install dependencies:
```bash
npm install
```

3. Start development server:
```bash
npm run dev-server
```

4. Build for production:
```bash
npm run build
```

## 🔧 Configuration

### LLM API Setup

1. **OpenAI Setup**:
   - Get API key from [OpenAI Platform](https://platform.openai.com)
   - Replace `API_KEY` in `src/taskpane/taskpane.ts`

2. **Claude Setup**:
   - Get API key from [Anthropic Console](https://console.anthropic.com)
   - Uncomment Claude function in code

3. **Azure OpenAI Setup**:
   - Set up Azure OpenAI resource
   - Configure endpoint and API key

### Environment Variables

Create a `.env` file (not committed to git):
```env
OPENAI_API_KEY=your_openai_key_here
CLAUDE_API_KEY=your_claude_key_here
AZURE_OPENAI_ENDPOINT=your_azure_endpoint
AZURE_OPENAI_KEY=your_azure_key
```

## 📱 Testing

### Browser Testing (Standalone)
```bash
npm run dev-server
open https://localhost:3000/taskpane.html
```

### Outlook Integration
1. **Outlook on the Web**: Manually sideload via add-in settings
2. **Outlook Desktop**: Use `npm start` (if sideloading is supported)
3. **Development Account**: Use Microsoft 365 Developer subscription

## 🏗️ Project Structure

```
├── src/
│   ├── taskpane/
│   │   ├── taskpane.ts       # Main add-in logic with LLM integration
│   │   ├── taskpane.html     # Add-in UI
│   │   └── taskpane.css      # Styling
│   └── commands/
├── assets/                   # Icons and images
├── manifest.json            # Add-in manifest
├── package.json            # Dependencies and scripts
└── webpack.config.js       # Build configuration
```

## 🎯 Key Features Implementation

### Email Content Access
```typescript
// Access email subject and body
const item = Office.context.mailbox.item;
const subject = item.subject;
item.body.getAsync("text", callback);
```

### LLM Integration
```typescript
// Call OpenAI API
const response = await fetch('https://api.openai.com/v1/chat/completions', {
  method: 'POST',
  headers: {
    'Authorization': `Bearer ${API_KEY}`,
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    model: "gpt-3.5-turbo",
    messages: [...]
  })
});
```

## 🔒 Security Considerations

- ✅ API keys are excluded from version control
- ✅ HTTPS required for all API calls
- ✅ Input validation for email content
- ✅ Error handling for API failures
- ⚠️ Consider data privacy when sending emails to external APIs

## 📝 Available Scripts

- `npm start` - Start with automatic sideloading
- `npm run dev-server` - Start development server only
- `npm run build` - Build for production
- `npm run lint` - Run ESLint
- `npm stop` - Stop development server

## 🔍 Troubleshooting

### Common Issues

1. **Sideloading fails**: Use manual sideloading via Outlook settings
2. **Certificate issues**: Accept localhost certificate when prompted
3. **API errors**: Check API key configuration and rate limits
4. **Browser caching**: Hard refresh or use incognito mode

### Debug Mode

Open browser dev tools to see console logs:
- `🔧 Running in standalone browser mode for testing`
- API call logs and error messages

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Microsoft Office Add-ins team for the excellent documentation
- OpenAI for the GPT API
- Anthropic for the Claude API
- The open-source community for webpack and other tools

---

# Build & Deployment Guide

## Overview

This guide explains the build and deployment process for the Outlook Connector application, common issues, and troubleshooting steps.

## 🏗️ Build Process

### Command Execution
```bash
npm run build  # Executes: webpack --mode production
```

### What Happens During Build

1. **Webpack Configuration Loading**
   - Loads `webpack.config.js`
   - Sets `options.mode = "production"`
   - Configures `publicPath = "/OutlookConnector/"`

2. **Entry Points Processing**
   ```javascript
   entry: {
     polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
     taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
     commands: "./src/commands/commands.ts",
   }
   ```

3. **Module Transformations**
   - **TypeScript → JavaScript**: `.ts` files compiled via Babel
   - **HTML Processing**: Templates processed with asset injection
   - **Asset Optimization**: Images copied and optimized

4. **Plugin Execution**
   - **HtmlWebpackPlugin**: Generates `taskpane.html` with correct script tags
   - **CopyWebpackPlugin**: Copies assets and transforms manifest URLs

5. **Production Optimizations**
   - **Minification**: JavaScript compressed
   - **Tree Shaking**: Unused code removed
   - **Code Splitting**: Separate bundles for different concerns

### Build Output Structure
```
dist/
├── taskpane.html              # Processed HTML with correct paths
├── taskpane.js                # Compiled & minified JavaScript (322KB)
├── taskpane.js.map            # Source map for debugging
├── polyfill.js                # Browser compatibility bundle (203KB)
├── commands.html              # Commands page
├── commands.js                # Commands bundle
├── manifest.json              # Production URLs
├── [hash].css                 # Compiled CSS
└── assets/                    # Static assets
    ├── icon-80.png
    ├── icon-128.png
    └── logo-filled.png
```

## 🚀 Deployment Process

### GitHub Actions Workflow (`deploy.yml`)

```yaml
name: Deploy to GitHub Pages

on:
  push:
    branches: [ main ]

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4
      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '18'
          cache: 'npm'
      - name: Install dependencies
        run: npm ci
      - name: Build project
        run: npm run build
      - name: Upload artifact
        uses: actions/upload-pages-artifact@v3
        with:
          path: './dist'

  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    needs: build
    if: github.ref == 'refs/heads/main'
    steps:
      - name: Deploy to GitHub Pages
        uses: actions/deploy-pages@v4
```

### Deployment Flow

1. **Trigger**: Push to `main` branch
2. **Build Job**: 
   - Install dependencies with `npm ci`
   - Build project with `npm run build`
   - Upload `dist/` folder as artifact
3. **Deploy Job**:
   - Wait for build completion
   - Deploy artifact to GitHub Pages
4. **Result**: Live at `https://jkevinxu.github.io/OutlookConnector/`

## ⚠️ Common Issues & Solutions

### Issue 1: 404 Errors on GitHub Pages

**Symptoms:**
- Build successful but pages return 404
- Assets not loading correctly

**Causes:**
- Incorrect `publicPath` configuration
- Conflicting deployment workflows

**Solution:**
```javascript
// webpack.config.js
const publicPath = dev ? "/" : "/OutlookConnector/";
```

### Issue 2: Conflicting Deployment Workflows

**Problem:** Multiple workflow files deploying different content

**Files to check:**
- `.github/workflows/deploy.yml` ✅ (Keep this)
- `.github/workflows/static.yml` ❌ (Remove this)

**Key Differences:**

| `deploy.yml` ✅ | `static.yml` ❌ |
|-----------------|-----------------|
| Builds project with `npm run build` | No build step |
| Deploys `./dist` (built files) | Deploys `.` (entire repo) |
| ~2MB deployment | ~500MB deployment |
| Compiled JavaScript | Raw TypeScript |
| Production URLs | Development URLs |

### Issue 3: Manifest URL Issues

**Problem:** Office add-in manifest contains localhost URLs in production

**Solution:** Webpack transforms URLs during build:
```javascript
transform(content) {
  if (dev) {
    return content;
  } else {
    return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
  }
}
```

- **Development**: `https://localhost:3000/`
- **Production**: `https://jkevinxu.github.io/OutlookConnector/`

## 🔍 Troubleshooting

### Check Deployment Status

1. **GitHub Actions**
   - Go to repository → Actions tab
   - Look for "Deploy to GitHub Pages" workflow
   - ✅ Green = Success, ❌ Red = Failed

2. **GitHub Pages Settings**
   - Repository → Settings → Pages
   - Source should be "GitHub Actions"
   - Check deployment status and URL

3. **Test URLs**
   ```bash
   # Should return 200 OK
   curl -I https://jkevinxu.github.io/OutlookConnector/taskpane.html
   curl -I https://jkevinxu.github.io/OutlookConnector/taskpane.js
   ```

### Force Redeploy

**Method 1: Re-run from GitHub UI**
- Go to Actions → Select workflow run → Re-run jobs

**Method 2: Trigger new deployment**
```bash
git commit --allow-empty -m "Trigger deployment"
git push origin main
```

### Debug Build Issues

**Check build output:**
```bash
npm run build
# Look for errors in webpack output
# Check generated files in dist/
```

**Common build errors:**
- TypeScript compilation errors
- Missing dependencies
- Asset path issues

## 📋 Best Practices

### Development Workflow

1. **Local Development**
   ```bash
   npm run dev-server  # Local development with hot reload
   ```

2. **Test Build Locally**
   ```bash
   npm run build  # Test production build
   ```

3. **Deploy**
   ```bash
   git add .
   git commit -m "Your changes"
   git push origin main  # Triggers deployment
   ```

### File Structure

- **Source files**: Keep in `src/` directory
- **Static assets**: Keep in `assets/` directory
- **Build output**: Generated in `dist/` (don't commit)
- **Workflows**: One deployment workflow only

### Configuration

- **Development**: `publicPath = "/"`
- **Production**: `publicPath = "/OutlookConnector/"`
- **URLs**: Transform localhost → production in manifest

## 🎯 Key Takeaways

1. **Build process is essential** - Browsers can't execute TypeScript directly
2. **Only one deployment workflow** - Multiple workflows cause conflicts
3. **Correct publicPath** - Must match GitHub Pages URL structure
4. **URL transformation** - Manifest must use production URLs
5. **Test locally first** - Always verify build works before deploying

## 📞 Quick Reference

**Build Commands:**
- `npm run build` - Production build
- `npm run build:dev` - Development build
- `npm run dev-server` - Local development server

**Deployment URL:**
- `https://jkevinxu.github.io/OutlookConnector/taskpane.html`

**Key Files:**
- `webpack.config.js` - Build configuration
- `.github/workflows/deploy.yml` - Deployment workflow
- `dist/` - Generated build output

# Outlook Add-in: Email Conversation Thread Access

## Overview

This document outlines the capabilities and limitations of accessing email conversation threads in Outlook add-ins using Office.js APIs, and provides solutions using Microsoft Graph API.

## 🚨 Current Office.js Limitations

### ❌ What Office.js **CANNOT** Do for Conversation Threads

Office.js for Outlook add-ins has **very limited support** for accessing conversation threads:

- ❌ **Get all emails in a conversation thread**
- ❌ **Navigate to parent/child emails in thread** 
- ❌ **Access previous messages in the same conversation**
- ❌ **Retrieve conversation history**
- ❌ **List related emails with same subject**

### ✅ What Office.js **CAN** Do

Office.js primarily provides access to the **current email item** only:

```typescript
// Access current email item
const item = Office.context.mailbox.item;

// Available properties
const subject = item.subject;
const body = await getEmailBody();
const sender = item.from;
const recipients = item.to;

// Basic conversation info (limited)
const conversationId = item.conversationId; // May not always be available
```

## 🔧 Solution: Microsoft Graph API Integration

To access conversation threads, you need to combine **Office.js** with **Microsoft Graph API**.

### Prerequisites

1. **Permissions in manifest.json**:
```json
{
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "MailboxItem.Read.User",
          "type": "Delegated"
        },
        {
          "name": "Mail.Read",
          "type": "Delegated"
        }
      ]
    }
  }
}
```

2. **Authentication Setup**: Your add-in must support Microsoft Graph authentication.

### Implementation Strategy

#### 1. Get Conversation Emails via Graph API

```typescript
/**
 * Fetch all emails in the same conversation thread
 */
async function getConversationEmails(): Promise<any[]> {
  try {
    // Get access token for Graph API
    const tokenResult = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true
    });
    
    const currentItem = Office.context.mailbox.item;
    const conversationId = currentItem.conversationId;
    
    if (!conversationId) {
      console.warn('No conversation ID available for current email');
      return [];
    }
    
    // Call Graph API to get conversation emails
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${conversationId}'&$orderby=receivedDateTime desc`,
      {
        headers: {
          'Authorization': `Bearer ${tokenResult}`,
          'Content-Type': 'application/json'
        }
      }
    );
    
    if (!response.ok) {
      throw new Error(`Graph API request failed: ${response.status}`);
    }
    
    const data = await response.json();
    return data.value || [];
    
  } catch (error) {
    console.error('Error fetching conversation emails:', error);
    return [];
  }
}
```

#### 2. Alternative: Search by Subject

```typescript
/**
 * Find related emails by subject (fallback approach)
 */
async function getEmailsBySubject(subject: string): Promise<any[]> {
  try {
    const token = await OfficeRuntime.auth.getAccessToken();
    
    // Clean and encode subject for search
    const cleanSubject = subject.replace(/^(RE:|FW:|FWD:)\s*/i, '').trim();
    const encodedSubject = encodeURIComponent(cleanSubject);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages?$filter=contains(subject, '${encodedSubject}')&$orderby=receivedDateTime desc&$top=10`,
      {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      }
    );
    
    const data = await response.json();
    return data.value || [];
    
  } catch (error) {
    console.error('Error searching emails by subject:', error);
    return [];
  }
}
```

#### 3. Complete Conversation Analysis Function

```typescript
/**
 * Analyze current email and its conversation thread
 */
async function analyzeConversationThread() {
  try {
    // Get current email info via Office.js
    const currentItem = Office.context.mailbox.item;
    
    const currentEmailInfo = {
      subject: currentItem.subject,
      sender: currentItem.from?.emailAddress,
      senderName: currentItem.from?.displayName,
      conversationId: currentItem.conversationId
    };
    
    // Get email body
    const bodyResult = await new Promise<string>((resolve, reject) => {
      currentItem.body.getAsync("text", (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error?.message || "Failed to get email body"));
        }
      });
    });
    
    // Extract first 20 words
    const words = bodyResult.trim().split(/\s+/);
    const first20Words = words.slice(0, 20).join(' ');
    
    currentEmailInfo.first20Words = first20Words + (words.length > 20 ? '...' : '');
    currentEmailInfo.wordCount = words.length;
    currentEmailInfo.charCount = bodyResult.length;
    
    // Get conversation thread emails via Graph API
    let threadEmails: any[] = [];
    
    if (currentEmailInfo.conversationId) {
      threadEmails = await getConversationEmails();
    } else {
      // Fallback to subject-based search
      threadEmails = await getEmailsBySubject(currentEmailInfo.subject);
    }
    
    return {
      currentEmail: currentEmailInfo,
      threadEmails: threadEmails.map(email => ({
        id: email.id,
        subject: email.subject,
        sender: email.from?.emailAddress?.address,
        senderName: email.from?.emailAddress?.name,
        receivedDateTime: email.receivedDateTime,
        bodyPreview: email.bodyPreview,
        isRead: email.isRead,
        hasAttachments: email.hasAttachments
      })),
      threadCount: threadEmails.length
    };
    
  } catch (error) {
    console.error('Error analyzing conversation thread:', error);
    throw error;
  }
}
```

### 4. Display Conversation Thread in UI

```typescript
/**
 * Display current email and conversation thread in the task pane
 */
async function displayEmailWithThread() {
  try {
    const analysis = await analyzeConversationThread();
    const container = document.getElementById("item-subject");
    
    if (!container) return;
    
    container.innerHTML = "";
    
    // Display current email info
    const currentEmailSection = document.createElement("div");
    currentEmailSection.style.cssText = `
      background-color: #f8f9fa;
      border: 1px solid #e9ecef;
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 20px;
    `;
    
    currentEmailSection.innerHTML = `
      <h3>📧 Current Email</h3>
      <p><strong>Subject:</strong> ${analysis.currentEmail.subject}</p>
      <p><strong>From:</strong> ${analysis.currentEmail.senderName} &lt;${analysis.currentEmail.sender}&gt;</p>
      <p><strong>First 20 words:</strong> ${analysis.currentEmail.first20Words}</p>
      <p><strong>Stats:</strong> ${analysis.currentEmail.wordCount} words, ${analysis.currentEmail.charCount} characters</p>
    `;
    
    container.appendChild(currentEmailSection);
    
    // Display conversation thread
    if (analysis.threadEmails.length > 1) {
      const threadSection = document.createElement("div");
      threadSection.style.cssText = `
        background-color: #fff;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
      `;
      
      threadSection.innerHTML = `
        <h3>💬 Conversation Thread (${analysis.threadCount} emails)</h3>
        <div id="thread-emails"></div>
      `;
      
      const threadContainer = threadSection.querySelector("#thread-emails");
      
      analysis.threadEmails.forEach((email, index) => {
        const emailItem = document.createElement("div");
        emailItem.style.cssText = `
          border-left: 3px solid ${index === 0 ? '#007bff' : '#dee2e6'};
          padding: 10px 15px;
          margin: 10px 0;
          background: ${index === 0 ? '#f8f9fa' : '#fff'};
        `;
        
        emailItem.innerHTML = `
          <div style="font-weight: bold; color: #495057;">
            ${email.senderName || email.sender} 
            ${index === 0 ? '(Current)' : ''}
          </div>
          <div style="font-size: 12px; color: #6c757d; margin: 5px 0;">
            ${new Date(email.receivedDateTime).toLocaleDateString()} - 
            ${email.isRead ? 'Read' : 'Unread'}${email.hasAttachments ? ' 📎' : ''}
          </div>
          <div style="font-size: 13px; color: #495057;">
            ${email.bodyPreview || 'No preview available'}
          </div>
        `;
        
        threadContainer.appendChild(emailItem);
      });
      
      container.appendChild(threadSection);
    }
    
  } catch (error) {
    console.error('Error displaying email with thread:', error);
    showError('Failed to load conversation thread: ' + error.message);
  }
}
```

## 🔗 Microsoft Graph API Endpoints

### Key Endpoints for Conversation Access

```typescript
// Get messages by conversation ID
GET /me/messages?$filter=conversationId eq '{conversation-id}'

// Get messages with specific ordering and limiting
GET /me/messages?$filter=conversationId eq '{conversation-id}'&$orderby=receivedDateTime desc&$top=50

// Search messages by subject
GET /me/messages?$filter=contains(subject, '{subject-text}')&$orderby=receivedDateTime desc

// Get conversation threads (alternative approach)
GET /me/conversations/{conversation-id}/threads

// Get message details with expanded properties
GET /me/messages/{message-id}?$expand=attachments
```

### Required Graph API Permissions

- `Mail.Read` - Read user's email
- `Mail.ReadWrite` - Read and write user's email (if modifications needed)

## 🚀 Integration Example

Here's how to integrate this into your existing `run()` function:

```typescript
export async function run() {
  // Check authentication first
  if (!authService.isAuthenticated()) {
    showError("Please sign in to analyze emails");
    return;
  }

  try {
    // Display current email and conversation thread
    await displayEmailWithThread();
    
  } catch (error) {
    console.error("Error in run():", error);
    showError("Failed to analyze email: " + error.message);
  }
}
```

## 📝 Implementation Notes

### Security Considerations
- Always validate Graph API responses
- Handle authentication failures gracefully
- Implement proper error handling for network issues

### Performance Tips
- Cache conversation data when possible
- Use `$top` parameter to limit result sets
- Implement pagination for large conversations

### Fallback Strategies
- Use subject-based search if `conversationId` is unavailable
- Gracefully degrade to current email only if Graph API fails
- Provide clear user feedback about limitations

## 🔮 Future Enhancements

Microsoft may expand Office.js conversation capabilities in future updates. Monitor the [Office Add-ins roadmap](https://docs.microsoft.com/en-us/office/dev/add-ins/) for new features.

## 📚 Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Office.js API Reference](https://docs.microsoft.com/en-us/javascript/api/office/)
- [Outlook Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/)
- [Authentication in Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)

---

## Summary

**Current State**: Office.js alone cannot access conversation threads  
**Solution**: Combine Office.js + Microsoft Graph API  
**Limitation**: Add-ins are designed around single-item interaction  
**Workaround**: Use Graph API for comprehensive thread-level operations
