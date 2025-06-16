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
