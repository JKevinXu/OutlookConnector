# 🤖 AI-Powered Outlook Add-in

An intelligent Outlook add-in that integrates with Large Language Models (LLMs) to provide email summarization and analysis capabilities.

## ✨ Features

- 📧 **Email Content Access**: Reads email subject and body content
- 🤖 **AI Summarization**: Integrates with popular LLM APIs for intelligent email summarization
- 🔍 **Key Point Extraction**: Identifies action items and important information
- 📊 **Content Analysis**: Displays email metrics and statistics
- 🎯 **Multiple LLM Support**: Works with OpenAI, Claude, Azure OpenAI, and other providers

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
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

MIT License - see LICENSE file for details

## 🌟 Future Enhancements

- [ ] Support for more LLM providers
- [ ] Sentiment analysis
- [ ] Email categorization
- [ ] Reply generation
- [ ] Multi-language support
- [ ] Offline mode with cached summaries
- [ ] Integration with Teams/SharePoint

## 📞 Support

For issues and questions:
- Check the troubleshooting section
- Review Office Add-ins documentation
- Open GitHub issues for bugs or feature requests 