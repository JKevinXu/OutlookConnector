// Add-in Configuration
// Change these URLs to match your deployment setup

module.exports = {
  // Development server URL (for local testing)
  development: "https://localhost:3000/",
  
  // Production deployment URL (GitHub Pages, Azure, etc.)
  production: "https://jkevinxu.github.io/OutlookConnector/",
  
  // Office Add-in settings
  addin: {
    name: "AM Personal Assistant",
    version: "1.0.0",
    developer: {
      name: "AM Solutions",
      website: "https://www.amsolutions.com"
    }
  }
}; 