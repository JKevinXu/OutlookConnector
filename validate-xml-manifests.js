#!/usr/bin/env node

/**
 * XML Manifest Validation Script
 * Validates Office Add-in XML manifest files
 */

const fs = require('fs');
const path = require('path');

// Simple XML validation function
function validateXMLStructure(xmlContent) {
  const errors = [];
  
  // Check for required elements
  const requiredElements = [
    '<Id>',
    '<Version>',
    '<ProviderName>',
    '<DisplayName',
    '<Description',
    '<IconUrl',
    '<Hosts>',
    '<Requirements>',
    '<Permissions>',
    '<VersionOverrides'
  ];
  
  requiredElements.forEach(element => {
    if (!xmlContent.includes(element)) {
      errors.push(`Missing required element: ${element}`);
    }
  });
  
  // Check for balanced tags (basic check)
  const openTags = (xmlContent.match(/<[^/][^>]*>/g) || []).length;
  const closeTags = (xmlContent.match(/<\/[^>]*>/g) || []).length;
  const selfClosingTags = (xmlContent.match(/<[^>]*\/>/g) || []).length;
  
  if (openTags !== closeTags + selfClosingTags) {
    errors.push('XML tags may not be properly balanced');
  }
  
  // Check for XML declaration
  if (!xmlContent.trim().startsWith('<?xml')) {
    errors.push('Missing XML declaration');
  }
  
  // Check for namespace declarations
  if (!xmlContent.includes('xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"')) {
    errors.push('Missing required Office namespace');
  }
  
  return errors;
}

// Validate URL accessibility (basic check)
function validateUrls(xmlContent) {
  const errors = [];
  const urlPattern = /DefaultValue="(https?:\/\/[^"]+)"/g;
  const urls = [];
  let match;
  
  while ((match = urlPattern.exec(xmlContent)) !== null) {
    urls.push(match[1]);
  }
  
  // Check for localhost in production manifests
  urls.forEach(url => {
    if (url.includes('localhost') && !xmlContent.includes('(Beta)') && !xmlContent.includes('.dev.xml')) {
      errors.push(`Production manifest should not contain localhost URL: ${url}`);
    }
  });
  
  return errors;
}

// Main validation function
function validateManifest(filePath) {
  console.log(`\n🔍 Validating: ${filePath}`);
  
  if (!fs.existsSync(filePath)) {
    console.log(`❌ File not found: ${filePath}`);
    return false;
  }
  
  try {
    const xmlContent = fs.readFileSync(filePath, 'utf8');
    
    // Validate XML structure
    const structureErrors = validateXMLStructure(xmlContent);
    const urlErrors = validateUrls(xmlContent);
    
    const allErrors = [...structureErrors, ...urlErrors];
    
    if (allErrors.length === 0) {
      console.log(`✅ ${path.basename(filePath)} is valid`);
      return true;
    } else {
      console.log(`❌ ${path.basename(filePath)} has errors:`);
      allErrors.forEach(error => console.log(`   - ${error}`));
      return false;
    }
    
  } catch (error) {
    console.log(`❌ Error reading file: ${error.message}`);
    return false;
  }
}

// Find and validate all XML manifest files
function validateAllManifests() {
  console.log('🔍 Office Add-in XML Manifest Validator');
  console.log('=====================================');
  
  const manifestFiles = [
    'manifest.xml',
    'manifest.beta.xml',
    'manifest.dev.xml'
  ];
  
  let allValid = true;
  
  manifestFiles.forEach(file => {
    const isValid = validateManifest(file);
    if (!isValid) allValid = false;
  });
  
  console.log('\n📊 Validation Summary:');
  console.log('======================');
  
  if (allValid) {
    console.log('✅ All XML manifests are valid!');
    process.exit(0);
  } else {
    console.log('❌ Some manifests have validation errors');
    console.log('💡 Please fix the errors above before deployment');
    process.exit(1);
  }
}

// Run validation
if (require.main === module) {
  validateAllManifests();
}

module.exports = { validateManifest, validateXMLStructure, validateUrls }; 