---

## 🔒 **POST-MIGRATION: Secure Authentication Flow**

### **Enhanced Security Architecture (After Migration)**

The following diagram shows how the authentication flow will work **AFTER** implementing the Authorization Code + PKCE migration:

```mermaid
sequenceDiagram
    participant User
    participant AddIn as "Office Add-in<br/>(Client)"
    participant AuthService as "AuthService<br/>(oidc-client-ts)"
    participant Amazon as "Amazon IDP<br/>idp.federate.amazon.com"
    participant TokenStore as "Secure Token Storage<br/>(Encrypted)"

    Note over AddIn,Amazon: 🔒 SECURE POST-MIGRATION FLOW (Authorization Code + PKCE)

    Note over AddIn,Amazon: 1. Secure Configuration & PKCE Setup
    AddIn->>AuthService: Initialize with secure config
    Note right of AuthService: responseType: 'code'<br/>usePkce: true<br/>sessionStorage + encryption
    AuthService->>AuthService: Generate code_verifier (random 128 chars)
    AuthService->>AuthService: Generate code_challenge (SHA256 hash)

    Note over AddIn,Amazon: 2. Secure Authorization Request
    User->>AddIn: Click Login
    AddIn->>AuthService: login()
    AuthService->>Amazon: GET /.well-known/openid_configuration
    Amazon-->>AuthService: Return IDP metadata & endpoints
    AuthService->>Amazon: Redirect to /authorize with PKCE
    Note right of Amazon: URL: https://idp.federate.amazon.com/authorize<br/>?client_id=amc-qbiz-aud<br/>&response_type=code<br/>&scope=openid profile email<br/>&redirect_uri=.../taskpane.html<br/>&code_challenge=SHA256_HASH<br/>&code_challenge_method=S256<br/>&state=random_state<br/>&nonce=random_nonce

    Note over AddIn,Amazon: 3. User Authentication (Same as Before)
    Amazon->>User: Show Amazon login page
    User->>Amazon: Enter Amazon credentials + MFA
    Amazon->>Amazon: Validate credentials
    Amazon->>Amazon: Generate authorization code

    Note over AddIn,Amazon: 4. Secure Code Response (No Token Exposure)
    Amazon->>AddIn: Redirect with authorization code
    Note right of AddIn: URL: .../taskpane.html<br/>?code=AUTHORIZATION_CODE<br/>&state=random_state<br/>✅ NO TOKENS IN URL!

    Note over AddIn,Amazon: 5. Secure Token Exchange (Backend-to-Backend)
    AddIn->>AuthService: handleCallback()
    AuthService->>AuthService: Validate state parameter
    AuthService->>AuthService: Extract authorization code
    AuthService->>Amazon: POST /token with PKCE verification
    Note right of Amazon: Body: grant_type=authorization_code<br/>client_id=amc-qbiz-aud<br/>code=AUTHORIZATION_CODE<br/>redirect_uri=.../taskpane.html<br/>code_verifier=ORIGINAL_VERIFIER
    Amazon->>Amazon: Verify code_challenge matches code_verifier
    Amazon-->>AuthService: Return token response (JSON)
    Note left of Amazon: Response: {<br/>  "access_token": "...",<br/>  "id_token": "...",<br/>  "refresh_token": "...",<br/>  "token_type": "Bearer",<br/>  "expires_in": 3600<br/>}

    Note over AddIn,Amazon: 6. Secure Token Storage & Validation
    AuthService->>AuthService: Validate JWT signature with JWKS
    AuthService->>AuthService: Verify nonce, iat, exp claims
    AuthService->>AuthService: Extract user profile from JWT
    AuthService->>TokenStore: Encrypt and store tokens securely
    Note right of TokenStore: sessionStorage with AES encryption<br/>✅ No localStorage exposure!
    AuthService->>AddIn: Authentication complete (secure)

    Note over AddIn,Amazon: 7. Enhanced Session Management
    loop Secure Token Lifecycle
        AuthService->>TokenStore: Check encrypted token expiration
        alt Token expired but refresh available
            AuthService->>Amazon: POST /token (refresh_token grant)
            Amazon-->>AuthService: New access_token + id_token
            AuthService->>TokenStore: Update encrypted tokens
            AuthService->>AddIn: Transparent renewal complete
        else Token expired, no refresh
            AuthService->>TokenStore: Clear all encrypted tokens
            AuthService->>AddIn: Re-authentication required
        else Token valid
            AuthService->>AddIn: Continue with authenticated session
        end
    end

    Note over AddIn,Amazon: 8. Secure Logout
    User->>AddIn: Click Logout
    AddIn->>AuthService: logout()
    AuthService->>TokenStore: Clear all encrypted tokens
    AuthService->>Amazon: GET /logout (optional)
    Amazon-->>AuthService: Logout confirmation
    AuthService->>AddIn: Logout complete
```

### **🔒 Key Security Improvements in New Flow**

#### **1. PKCE Protection**
```typescript
// Code verifier generation (cryptographically secure random)
const codeVerifier = generateRandomString(128); // Base64URL-encoded

// Code challenge generation (SHA256 hash)
const codeChallenge = base64URLEncode(sha256(codeVerifier));

// Authorization request includes challenge
const authUrl = `${authority}/authorize?` +
    `code_challenge=${codeChallenge}&` +
    `code_challenge_method=S256&` +
    // ... other parameters
```

#### **2. No Token Exposure**
- ✅ **Before**: `#id_token=eyJ...` (exposed in URL)
- ✅ **After**: `?code=AUTH_CODE` (short-lived, single-use code)

#### **3. Secure Token Exchange**
```typescript
// Backend-to-backend token exchange (no client secret needed)
const tokenResponse = await fetch(`${authority}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
        grant_type: 'authorization_code',
        client_id: clientId,
        code: authorizationCode,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier // PKCE verification
    })
});
```

#### **4. Encrypted Token Storage**
```typescript
// Secure token storage implementation
class SecureTokenService {
    private readonly ENCRYPTION_KEY = 'derived-from-secure-source';
    
    async storeTokens(tokens: TokenResponse): Promise<void> {
        const encrypted = CryptoJS.AES.encrypt(
            JSON.stringify(tokens), 
            this.ENCRYPTION_KEY
        ).toString();
        
        // Use sessionStorage instead of localStorage
        sessionStorage.setItem('encrypted_auth_tokens', encrypted);
    }
    
    async getTokens(): Promise<TokenResponse | null> {
        const encrypted = sessionStorage.getItem('encrypted_auth_tokens');
        if (!encrypted) return null;
        
        const decrypted = CryptoJS.AES.decrypt(encrypted, this.ENCRYPTION_KEY);
        return JSON.parse(decrypted.toString(CryptoJS.enc.Utf8));
    }
}
```

#### **5. Enhanced Token Refresh**
```typescript
// Automatic token refresh with refresh_token
async refreshTokens(): Promise<void> {
    const currentTokens = await this.tokenService.getTokens();
    if (!currentTokens?.refresh_token) {
        throw new Error('No refresh token available');
    }
    
    const refreshResponse = await fetch(`${authority}/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'refresh_token',
            client_id: clientId,
            refresh_token: currentTokens.refresh_token
        })
    });
    
    const newTokens = await refreshResponse.json();
    await this.tokenService.storeTokens(newTokens);
}
```

### **🛡️ Security Benefits Summary**

| Security Aspect | Before (Implicit) | After (Code + PKCE) |
|------------------|-------------------|---------------------|
| **Token Exposure** | ❌ Tokens in URL hash | ✅ No token exposure |
| **CSRF Protection** | ⚠️ Basic state param | ✅ PKCE + state validation |
| **Token Storage** | ❌ Plain localStorage | ✅ Encrypted sessionStorage |
| **Token Refresh** | ❌ No refresh tokens | ✅ Secure refresh mechanism |
| **Replay Attacks** | ⚠️ Limited protection | ✅ Single-use codes + PKCE |
| **XSS Resistance** | ❌ Vulnerable | ✅ Encrypted storage |
| **Browser History** | ❌ Tokens logged | ✅ Clean URLs |
| **Network Logs** | ❌ Token visibility | ✅ Code-only visibility |

### **📋 Migration Validation Checklist**

- [ ] ✅ Authorization code received instead of tokens
- [ ] ✅ PKCE code_challenge/code_verifier flow working
- [ ] ✅ Token exchange via POST request (not URL)
- [ ] ✅ Encrypted token storage implemented
- [ ] ✅ Refresh token mechanism functional
- [ ] ✅ No tokens visible in browser developer tools
- [ ] ✅ No tokens in browser history
- [ ] ✅ Silent token renewal working
- [ ] ✅ Office Add-in compatibility maintained
- [ ] ✅ Amazon IDP integration successful

--- 