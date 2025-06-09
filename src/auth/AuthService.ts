/*
 * OIDC Authentication Service
 * Based on Amazon-style implicit flow implementation
 */

import {
    UserManager,
    User,
    WebStorageStateStore,
    UserManagerSettings
} from 'oidc-client';
import { OIDCConfig, UserProfile, AuthState, AuthenticationError } from '../types/auth';

export class AuthService {
    private userManager: UserManager;
    private authState: AuthState;
    private eventListeners: { [key: string]: Function[] } = {};

    constructor(config?: OIDCConfig) {
        // Amazon federated configuration
        const baseUrl = window.location.origin;
        const defaultConfig: OIDCConfig = {
            authority: 'https://idp.federate.amazon.com',
            clientId: 'amc-qbiz-aud',
            redirectUri: `${baseUrl}/taskpane/callback`,
            postLogoutRedirectUri: `${baseUrl}/taskpane/callback`,
            scope: 'openid profile email',
            responseType: 'id_token' // Implicit flow - TODO: should use auth code with PKCE, and upgrade from oidc-client to oidc-client-ts
        };

        const finalConfig = config || defaultConfig;
        
        // Detect Office Add-in environment
        const isOfficeAddIn = typeof Office !== 'undefined';
        
        const settings: UserManagerSettings = {
            authority: finalConfig.authority,
            client_id: finalConfig.clientId,
            redirect_uri: finalConfig.redirectUri,
            post_logout_redirect_uri: finalConfig.postLogoutRedirectUri,
            response_type: finalConfig.responseType,
            scope: finalConfig.scope,
            stateStore: new WebStorageStateStore({ store: window.localStorage }),
            loadUserInfo: true,
            
            // Disable automatic silent renewal in Office Add-in environments
            automaticSilentRenew: !isOfficeAddIn,
            includeIdTokenInSilentRenew: !isOfficeAddIn,
            monitorSession: !isOfficeAddIn,
            
            // Shorter expiration notification time for better UX
            accessTokenExpiringNotificationTime: 300, // 5 minutes before expiration
            
            filterProtocolClaims: true,
        };

        this.userManager = new UserManager(settings);
        this.authState = {
            isAuthenticated: false,
            user: null,
            isLoading: false,
            error: null
        };

        this.setupEventHandlers();
    }

    private setupEventHandlers(): void {
        this.userManager.events.addUserLoaded((user: User) => {
            console.log('🔓 User loaded:', user);
            this.authState.isAuthenticated = true;
            this.authState.user = this.extractUserProfile(user);
            this.authState.isLoading = false;
            this.authState.error = null;
            this.emit('userLoaded', this.authState.user);
        });

        this.userManager.events.addUserUnloaded(() => {
            console.log('🔒 User unloaded');
            this.authState.isAuthenticated = false;
            this.authState.user = null;
            this.authState.isLoading = false;
            this.emit('userUnloaded');
        });

        this.userManager.events.addAccessTokenExpiring(() => {
            console.log('⏰ Access token expiring');
            this.emit('accessTokenExpiring');
        });

        this.userManager.events.addAccessTokenExpired(() => {
            console.log('❌ Access token expired');
            this.emit('accessTokenExpired');
        });

        this.userManager.events.addSilentRenewError((error: Error) => {
            console.error('🔄 Silent renew error:', error);
            this.authState.error = error.message;
            this.emit('silentRenewError', error);
        });
    }

    // Event system
    public on(event: string, callback: Function): void {
        if (!this.eventListeners[event]) {
            this.eventListeners[event] = [];
        }
        this.eventListeners[event].push(callback);
    }

    private emit(event: string, data?: any): void {
        if (this.eventListeners[event]) {
            this.eventListeners[event].forEach(callback => callback(data));
        }
    }

    // Initialize authentication state
    public async initialize(): Promise<void> {
        try {
            this.authState.isLoading = true;
            
            // Handle callback if we're returning from authentication
            if (window.location.hash && window.location.hash.includes('id_token')) {
                await this.handleCallback();
                return;
            }

            // Check for existing user
            const user = await this.userManager.getUser();
            if (user && !user.expired) {
                this.authState.isAuthenticated = true;
                this.authState.user = this.extractUserProfile(user);
                console.log('✅ Existing user found:', this.authState.user);
            }
        } catch (error) {
            console.error('❌ Initialization error:', error);
            this.authState.error = error instanceof Error ? error.message : 'Initialization failed';
        } finally {
            this.authState.isLoading = false;
        }
    }

    // Trigger login
    public async login(): Promise<void> {
        try {
            this.authState.isLoading = true;
            this.authState.error = null;
            console.log('🚀 Starting login...');
            
            await this.userManager.signinRedirect();
        } catch (error) {
            console.error('❌ Login error:', error);
            this.authState.error = error instanceof Error ? error.message : 'Login failed';
            this.authState.isLoading = false;
            throw new AuthenticationError('Login failed', 'LOGIN_ERROR');
        }
    }

    // Handle the callback after login
    public async handleCallback(): Promise<void> {
        try {
            console.log('🔄 Handling authentication callback...');
            const user = await this.userManager.signinRedirectCallback();
            console.log('✅ Callback handled successfully:', user);
            
            // Clean up URL only if history API is available
            if (window.history && typeof window.history.replaceState === 'function') {
                window.history.replaceState(null, '', window.location.pathname);
            } else {
                console.log('ℹ️ History API not available in this environment (Office Add-in)');
            }
        } catch (error) {
            console.error('❌ Callback error:', error);
            this.authState.error = error instanceof Error ? error.message : 'Callback failed';
            throw new AuthenticationError('Authentication callback failed', 'CALLBACK_ERROR');
        }
    }

    // Get the current user
    public async getUser(): Promise<UserProfile | null> {
        try {
            const user = await this.userManager.getUser();
            return user ? this.extractUserProfile(user) : null;
        } catch (error) {
            console.error('❌ Error getting user:', error);
            return null;
        }
    }

    // Get access token
    public async getAccessToken(): Promise<string | null> {
        try {
            const user = await this.userManager.getUser();
            return user?.access_token || null;
        } catch (error) {
            console.error('❌ Error getting access token:', error);
            return null;
        }
    }

    // Get ID token with automatic renewal if expired
    public async getIdToken(): Promise<string | null> {
        try {
            let user = await this.userManager.getUser();
            
            // Check if token is expired or about to expire (within 5 minutes)
            if (!user || this.isTokenExpired(user)) {
                console.log('🔄 Token expired or about to expire, attempting renewal...');
                try {
                    await this.renewToken();
                    user = await this.userManager.getUser();
                } catch (renewError) {
                    console.error('❌ Token renewal failed:', renewError);
                    // Token renewal failed, user needs to re-authenticate
                    this.authState.isAuthenticated = false;
                    this.authState.user = null;
                    this.emit('tokenExpired');
                    return null;
                }
            }
            
            return user?.id_token || null;
        } catch (error) {
            console.error('❌ Error getting ID token:', error);
            return null;
        }
    }

    // Check if token is expired or about to expire
    private isTokenExpired(user: User): boolean {
        if (!user) {
            console.log('🚫 Token check: No user object');
            return true;
        }
        
        // Try to get expiration from user.expires_at first
        let expiresAt = user.expires_at;
        
        // If expires_at is not set (common with implicit flow), decode the ID token
        if (!expiresAt && user.id_token) {
            try {
                // Decode the JWT token to get the exp claim
                const tokenParts = user.id_token.split('.');
                if (tokenParts.length === 3) {
                    const payload = JSON.parse(atob(tokenParts[1]));
                    expiresAt = payload.exp;
                    console.log('🔍 Extracted expiration from ID token:', expiresAt);
                }
            } catch (error) {
                console.log('❌ Failed to decode ID token for expiration:', error);
            }
        }
        
        if (!expiresAt) {
            console.log('🚫 Token check: No expiration time available');
            console.log('🔍 User object keys:', Object.keys(user));
            console.log('🔍 User expires_at:', user.expires_at);
            console.log('🔍 Has ID token:', !!user.id_token);
            return true;
        }
        
        const isOfficeAddIn = typeof Office !== 'undefined';
        // In Office Add-ins, only check for actual expiration since renewal isn't supported
        // In regular web apps, use a 1-minute buffer for proactive renewal
        const expirationBuffer = isOfficeAddIn ? 0 : 60;
        const currentTime = Math.floor(Date.now() / 1000);
        const timeUntilExpiry = expiresAt - currentTime;
        const willExpireSoon = timeUntilExpiry <= expirationBuffer;
        
        console.log('🕐 Token expiration check:', {
            isOfficeAddIn: isOfficeAddIn,
            currentTime: currentTime,
            expiresAt: expiresAt,
            timeUntilExpiry: timeUntilExpiry,
            timeUntilExpiryMinutes: Math.round(timeUntilExpiry / 60),
            expirationBuffer: expirationBuffer,
            willExpireSoon: willExpireSoon,
            currentTimeReadable: new Date(currentTime * 1000).toISOString(),
            expiresAtReadable: new Date(expiresAt * 1000).toISOString()
        });
        
        return willExpireSoon;
    }

    // Enhanced token renewal with retry logic
    public async renewToken(): Promise<void> {
        const isOfficeAddIn = typeof Office !== 'undefined';
        
        // In Office Add-in environments, silent renewal is not supported
        if (isOfficeAddIn) {
            console.log('⚠️ Office Add-in environment detected - silent renewal not supported');
            console.log('🔄 Token expired. User needs to sign in again.');
            
            // Clear auth state and require interactive login
            this.authState.isAuthenticated = false;
            this.authState.user = null;
            this.authState.error = 'Your session has expired. Please sign in again.';
            this.emit('tokenExpired');
            throw new AuthenticationError('Your session has expired. Please sign in again to continue.', 'TOKEN_EXPIRED');
        }
        
        try {
            console.log('🔄 Attempting token renewal...');
            const renewedUser = await this.userManager.signinSilent();
            
            if (renewedUser) {
                console.log('✅ Token renewed successfully');
                this.authState.user = this.extractUserProfile(renewedUser);
                this.emit('tokenRenewed', this.authState.user);
            }
        } catch (error) {
            console.error('❌ Token renewal failed:', error);
            
            // Handle specific renewal errors
            if (error.message?.includes('Frame window timed out')) {
                console.log('🕒 Silent renewal timed out');
                this.handleRenewalTimeout();
            } else if (error.message?.includes('login_required') || 
                       error.message?.includes('interaction_required')) {
                console.log('🔐 Interactive login required');
                this.handleInteractionRequired();
            } else {
                console.log('❌ Unexpected renewal error');
                this.handleRenewalError();
            }
            
            throw new AuthenticationError('Your session has expired. Please sign in again to continue.', 'TOKEN_EXPIRED');
        }
    }

    // Handle silent renewal timeout (common in Office Add-ins)
    private handleRenewalTimeout(): void {
        console.log('🔄 Handling renewal timeout...');
        this.authState.isAuthenticated = false;
        this.authState.user = null;
        this.authState.error = 'Your session has expired. Please sign in again.';
        this.emit('tokenExpired');
    }

    // Handle cases where user interaction is required
    private handleInteractionRequired(): void {
        console.log('🔐 User interaction required for renewal...');
        this.authState.isAuthenticated = false;
        this.authState.user = null;
        this.authState.error = 'Please sign in again to continue.';
        this.emit('loginRequired');
    }

    // Handle other renewal errors
    private handleRenewalError(): void {
        console.log('❌ General renewal error occurred...');
        this.authState.isAuthenticated = false;
        this.authState.user = null;
        this.authState.error = 'Authentication failed. Please sign in again.';
        this.emit('authenticationFailed');
    }

    // Logout
    public async logout(): Promise<void> {
        try {
            console.log('🚪 Starting logout...');
            await this.userManager.signoutRedirect();
        } catch (error) {
            console.error('❌ Logout error:', error);
            throw new AuthenticationError('Logout failed', 'LOGOUT_ERROR');
        }
    }
    
    // Sign out (clear local state without redirect - better for Office Add-ins)
    public async signOut(): Promise<void> {
        try {
            console.log('🚪 Signing out and clearing local state...');
            
            // Clear the user manager state
            await this.userManager.removeUser();
            
            // Clear auth state
            this.authState.isAuthenticated = false;
            this.authState.user = null;
            this.authState.error = null;
            this.authState.isLoading = false;
            
            // Clear any additional storage items related to OIDC
            const storageKeys = Object.keys(localStorage);
            storageKeys.forEach(key => {
                if (key.startsWith('oidc.') || key.includes('auth') || key.includes('token')) {
                    localStorage.removeItem(key);
                }
            });
            
            // Emit event
            this.emit('userSignedOut');
            
            console.log('✅ Successfully signed out');
        } catch (error) {
            console.error('❌ Sign out error:', error);
            // Don't throw error - we want to clear state even if there are issues
        }
    }

    // Check if user is authenticated
    public isAuthenticated(): boolean {
        return this.authState.isAuthenticated;
    }

    // Get current auth state
    public getAuthState(): AuthState {
        return { ...this.authState };
    }

    // Extract user profile from OIDC user
    private extractUserProfile(user: User): UserProfile {
        const profile = user.profile;
        return {
            sub: profile.sub,
            name: profile.name || profile.preferred_username || 'Unknown',
            email: profile.email || '',
            email_verified: profile.email_verified,
            org: profile.org || profile.organization,
            roles: profile.roles ? (Array.isArray(profile.roles) ? profile.roles : [profile.roles]) : [],
            iat: profile.iat || Math.floor(Date.now() / 1000),
            exp: profile.exp || Math.floor(Date.now() / 1000) + 3600
        };
    }
}

// Create a singleton instance
export const authService = new AuthService(); 