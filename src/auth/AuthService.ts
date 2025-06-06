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
        
        const settings: UserManagerSettings = {
            authority: finalConfig.authority,
            client_id: finalConfig.clientId,
            redirect_uri: finalConfig.redirectUri,
            post_logout_redirect_uri: finalConfig.postLogoutRedirectUri,
            response_type: finalConfig.responseType,
            scope: finalConfig.scope,
            stateStore: new WebStorageStateStore({ store: window.localStorage }),
            loadUserInfo: true,
            automaticSilentRenew: true,
            includeIdTokenInSilentRenew: true,
            monitorSession: true,
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
            
            // Clean up URL
            window.history.replaceState(null, '', window.location.pathname);
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

    // Get ID token
    public async getIdToken(): Promise<string | null> {
        try {
            const user = await this.userManager.getUser();
            return user?.id_token || null;
        } catch (error) {
            console.error('❌ Error getting ID token:', error);
            return null;
        }
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

    // Renew token silently
    public async renewToken(): Promise<void> {
        try {
            console.log('🔄 Renewing token...');
            await this.userManager.signinSilent();
        } catch (error) {
            console.error('❌ Token renewal error:', error);
            throw new AuthenticationError('Token renewal failed', 'RENEWAL_ERROR');
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