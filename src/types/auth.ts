/*
 * OIDC Authentication Type Definitions
 */

export interface OIDCConfig {
  authority: string;
  clientId: string;
  redirectUri: string;
  postLogoutRedirectUri?: string;
  scope: string;
  responseType: string;
  popupWindowFeatures?: string;
  popupWindowTarget?: string;
}

export interface UserProfile {
  sub: string;                // Subject identifier
  name: string;               // Full name
  email: string;              // Email address
  email_verified?: boolean;   // Email verification status
  org?: string;               // Organization claim
  roles?: string[];           // User roles
  iat: number;                // Issued at
  exp: number;                // Expiration time
}

export interface AuthState {
  isAuthenticated: boolean;
  user: UserProfile | null;
  isLoading: boolean;
  error: string | null;
}

export interface AuthEvents {
  onUserLoaded: (user: UserProfile) => void;
  onUserUnloaded: () => void;
  onAccessTokenExpiring: () => void;
  onAccessTokenExpired: () => void;
  onSilentRenewError: (error: Error) => void;
}

export class AuthenticationError extends Error {
  constructor(message: string, public code?: string) {
    super(message);
    this.name = 'AuthenticationError';
  }
} 