/*
 * Seller History API Service
 * Integrates with Office Add-in authentication
 */

import { authService } from '../auth/AuthService';

// Types for the API response
interface SellerInfo {
  [key: string]: any; // The schema doesn't specify exact structure
}

interface Activity {
  [key: string]: any; // The schema doesn't specify exact structure
}

interface Opportunity {
  [key: string]: any; // The schema doesn't specify exact structure
}

interface SellerHistoryResponse {
  seller_info: SellerInfo;
  activities: Activity[];
  opportunities: Opportunity[];
}

interface ErrorResponse {
  message: string;
}

// API configuration
const API_CONFIG = {
  prod: {
    baseUrl: (typeof window !== 'undefined' && window.location.hostname === 'localhost')
      ? '/api'  // Use webpack proxy during development on localhost
      : 'https://bwzo9wnhy3.execute-api.us-west-2.amazonaws.com/seller-history-prod-auth', // Direct API call for all environments
    region: 'us-west-2'
  }
};

// Parameters for the seller history API
interface SellerHistoryParams {
  marketplaceId?: string;      // Optional, defaults to JP marketplace
}

export class SellerHistoryService {
  private environment: 'prod' = 'prod';
  private readonly MERCHANT_ID = '7489395755'; // Hardcoded merchant ID

  constructor() {
    this.environment = 'prod';
  }

  // Get seller history using authenticated user's token
  async getSellerHistory(params: SellerHistoryParams = {}): Promise<SellerHistoryResponse> {
    try {
      // Check if user is authenticated
      if (!authService.isAuthenticated()) {
        throw new Error('User not authenticated. Please sign in first.');
      }

      return await this.makeApiCallWithRetry(params);
      
    } catch (error) {
      console.error('❌ Error calling seller history API:', error);
      throw error;
    }
  }

  // Make API call with automatic retry on token expiration
  private async makeApiCallWithRetry(params: SellerHistoryParams, retryCount = 0): Promise<SellerHistoryResponse> {
    const maxRetries = 1; // Only retry once for token refresh
    
    try {
      // Get access token from auth service (with automatic renewal)
      const accessToken = await authService.getIdToken();
      if (!accessToken) {
        throw new Error('Your session has expired and could not be renewed. Please sign in again to continue.');
      }

      console.log('🔍 Calling seller history API with hardcoded merchant ID:', this.MERCHANT_ID);
      
      const config = API_CONFIG.prod;
      console.log('🌐 Using API base URL:', config.baseUrl);
      
      // Build query parameters with hardcoded merchant ID
      const queryParams = new URLSearchParams({
        merchantId: this.MERCHANT_ID,
        ...(params.marketplaceId && { marketplaceId: params.marketplaceId })
      });
      
      const url = `${config.baseUrl}/seller-history?${queryParams.toString()}`;
      console.log('📡 API URL:', url);
      
      const requestOptions = {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json',
        },
        mode: 'cors' as RequestMode,
      };
      
      console.log('📝 Request options:', requestOptions);
      
      const response = await fetch(url, requestOptions);
      
      console.log('📊 API Response status:', response.status);
      
      if (!response.ok) {
        // Handle authentication/authorization errors with retry
        if ((response.status === 401 || response.status === 403) && retryCount < maxRetries) {
          console.log('🔄 Token may be expired, attempting renewal and retry...');
          
          try {
            // Force token renewal
            await authService.renewToken();
            console.log('✅ Token renewed, retrying API call...');
            
            // Retry the API call with the new token
            return await this.makeApiCallWithRetry(params, retryCount + 1);
          } catch (renewError) {
            console.error('❌ Token renewal failed during retry:', renewError);
            throw new Error('Your session has expired and could not be renewed. Please sign in again.');
          }
        }
        
        // Handle other errors
        if (response.status === 502) {
          const errorData: ErrorResponse = await response.json().catch(() => ({ message: 'Bad Gateway' }));
          throw new Error(`Server error: ${errorData.message}`);
        } else if (response.status === 401) {
          throw new Error('Your session has expired. Please sign in again.');
        } else if (response.status === 403) {
          throw new Error('Access denied. Your session may have expired or you may not have permission. Please sign in again.');
        }
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      
      const data: SellerHistoryResponse = await response.json();
      console.log('✅ Seller history data received:', data);
      return data;
      
    } catch (error) {
      // If this is not a retry and the error suggests token issues, don't retry again
      if (retryCount >= maxRetries) {
        console.error('❌ API call failed after retry:', error);
      }
      throw error;
    }
  }

  // Get current merchant ID (now hardcoded)
  getCurrentMerchantId(): string {
    return this.MERCHANT_ID;
  }

  // Convenience method to get seller history (simplified since merchant ID is hardcoded)
  async getCurrentUserSellerHistory(marketplaceId?: string): Promise<SellerHistoryResponse> {
    return this.getSellerHistory({
      marketplaceId
    });
  }

  // Get current environment
  getEnvironment(): 'prod' {
    return this.environment;
  }
}

// Create a singleton instance
export const sellerHistoryService = new SellerHistoryService(); 