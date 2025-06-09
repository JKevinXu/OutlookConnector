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
      : 'https://g3kt5l83j8.execute-api.us-east-1.amazonaws.com/prod',
    region: 'us-east-1'
  }
};

// Parameters for the seller history API
interface SellerHistoryParams {
  marketplaceId?: string;      // Optional, defaults to JP marketplace
}

export class SellerHistoryService {
  private environment: 'prod' = 'prod';
  private readonly MERCHANT_ID = '7956983745'; // Hardcoded merchant ID

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

      // Get access token from auth service
      const accessToken = await authService.getIdToken(); // Use ID token for now
      if (!accessToken) {
        throw new Error('No access token available. Please sign in again.');
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
      
      // Try different approaches to handle CORS
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
        if (response.status === 502) {
          const errorData: ErrorResponse = await response.json();
          throw new Error(`Bad Gateway: ${errorData.message}`);
        } else if (response.status === 401) {
          throw new Error('Unauthorized: Please sign in again.');
        } else if (response.status === 403) {
          throw new Error('Forbidden: You don\'t have permission to access this resource.');
        }
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      
      const data: SellerHistoryResponse = await response.json();
      console.log('✅ Seller history data received:', data);
      return data;
      
    } catch (error) {
      console.error('❌ Error calling seller history API:', error);
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