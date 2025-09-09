import { authService } from '../auth/AuthService';

export interface LambdaAgentResponse {
  response: string;
  mcp_enabled?: boolean;
  mcp_tools_count?: number;
  status?: string;
}

export interface LambdaAgentRequest {
  prompt: string;
  mcp_authorization_token: string;
}

export class LambdaAgentClient {
  private lambdaFunctionUrl: string;

  constructor(lambdaFunctionUrl?: string) {
    this.lambdaFunctionUrl = lambdaFunctionUrl || 'https://zr5sblu3idcilhcthrpfzulrg40dlpss.lambda-url.us-west-2.on.aws/';
  }

  async invoke(prompt: string): Promise<LambdaAgentResponse> {
    const user = await authService.getUser();
    if (!user) {
      throw new Error('User not authenticated');
    }

    // Get authentication token
    let authToken = '';
    try {
      const accessToken = await authService.getAccessToken();
      if (accessToken) {
        authToken = accessToken;
      } else {
        // Fallback to ID token for dev/demo purposes
        const idToken = await authService.getIdToken();
        if (idToken) {
          authToken = idToken;
        }
      }
    } catch (error) {
      console.error('Failed to get auth token:', error);
      throw new Error('Authentication failed - unable to get token');
    }

    if (!authToken) {
      throw new Error('No authentication token available');
    }

    const requestBody: LambdaAgentRequest = {
      prompt,
      mcp_authorization_token: authToken
    };

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };

    console.log('🚀 Calling Lambda Function URL for email summarization...');
    console.log('🌐 Lambda URL:', this.lambdaFunctionUrl);
    console.log('📝 Prompt length:', prompt.length);

    // Add timeout support (90 seconds to match the Python test script)
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 90000);

    try {
      const response = await fetch(this.lambdaFunctionUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(requestBody),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      console.log('📊 Lambda response status:', response.status);

      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
        
        try {
          const errorData = JSON.parse(errorText);
          errorMessage = errorData.message || errorData.error || errorMessage;
        } catch {
          // If can't parse JSON, use the raw text
          if (errorText) {
            errorMessage = errorText;
          }
        }
        
        console.error('❌ Lambda Function error:', errorMessage);
        throw new Error(errorMessage);
      }

      const data = await response.json();
      console.log('✅ Lambda Function response received');
      console.log('📊 MCP Enabled:', data.mcp_enabled);
      console.log('🔧 MCP Tools Count:', data.mcp_tools_count);
      
      // Return response in expected format
      return {
        response: data.response || 'Lambda response received but no content.',
        mcp_enabled: data.mcp_enabled,
        mcp_tools_count: data.mcp_tools_count,
        status: data.status
      };

    } catch (error) {
      clearTimeout(timeoutId);
      
      if (error.name === 'AbortError') {
        console.error('❌ Lambda Function request timeout (90 seconds)');
        throw new Error('Request timed out after 90 seconds');
      }
      
      console.error('❌ Error calling Lambda Function:', error);
      throw new Error(`Lambda Function call failed: ${(error as Error).message}`);
    }
  }

  async health(): Promise<{ status: string; timestamp: string; version?: string }> {
    // Simple health check - just try to make a basic request
    try {
      const testResponse = await this.invoke('health check');
      return {
        status: 'healthy',
        timestamp: new Date().toISOString(),
        version: 'lambda-function-url'
      };
    } catch (error) {
      throw new Error(`Lambda Function health check failed: ${(error as Error).message}`);
    }
  }
}

// Singleton instance
export const lambdaAgentClient = new LambdaAgentClient();
