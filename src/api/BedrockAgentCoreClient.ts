import { authService } from '../auth/AuthService';

export interface BedrockAgentCoreResponse {
  response: string;
  mcp_enabled?: boolean;
  mcp_tools_count?: number;
  status?: string;
  sessionId?: string;
}

export interface BedrockAgentCoreRequest {
  prompt: string;
  mcp_authorization_token: string;
  sessionId?: string;
  enableTrace?: boolean;
  endSession?: boolean;
}

export class BedrockAgentCoreClient {
  private baseUrl: string;
  private agentArn: string;
  private awsRegion: string;

  constructor(agentArn?: string, awsRegion?: string) {
    this.agentArn = agentArn || 'arn:aws:bedrock-agentcore:us-west-2:925509123747:runtime/amc_pa_strands_beta-MItdGrBq0E';
    this.awsRegion = awsRegion || 'us-west-2';
    
    // Bedrock Agent Core HTTP endpoint format
    const escapedAgentArn = encodeURIComponent(this.agentArn);
    this.baseUrl = `https://bedrock-agentcore.${this.awsRegion}.amazonaws.com/runtimes/${escapedAgentArn}/invocations?qualifier=DEFAULT`;
  }

  async invoke(prompt: string, sessionId?: string, enableTrace = false): Promise<BedrockAgentCoreResponse> {
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

    const requestBody: BedrockAgentCoreRequest = {
      prompt,
      mcp_authorization_token: authToken,
      sessionId,
      enableTrace,
      endSession: false
    };

    const headers: Record<string, string> = {
      'Authorization': `Bearer ${authToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };

    console.log('🚀 Calling Bedrock Agent Core for email analysis...');
    console.log('🤖 Agent ARN:', this.agentArn);
    console.log('🌐 AWS Region:', this.awsRegion);
    console.log('🔗 Endpoint URL:', this.baseUrl);
    console.log('📝 Prompt length:', prompt.length);

    // Add timeout support (120 seconds for Bedrock Agent Core)
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 120000);

    try {
      const response = await fetch(this.baseUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(requestBody),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      console.log('📊 Bedrock Agent Core response status:', response.status);
      console.log('📊 Response headers:', Object.fromEntries(response.headers.entries()));

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
        
        console.error('❌ Bedrock Agent Core error:', errorMessage);
        throw new Error(errorMessage);
      }

      const data = await response.json();

      // Handle Bedrock Agent Core response format (may be wrapped in API Gateway format)
      let actualResponse = data;
      if (data.body && typeof data.body === 'string') {
        try {
          actualResponse = JSON.parse(data.body);
          console.log('📦 Unwrapped API Gateway response body');
        } catch (e) {
          console.warn("Could not parse response body as JSON:", e);
          actualResponse = data;
        }
      }

      console.log('✅ Bedrock Agent Core response received');
      console.log('📊 MCP Enabled:', actualResponse.mcp_enabled);
      console.log('🔧 MCP Tools Count:', actualResponse.mcp_tools_count);
      console.log('📋 Response keys:', Object.keys(actualResponse));
      
      // Return response in expected format
      return {
        response: actualResponse.response || 'Bedrock Agent Core response received but no content.',
        mcp_enabled: actualResponse.mcp_enabled,
        mcp_tools_count: actualResponse.mcp_tools_count,
        status: actualResponse.status,
        sessionId: actualResponse.sessionId
      };

    } catch (error) {
      clearTimeout(timeoutId);
      
      if (error.name === 'AbortError') {
        console.error('❌ Bedrock Agent Core request timeout (120 seconds)');
        throw new Error('Request timed out after 120 seconds');
      }
      
      console.error('❌ Error calling Bedrock Agent Core:', error);
      throw new Error(`Bedrock Agent Core call failed: ${(error as Error).message}`);
    }
  }

  async health(): Promise<{ status: string; timestamp: string; version?: string }> {
    // Simple health check - just try to make a basic request
    try {
      const testResponse = await this.invoke('health check');
      return {
        status: 'healthy',
        timestamp: new Date().toISOString(),
        version: 'bedrock-agent-core'
      };
    } catch (error) {
      throw new Error(`Bedrock Agent Core health check failed: ${(error as Error).message}`);
    }
  }
}

// Singleton instance
export const bedrockAgentCoreClient = new BedrockAgentCoreClient();
