import { authService } from '../auth/AuthService';

export interface BedrockAgentResponse {
  response: string;
  sessionId?: string;
  citations?: Array<{
    generatedResponsePart: {
      textResponsePart: {
        text: string;
        span: {
          start: number;
          end: number;
        };
      };
    };
    retrievedReferences: Array<{
      content: {
        text: string;
      };
      location: {
        type: string;
        s3Location?: {
          uri: string;
        };
      };
      metadata: Record<string, any>;
    }>;
  }>;
  trace?: {
    failureTrace?: {
      failureReason: string;
      traceId: string;
    };
    orchestrationTrace?: {
      invocationInput?: any;
      modelInvocationInput?: any;
      modelInvocationOutput?: any;
      observation?: any;
      rationale?: {
        text: string;
      };
    };
  };
}

export interface BedrockAgentRequest {
  input: string;
  sessionId?: string;
  enableTrace?: boolean;
  endSession?: boolean;
}

export class BedrockAgentClient {
  private baseUrl: string;
  private apiKey?: string;

  constructor(baseUrl?: string, apiKey?: string) {
    this.baseUrl = baseUrl || 'https://i03qauf1s6.execute-api.us-west-2.amazonaws.com/prod';
    this.apiKey = apiKey;
  }

  async invoke(input: string, sessionId?: string, enableTrace = false): Promise<BedrockAgentResponse> {
    const user = await authService.getUser();
    if (!user) {
      throw new Error('User not authenticated');
    }

    const requestBody: BedrockAgentRequest = {
      input,
      sessionId,
      enableTrace
    };

    const headers: Record<string, string> = {
      'Content-Type': 'application/json'
    };

    // Add authentication - prefer API key, fallback to access token
    if (this.apiKey) {
      headers['X-Api-Key'] = this.apiKey;
    } else {
      const accessToken = await authService.getAccessToken();
      if (accessToken) {
        headers['Authorization'] = `Bearer ${accessToken}`;
      } else {
        // Fallback to ID token for dev/demo purposes
        const idToken = await authService.getIdToken();
        if (idToken) {
          headers['Authorization'] = `Bearer ${idToken}`;
        }
      }
    }

    const response = await fetch(`${this.baseUrl}/agent/invoke`, {
      method: 'POST',
      headers,
      body: JSON.stringify(requestBody)
    });

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
      
      throw new Error(errorMessage);
    }

    const data = await response.json();
    
    // Handle both direct response and wrapped response formats
    if (data.response) {
      return data;
    } else if (data.output) {
      // AWS SDK format - extract the response text
      return {
        response: data.output.text || 'Response received but no text content',
        sessionId: data.sessionId,
        citations: data.citations,
        trace: data.trace
      };
    } else {
      // Fallback for unexpected response format
      return {
        response: JSON.stringify(data),
        sessionId
      };
    }
  }

  async health(): Promise<{ status: string; timestamp: string; version?: string }> {
    const response = await fetch(`${this.baseUrl}/agent/health`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        ...(this.apiKey && { 'X-Api-Key': this.apiKey })
      }
    });

    if (!response.ok) {
      throw new Error(`Health check failed: ${response.status} ${response.statusText}`);
    }

    return await response.json();
  }

  async getSession(sessionId: string): Promise<any> {
    const response = await fetch(`${this.baseUrl}/agent/sessions/${sessionId}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        ...(this.apiKey && { 'X-Api-Key': this.apiKey })
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to get session: ${response.status} ${response.statusText}`);
    }

    return await response.json();
  }

  // Convenience method for email analysis
  async analyzeEmail(emailData: {
    subject: string;
    sender: string;
    recipients: string[];
    body: string;
    ccRecipients?: string[];
  }, sessionId?: string): Promise<BedrockAgentResponse> {
    const emailContext = `
Email Analysis Request:

Subject: ${emailData.subject}
From: ${emailData.sender}
To: ${emailData.recipients.join(', ')}
${emailData.ccRecipients?.length ? `CC: ${emailData.ccRecipients.join(', ')}` : ''}

Email Content:
${emailData.body}

Please analyze this email and provide:
1. A concise summary of the key points
2. Identified action items or requests
3. Sentiment analysis
4. Priority level assessment
5. Suggested response approach
`.trim();

    return this.invoke(emailContext, sessionId, true);
  }

  // Convenience method for seller/business intelligence
  async analyzeBusinessData(data: any, query: string, sessionId?: string): Promise<BedrockAgentResponse> {
    const context = `
Business Data Analysis Request:

Query: ${query}

Data Context:
${JSON.stringify(data, null, 2)}

Please analyze the provided business data and respond to the query with insights, trends, and actionable recommendations.
`.trim();

    return this.invoke(context, sessionId, true);
  }
}

// Singleton instance
export const bedrockAgentClient = new BedrockAgentClient();