import { Attachment } from '../types/chat';
import { AppSettings } from '../types/settings';

export interface FunctionDeclaration {
  name: string;
  description: string;
  parameters: {
    type: string;
    properties: Record<string, unknown>;
    required?: string[];
  };
}

export interface ToolCall {
  name: string;
  args: Record<string, unknown>;
}

export interface GeminiStreamChunk {
  textDelta?: string;
  toolCalls?: ToolCall[];
  tokenCount?: number;
}

export const geminiService = {
  /**
   * Builds the API endpoint URL taking into account direct Google API or Cloudflare team proxy
   */
  getEndpoint(action: 'streamGenerateContent' | 'generateContent' | 'countTokens', settings: AppSettings): string {
    const model = settings.model || 'gemini-2.5-flash';
    if (settings.teamProxyUrl && settings.teamProxyUrl.trim()) {
      let base = settings.teamProxyUrl.trim().replace(/\/+$/, '');
      return `${base}/v1beta/models/${model}:${action}`;
    }
    return `https://generativelanguage.googleapis.com/v1beta/models/${model}:${action}?key=${settings.geminiApiKey.trim()}`;
  },

  /**
   * Prepares request headers (supports Team Proxy token or Google OAuth bearer token)
   */
  getHeaders(settings: AppSettings): HeadersInit {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json'
    };

    if (settings.teamProxyToken && settings.teamProxyToken.trim()) {
      headers['X-Proxy-Token'] = settings.teamProxyToken.trim();
      headers['Authorization'] = `Bearer ${settings.teamProxyToken.trim()}`;
    } else if (settings.googleAuthToken && settings.googleAuthToken.trim()) {
      headers['Authorization'] = `Bearer ${settings.googleAuthToken.trim()}`;
    }

    return headers;
  },

  /**
   * Estimates prompt token count using API or character heuristic
   */
  async estimateTokens(
    prompt: string,
    attachments: Attachment[] = [],
    settings: AppSettings
  ): Promise<number> {
    try {
      if (settings.geminiApiKey || settings.teamProxyUrl) {
        const url = this.getEndpoint('countTokens', settings);
        const headers = this.getHeaders(settings);

        const parts: unknown[] = [{ text: prompt }];
        for (const att of attachments) {
          parts.push({
            inlineData: {
              mimeType: att.mimeType,
              data: att.base64Data
            }
          });
        }

        const res = await fetch(url, {
          method: 'POST',
          headers,
          body: JSON.stringify({
            contents: [{ role: 'user', parts }]
          })
        });

        if (res.ok) {
          const data = await res.json();
          if (data && typeof data.totalTokens === 'number') {
            return data.totalTokens;
          }
        }
      }
    } catch (e) {
      console.warn('CountTokens call failed, falling back to heuristic:', e);
    }

    // Heuristic: ~4 chars per token for English, ~1.5 chars per token for Arabic
    let charTokens = Math.ceil(prompt.length / 3);
    for (const att of attachments) {
      if (att.type === 'image') charTokens += 1000;
      if (att.type === 'pdf') charTokens += 2000;
    }
    return charTokens;
  },

  /**
   * Streams chat completions with function calling, search, code execution, and attachments
   */
  async *streamChat(options: {
    messages: { role: string; content: string; attachments?: Attachment[] }[];
    systemInstruction?: string;
    tools?: FunctionDeclaration[];
    settings: AppSettings;
    signal?: AbortSignal;
  }): AsyncGenerator<GeminiStreamChunk> {
    const { messages, systemInstruction, tools, settings, signal } = options;

    if (!settings.geminiApiKey && !settings.teamProxyUrl && !settings.googleAuthToken) {
      throw new Error('Please enter your Gemini API key or configure Team Proxy in Settings.');
    }

    const url = `${this.getEndpoint('streamGenerateContent', settings)}&alt=sse`;
    const headers = this.getHeaders(settings);

    // Format contents
    const contents = messages.map(m => {
      const parts: unknown[] = [];
      if (m.attachments && m.attachments.length > 0) {
        for (const att of m.attachments) {
          parts.push({
            inlineData: {
              mimeType: att.mimeType,
              data: att.base64Data
            }
          });
        }
      }
      if (m.content) {
        parts.push({ text: m.content });
      }
      return {
        role: m.role === 'user' ? 'user' : 'model',
        parts
      };
    });

    // Assemble tools
    const geminiTools: unknown[] = [];

    if (tools && tools.length > 0) {
      geminiTools.push({
        functionDeclarations: tools
      });
    }

    if (settings.webSearch) {
      geminiTools.push({ googleSearch: {} });
    }

    if (settings.codeExecution) {
      geminiTools.push({ codeExecution: {} });
    }

    const requestBody: Record<string, unknown> = {
      contents,
      generationConfig: {
        temperature: 0.3,
        topP: 0.95
      }
    };

    if (systemInstruction) {
      requestBody.systemInstruction = {
        role: 'system',
        parts: [{ text: systemInstruction }]
      };
    }

    if (geminiTools.length > 0) {
      requestBody.tools = geminiTools;
    }

    const response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(requestBody),
      signal
    });

    if (!response.ok) {
      const errText = await response.text();
      let errorMsg = `Gemini API error (${response.status}): ${errText}`;
      try {
        const errJson = JSON.parse(errText);
        if (errJson.error?.message) {
          errorMsg = errJson.error.message;
        }
      } catch {
        // use raw text
      }
      throw new Error(errorMsg);
    }

    if (!response.body) {
      throw new Error('No response body returned from Gemini.');
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || !trimmed.startsWith('data:')) continue;

        const jsonStr = trimmed.replace(/^data:\s*/, '');
        if (jsonStr === '[DONE]') continue;

        try {
          const chunk = JSON.parse(jsonStr);
          const candidate = chunk.candidates?.[0];
          if (!candidate) continue;

          const parts = candidate.content?.parts;
          if (parts && Array.isArray(parts)) {
            for (const part of parts) {
              if (part.text) {
                yield { textDelta: part.text };
              }
              if (part.functionCall) {
                yield {
                  toolCalls: [
                    {
                      name: part.functionCall.name,
                      args: part.functionCall.args || {}
                    }
                  ]
                };
              }
            }
          }
        } catch {
          // partial line or parsing skip
        }
      }
    }
  }
};
