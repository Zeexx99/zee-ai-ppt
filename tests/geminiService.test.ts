import { describe, it, expect } from 'vitest';
import { geminiService } from '../src/services/geminiService';
import { DEFAULT_SETTINGS } from '../src/types/settings';

describe('geminiService', () => {
  it('constructs direct Google API endpoint with key', () => {
    const settings = {
      ...DEFAULT_SETTINGS,
      geminiApiKey: 'test-key-abc',
      model: 'gemini-2.5-flash'
    };

    const endpoint = geminiService.getEndpoint('streamGenerateContent', settings);
    expect(endpoint).toBe(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?key=test-key-abc'
    );
  });

  it('constructs Cloudflare Team Proxy endpoint when proxy URL is configured', () => {
    const settings = {
      ...DEFAULT_SETTINGS,
      teamProxyUrl: 'https://my-team-proxy.workers.dev/',
      model: 'gemini-1.5-pro'
    };

    const endpoint = geminiService.getEndpoint('generateContent', settings);
    expect(endpoint).toBe(
      'https://my-team-proxy.workers.dev/v1beta/models/gemini-1.5-pro:generateContent'
    );
  });

  it('populates team proxy authorization headers', () => {
    const settings = {
      ...DEFAULT_SETTINGS,
      teamProxyToken: 'secret-team-token-999'
    };

    const headers = geminiService.getHeaders(settings) as Record<string, string>;
    expect(headers['X-Proxy-Token']).toBe('secret-team-token-999');
    expect(headers['Authorization']).toBe('Bearer secret-team-token-999');
  });

  it('estimates prompt tokens with heuristic', async () => {
    const settings = { ...DEFAULT_SETTINGS, geminiApiKey: '' };
    const prompt = 'Short test prompt for token estimation';
    const count = await geminiService.estimateTokens(prompt, [], settings);
    expect(count).toBeGreaterThan(5);
  });
});
