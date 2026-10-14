import { describe, expect, it } from "vitest";
import {
  buildGeminiEndpoint,
  buildRelativeAuthUrl,
  isAuthConfigured,
} from "../src/services/authService";
import {
  buildSystemInstruction,
  buildToolsPayload,
  estimatePromptTokens,
  evaluateCostGuard,
  extractResultFromCandidate,
} from "../src/services/geminiService";
import { DEFAULT_SETTINGS } from "../src/types/settings";

describe("authService & geminiService", () => {
  it("builds relative auth URLs that work on both localhost and GitHub Pages subpaths", () => {
    const localUrl = buildRelativeAuthUrl(
      "auth/start.html",
      { client_id: "abc" },
      "https://localhost:3300/index.html"
    );
    expect(localUrl).toBe("https://localhost:3300/auth/start.html?client_id=abc");

    const pagesUrl = buildRelativeAuthUrl(
      "auth/start.html",
      { client_id: "abc" },
      "https://zeexx99.github.io/zee-ai-ppt/index.html"
    );
    expect(pagesUrl).toBe(
      "https://zeexx99.github.io/zee-ai-ppt/auth/start.html?client_id=abc"
    );
  });

  it("constructs Gemini endpoints for API key, OAuth, and Cloudflare Team Proxy", () => {
    const apiKeySettings = {
      ...DEFAULT_SETTINGS,
      authMode: "apiKey" as const,
      geminiApiKey: "AIzaTestKey",
    };
    expect(isAuthConfigured(apiKeySettings)).toBe(true);
    const ep1 = buildGeminiEndpoint(apiKeySettings, "streamGenerateContent", true);
    expect(ep1.url).toContain(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?alt=sse"
    );
    expect(ep1.headers["x-goog-api-key"]).toBe("AIzaTestKey");

    const proxySettings = {
      ...DEFAULT_SETTINGS,
      authMode: "proxy" as const,
      proxyUrl: "https://zee-proxy.workers.dev/",
      proxyToken: "team-secret-token",
    };
    expect(isAuthConfigured(proxySettings)).toBe(true);
    const ep2 = buildGeminiEndpoint(proxySettings, "generateContent", false);
    expect(ep2.url).toBe(
      "https://zee-proxy.workers.dev/v1beta/models/gemini-2.5-flash:generateContent"
    );
    expect(ep2.headers["Authorization"]).toBe("Bearer team-secret-token");
    expect(ep2.headers["X-Zee-Token"]).toBe("team-secret-token");
  });

  it("estimates tokens and triggers Cost Guard when threshold is exceeded", () => {
    const smallTokens = estimatePromptTokens("Short prompt", [], []);
    expect(smallTokens).toBeGreaterThan(0);
    expect(smallTokens).toBeLessThan(100);

    const guard = evaluateCostGuard(
      { ...DEFAULT_SETTINGS, costGuardEnabled: true, costGuardTokenLimit: 1000 },
      "A".repeat(6000),
      []
    );
    expect(guard.exceedsLimit).toBe(true);
  });

  it("includes protected terms and Company Template rules in systemInstruction", () => {
    const instruction = buildSystemInstruction(
      {
        ...DEFAULT_SETTINGS,
        bilingualMode: true,
        companyTemplate: {
          ...DEFAULT_SETTINGS.companyTemplate,
          enabled: true,
          layoutName: "Two Content",
        },
      },
      "Slide Count: 3"
    );
    expect(instruction).toContain("BILINGUAL MODE IS ENABLED");
    expect(instruction).toContain("ESP");
    expect(instruction).toContain("TDH");
    expect(instruction).toContain("COMPANY TEMPLATE MODE IS ACTIVE");
    expect(instruction).toContain("Slide Count: 3");
  });

  it("builds tools payload for web search and code execution and extracts candidates", () => {
    const tools = buildToolsPayload({
      ...DEFAULT_SETTINGS,
      webSearchEnabled: true,
      codeExecutionEnabled: true,
    });
    expect(tools).toEqual([{ google_search: {} }, { code_execution: {} }]);

    const parsed = extractResultFromCandidate({
      content: {
        parts: [
          { text: "Slide summary ready." },
          {
            functionCall: {
              name: "setSlideTitle",
              args: { slideNumber: 1, title: "Updated Title" },
            },
          },
        ],
      },
      groundingMetadata: {
        groundingChunks: [{ web: { title: "SPE Paper", uri: "https://onepetro.org" } }],
      },
    });
    expect(parsed.text).toBe("Slide summary ready.");
    expect(parsed.functionCalls.length).toBe(1);
    expect(parsed.functionCalls[0].name).toBe("setSlideTitle");
    expect(parsed.citations[0].uri).toBe("https://onepetro.org");
  });
});
