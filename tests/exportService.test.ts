import { describe, expect, it } from "vitest";
import {
  exportTextToSlideElement,
  formatChatHistoryForExport,
} from "../src/services/exportService";
import { ChatMessage } from "../src/types/chat";
import { DEFAULT_SETTINGS } from "../src/types/settings";

describe("exportService", () => {
  it("formats chat history into structured Markdown with agent steps and citations", () => {
    const messages: ChatMessage[] = [
      {
        id: "m1",
        role: "user",
        content: "Summarize Slide 1 ESP metrics",
        timestamp: 1700000000000,
      },
      {
        id: "m2",
        role: "assistant",
        content: "WELL-04 operates at PIP 1,120 psi and TDH 5,400 ft.",
        timestamp: 1700000005000,
        citations: [{ title: "ESP Handbook", uri: "https://example.com/esp" }],
        agentSteps: [
          {
            id: "s1",
            stepNumber: 1,
            toolName: "readSlide",
            args: { slideNumber: 1 },
            description: "Read Slide 1",
            requiresApproval: false,
            status: "completed",
            resultSummary: "Read 2 shapes",
            retryCount: 0,
          },
        ],
      },
    ];

    const md = formatChatHistoryForExport(messages, "North Field Deck");
    expect(md).toContain("# Zee AI for PowerPoint — Chat Export");
    expect(md).toContain("North Field Deck");
    expect(md).toContain("WELL-04 operates at PIP 1,120 psi");
    expect(md).toContain("ESP Handbook");
    expect(md).toContain("Step 1: Read Slide 1");
  });

  it("exports assistant response text to a new slide or current slide text box", async () => {
    const res1 = await exportTextToSlideElement(
      "ESP Optimization Takeaways\n• Increase frequency by 2 Hz\n• Monitor BHP drawdown",
      "new_slide",
      DEFAULT_SETTINGS
    );
    expect(res1).toContain("Added Slide");

    const res2 = await exportTextToSlideElement(
      "• Key note: GOR stable at 420 scf/stb",
      "current_slide_box",
      DEFAULT_SETTINGS
    );
    expect(res2).toContain("Inserted text box");
  });
});
