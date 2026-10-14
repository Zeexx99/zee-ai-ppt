import { describe, expect, it } from "vitest";
import {
  describeToolStep,
  executeAgentToolStep,
  getSupportedAgentTools,
  isReadOnlyTool,
  requiresToolApproval,
  undoCompletedAgentStep,
  validateToolArgs,
} from "../src/services/agentService";
import {
  analyzeSlidesConsistency,
  buildFallbackOutlineFromText,
  detectTextFitIssuesInSlides,
  parseTabularText,
} from "../src/services/pptFeaturesService";
import { listSlides, readSlide } from "../src/services/pptService";
import { DEFAULT_SETTINGS } from "../src/types/settings";
import { PowerPointApiRequirements, SlideDetail } from "../src/types/ppt";

describe("agentService & pptFeaturesService", () => {
  it("distinguishes read-only tools from write tools requiring approval", () => {
    expect(isReadOnlyTool("listSlides")).toBe(true);
    expect(isReadOnlyTool("readSlide")).toBe(true);
    expect(isReadOnlyTool("getSelectedText")).toBe(true);
    expect(isReadOnlyTool("readSpeakerNotes")).toBe(true);
    expect(requiresToolApproval("listSlides")).toBe(false);

    expect(requiresToolApproval("addSlide")).toBe(true);
    expect(requiresToolApproval("deleteSlide")).toBe(true);
    expect(requiresToolApproval("reorderSlide")).toBe(true);
    expect(requiresToolApproval("setSlideTitle")).toBe(true);
    expect(requiresToolApproval("insertTable")).toBe(true);
    expect(requiresToolApproval("writeSpeakerNotes")).toBe(true);
  });

  it("filters out tools when PowerPointApi requirement sets are unsupported", () => {
    const olderHostReqs: PowerPointApiRequirements = {
      isPowerPointHost: true,
      pptApi1_1: true,
      pptApi1_2: false,
      pptApi1_3: true,
      pptApi1_4: true,
      pptApi1_5: true,
      pptApi1_6: false,
      pptApi1_7: false,
      pptApi1_8: false,
      selectedTextCommonApi: true,
      imageCoercionSupported: false,
      tableCoercionSupported: true,
    };

    const tools = getSupportedAgentTools(olderHostReqs);
    const names = tools.map((t) => t.name);
    expect(names).not.toContain("reorderSlide");
    expect(names).not.toContain("deleteSlide");
    expect(names).not.toContain("insertImage");
    expect(names).toContain("addSlide");
    expect(names).toContain("setSlideTitle");
  });

  it("validates tool arguments and rejects invalid parameters before calling Office.js", () => {
    expect(validateToolArgs("readSlide", { slideNumber: 0 }).valid).toBe(false);
    expect(validateToolArgs("readSlide", { slideNumber: 1 }).valid).toBe(true);
    expect(validateToolArgs("addSlide", {}).valid).toBe(false);
    expect(validateToolArgs("addSlide", { title: "ESP Summary" }).valid).toBe(true);
    expect(validateToolArgs("replaceSlideText", { findText: "", replaceWith: "x" }).valid).toBe(
      false
    );
    expect(describeToolStep("setSlideTitle", { slideNumber: 2, title: "WELL-04 DCA" })).toContain(
      "Slide 2"
    );
  });

  it("executes slide tools and supports multi-step undo via snapshots", async () => {
    const initialSlides = await listSlides();
    const initialCount = initialSlides.length;

    const addStepResult = await executeAgentToolStep(
      {
        id: "step-add-1",
        stepNumber: 1,
        toolName: "addSlide",
        args: {
          title: "WELL-09 ESP Diagnosis",
          bodyText: "• PIP: 1,150 psi\n• TDH: 5,200 ft",
          notes: "Discuss WELL-09 intake pressure.",
        },
        description: "Add new slide",
        requiresApproval: true,
        status: "approved",
        retryCount: 0,
      },
      DEFAULT_SETTINGS
    );

    expect(addStepResult.status).toBe("completed");
    expect(addStepResult.undoSnapshot).toBeDefined();

    const afterAdd = await listSlides();
    expect(afterAdd.length).toBe(initialCount + 1);

    // Undo the added slide
    const undone = await undoCompletedAgentStep(addStepResult);
    expect(undone.status).toBe("undone");

    const afterUndo = await listSlides();
    expect(afterUndo.length).toBe(initialCount);
  });

  it("detects consistency issues (fonts, title periods, units, and protected terms)", () => {
    const sampleSlides: SlideDetail[] = [
      {
        index: 1,
        id: "s-1",
        title: "ESP Field Review.",
        fullText: "ESP Field Review.\nWell WELL-01 rate 1200 bbld with esp and tdh check",
        speakerNotes: "",
        shapes: [
          {
            id: "sh-1",
            name: "Title 1",
            type: "Shape",
            text: "ESP Field Review.",
            left: 48,
            top: 32,
            width: 600,
            height: 50,
            fontName: "Calibri Light",
            fontSize: 28,
            isTitle: true,
          },
          {
            id: "sh-2",
            name: "Content 2",
            type: "Shape",
            text: "Well WELL-01 rate 1200 bbld with esp and tdh check",
            left: 48,
            top: 110,
            width: 600,
            height: 250,
            fontName: "Comic Sans MS",
            fontSize: 18,
            isTitle: false,
          },
        ],
      },
    ];

    const issues = analyzeSlidesConsistency(sampleSlides, {
      ...DEFAULT_SETTINGS,
      companyTemplate: {
        ...DEFAULT_SETTINGS.companyTemplate,
        enabled: true,
        bodyFont: "Calibri",
      },
    });

    expect(issues.some((i) => i.category === "title_case")).toBe(true);
    expect(issues.some((i) => i.category === "font_family")).toBe(true);
    expect(issues.some((i) => i.category === "units" && i.suggestedValue === "bbl/d")).toBe(
      true
    );
    expect(
      issues.some((i) => i.category === "protected_term" && i.suggestedValue === "ESP")
    ).toBe(true);
  });

  it("detects text-fit overflow issues and parses Excel tabular data and report outlines", () => {
    const overflowingSlide: SlideDetail = {
      index: 1,
      id: "s-overflow",
      title: "Detailed Report",
      fullText: "",
      speakerNotes: "",
      shapes: [
        {
          id: "sh-body",
          name: "Content Placeholder",
          type: "Shape",
          text: Array.from(
            { length: 12 },
            (_, i) =>
              `• Bullet ${i + 1}: WELL-04 operating at PIP 1,120 psi and TDH 5,400 ft with liquid production 2,450 bbl/d.`
          ).join("\n"),
          left: 48,
          top: 110,
          width: 400,
          height: 150,
          fontSize: 18,
          isTitle: false,
        },
      ],
    };

    const fitIssues = detectTextFitIssuesInSlides([overflowingSlide]);
    expect(fitIssues.length).toBe(1);
    expect(fitIssues[0].estimatedOverflowRatio).toBeGreaterThan(1);

    const table = parseTabularText("Well\tRate (bbl/d)\tPIP (psi)\nWELL-01\t1,850\t1,120");
    expect(table.headers).toEqual(["Well", "Rate (bbl/d)", "PIP (psi)"]);
    expect(table.rows[0]).toEqual(["WELL-01", "1,850", "1,120"]);

    const outline = buildFallbackOutlineFromText(
      "# Q3 Field Review\n\nSection 1: ESP Surveillance\n• WELL-04 PIP stable\n• TDH optimized",
      DEFAULT_SETTINGS
    );
    expect(outline.slides.length).toBeGreaterThanOrEqual(2);
  });
});
