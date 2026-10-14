import React, { useEffect, useRef, useState } from "react";
import {
  Button,
  FluentProvider,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { Header } from "./components/Header";
import { QuickActions } from "./components/QuickActions";
import { ChatFeed } from "./components/ChatFeed";
import { ChatInput } from "./components/ChatInput";
import { SettingsDrawer } from "./components/SettingsDrawer";
import { DeckToolMode, PowerPointDeckModal } from "./components/PowerPointDeckModal";
import { ResultsSlideModal } from "./components/ResultsSlideModal";
import {
  AgentToolStep,
  ChatAttachment,
  ChatMessage,
  QuickActionType,
  SavedPrompt,
  SlideUndoSnapshot,
} from "./types/chat";
import { PowerPointApiRequirements } from "./types/ppt";
import { ZeeSettings } from "./types/settings";
import {
  clearChatHistory,
  loadChatHistory,
  loadSavedPrompts,
  loadSettings,
  saveChatHistory,
  saveCustomPrompts,
  saveSettings,
} from "./services/storageService";
import {
  detectPowerPointRequirements,
  formatCompactPresentationSummary,
  getActiveSlideIndex,
  getDocumentIdentifier,
  getPresentationSummary,
  getSelectedText,
  listAvailableLayouts,
  readSlide,
  restoreUndoSnapshot,
  startOfficeReadinessMonitor,
} from "./services/pptService";
import {
  appendConvertedUnitsToText,
  buildDirectionAndLanguageInstruction,
  detectTextDirection,
} from "./services/arabicService";
import { isAuthConfigured } from "./services/authService";
import {
  evaluateCostGuard,
  streamGeminiChat,
} from "./services/geminiService";
import {
  AgentLoopSession,
  executeAgentToolStep,
  runAgentLoop,
  undoCompletedAgentStep,
} from "./services/agentService";
import {
  copyTextToClipboard,
  downloadChatTranscript,
  exportTextToSlideElement,
} from "./services/exportService";
import { generateSpeakerNotesForSlide } from "./services/pptFeaturesService";

export const App: React.FC = () => {
  const [settings, setSettings] = useState<ZeeSettings>(() => loadSettings());
  const [requirements, setRequirements] = useState<PowerPointApiRequirements>(() =>
    detectPowerPointRequirements()
  );
  const [docId, setDocId] = useState<string>(() => getDocumentIdentifier());
  const [messages, setMessages] = useState<ChatMessage[]>(() =>
    loadChatHistory(getDocumentIdentifier())
  );
  const [savedPrompts, setSavedPrompts] = useState<SavedPrompt[]>(() =>
    loadSavedPrompts()
  );
  const [availableLayouts, setAvailableLayouts] = useState<string[]>([
    "Title Slide",
    "Title and Content",
    "Section Header",
    "Two Content",
  ]);

  const [settingsOpen, setSettingsOpen] = useState(false);
  const [deckModalMode, setDeckModalMode] = useState<DeckToolMode>(null);
  const [excelModalOpen, setExcelModalOpen] = useState(false);
  const [isBusy, setIsBusy] = useState(false);
  const [statusBanner, setStatusBanner] = useState<string>("");
  const [undoStack, setUndoStack] = useState<SlideUndoSnapshot[]>([]);

  // Cost guard confirmation state
  const [pendingCostGuard, setPendingCostGuard] = useState<{
    text: string;
    attachments: ChatAttachment[];
    estimatedTokens: number;
    limit: number;
  } | null>(null);

  // Agent session & approval resolvers
  const activeAgentSessionRef = useRef<{
    messageId: string;
    session: AgentLoopSession;
    userPrompt: string;
    attachments: ChatAttachment[];
  } | null>(null);

  const approvalResolversRef = useRef<Map<string, (approved: boolean) => void>>(
    new Map()
  );
  const abortControllerRef = useRef<AbortController | null>(null);

  // Repeated Office readiness checks (fixes stuck "Preview mode" banner bug)
  useEffect(() => {
    const stopMonitor = startOfficeReadinessMonitor(async (reqs) => {
      setRequirements(reqs);
      const currentDoc = getDocumentIdentifier();
      setDocId(currentDoc);
      const layouts = await listAvailableLayouts();
      if (layouts.length > 0) setAvailableLayouts(layouts);
    });
    return stopMonitor;
  }, []);

  // Persist settings
  const handleUpdateSettings = (partial: Partial<ZeeSettings>) => {
    setSettings((prev) => {
      const updated: ZeeSettings = {
        ...prev,
        ...partial,
        companyTemplate: partial.companyTemplate
          ? { ...prev.companyTemplate, ...partial.companyTemplate }
          : prev.companyTemplate,
      };
      saveSettings(updated);
      return updated;
    });
  };

  // Persist chat history with automatic oldest-message trimming
  useEffect(() => {
    saveChatHistory(messages, docId);
  }, [messages, docId]);

  const systemPrefersDark =
    typeof window !== "undefined" &&
    window.matchMedia &&
    window.matchMedia("(prefers-color-scheme: dark)").matches;

  const isDark =
    settings.appearance === "dark" ||
    (settings.appearance === "auto" && systemPrefersDark);

  const showTransientBanner = (msg: string) => {
    setStatusBanner(msg);
    window.setTimeout(() => {
      setStatusBanner((curr) => (curr === msg ? "" : curr));
    }, 4000);
  };

  const pushUndoSnapshots = (snaps: SlideUndoSnapshot[]) => {
    if (!snaps || snaps.length === 0) return;
    setUndoStack((prev) => [...prev, ...snaps]);
  };

  const handleUndoLast = async () => {
    if (undoStack.length === 0) return;
    const last = undoStack[undoStack.length - 1];
    try {
      const msg = await restoreUndoSnapshot(last);
      setUndoStack((prev) => prev.slice(0, -1));
      showTransientBanner(`↶ ${msg}`);
    } catch (err) {
      showTransientBanner(
        `Undo error: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  };

  const executeSendPrompt = async (
    text: string,
    attachments: ChatAttachment[] = [],
    bypassCostGuard: boolean = false
  ) => {
    if (!isAuthConfigured(settings)) {
      setSettingsOpen(true);
      showTransientBanner("Please enter your Gemini API Key or Proxy settings first.");
      return;
    }

    if (!bypassCostGuard) {
      const guard = evaluateCostGuard(settings, text, messages, attachments);
      if (guard.exceedsLimit) {
        setPendingCostGuard({
          text,
          attachments,
          estimatedTokens: guard.estimatedTokens,
          limit: guard.limit,
        });
        return;
      }
    }

    setPendingCostGuard(null);
    const userMsg: ChatMessage = {
      id: `user-${Date.now()}`,
      role: "user",
      content: text,
      timestamp: Date.now(),
      attachments: attachments.length > 0 ? attachments : undefined,
      direction: detectTextDirection(text, settings.directionMode),
    };

    const assistantMsgId = `assistant-${Date.now() + 1}`;
    const placeholderAssistant: ChatMessage = {
      id: assistantMsgId,
      role: "assistant",
      content: settings.agentMode ? "Inspecting presentation and planning slide actions..." : "",
      timestamp: Date.now(),
      isStreaming: true,
      agentSteps: settings.agentMode ? [] : undefined,
    };

    const historySnapshot = [...messages];
    setMessages((prev) => [...prev, userMsg, placeholderAssistant]);
    setIsBusy(true);

    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    try {
      if (settings.agentMode) {
        const session = await runAgentLoop({
          settings,
          history: historySnapshot,
          userPrompt: text,
          attachments,
          requirements,
          signal: abortController.signal,
          onUpdate: (updatedSession) => {
            activeAgentSessionRef.current = {
              messageId: assistantMsgId,
              session: updatedSession,
              userPrompt: text,
              attachments,
            };
            const completedSnaps = updatedSession.steps
              .filter((s) => s.status === "completed" && s.undoSnapshot)
              .map((s) => s.undoSnapshot!);
            if (completedSnaps.length > 0) {
              setUndoStack((prev) => {
                const existingIds = new Set(prev.map((p) => p.stepId));
                const fresh = completedSnaps.filter((s) => !existingIds.has(s.stepId));
                return fresh.length > 0 ? [...prev, ...fresh] : prev;
              });
            }

            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantMsgId
                  ? {
                      ...m,
                      content:
                        updatedSession.finalText ||
                        `Executed ${updatedSession.steps.length} slide step(s).`,
                      agentSteps: [...updatedSession.steps],
                      awaitingContinue: updatedSession.awaitingContinue,
                    }
                  : m
              )
            );
          },
          requestApproval: (step) =>
            new Promise<boolean>((resolve) => {
              approvalResolversRef.current.set(step.id, resolve);
            }),
        });

        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantMsgId
              ? {
                  ...m,
                  isStreaming: false,
                  content:
                    session.finalText ||
                    `Completed ${session.steps.filter((s) => s.status === "completed").length} slide action(s).`,
                  agentSteps: session.steps,
                  awaitingContinue: session.awaitingContinue,
                  direction: detectTextDirection(
                    session.finalText || "",
                    settings.directionMode
                  ),
                }
              : m
          )
        );
      } else {
        const summary = await getPresentationSummary();
        const compactContext = formatCompactPresentationSummary(summary);

        const finalRes = await streamGeminiChat(settings, historySnapshot, text, {
          attachments,
          presentationContextSummary: compactContext,
          signal: abortController.signal,
          onChunk: (accumulatedText, citations, codeOutput) => {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantMsgId
                  ? {
                      ...m,
                      content: accumulatedText,
                      citations,
                      codeExecutionOutput: codeOutput,
                      direction: detectTextDirection(
                        accumulatedText,
                        settings.directionMode
                      ),
                    }
                  : m
              )
            );
          },
        });

        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantMsgId
              ? {
                  ...m,
                  content: finalRes.text || "(No text returned)",
                  citations: finalRes.citations,
                  codeExecutionOutput: finalRes.codeExecutionOutput,
                  isStreaming: false,
                  direction: detectTextDirection(finalRes.text, settings.directionMode),
                }
              : m
          )
        );
      }
    } catch (err) {
      const errMessage = err instanceof Error ? err.message : String(err);
      setMessages((prev) =>
        prev.map((m) =>
          m.id === assistantMsgId
            ? {
                ...m,
                isStreaming: false,
                content: m.content
                  ? `${m.content}\n\n⚠️ ${errMessage}`
                  : `⚠️ ${errMessage}`,
              }
            : m
        )
      );
    } finally {
      setIsBusy(false);
      abortControllerRef.current = null;
    }
  };

  const handleApproveStep = (stepId: string) => {
    const resolver = approvalResolversRef.current.get(stepId);
    if (resolver) {
      approvalResolversRef.current.delete(stepId);
      resolver(true);
    }
  };

  const handleRejectStep = (stepId: string) => {
    const resolver = approvalResolversRef.current.get(stepId);
    if (resolver) {
      approvalResolversRef.current.delete(stepId);
      resolver(false);
    }
  };

  const handleUndoStep = async (step: AgentToolStep) => {
    try {
      const updated = await undoCompletedAgentStep(step);
      setMessages((prev) =>
        prev.map((m) => ({
          ...m,
          agentSteps: m.agentSteps?.map((s) => (s.id === step.id ? updated : s)),
        }))
      );
      showTransientBanner(`↶ Undid step #${step.stepNumber} (${step.toolName})`);
    } catch (err) {
      showTransientBanner(
        `Undo failed: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  };

  const handleRetryStep = async (step: AgentToolStep) => {
    const retried = await executeAgentToolStep(step, settings);
    setMessages((prev) =>
      prev.map((m) => ({
        ...m,
        agentSteps: m.agentSteps?.map((s) => (s.id === step.id ? retried : s)),
      }))
    );
  };

  const handleContinueAgent = async () => {
    const active = activeAgentSessionRef.current;
    if (!active || isBusy) return;
    setIsBusy(true);
    try {
      const session = await runAgentLoop({
        settings,
        history: messages,
        userPrompt: active.userPrompt,
        attachments: active.attachments,
        requirements,
        existingSession: active.session,
        maxAdditionalSteps: settings.agentMaxSteps || 8,
        onUpdate: (updated) => {
          activeAgentSessionRef.current = { ...active, session: updated };
          setMessages((prev) =>
            prev.map((m) =>
              m.id === active.messageId
                ? {
                    ...m,
                    content:
                      updated.finalText ||
                      `Executed ${updated.steps.length} slide step(s).`,
                    agentSteps: [...updated.steps],
                    awaitingContinue: updated.awaitingContinue,
                  }
                : m
            )
          );
        },
        requestApproval: (step) =>
          new Promise<boolean>((resolve) => {
            approvalResolversRef.current.set(step.id, resolve);
          }),
      });
      activeAgentSessionRef.current = { ...active, session };
    } finally {
      setIsBusy(false);
    }
  };

  const handleRunQuickAction = async (action: QuickActionType) => {
    const selectedOrSlideText =
      (await getSelectedText()) ||
      (await readSlide(await getActiveSlideIndex())).fullText;

    if (action === "convert_units") {
      const { convertedText, conversions } =
        appendConvertedUnitsToText(selectedOrSlideText);
      if (conversions.length > 0) {
        const summaryLines = conversions.map((c) => `• ${c.formatted}`).join("\n");
        setMessages((prev) => [
          ...prev,
          {
            id: `sys-units-${Date.now()}`,
            role: "assistant",
            content: `🛢️ **Petroleum Unit Conversions Found (${conversions.length}):**\n${summaryLines}\n\n**Converted Text:**\n${convertedText}`,
            timestamp: Date.now(),
          },
        ]);
        return;
      }
      await executeSendPrompt(
        `Convert all petroleum engineering units (bbl/d, stb/d, psi, ft, °F, scf/stb) in the following slide text into their metric/field equivalents:\n\n${selectedOrSlideText}`
      );
      return;
    }

    if (action === "speaker_notes") {
      if (!isAuthConfigured(settings)) {
        setSettingsOpen(true);
        return;
      }
      setIsBusy(true);
      try {
        const activeSlide = await getActiveSlideIndex();
        const { notes, undoSnapshot } = await generateSpeakerNotesForSlide(
          activeSlide,
          settings,
          true
        );
        if (undoSnapshot) pushUndoSnapshots([undoSnapshot]);
        setMessages((prev) => [
          ...prev,
          {
            id: `notes-${Date.now()}`,
            role: "assistant",
            content: `🎙️ **Speaker Notes Written to Slide #${activeSlide}:**\n\n${notes}`,
            timestamp: Date.now(),
            direction: detectTextDirection(notes, settings.directionMode),
          },
        ]);
      } catch (err) {
        showTransientBanner(
          `Speaker notes error: ${err instanceof Error ? err.message : String(err)}`
        );
      } finally {
        setIsBusy(false);
      }
      return;
    }

    const langRules = buildDirectionAndLanguageInstruction(settings);
    const promptsByAction: Record<QuickActionType, string> = {
      rewrite: `Rewrite the following slide text to be crisp, executive-ready, and high-impact for PowerPoint:\n${langRules}\n\nTEXT:\n${selectedOrSlideText}`,
      translate_ar: `Translate the following slide text into professional Arabic (Right-to-Left), keeping all protected engineering terms (ESP, TDH, PIP, GOR, BHP, bbl/d, well names, and numbers with units) unchanged in Latin script:\n\nTEXT:\n${selectedOrSlideText}`,
      translate_en: `Translate the following slide text into crisp professional English, keeping all protected engineering terms (ESP, TDH, PIP, GOR, BHP, bbl/d, well names, and units) unchanged:\n\nTEXT:\n${selectedOrSlideText}`,
      translate_bilingual: `Create a Bilingual (English + Arabic) version of this slide content. Put each English bullet/heading followed by its Arabic counterpart, keeping all protected terms (ESP, TDH, PIP, GOR, BHP, bbl/d, well names, units) unchanged:\n\nTEXT:\n${selectedOrSlideText}`,
      summarize: `Summarize the following slide content into 4 concise bullet points with key quantitative takeaways:\n${langRules}\n\nTEXT:\n${selectedOrSlideText}`,
      shorten_fit: `Shorten these bullets so they fit cleanly inside a single PowerPoint placeholder (max 12 words per bullet) while keeping all numbers and protected terms:\n\nTEXT:\n${selectedOrSlideText}`,
      speaker_notes: "",
      convert_units: "",
    };

    await executeSendPrompt(promptsByAction[action]);
  };

  return (
    <FluentProvider
      theme={isDark ? webDarkTheme : webLightTheme}
      className="zee-app-root"
    >
      <Header
        settings={settings}
        requirements={requirements}
        isDark={isDark}
        onToggleTheme={() =>
          handleUpdateSettings({ appearance: isDark ? "light" : "dark" })
        }
        onToggleAgentMode={(enabled) => handleUpdateSettings({ agentMode: enabled })}
        onOpenSettings={() => setSettingsOpen(true)}
        onExportChat={() => {
          downloadChatTranscript(messages);
          showTransientBanner("Exported chat transcript (.md)");
        }}
        onClearChat={() => {
          clearChatHistory(docId);
          setMessages([]);
          showTransientBanner("Cleared chat history for this presentation.");
        }}
        canUndoLast={undoStack.length > 0}
        onUndoLast={handleUndoLast}
      />

      <QuickActions
        requirements={requirements}
        settings={settings}
        disabled={isBusy}
        onRunQuickAction={handleRunQuickAction}
        onOpenReportToDeck={() => setDeckModalMode("report_to_deck")}
        onOpenExcelToSlide={() => setExcelModalOpen(true)}
        onOpenConsistencyChecker={() => setDeckModalMode("consistency")}
        onOpenTextFitFixer={() => setDeckModalMode("text_fit")}
        onOpenDeckReviewer={() => setDeckModalMode("deck_review")}
        onOpenBeautifier={() => setDeckModalMode("beautify")}
        onOpenDiagrams={() => setDeckModalMode("diagrams")}
        onOpenAgendaBuilder={() => setDeckModalMode("agenda_actions")}
      />

      {statusBanner && (
        <div
          style={{
            padding: "6px 10px",
            background: "rgba(183, 54, 24, 0.12)",
            borderBottom: "1px solid rgba(183, 54, 24, 0.3)",
            fontSize: 11,
            fontWeight: 600,
            flexShrink: 0,
          }}
        >
          {statusBanner}
        </div>
      )}

      {/* Cost Guard Warning Banner */}
      {pendingCostGuard && (
        <div
          style={{
            padding: "8px 10px",
            background: "rgba(216, 59, 1, 0.14)",
            borderBottom: "1px solid #d83b01",
            fontSize: 12,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 8,
            flexShrink: 0,
          }}
        >
          <span>
            ⚠️ <b>Cost Guard:</b> Estimated prompt size is ~
            {pendingCostGuard.estimatedTokens.toLocaleString()} tokens (limit:{" "}
            {pendingCostGuard.limit.toLocaleString()}). Send anyway?
          </span>
          <div style={{ display: "flex", gap: 4 }}>
            <Button
              size="small"
              appearance="primary"
              onClick={() =>
                executeSendPrompt(
                  pendingCostGuard.text,
                  pendingCostGuard.attachments,
                  true
                )
              }
              style={{ backgroundColor: "#b73618", fontSize: 11 }}
            >
              Send
            </Button>
            <Button
              size="small"
              appearance="outline"
              onClick={() => setPendingCostGuard(null)}
              style={{ fontSize: 11 }}
            >
              Cancel
            </Button>
          </div>
        </div>
      )}

      <ChatFeed
        messages={messages}
        settings={settings}
        isDark={isDark}
        onCopyMessage={async (text) => {
          await copyTextToClipboard(text);
          showTransientBanner("Copied to clipboard!");
        }}
        onInsertIntoSlide={async (text) => {
          const msg = await exportTextToSlideElement(
            text,
            "current_slide_box",
            settings
          );
          showTransientBanner(msg);
        }}
        onCreateSlideFromMessage={async (text) => {
          const msg = await exportTextToSlideElement(text, "new_slide", settings);
          showTransientBanner(msg);
        }}
        onApproveStep={handleApproveStep}
        onRejectStep={handleRejectStep}
        onUndoStep={handleUndoStep}
        onRetryStep={handleRetryStep}
        onContinueAgent={handleContinueAgent}
      />

      <ChatInput
        settings={settings}
        isBusy={isBusy}
        savedPrompts={savedPrompts}
        onSend={(text, attachments) => executeSendPrompt(text, attachments)}
        onStop={() => abortControllerRef.current?.abort()}
        onUpdateSettings={handleUpdateSettings}
        onAddCustomPrompt={(title, prompt) => {
          const next: SavedPrompt[] = [
            ...savedPrompts,
            {
              id: `custom-${Date.now()}`,
              title,
              prompt,
              category: "custom",
            },
          ];
          setSavedPrompts(next);
          saveCustomPrompts(next);
        }}
        onDeleteCustomPrompt={(id) => {
          const next = savedPrompts.filter((p) => p.id !== id);
          setSavedPrompts(next);
          saveCustomPrompts(next);
        }}
      />

      {/* Eagerly loaded Settings Drawer & PowerPoint Modals */}
      <SettingsDrawer
        open={settingsOpen}
        settings={settings}
        requirements={requirements}
        availableLayouts={availableLayouts}
        onClose={() => setSettingsOpen(false)}
        onUpdateSettings={handleUpdateSettings}
      />

      <PowerPointDeckModal
        mode={deckModalMode}
        settings={settings}
        onClose={() => setDeckModalMode(null)}
        onCompleted={(summaryMessage, snaps) => {
          pushUndoSnapshots(snaps);
          setMessages((prev) => [
            ...prev,
            {
              id: `deck-action-${Date.now()}`,
              role: "assistant",
              content: `✅ ${summaryMessage}`,
              timestamp: Date.now(),
            },
          ]);
        }}
      />

      <ResultsSlideModal
        open={excelModalOpen}
        settings={settings}
        onClose={() => setExcelModalOpen(false)}
        onCompleted={(summaryMessage, snaps) => {
          pushUndoSnapshots(snaps);
          setMessages((prev) => [
            ...prev,
            {
              id: `excel-action-${Date.now()}`,
              role: "assistant",
              content: `📊 ${summaryMessage}`,
              timestamp: Date.now(),
            },
          ]);
        }}
      />
    </FluentProvider>
  );
};
