import React, { useState, useEffect, useRef } from 'react';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  Theme
} from '@fluentui/react-components';
import { Header } from './components/Header';
import { ChatFeed } from './components/ChatFeed';
import { ChatInput } from './components/ChatInput';
import { SettingsDrawer } from './components/SettingsDrawer';
import { QuickActions } from './components/QuickActions';
import { PowerPointDeckModal } from './components/PowerPointDeckModal';
import { ResultsSlideModal } from './components/ResultsSlideModal';

import { AppSettings, DEFAULT_SETTINGS, DarkModeSetting } from './types/settings';
import { ChatMessage, Attachment, AgentStep, ToolApproval, UndoRecord, SavedPrompt } from './types/chat';
import { RequirementSetStatus } from './types/ppt';

import { storageService } from './services/storageService';
import { pptService } from './services/pptService';
import { geminiService, ToolCall } from './services/geminiService';
import { agentService, AGENT_TOOLS } from './services/agentService';
import { pptFeaturesService } from './services/pptFeaturesService';
import { exportService } from './services/exportService';

interface AppProps {
  isOfficeReady: boolean;
}

export const App: React.FC<AppProps> = ({ isOfficeReady }) => {
  // Eagerly loaded state
  const [settings, setSettings] = useState<AppSettings>(() => storageService.getSettings());
  const [messages, setMessages] = useState<ChatMessage[]>(() => storageService.getChatHistory());
  const [savedPrompts, setSavedPrompts] = useState<SavedPrompt[]>(() => storageService.getSavedPrompts());
  const [requirementStatus, setRequirementStatus] = useState<RequirementSetStatus>(() =>
    pptService.getRequirementStatus()
  );

  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isDeckModalOpen, setIsDeckModalOpen] = useState(false);
  const [isResultsModalOpen, setIsResultsModalOpen] = useState(false);

  const [isStreaming, setIsStreaming] = useState(false);
  const [isStepLimitReached, setIsStepLimitReached] = useState(false);
  const [undoStack, setUndoStack] = useState<UndoRecord[]>([]);

  const abortControllerRef = useRef<AbortController | null>(null);
  const currentStepCountRef = useRef<number>(0);

  // Sync settings and chat history to storage
  useEffect(() => {
    storageService.saveSettings(settings);
  }, [settings]);

  useEffect(() => {
    storageService.saveChatHistory(messages);
  }, [messages]);

  useEffect(() => {
    setRequirementStatus(pptService.getRequirementStatus());
  }, [isOfficeReady]);

  // Dark mode resolution
  const [resolvedTheme, setResolvedTheme] = useState<Theme>(webLightTheme);

  useEffect(() => {
    if (settings.darkMode === 'dark') {
      setResolvedTheme(webDarkTheme);
    } else if (settings.darkMode === 'light') {
      setResolvedTheme(webLightTheme);
    } else {
      // Auto mode
      const mq = window.matchMedia('(prefers-color-scheme: dark)');
      setResolvedTheme(mq.matches ? webDarkTheme : webLightTheme);
      const listener = (e: MediaQueryListEvent) => {
        setResolvedTheme(e.matches ? webDarkTheme : webLightTheme);
      };
      mq.addEventListener('change', listener);
      return () => mq.removeEventListener('change', listener);
    }
  }, [settings.darkMode]);

  const toggleDarkMode = () => {
    const next: DarkModeSetting = settings.darkMode === 'dark' ? 'light' : 'dark';
    setSettings(prev => ({ ...prev, darkMode: next }));
  };

  const handleStopStreaming = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      abortControllerRef.current = null;
    }
    setIsStreaming(false);
  };

  /**
   * Main Agent Execution Loop with approval cards, step limits, and streaming
   */
  const runAgentTurn = async (conversation: ChatMessage[], newPrompt?: string) => {
    setIsStreaming(true);
    setIsStepLimitReached(false);
    abortControllerRef.current = new AbortController();

    const assistantMsgId = `msg-assistant-${Date.now()}`;
    const initialAssistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      agentSteps: []
    };

    setMessages(prev => [...prev, initialAssistantMsg]);

    try {
      const overview = await pptService.getPresentationOverview();
      const systemInstruction = agentService.buildSystemPrompt(overview, settings);

      const streamMessages = conversation.map(m => ({
        role: m.role,
        content: m.content,
        attachments: m.attachments
      }));

      if (newPrompt) {
        streamMessages.push({ role: 'user', content: newPrompt, attachments: [] });
      }

      let accumulatedText = '';
      const steps: AgentStep[] = [];
      let pendingApproval: ToolApproval | null = null;

      for await (const chunk of geminiService.streamChat({
        messages: streamMessages,
        systemInstruction,
        tools: AGENT_TOOLS,
        settings,
        signal: abortControllerRef.current.signal
      })) {
        if (chunk.textDelta) {
          accumulatedText += chunk.textDelta;
          setMessages(prev =>
            prev.map(m => (m.id === assistantMsgId ? { ...m, content: accumulatedText } : m))
          );
        }

        if (chunk.toolCalls && chunk.toolCalls.length > 0) {
          for (const tc of chunk.toolCalls) {
            currentStepCountRef.current++;
            const stepNum = currentStepCountRef.current;
            const isRead = agentService.isReadOnlyTool(tc.name);
            const description = agentService.formatApprovalDescription(tc.name, tc.args);

            const stepId = `step-${Date.now()}-${stepNum}`;
            const newStep: AgentStep = {
              id: stepId,
              stepNumber: stepNum,
              toolName: tc.name,
              description,
              status: isRead ? 'running' : 'approving',
              params: tc.args,
              timestamp: Date.now()
            };

            steps.push(newStep);

            if (isRead) {
              // Execute read-only tool immediately
              setMessages(prev =>
                prev.map(m => (m.id === assistantMsgId ? { ...m, agentSteps: [...steps] } : m))
              );

              try {
                const res = await agentService.executeTool(tc.name, tc.args, settings);
                newStep.status = 'success';
                newStep.result = res.result;
              } catch (err) {
                newStep.status = 'failed';
                newStep.error = err instanceof Error ? err.message : String(err);
              }

              setMessages(prev =>
                prev.map(m => (m.id === assistantMsgId ? { ...m, agentSteps: [...steps] } : m))
              );
            } else {
              // Write tool requires Approval Card!
              pendingApproval = {
                stepId,
                toolName: tc.name,
                description,
                params: tc.args,
                isPending: true
              };

              setMessages(prev =>
                prev.map(m =>
                  m.id === assistantMsgId
                    ? { ...m, agentSteps: [...steps], toolApproval: pendingApproval! }
                    : m
                )
              );

              // Pause for user approval
              setIsStreaming(false);
              return;
            }

            // Step limit check
            if (currentStepCountRef.current >= settings.maxAgentSteps) {
              setIsStepLimitReached(true);
              setIsStreaming(false);
              return;
            }
          }
        }
      }
    } catch (err: unknown) {
      if (err instanceof Error && err.name === 'AbortError') {
        // user aborted
      } else {
        const errorMsg = err instanceof Error ? err.message : 'An error occurred during AI processing.';
        setMessages(prev =>
          prev.map(m => (m.id === assistantMsgId ? { ...m, error: errorMsg } : m))
        );
      }
    } finally {
      setIsStreaming(false);
    }
  };

  /**
   * Sends user message
   */
  const handleSendMessage = async (text: string, attachments: Attachment[]) => {
    // Cost guard token estimate
    if (settings.costGuardEnabled) {
      const tokens = await geminiService.estimateTokens(text, attachments, settings);
      if (tokens > settings.costGuardThreshold) {
        const proceed = window.confirm(
          `Cost Guard Warning:\nThis prompt is estimated at ~${tokens.toLocaleString()} tokens (threshold: ${settings.costGuardThreshold.toLocaleString()}).\nProceed anyway?`
        );
        if (!proceed) return;
      }
    }

    currentStepCountRef.current = 0;
    const userMsg: ChatMessage = {
      id: `msg-user-${Date.now()}`,
      role: 'user',
      content: text,
      timestamp: Date.now(),
      attachments
    };

    const nextConversation = [...messages, userMsg];
    setMessages(nextConversation);
    await runAgentTurn(nextConversation);
  };

  /**
   * Handles tool approval by user
   */
  const handleApproveTool = async (messageId: string, approval: ToolApproval) => {
    // Mark approval
    setMessages(prev =>
      prev.map(m =>
        m.id === messageId
          ? {
              ...m,
              toolApproval: { ...approval, isPending: false, approved: true },
              agentSteps: m.agentSteps?.map(s =>
                s.id === approval.stepId ? { ...s, status: 'running' } : s
              )
            }
          : m
      )
    );

    try {
      const res = await agentService.executeTool(approval.toolName, approval.params, settings);

      if (res.undoRecord) {
        setUndoStack(prev => [...prev, res.undoRecord!]);
      }

      setMessages(prev =>
        prev.map(m =>
          m.id === messageId
            ? {
                ...m,
                agentSteps: m.agentSteps?.map(s =>
                  s.id === approval.stepId
                    ? { ...s, status: 'success', result: res.result }
                    : s
                )
              }
            : m
        )
      );

      // Continue conversation with result
      await runAgentTurn(messages, `Tool ${approval.toolName} succeeded: ${res.result}`);
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      setMessages(prev =>
        prev.map(m =>
          m.id === messageId
            ? {
                ...m,
                agentSteps: m.agentSteps?.map(s =>
                  s.id === approval.stepId ? { ...s, status: 'failed', error: errorMsg } : s
                )
              }
            : m
        )
      );
    }
  };

  const handleRejectTool = (messageId: string, approval: ToolApproval) => {
    setMessages(prev =>
      prev.map(m =>
        m.id === messageId
          ? {
              ...m,
              toolApproval: { ...approval, isPending: false, approved: false },
              agentSteps: m.agentSteps?.map(s =>
                s.id === approval.stepId
                  ? { ...s, status: 'failed', error: 'Action rejected by user.' }
                  : s
              )
            }
          : m
      )
    );
  };

  const handleContinueAgent = () => {
    currentStepCountRef.current = 0;
    setIsStepLimitReached(false);
    runAgentTurn(messages, 'Please continue with the remaining steps.');
  };

  const handleUndoLastAction = async () => {
    if (undoStack.length === 0) return;
    const record = undoStack[undoStack.length - 1];
    try {
      const resultMsg = await agentService.undoAction(record);
      setUndoStack(prev => prev.slice(0, -1));
      setMessages(prev => [
        ...prev,
        {
          id: `msg-undo-${Date.now()}`,
          role: 'assistant',
          content: `🔄 **Undo successful:** ${resultMsg}`,
          timestamp: Date.now()
        }
      ]);
    } catch (err) {
      alert(`Undo failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  /**
   * Quick Actions dispatch
   */
  const handleQuickAction = async (actionKey: string) => {
    switch (actionKey) {
      case 'rewrite': {
        const selected = await pptService.getSelectedText();
        if (!selected) {
          alert('Please select text inside a slide textbox to rewrite.');
          return;
        }
        await handleSendMessage(
          `Please rewrite and polish the following selected text for clarity and impact, preserving all technical petroleum terms and units:\n"""\n${selected}\n"""`,
          []
        );
        break;
      }

      case 'translate_ar': {
        const selected = await pptService.getSelectedText();
        if (!selected) {
          alert('Please select text inside a slide textbox to translate to Arabic.');
          return;
        }
        await handleSendMessage(
          `Translate the following text into professional Arabic with RTL direction. Keep all engineering acronyms, well names, and units (ESP, TDH, PIP, GOR, BHP, bbl/d) untouched in English:\n"""\n${selected}\n"""`,
          []
        );
        break;
      }

      case 'bilingual': {
        const selected = await pptService.getSelectedText();
        if (!selected) {
          alert('Please select text to format as bilingual.');
          return;
        }
        await handleSendMessage(
          `Format the following text in Bilingual mode (English LTR followed by Arabic RTL), keeping technical terms intact:\n"""\n${selected}\n"""`,
          []
        );
        break;
      }

      case 'summarize': {
        const overview = await pptService.getPresentationOverview();
        const activeSlide = await pptService.readSlide(overview.activeSlideIndex || 1);
        await handleSendMessage(
          `Summarize the key takeaways of Slide ${overview.activeSlideIndex || 1} ("${activeSlide.title}") into 3 high-impact bullet points:\n${activeSlide.allText}`,
          []
        );
        break;
      }

      case 'speaker_notes': {
        const overview = await pptService.getPresentationOverview();
        const idx = overview.activeSlideIndex || 1;
        setIsStreaming(true);
        try {
          const notes = await pptFeaturesService.generateSpeakerNotesForSlide(idx, settings);
          setMessages(prev => [
            ...prev,
            {
              id: `msg-notes-${Date.now()}`,
              role: 'assistant',
              content: `🎙️ **Speaker notes generated and added to Slide ${idx}:**\n\n${notes}`,
              timestamp: Date.now()
            }
          ]);
        } catch (err) {
          alert(err instanceof Error ? err.message : 'Failed to generate speaker notes.');
        } finally {
          setIsStreaming(false);
        }
        break;
      }

      case 'text_fit': {
        const overview = await pptService.getPresentationOverview();
        const idx = overview.activeSlideIndex || 1;
        setIsStreaming(true);
        try {
          const res = await pptFeaturesService.shortenSlideText(idx, settings);
          setMessages(prev => [
            ...prev,
            {
              id: `msg-fit-${Date.now()}`,
              role: 'assistant',
              content: `✂️ **Shortened Text for Slide ${idx}:**\n\n${res.shortenedText}\n\n*Would you like me to replace the bullets on Slide ${idx}?*`,
              timestamp: Date.now()
            }
          ]);
        } catch (err) {
          alert(err instanceof Error ? err.message : 'Failed to fit text.');
        } finally {
          setIsStreaming(false);
        }
        break;
      }

      case 'report_to_deck':
        setIsDeckModalOpen(true);
        break;

      case 'excel_to_slide':
        setIsResultsModalOpen(true);
        break;

      case 'consistency_audit': {
        setIsStreaming(true);
        try {
          const issues = await pptFeaturesService.checkDeckConsistency(settings);
          if (issues.length === 0) {
            setMessages(prev => [
              ...prev,
              {
                id: `msg-audit-${Date.now()}`,
                role: 'assistant',
                content: '✅ **Consistency Audit Complete:** No issues found! Font styles, title capitalization, and protected terms are aligned across all slides.',
                timestamp: Date.now()
              }
            ]);
          } else {
            const report = issues
              .map(iss => `- **Slide ${iss.slideIndex} (${iss.type}):** ${iss.description}`)
              .join('\n');
            setMessages(prev => [
              ...prev,
              {
                id: `msg-audit-${Date.now()}`,
                role: 'assistant',
                content: `🔍 **Consistency Audit Findings (${issues.length} item(s)):**\n\n${report}\n\n*Ask me to fix any of these slides!*`,
                timestamp: Date.now()
              }
            ]);
          }
        } catch (err) {
          alert(err instanceof Error ? err.message : 'Audit failed.');
        } finally {
          setIsStreaming(false);
        }
        break;
      }
    }
  };

  const handleExportChat = () => {
    exportService.downloadAsFile(messages);
  };

  const handleClearChat = () => {
    if (window.confirm('Clear current chat history?')) {
      setMessages([]);
      storageService.clearChatHistory();
    }
  };

  const handleAddSavedPrompt = (prompt: SavedPrompt) => {
    const updated = [...savedPrompts, prompt];
    setSavedPrompts(updated);
    storageService.saveSavedPrompts(updated);
  };

  return (
    <FluentProvider theme={resolvedTheme} className="zee-app-root">
      {/* Header */}
      <Header
        darkMode={settings.darkMode}
        onToggleDarkMode={toggleDarkMode}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onExportChat={handleExportChat}
        onClearChat={handleClearChat}
        hasMessages={messages.length > 0}
        isOfficeReady={isOfficeReady}
      />

      {/* Quick Actions Bar */}
      <QuickActions
        onAction={handleQuickAction}
        isLoading={isStreaming}
        requirementStatus={requirementStatus}
      />

      {/* Chat Messages */}
      <ChatFeed
        messages={messages}
        onApproveTool={handleApproveTool}
        onRejectTool={handleRejectTool}
        onContinueAgent={handleContinueAgent}
        isStepLimitReached={isStepLimitReached}
        onUndoLastAction={handleUndoLastAction}
        canUndo={undoStack.length > 0}
        lastUndoDescription={undoStack[undoStack.length - 1]?.description}
        isStreaming={isStreaming}
      />

      {/* Chat Input */}
      <ChatInput
        onSendMessage={handleSendMessage}
        isLoading={isStreaming}
        onStop={handleStopStreaming}
        webSearchEnabled={settings.webSearch}
        onToggleWebSearch={() => setSettings(p => ({ ...p, webSearch: !p.webSearch }))}
        codeExecutionEnabled={settings.codeExecution}
        onToggleCodeExecution={() => setSettings(p => ({ ...p, codeExecution: !p.codeExecution }))}
        savedPrompts={savedPrompts}
        onAddSavedPrompt={handleAddSavedPrompt}
      />

      {/* Eagerly Loaded Settings Drawer */}
      <SettingsDrawer
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        settings={settings}
        onSave={setSettings}
        requirementStatus={requirementStatus}
      />

      {/* Report to Deck Modal */}
      <PowerPointDeckModal
        isOpen={isDeckModalOpen}
        onClose={() => setIsDeckModalOpen(false)}
        settings={settings}
        onDeckBuilt={count => {
          setMessages(prev => [
            ...prev,
            {
              id: `msg-deck-${Date.now()}`,
              role: 'assistant',
              content: `🎉 **Presentation Created:** Successfully built **${count} slides** with speaker notes into your presentation!`,
              timestamp: Date.now()
            }
          ]);
        }}
      />

      {/* Results Slide Modal */}
      <ResultsSlideModal
        isOpen={isResultsModalOpen}
        onClose={() => setIsResultsModalOpen(false)}
        settings={settings}
        onSlideCreated={slideIdx => {
          setMessages(prev => [
            ...prev,
            {
              id: `msg-results-${Date.now()}`,
              role: 'assistant',
              content: `📊 **Results Slide Created:** Added executive results summary at Slide ${slideIdx}.`,
              timestamp: Date.now()
            }
          ]);
        }}
      />
    </FluentProvider>
  );
};
