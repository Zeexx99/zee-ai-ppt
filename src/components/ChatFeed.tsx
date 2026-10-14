import React, { useEffect, useRef } from "react";
import { Badge, Button, Spinner, Tooltip } from "@fluentui/react-components";
import {
  Copy16Regular,
  SlideAdd16Regular,
  TextBoxSettingsRegular,
} from "@fluentui/react-icons";
import { AgentToolStep, ChatMessage } from "../types/chat";
import { ZeeSettings } from "../types/settings";
import { detectTextDirection, resolveArabicFontFamily } from "../services/arabicService";
import { ApprovalCard } from "./ApprovalCard";

interface ChatFeedProps {
  messages: ChatMessage[];
  settings: ZeeSettings;
  isDark: boolean;
  onCopyMessage: (text: string) => void;
  onInsertIntoSlide: (text: string) => void;
  onCreateSlideFromMessage: (text: string) => void;
  onApproveStep: (stepId: string) => void;
  onRejectStep: (stepId: string) => void;
  onUndoStep: (step: AgentToolStep) => void;
  onRetryStep: (step: AgentToolStep) => void;
  onContinueAgent: () => void;
}

export const ChatFeed: React.FC<ChatFeedProps> = ({
  messages,
  settings,
  isDark,
  onCopyMessage,
  onInsertIntoSlide,
  onCreateSlideFromMessage,
  onApproveStep,
  onRejectStep,
  onUndoStep,
  onRetryStep,
  onContinueAgent,
}) => {
  const bottomRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  if (messages.length === 0) {
    return (
      <div
        className="zee-scroll-area"
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
          textAlign: "center",
          padding: 20,
          opacity: 0.85,
        }}
      >
        <img
          src="assets/icon-64.png"
          alt="Zee AI"
          width={44}
          height={44}
          style={{ borderRadius: 10, marginBottom: 10 }}
        />
        <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 4 }}>
          Zee AI for PowerPoint
        </div>
        <div style={{ fontSize: 12, maxWidth: 280, lineHeight: 1.5, opacity: 0.85 }}>
          Chat with Gemini, use <b>Agent Mode</b> to inspect &amp; build slides with approval cards,
          or launch <b>Report → Deck</b>, <b>Excel → Slide</b>, <b>Consistency</b>, and{" "}
          <b>Text-Fit</b> tools above.
        </div>
      </div>
    );
  }

  return (
    <div className="zee-scroll-area">
      {messages.map((msg) => {
        const dir = msg.direction || detectTextDirection(msg.content, settings.directionMode);
        const fontFamily = resolveArabicFontFamily(
          msg.content,
          settings.arabicFont,
          undefined
        );
        const isUser = msg.role === "user";
        const isSystem = msg.role === "system";

        return (
          <div
            key={msg.id}
            className="zee-message-bubble"
            dir={dir}
            style={{
              fontFamily,
              background: isUser
                ? isDark
                  ? "#3b231d"
                  : "#fdf3f0"
                : isSystem
                  ? isDark
                    ? "#292929"
                    : "#f3f2f1"
                  : isDark
                    ? "#252525"
                    : "#f8f8f8",
              border: isUser
                ? "1px solid rgba(183, 54, 24, 0.35)"
                : "1px solid var(--colorNeutralStroke2, #e1dfdd)",
              marginLeft: isUser ? 18 : 0,
              marginRight: isUser ? 0 : 8,
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                fontSize: 11,
                fontWeight: 600,
                marginBottom: 4,
                opacity: 0.75,
              }}
              dir="ltr"
            >
              <span>
                {isUser ? "You" : isSystem ? "System" : "Zee AI"}
                {dir === "rtl" ? " • RTL" : ""}
              </span>
              <span>
                {new Date(msg.timestamp).toLocaleTimeString([], {
                  hour: "2-digit",
                  minute: "2-digit",
                })}
              </span>
            </div>

            {msg.attachments && msg.attachments.length > 0 && (
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: 4,
                  marginBottom: 6,
                }}
                dir="ltr"
              >
                {msg.attachments.map((att) => (
                  <Badge key={att.id} size="small" appearance="outline" color="informative">
                    📎 {att.name}
                  </Badge>
                ))}
              </div>
            )}

            <div>
              {msg.content}
              {msg.isStreaming && (
                <span style={{ display: "inline-flex", alignItems: "center", marginLeft: 6 }}>
                  <Spinner size="extra-tiny" />
                </span>
              )}
            </div>

            {msg.codeExecutionOutput && (
              <div
                dir="ltr"
                style={{
                  marginTop: 8,
                  padding: 8,
                  borderRadius: 6,
                  background: isDark ? "#141414" : "#eee",
                  fontFamily: "Consolas, monospace",
                  fontSize: 11,
                  overflowX: "auto",
                }}
              >
                {msg.codeExecutionOutput}
              </div>
            )}

            {msg.citations && msg.citations.length > 0 && (
              <div
                dir="ltr"
                style={{
                  marginTop: 8,
                  paddingTop: 6,
                  borderTop: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                  fontSize: 11,
                }}
              >
                <div style={{ fontWeight: 600, marginBottom: 3 }}>🌐 Web Sources:</div>
                {msg.citations.map((c, idx) => (
                  <div key={idx} style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
                    •{" "}
                    <a
                      href={c.uri}
                      target="_blank"
                      rel="noreferrer"
                      style={{ color: "#0f6cbd" }}
                    >
                      {c.title}
                    </a>
                  </div>
                ))}
              </div>
            )}

            {msg.agentSteps && msg.agentSteps.length > 0 && (
              <div dir="ltr">
                <ApprovalCard
                  steps={msg.agentSteps}
                  awaitingContinue={msg.awaitingContinue}
                  onApproveStep={onApproveStep}
                  onRejectStep={onRejectStep}
                  onUndoStep={onUndoStep}
                  onRetryStep={onRetryStep}
                  onContinueAgent={onContinueAgent}
                />
              </div>
            )}

            {msg.role === "assistant" && !msg.isStreaming && msg.content.trim() && (
              <div
                dir="ltr"
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: 4,
                  marginTop: 8,
                  paddingTop: 6,
                  borderTop: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                }}
              >
                <Tooltip content="Copy response text" relationship="label">
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<Copy16Regular />}
                    onClick={() => onCopyMessage(msg.content)}
                    style={{ fontSize: 11, padding: "2px 6px" }}
                  >
                    Copy
                  </Button>
                </Tooltip>
                <Tooltip content="Insert into active slide" relationship="label">
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<TextBoxSettingsRegular />}
                    onClick={() => onInsertIntoSlide(msg.content)}
                    style={{ fontSize: 11, padding: "2px 6px" }}
                  >
                    To Slide
                  </Button>
                </Tooltip>
                <Tooltip content="Create a new slide from this response" relationship="label">
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<SlideAdd16Regular />}
                    onClick={() => onCreateSlideFromMessage(msg.content)}
                    style={{ fontSize: 11, padding: "2px 6px" }}
                  >
                    New Slide
                  </Button>
                </Tooltip>
              </div>
            )}
          </div>
        );
      })}
      <div ref={bottomRef} />
    </div>
  );
};
