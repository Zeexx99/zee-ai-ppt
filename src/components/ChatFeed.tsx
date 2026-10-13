import React, { useEffect, useRef } from 'react';
import {
  Button,
  Badge,
  Spinner,
  Card,
  Tooltip
} from '@fluentui/react-components';
import {
  ArrowUndo24Regular,
  Play24Regular,
  Copy24Regular,
  Warning24Regular,
  Checkmark24Regular,
  Dismiss24Regular
} from '@fluentui/react-icons';
import { ChatMessage, AgentStep, ToolApproval, UndoRecord } from '../types/chat';
import { ApprovalCard } from './ApprovalCard';
import { arabicService } from '../services/arabicService';

interface ChatFeedProps {
  messages: ChatMessage[];
  onApproveTool: (messageId: string, approval: ToolApproval) => void;
  onRejectTool: (messageId: string, approval: ToolApproval) => void;
  onContinueAgent: () => void;
  isStepLimitReached: boolean;
  onUndoLastAction: () => void;
  canUndo: boolean;
  lastUndoDescription?: string;
  isStreaming: boolean;
}

export const ChatFeed: React.FC<ChatFeedProps> = ({
  messages,
  onApproveTool,
  onRejectTool,
  onContinueAgent,
  isStepLimitReached,
  onUndoLastAction,
  canUndo,
  lastUndoDescription,
  isStreaming
}) => {
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isStreaming, isStepLimitReached]);

  const copyToClipboard = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // ignore
    }
  };

  /**
   * Renders simple markdown (code blocks, bold, bullets, headers) with RTL support
   */
  const renderFormattedContent = (content: string) => {
    if (!content) return null;

    // Check if primarily Arabic
    const isArabic = arabicService.hasArabic(content);

    // Split code blocks
    const codeBlockRegex = /```([a-z]*)\n([\s\S]*?)```/g;
    const parts: React.ReactNode[] = [];
    let lastIndex = 0;
    let match: RegExpExecArray | null;

    while ((match = codeBlockRegex.exec(content)) !== null) {
      if (match.index > lastIndex) {
        const textSegment = content.substring(lastIndex, match.index);
        parts.push(renderTextSegment(textSegment, isArabic, `text-${lastIndex}`));
      }

      const lang = match[1] || 'code';
      const code = match[2];
      const codeIndex = match.index;

      parts.push(
        <div
          key={`code-${codeIndex}`}
          style={{
            position: 'relative',
            background: 'var(--colorNeutralBackground4)',
            borderRadius: '6px',
            padding: '10px',
            margin: '8px 0',
            fontFamily: 'Consolas, monospace',
            fontSize: '12px',
            overflowX: 'auto'
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px', opacity: 0.7 }}>
            <span style={{ fontSize: '10px', textTransform: 'uppercase' }}>{lang}</span>
            <Button
              size="small"
              appearance="subtle"
              icon={<Copy24Regular style={{ fontSize: '12px' }} />}
              onClick={() => copyToClipboard(code)}
              aria-label="Copy code block"
              style={{ minWidth: 'auto', height: '20px', padding: '0 4px' }}
            />
          </div>
          <pre style={{ margin: 0, whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{code}</pre>
        </div>
      );

      lastIndex = match.index + match[0].length;
    }

    if (lastIndex < content.length) {
      parts.push(renderTextSegment(content.substring(lastIndex), isArabic, `text-${lastIndex}`));
    }

    return parts;
  };

  const renderTextSegment = (text: string, isRtl: boolean, key: string) => {
    const lines = text.split('\n');
    return (
      <div
        key={key}
        style={{
          direction: isRtl ? 'rtl' : 'ltr',
          textAlign: isRtl ? 'right' : 'left',
          lineHeight: '1.5'
        }}
      >
        {lines.map((line, idx) => {
          if (!line.trim()) {
            return <div key={idx} style={{ height: '8px' }} />;
          }

          // Headers
          if (line.startsWith('### ')) {
            return (
              <h4 key={idx} style={{ margin: '6px 0 2px 0', fontWeight: 600, fontSize: '13px' }}>
                {line.replace('### ', '')}
              </h4>
            );
          }
          if (line.startsWith('## ')) {
            return (
              <h3 key={idx} style={{ margin: '8px 0 4px 0', fontWeight: 700, fontSize: '14px' }}>
                {line.replace('## ', '')}
              </h3>
            );
          }

          // Bullets
          if (line.trim().startsWith('•') || line.trim().startsWith('- ') || line.trim().startsWith('* ')) {
            const cleanBullet = line.trim().replace(/^[-*•]\s*/, '');
            return (
              <div key={idx} style={{ display: 'flex', gap: '6px', margin: '2px 0' }}>
                <span>•</span>
                <span>{renderInlineStyles(cleanBullet)}</span>
              </div>
            );
          }

          return <div key={idx}>{renderInlineStyles(line)}</div>;
        })}
      </div>
    );
  };

  const renderInlineStyles = (line: string) => {
    // Bold **text**
    const parts = line.split(/(\*\*.*?\*\*)/g);
    return parts.map((part, pIdx) => {
      if (part.startsWith('**') && part.endsWith('**')) {
        return <strong key={pIdx}>{part.slice(2, -2)}</strong>;
      }
      return part;
    });
  };

  return (
    <div
      style={{
        flex: 1,
        overflowY: 'auto',
        padding: '12px',
        display: 'flex',
        flexDirection: 'column',
        gap: '12px'
      }}
    >
      {/* Empty State Banner */}
      {messages.length === 0 && (
        <div
          style={{
            margin: 'auto',
            textAlign: 'center',
            padding: '24px 16px',
            color: 'var(--colorNeutralForeground3)'
          }}
        >
          <div
            style={{
              width: '48px',
              height: '48px',
              margin: '0 auto 12px auto',
              borderRadius: '12px',
              background: 'linear-gradient(135deg, #D24726 0%, #E8633A 100%)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: '#fff',
              fontSize: '24px',
              fontWeight: 800
            }}
          >
            Z
          </div>
          <h3 style={{ margin: '0 0 4px 0', fontSize: '15px', color: 'var(--colorNeutralForeground1)' }}>
            Zee AI for PowerPoint
          </h3>
          <p style={{ margin: 0, fontSize: '12px', maxWidth: '240px' }}>
            Generate presentations, rewrite selected text, create speaker notes, or turn Excel results into slides.
          </p>
        </div>
      )}

      {/* Messages */}
      {messages.map(msg => {
        const isUser = msg.role === 'user';

        return (
          <div
            key={msg.id}
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: isUser ? 'flex-end' : 'flex-start',
              width: '100%'
            }}
          >
            {/* Message Bubble */}
            <div
              style={{
                maxWidth: '88%',
                padding: '10px 12px',
                borderRadius: isUser ? '12px 12px 2px 12px' : '12px 12px 12px 2px',
                background: isUser ? '#D24726' : 'var(--colorNeutralBackground2)',
                color: isUser ? '#ffffff' : 'var(--colorNeutralForeground1)',
                border: isUser ? 'none' : '1px solid rgba(128,128,128,0.15)',
                boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                fontSize: '13px'
              }}
            >
              {/* Attachments */}
              {msg.attachments && msg.attachments.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginBottom: '8px' }}>
                  {msg.attachments.map(att => (
                    <div
                      key={att.id}
                      style={{
                        fontSize: '11px',
                        background: isUser ? 'rgba(255,255,255,0.2)' : 'var(--colorNeutralBackground3)',
                        padding: '4px 6px',
                        borderRadius: '4px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '4px'
                      }}
                    >
                      <span>{att.type === 'pdf' ? '📄' : '🖼️'}</span>
                      <span style={{ maxWidth: '140px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {att.name}
                      </span>
                    </div>
                  ))}
                </div>
              )}

              {/* Text content */}
              {renderFormattedContent(msg.content)}

              {/* Error display */}
              {msg.error && (
                <div
                  style={{
                    color: '#d13438',
                    fontSize: '12px',
                    marginTop: '6px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px'
                  }}
                >
                  <Warning24Regular style={{ fontSize: '14px' }} />
                  <span>{msg.error}</span>
                </div>
              )}
            </div>

            {/* Agent Steps Container */}
            {msg.agentSteps && msg.agentSteps.length > 0 && (
              <div
                style={{
                  width: '95%',
                  marginTop: '8px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '6px'
                }}
              >
                {msg.agentSteps.map(step => (
                  <div
                    key={step.id}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      fontSize: '11px',
                      padding: '6px 8px',
                      borderRadius: '4px',
                      background: 'var(--colorNeutralBackground3)',
                      border: '1px solid rgba(128,128,128,0.1)'
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', overflow: 'hidden' }}>
                      {step.status === 'running' && <Spinner size="extra-tiny" />}
                      {step.status === 'success' && <Checkmark24Regular style={{ color: '#107C41', fontSize: '13px' }} />}
                      {step.status === 'failed' && <Dismiss24Regular style={{ color: '#d13438', fontSize: '13px' }} />}
                      <span style={{ fontWeight: 600 }}>Step {step.stepNumber}:</span>
                      <span style={{ textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }}>
                        {step.description}
                      </span>
                    </div>
                    <Badge
                      size="small"
                      appearance="tint"
                      color={
                        step.status === 'success' ? 'success' : step.status === 'failed' ? 'danger' : 'informative'
                      }
                    >
                      {step.status}
                    </Badge>
                  </div>
                ))}
              </div>
            )}

            {/* Approval Card */}
            {msg.toolApproval && (
              <div style={{ width: '95%', marginTop: '6px' }}>
                <ApprovalCard
                  approval={msg.toolApproval}
                  onApprove={() => onApproveTool(msg.id, msg.toolApproval!)}
                  onReject={() => onRejectTool(msg.id, msg.toolApproval!)}
                />
              </div>
            )}
          </div>
        );
      })}

      {/* Streaming spinner */}
      {isStreaming && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '6px 0', fontSize: '12px', color: 'var(--colorNeutralForeground3)' }}>
          <Spinner size="tiny" />
          <span>Zee AI is drafting response...</span>
        </div>
      )}

      {/* Step Limit Reached Warning & Continue Button */}
      {isStepLimitReached && (
        <Card
          style={{
            padding: '12px',
            border: '1px solid #D24726',
            background: 'var(--colorNeutralBackground2)',
            borderRadius: '6px',
            margin: '8px 0'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
            <Warning24Regular style={{ color: '#D24726', fontSize: '18px' }} />
            <span style={{ fontWeight: 600, fontSize: '13px' }}>Step Limit Reached</span>
          </div>
          <p style={{ margin: '0 0 10px 0', fontSize: '12px', color: 'var(--colorNeutralForeground2)' }}>
            Zee AI completed the current action batch. Would you like to continue to the next steps?
          </p>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
            <Button
              size="small"
              appearance="primary"
              icon={<Play24Regular />}
              onClick={onContinueAgent}
              style={{ background: '#D24726', borderColor: '#D24726' }}
            >
              Continue Next Steps
            </Button>
          </div>
        </Card>
      )}

      {/* Undo Button floating banner */}
      {canUndo && (
        <div
          style={{
            alignSelf: 'center',
            margin: '4px 0',
            display: 'flex',
            alignItems: 'center',
            gap: '8px'
          }}
        >
          <Tooltip content={lastUndoDescription ? `Revert: ${lastUndoDescription}` : 'Undo last slide action'} relationship="label">
            <Button
              size="small"
              appearance="secondary"
              icon={<ArrowUndo24Regular style={{ fontSize: '14px' }} />}
              onClick={onUndoLastAction}
            >
              Undo Last Action
            </Button>
          </Tooltip>
        </div>
      )}

      <div ref={bottomRef} style={{ height: '2px' }} />
    </div>
  );
};
