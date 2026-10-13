import React, { useRef, useState } from 'react';
import {
  Button,
  Textarea,
  Tooltip,
  Tag,
  Badge
} from '@fluentui/react-components';
import {
  Send24Regular,
  Attach24Regular,
  Dismiss16Regular,
  Globe24Regular,
  Code24Regular,
  Stop24Regular
} from '@fluentui/react-icons';
import { Attachment, SavedPrompt } from '../types/chat';
import { SavedPromptsMenu } from './SavedPromptsMenu';

interface ChatInputProps {
  onSendMessage: (text: string, attachments: Attachment[]) => void;
  isLoading: boolean;
  onStop: () => void;
  webSearchEnabled: boolean;
  onToggleWebSearch: () => void;
  codeExecutionEnabled: boolean;
  onToggleCodeExecution: () => void;
  savedPrompts: SavedPrompt[];
  onAddSavedPrompt: (p: SavedPrompt) => void;
}

export const ChatInput: React.FC<ChatInputProps> = ({
  onSendMessage,
  isLoading,
  onStop,
  webSearchEnabled,
  onToggleWebSearch,
  codeExecutionEnabled,
  onToggleCodeExecution,
  savedPrompts,
  onAddSavedPrompt
}) => {
  const [text, setText] = useState('');
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleSend = () => {
    if ((text.trim() || attachments.length > 0) && !isLoading) {
      onSendMessage(text.trim(), attachments);
      setText('');
      setAttachments([]);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const processFile = async (file: File) => {
    if (!file.type.startsWith('image/') && file.type !== 'application/pdf') {
      alert('Only images and PDF files are supported.');
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      const base64Data = dataUrl.split(',')[1];
      const newAtt: Attachment = {
        id: `att-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
        name: file.name,
        type: file.type === 'application/pdf' ? 'pdf' : 'image',
        mimeType: file.type,
        base64Data,
        size: file.size
      };
      setAttachments(prev => [...prev, newAtt]);
    };
    reader.readAsDataURL(file);
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files) {
      for (let i = 0; i < files.length; i++) {
        processFile(files[i]);
      }
    }
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handlePaste = (e: React.ClipboardEvent) => {
    const items = e.clipboardData?.items;
    if (items) {
      for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
          const file = items[i].getAsFile();
          if (file) {
            processFile(file);
          }
        }
      }
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer?.files;
    if (files) {
      for (let i = 0; i < files.length; i++) {
        processFile(files[i]);
      }
    }
  };

  const removeAttachment = (id: string) => {
    setAttachments(prev => prev.filter(a => a.id !== id));
  };

  return (
    <div
      style={{
        padding: '8px 12px 12px 12px',
        borderTop: '1px solid rgba(128,128,128,0.15)',
        background: 'var(--colorNeutralBackground1)',
        position: 'relative'
      }}
      onDragOver={e => {
        e.preventDefault();
        setIsDragging(true);
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
    >
      {/* Drag overlay */}
      {isDragging && (
        <div
          style={{
            position: 'absolute',
            inset: 0,
            background: 'rgba(210, 71, 38, 0.1)',
            border: '2px dashed #D24726',
            borderRadius: '6px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 600,
            color: '#D24726',
            zIndex: 10
          }}
        >
          Drop image or PDF to attach
        </div>
      )}

      {/* Attachments preview list */}
      {attachments.length > 0 && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', marginBottom: '8px' }}>
          {attachments.map(att => (
            <Tag
              key={att.id}
              size="small"
              appearance="brand"
              dismissible
              dismissIcon={<Dismiss16Regular />}
              onClick={() => removeAttachment(att.id)}
            >
              {att.type === 'pdf' ? '📄 ' : '🖼️ '}
              {att.name}
            </Tag>
          ))}
        </div>
      )}

      {/* Input row */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
        <Textarea
          placeholder="Ask Zee AI to create slides, rewrite text, or analyze data..."
          value={text}
          onChange={(_, d) => setText(d.value)}
          onKeyDown={handleKeyDown}
          onPaste={handlePaste}
          rows={2}
          style={{
            width: '100%',
            resize: 'none',
            fontFamily: 'inherit',
            fontSize: '13px'
          }}
        />

        {/* Toolbar & Send Button */}
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
            <input
              type="file"
              ref={fileInputRef}
              style={{ display: 'none' }}
              accept="image/*,application/pdf"
              multiple
              onChange={handleFileInputChange}
            />

            <Tooltip content="Attach Image or PDF" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<Attach24Regular style={{ fontSize: '16px' }} />}
                onClick={() => fileInputRef.current?.click()}
                aria-label="Attach file"
              />
            </Tooltip>

            <SavedPromptsMenu
              prompts={savedPrompts}
              onSelectPrompt={p => setText(p)}
              onAddPrompt={onAddSavedPrompt}
            />

            <Tooltip content={webSearchEnabled ? 'Web Search: ON' : 'Web Search: OFF'} relationship="label">
              <Button
                appearance={webSearchEnabled ? 'primary' : 'subtle'}
                size="small"
                icon={<Globe24Regular style={{ fontSize: '15px' }} />}
                onClick={onToggleWebSearch}
                aria-label="Toggle web search"
                style={{ padding: '0 6px', minWidth: 'auto', height: '24px' }}
              />
            </Tooltip>

            <Tooltip content={codeExecutionEnabled ? 'Code Execution: ON' : 'Code Execution: OFF'} relationship="label">
              <Button
                appearance={codeExecutionEnabled ? 'primary' : 'subtle'}
                size="small"
                icon={<Code24Regular style={{ fontSize: '15px' }} />}
                onClick={onToggleCodeExecution}
                aria-label="Toggle code execution"
                style={{ padding: '0 6px', minWidth: 'auto', height: '24px' }}
              />
            </Tooltip>
          </div>

          <div>
            {isLoading ? (
              <Button
                size="small"
                appearance="secondary"
                icon={<Stop24Regular style={{ fontSize: '16px', color: '#D24726' }} />}
                onClick={onStop}
              >
                Stop
              </Button>
            ) : (
              <Button
                size="small"
                appearance="primary"
                icon={<Send24Regular style={{ fontSize: '16px' }} />}
                disabled={!text.trim() && attachments.length === 0}
                onClick={handleSend}
                style={{ background: '#D24726', borderColor: '#D24726' }}
              >
                Send
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
