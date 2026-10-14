import React, { useRef, useState } from "react";
import { Badge, Button, Textarea, Tooltip } from "@fluentui/react-components";
import {
  Attach20Regular,
  Send20Regular,
  Stop20Regular,
  Dismiss12Regular,
} from "@fluentui/react-icons";
import { ChatAttachment, SavedPrompt } from "../types/chat";
import { ZeeSettings } from "../types/settings";
import { detectTextDirection } from "../services/arabicService";
import { SavedPromptsMenu } from "./SavedPromptsMenu";

interface ChatInputProps {
  settings: ZeeSettings;
  isBusy: boolean;
  savedPrompts: SavedPrompt[];
  onSend: (text: string, attachments: ChatAttachment[]) => void;
  onStop: () => void;
  onUpdateSettings: (partial: Partial<ZeeSettings>) => void;
  onAddCustomPrompt: (title: string, prompt: string) => void;
  onDeleteCustomPrompt: (id: string) => void;
}

export async function fileToChatAttachment(file: File): Promise<ChatAttachment> {
  const isPdf = file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
  const isImage =
    file.type.startsWith("image/") ||
    /\.(png|jpe?g|webp|gif|bmp)$/i.test(file.name);

  if (!isPdf && !isImage) {
    const textContent = await file.text();
    return {
      id: `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      name: file.name,
      mimeType: file.type || "text/plain",
      kind: "text",
      base64Data: "",
      sizeBytes: file.size,
      textContent,
    };
  }

  const base64Data = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const res = String(reader.result || "");
      const idx = res.indexOf(",");
      resolve(idx >= 0 ? res.slice(idx + 1) : res);
    };
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsDataURL(file);
  });

  return {
    id: `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    name: file.name,
    mimeType: isPdf ? "application/pdf" : file.type || "image/png",
    kind: isPdf ? "pdf" : "image",
    base64Data,
    sizeBytes: file.size,
  };
}

export const ChatInput: React.FC<ChatInputProps> = ({
  settings,
  isBusy,
  savedPrompts,
  onSend,
  onStop,
  onUpdateSettings,
  onAddCustomPrompt,
  onDeleteCustomPrompt,
}) => {
  const [input, setInput] = useState("");
  const [attachments, setAttachments] = useState<ChatAttachment[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const handleFiles = async (files: FileList | File[]) => {
    const list = Array.from(files);
    const converted: ChatAttachment[] = [];
    for (const f of list) {
      try {
        converted.push(await fileToChatAttachment(f));
      } catch {
        // ignore unreadable file
      }
    }
    if (converted.length > 0) {
      setAttachments((prev) => [...prev, ...converted]);
    }
  };

  const handlePaste = async (e: React.ClipboardEvent) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    const files: File[] = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.kind === "file") {
        const f = item.getAsFile();
        if (f) files.push(f);
      }
    }
    if (files.length > 0) {
      await handleFiles(files);
    }
  };

  const handleSubmit = () => {
    if (isBusy) return;
    if (!input.trim() && attachments.length === 0) return;
    onSend(input.trim() || "Analyze the attached file(s) for this presentation.", attachments);
    setInput("");
    setAttachments([]);
  };

  const inputDir = detectTextDirection(input, settings.directionMode);

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setIsDragging(true);
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={async (e) => {
        e.preventDefault();
        setIsDragging(false);
        if (e.dataTransfer?.files?.length) {
          await handleFiles(e.dataTransfer.files);
        }
      }}
      style={{
        padding: "8px 10px",
        borderTop: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
        background: isDragging ? "rgba(183, 54, 24, 0.08)" : undefined,
        flexShrink: 0,
      }}
    >
      {/* Attachments preview bar */}
      {attachments.length > 0 && (
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: 4,
            marginBottom: 6,
          }}
        >
          {attachments.map((att) => (
            <Badge
              key={att.id}
              size="small"
              appearance="tint"
              color="brand"
              style={{ display: "inline-flex", alignItems: "center", gap: 4 }}
            >
              <span>
                {att.kind === "pdf" ? "📄" : att.kind === "image" ? "🖼️" : "📝"} {att.name}
              </span>
              <Dismiss12Regular
                style={{ cursor: "pointer" }}
                onClick={() =>
                  setAttachments((prev) => prev.filter((item) => item.id !== att.id))
                }
              />
            </Badge>
          ))}
        </div>
      )}

      {/* Tool Toggles Row */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 4,
          marginBottom: 6,
          flexWrap: "wrap",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap" }}>
          <SavedPromptsMenu
            prompts={savedPrompts}
            onSelectPrompt={(p) => setInput((prev) => (prev ? `${prev}\n${p}` : p))}
            onAddCustomPrompt={onAddCustomPrompt}
            onDeleteCustomPrompt={onDeleteCustomPrompt}
          />

          <Tooltip content="Toggle Google Search grounding" relationship="description">
            <Button
              size="small"
              appearance={settings.webSearchEnabled ? "primary" : "subtle"}
              onClick={() =>
                onUpdateSettings({ webSearchEnabled: !settings.webSearchEnabled })
              }
              style={{
                fontSize: 11,
                padding: "2px 6px",
                backgroundColor: settings.webSearchEnabled ? "#0f6cbd" : undefined,
              }}
            >
              🌐 Web
            </Button>
          </Tooltip>

          <Tooltip content="Toggle Gemini Python Code Execution" relationship="description">
            <Button
              size="small"
              appearance={settings.codeExecutionEnabled ? "primary" : "subtle"}
              onClick={() =>
                onUpdateSettings({ codeExecutionEnabled: !settings.codeExecutionEnabled })
              }
              style={{
                fontSize: 11,
                padding: "2px 6px",
                backgroundColor: settings.codeExecutionEnabled ? "#0f6cbd" : undefined,
              }}
            >
              🐍 Code
            </Button>
          </Tooltip>

          <Tooltip content="Toggle Bilingual Mode (English + Arabic)" relationship="description">
            <Button
              size="small"
              appearance={settings.bilingualMode ? "primary" : "subtle"}
              onClick={() => onUpdateSettings({ bilingualMode: !settings.bilingualMode })}
              style={{
                fontSize: 11,
                padding: "2px 6px",
                backgroundColor: settings.bilingualMode ? "#b73618" : undefined,
              }}
            >
              EN+AR
            </Button>
          </Tooltip>
        </div>

        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept="image/*,application/pdf,.pdf,.txt,.csv,.md"
          style={{ display: "none" }}
          onChange={async (e) => {
            if (e.target.files?.length) {
              await handleFiles(e.target.files);
              e.target.value = "";
            }
          }}
        />
      </div>

      {/* Input box + Paperclip + Send/Stop */}
      <div style={{ display: "flex", alignItems: "flex-end", gap: 6 }}>
        <Tooltip
          content="Attach image, PDF report, or CSV/text (or paste/drag-and-drop)"
          relationship="description"
        >
          <Button
            size="small"
            appearance="subtle"
            icon={<Attach20Regular />}
            onClick={() => fileInputRef.current?.click()}
            aria-label="Attach file"
          />
        </Tooltip>

        <Textarea
          style={{ flex: 1, fontSize: 13 }}
          dir={inputDir}
          rows={2}
          resize="vertical"
          placeholder={
            settings.agentMode
              ? "Ask Agent to read, add, edit, or format slides..."
              : "Ask Zee AI or drop a PDF/image..."
          }
          value={input}
          onPaste={handlePaste}
          onChange={(_, data) => setInput(data.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              handleSubmit();
            }
          }}
        />

        {isBusy ? (
          <Tooltip content="Stop current generation" relationship="label">
            <Button
              size="small"
              appearance="primary"
              icon={<Stop20Regular />}
              onClick={onStop}
              style={{ backgroundColor: "#d13438" }}
              aria-label="Stop"
            />
          </Tooltip>
        ) : (
          <Tooltip content="Send message (Enter)" relationship="label">
            <Button
              size="small"
              appearance="primary"
              icon={<Send20Regular />}
              onClick={handleSubmit}
              disabled={!input.trim() && attachments.length === 0}
              style={{ backgroundColor: "#b73618" }}
              aria-label="Send"
            />
          </Tooltip>
        )}
      </div>
    </div>
  );
};
