import React, { useState } from "react";
import {
  Button,
  Input,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Textarea,
  Badge,
} from "@fluentui/react-components";
import { Bookmark16Regular, Delete16Regular, Add16Regular } from "@fluentui/react-icons";
import { SavedPrompt } from "../types/chat";

interface SavedPromptsMenuProps {
  prompts: SavedPrompt[];
  onSelectPrompt: (promptText: string) => void;
  onAddCustomPrompt: (title: string, promptText: string) => void;
  onDeleteCustomPrompt: (id: string) => void;
}

export const SavedPromptsMenu: React.FC<SavedPromptsMenuProps> = ({
  prompts,
  onSelectPrompt,
  onAddCustomPrompt,
  onDeleteCustomPrompt,
}) => {
  const [open, setOpen] = useState(false);
  const [adding, setAdding] = useState(false);
  const [newTitle, setNewTitle] = useState("");
  const [newPrompt, setNewPrompt] = useState("");

  const handleSaveNew = () => {
    if (!newTitle.trim() || !newPrompt.trim()) return;
    onAddCustomPrompt(newTitle.trim(), newPrompt.trim());
    setNewTitle("");
    setNewPrompt("");
    setAdding(false);
  };

  return (
    <Popover open={open} onOpenChange={(_, data) => setOpen(data.open)} positioning="above-start">
      <PopoverTrigger disableButtonEnhancement>
        <Button
          size="small"
          appearance="subtle"
          icon={<Bookmark16Regular />}
          style={{ fontSize: 11, padding: "2px 6px" }}
        >
          Prompts
        </Button>
      </PopoverTrigger>
      <PopoverSurface style={{ width: 295, maxHeight: 380, overflowY: "auto", padding: 10 }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: 8,
          }}
        >
          <span style={{ fontWeight: 600, fontSize: 12 }}>Saved Prompts</span>
          <Button
            size="small"
            appearance="subtle"
            icon={<Add16Regular />}
            onClick={() => setAdding(!adding)}
            style={{ fontSize: 11 }}
          >
            {adding ? "Cancel" : "New"}
          </Button>
        </div>

        {adding && (
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 6,
              marginBottom: 10,
              padding: 8,
              borderRadius: 6,
              border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
            }}
          >
            <Input
              size="small"
              placeholder="Prompt title (e.g. Monthly Well Review)"
              value={newTitle}
              onChange={(_, d) => setNewTitle(d.value)}
            />
            <Textarea
              size="small"
              rows={3}
              placeholder="Prompt instructions..."
              value={newPrompt}
              onChange={(_, d) => setNewPrompt(d.value)}
            />
            <Button
              size="small"
              appearance="primary"
              onClick={handleSaveNew}
              style={{ backgroundColor: "#b73618", fontSize: 11 }}
            >
              Save Prompt
            </Button>
          </div>
        )}

        <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
          {prompts.map((p) => (
            <div
              key={p.id}
              style={{
                border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                borderRadius: 6,
                padding: "6px 8px",
                cursor: "pointer",
              }}
              onClick={() => {
                onSelectPrompt(p.prompt);
                setOpen(false);
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: 4,
                }}
              >
                <span style={{ fontWeight: 600, fontSize: 12 }}>{p.title}</span>
                <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                  <Badge size="extra-small" appearance="tint">
                    {p.category}
                  </Badge>
                  {!p.isBuiltIn && (
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<Delete16Regular />}
                      onClick={(e) => {
                        e.stopPropagation();
                        onDeleteCustomPrompt(p.id);
                      }}
                      aria-label="Delete prompt"
                    />
                  )}
                </div>
              </div>
              <div
                style={{
                  fontSize: 11,
                  opacity: 0.78,
                  marginTop: 2,
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  display: "-webkit-box",
                  WebkitLineClamp: 2,
                  WebkitBoxOrient: "vertical",
                }}
              >
                {p.prompt}
              </div>
            </div>
          ))}
        </div>
      </PopoverSurface>
    </Popover>
  );
};
