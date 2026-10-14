import React from "react";
import {
  Badge,
  Button,
  Switch,
  Tooltip,
} from "@fluentui/react-components";
import {
  WeatherMoon20Regular,
  WeatherSunny20Regular,
  Settings20Regular,
  Delete20Regular,
  ArrowDownload20Regular,
  Bot20Regular,
} from "@fluentui/react-icons";
import { ZeeSettings } from "../types/settings";
import { PowerPointApiRequirements } from "../types/ppt";

interface HeaderProps {
  settings: ZeeSettings;
  requirements: PowerPointApiRequirements;
  isDark: boolean;
  onToggleTheme: () => void;
  onToggleAgentMode: (enabled: boolean) => void;
  onOpenSettings: () => void;
  onExportChat: () => void;
  onClearChat: () => void;
  canUndoLast: boolean;
  onUndoLast: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  settings,
  requirements,
  isDark,
  onToggleTheme,
  onToggleAgentMode,
  onOpenSettings,
  onExportChat,
  onClearChat,
  canUndoLast,
  onUndoLast,
}) => {
  return (
    <header
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "8px 10px",
        borderBottom: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
        background: isDark ? "#1f1f1f" : "#faf9f8",
        gap: "6px",
        flexShrink: 0,
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: "8px", minWidth: 0 }}>
        {/* Relative icon path so it works on both localhost and GitHub Pages */}
        <img
          src="assets/icon-32.png"
          alt="Zee AI"
          width={24}
          height={24}
          style={{ borderRadius: 5, flexShrink: 0 }}
        />
        <div style={{ minWidth: 0 }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "6px",
              fontWeight: 600,
              fontSize: "13px",
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis",
            }}
          >
            <span>Zee AI</span>
            <Badge
              size="small"
              appearance="tint"
              color={requirements.isPowerPointHost ? "brand" : "warning"}
            >
              {requirements.isPowerPointHost ? "PPT" : "Preview"}
            </Badge>
            {settings.bilingualMode && (
              <Badge size="small" appearance="filled" color="informative">
                EN+AR
              </Badge>
            )}
          </div>
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: "2px", flexShrink: 0 }}>
        <Tooltip
          content="Agent Mode: AI reads and edits slides using approval cards and undo"
          relationship="description"
        >
          <div style={{ display: "flex", alignItems: "center", marginRight: 2 }}>
            <Bot20Regular
              style={{
                color: settings.agentMode ? "#b73618" : undefined,
                marginRight: -4,
              }}
            />
            <Switch
              checked={settings.agentMode}
              onChange={(_, data) => onToggleAgentMode(data.checked)}
              label={<span className="zee-hide-narrow" style={{ fontSize: 11 }}>Agent</span>}
            />
          </div>
        </Tooltip>

        {canUndoLast && (
          <Tooltip content="Undo last slide action" relationship="label">
            <Button
              size="small"
              appearance="subtle"
              onClick={onUndoLast}
              style={{ fontSize: 11, padding: "2px 6px" }}
            >
              ↶ Undo
            </Button>
          </Tooltip>
        )}

        <Tooltip content={isDark ? "Switch to Light Mode" : "Switch to Dark Mode"} relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={isDark ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />}
            onClick={onToggleTheme}
            aria-label="Toggle dark mode"
          />
        </Tooltip>

        <Tooltip content="Export chat transcript (.md)" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<ArrowDownload20Regular />}
            onClick={onExportChat}
            aria-label="Export chat"
          />
        </Tooltip>

        <Tooltip content="Clear chat history for this presentation" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Delete20Regular />}
            onClick={onClearChat}
            aria-label="Clear chat"
          />
        </Tooltip>

        <Tooltip content="Settings (API Key, Arabic/RTL, Protected Terms, Template)" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Settings20Regular />}
            onClick={onOpenSettings}
            aria-label="Open Settings"
          />
        </Tooltip>
      </div>
    </header>
  );
};
