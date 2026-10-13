import React from 'react';
import {
  Button,
  Toolbar,
  ToolbarButton,
  Badge,
  Tooltip
} from '@fluentui/react-components';
import {
  Settings24Regular,
  WeatherMoon24Regular,
  WeatherSunny24Regular,
  ArrowDownload24Regular,
  Delete24Regular,
  Sparkle24Regular
} from '@fluentui/react-icons';
import { DarkModeSetting } from '../types/settings';

interface HeaderProps {
  darkMode: DarkModeSetting;
  onToggleDarkMode: () => void;
  onOpenSettings: () => void;
  onExportChat: () => void;
  onClearChat: () => void;
  hasMessages: boolean;
  isOfficeReady: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  darkMode,
  onToggleDarkMode,
  onOpenSettings,
  onExportChat,
  onClearChat,
  hasMessages,
  isOfficeReady
}) => {
  const versionString = 'v1.0.0 (Oct 2026)';

  return (
    <header
      style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '8px 12px',
        borderBottom: '1px solid rgba(128, 128, 128, 0.2)',
        background: 'var(--colorNeutralBackground1)',
        minHeight: '44px',
        userSelect: 'none'
      }}
    >
      {/* Brand Title and Version */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
        <div
          style={{
            width: '24px',
            height: '24px',
            borderRadius: '6px',
            background: 'linear-gradient(135deg, #D24726 0%, #E8633A 100%)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            color: '#fff',
            fontWeight: 800,
            fontSize: '13px',
            boxShadow: '0 2px 4px rgba(210, 71, 38, 0.3)'
          }}
        >
          Z
        </div>
        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <span style={{ fontWeight: 600, fontSize: '13px', color: 'var(--colorNeutralForeground1)' }}>
              Zee AI
            </span>
            <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)' }}>
              PowerPoint
            </span>
            {isOfficeReady ? (
              <Badge size="small" appearance="filled" color="success">
                PPT
              </Badge>
            ) : (
              <Badge size="small" appearance="outline" color="warning">
                Preview
              </Badge>
            )}
          </div>
          <span style={{ fontSize: '9px', color: 'var(--colorNeutralForeground4)', lineHeight: 1 }}>
            {versionString}
          </span>
        </div>
      </div>

      {/* Actions Toolbar */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
        {hasMessages && (
          <>
            <Tooltip content="Export Chat" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<ArrowDownload24Regular style={{ fontSize: '16px' }} />}
                onClick={onExportChat}
                aria-label="Export chat transcript"
              />
            </Tooltip>
            <Tooltip content="Clear Chat" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<Delete24Regular style={{ fontSize: '16px' }} />}
                onClick={onClearChat}
                aria-label="Clear chat history"
              />
            </Tooltip>
          </>
        )}

        <Tooltip content={darkMode === 'dark' ? 'Switch to Light' : 'Switch to Dark'} relationship="label">
          <Button
            appearance="subtle"
            size="small"
            icon={
              darkMode === 'dark' ? (
                <WeatherSunny24Regular style={{ fontSize: '16px', color: '#F8D247' }} />
              ) : (
                <WeatherMoon24Regular style={{ fontSize: '16px' }} />
              )
            }
            onClick={onToggleDarkMode}
            aria-label="Toggle dark mode"
          />
        </Tooltip>

        <Tooltip content="Settings" relationship="label">
          <Button
            appearance="subtle"
            size="small"
            icon={<Settings24Regular style={{ fontSize: '17px' }} />}
            onClick={onOpenSettings}
            aria-label="Open settings drawer"
          />
        </Tooltip>
      </div>
    </header>
  );
};
