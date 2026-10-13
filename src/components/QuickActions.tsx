import React from 'react';
import { Button, Tooltip, Badge } from '@fluentui/react-components';
import {
  Sparkle24Regular,
  Translate24Regular,
  DocumentBulletList24Regular,
  Speaker224Regular,
  ArrowAutofitHeight24Regular,
  DocumentTable24Regular,
  CheckmarkStarburst24Regular,
  News24Regular
} from '@fluentui/react-icons';
import { RequirementSetStatus } from '../types/ppt';

interface QuickActionsProps {
  onAction: (actionKey: string) => void;
  isLoading: boolean;
  requirementStatus: RequirementSetStatus;
}

export const QuickActions: React.FC<QuickActionsProps> = ({
  onAction,
  isLoading,
  requirementStatus
}) => {
  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: '6px',
        overflowX: 'auto',
        padding: '6px 8px',
        borderBottom: '1px solid rgba(128,128,128,0.15)',
        background: 'var(--colorNeutralBackground2)',
        scrollbarWidth: 'none',
        msOverflowStyle: 'none'
      }}
    >
      <Tooltip content="Polish and improve selected text" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<Sparkle24Regular style={{ color: '#D24726', fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('rewrite')}
        >
          Rewrite
        </Button>
      </Tooltip>

      <Tooltip content="Translate selected text to Arabic (RTL)" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<Translate24Regular style={{ fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('translate_ar')}
        >
          عربي (AR)
        </Button>
      </Tooltip>

      <Tooltip content="Format as bilingual English + Arabic" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          disabled={isLoading}
          onClick={() => onAction('bilingual')}
        >
          EN/AR
        </Button>
      </Tooltip>

      <Tooltip content="Summarize slide content into 3 key points" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<DocumentBulletList24Regular style={{ fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('summarize')}
        >
          Summarize
        </Button>
      </Tooltip>

      {/* Speaker Notes Generator: only visible if notes API is supported */}
      {requirementStatus.canManageNotes ? (
        <Tooltip content="Generate natural presenter speaker notes" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Speaker224Regular style={{ fontSize: '15px' }} />}
            disabled={isLoading}
            onClick={() => onAction('speaker_notes')}
          >
            Notes
          </Button>
        </Tooltip>
      ) : null}

      <Tooltip content="Condense bullets that overflow the slide placeholder" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<ArrowAutofitHeight24Regular style={{ fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('text_fit')}
        >
          Text-Fit
        </Button>
      </Tooltip>

      <Tooltip content="Generate presentation outline from report or PDF" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<News24Regular style={{ fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('report_to_deck')}
        >
          Report→Deck
        </Button>
      </Tooltip>

      <Tooltip content="Format Excel or petroleum test data into a Results slide" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<DocumentTable24Regular style={{ fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('excel_to_slide')}
        >
          Excel→Slide
        </Button>
      </Tooltip>

      <Tooltip content="Audit presentation for font, casing, and protected terms consistency" relationship="label">
        <Button
          size="small"
          appearance="subtle"
          icon={<CheckmarkStarburst24Regular style={{ fontSize: '15px' }} />}
          disabled={isLoading}
          onClick={() => onAction('consistency_audit')}
        >
          Consistency
        </Button>
      </Tooltip>
    </div>
  );
};
