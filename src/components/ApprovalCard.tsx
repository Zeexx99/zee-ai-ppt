import React from 'react';
import { Card, Button, Badge } from '@fluentui/react-components';
import { Checkmark24Regular, Dismiss24Regular, Sparkle24Regular } from '@fluentui/react-icons';
import { ToolApproval } from '../types/chat';

interface ApprovalCardProps {
  approval: ToolApproval;
  onApprove: () => void;
  onReject: () => void;
}

export const ApprovalCard: React.FC<ApprovalCardProps> = ({ approval, onApprove, onReject }) => {
  return (
    <Card
      style={{
        margin: '8px 0',
        padding: '12px',
        border: '1px solid var(--colorBrandStroke1)',
        background: 'var(--colorNeutralBackground2)',
        borderRadius: '8px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.06)'
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
          <Sparkle24Regular style={{ color: '#D24726', fontSize: '18px' }} />
          <span style={{ fontWeight: 600, fontSize: '13px', color: 'var(--colorNeutralForeground1)' }}>
            Action Approval Required
          </span>
        </div>
        <Badge appearance="tint" color="informative">
          {approval.toolName.replace(/_/g, ' ')}
        </Badge>
      </div>

      <div style={{ fontSize: '13px', color: 'var(--colorNeutralForeground2)', marginBottom: '12px' }}>
        {approval.description}
      </div>

      {approval.params && Object.keys(approval.params).length > 0 && (
        <div
          style={{
            fontSize: '11px',
            fontFamily: 'Consolas, monospace',
            background: 'var(--colorNeutralBackground3)',
            padding: '6px 8px',
            borderRadius: '4px',
            marginBottom: '12px',
            maxHeight: '100px',
            overflowY: 'auto'
          }}
        >
          {JSON.stringify(approval.params, null, 2)}
        </div>
      )}

      {approval.isPending ? (
        <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
          <Button
            size="small"
            appearance="subtle"
            icon={<Dismiss24Regular />}
            onClick={onReject}
          >
            Reject
          </Button>
          <Button
            size="small"
            appearance="primary"
            icon={<Checkmark24Regular />}
            onClick={onApprove}
            style={{ background: '#107C41', borderColor: '#107C41' }}
          >
            Approve Action
          </Button>
        </div>
      ) : (
        <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
          <Badge
            appearance="filled"
            color={approval.approved ? 'success' : 'danger'}
          >
            {approval.approved ? 'Approved & Executed' : 'Rejected'}
          </Badge>
        </div>
      )}
    </Card>
  );
};
