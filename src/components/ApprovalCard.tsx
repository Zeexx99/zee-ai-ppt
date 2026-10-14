import React from "react";
import { Badge, Button, Spinner } from "@fluentui/react-components";
import { AgentToolStep } from "../types/chat";

interface ApprovalCardProps {
  steps: AgentToolStep[];
  awaitingContinue?: boolean;
  onApproveStep: (stepId: string) => void;
  onRejectStep: (stepId: string) => void;
  onUndoStep: (step: AgentToolStep) => void;
  onRetryStep: (step: AgentToolStep) => void;
  onContinueAgent?: () => void;
}

export const ApprovalCard: React.FC<ApprovalCardProps> = ({
  steps,
  awaitingContinue,
  onApproveStep,
  onRejectStep,
  onUndoStep,
  onRetryStep,
  onContinueAgent,
}) => {
  if (!steps || steps.length === 0) return null;

  const statusColor = (status: AgentToolStep["status"]) => {
    switch (status) {
      case "completed":
        return "success";
      case "failed":
        return "danger";
      case "pending_approval":
        return "warning";
      case "running":
      case "approved":
        return "brand";
      default:
        return "subtle";
    }
  };

  return (
    <div
      style={{
        marginTop: 8,
        display: "flex",
        flexDirection: "column",
        gap: 6,
      }}
    >
      <div style={{ fontSize: 11, fontWeight: 600, opacity: 0.8 }}>
        🤖 Agent Slide Steps ({steps.filter((s) => s.status === "completed").length}/{steps.length})
      </div>

      {steps.map((step) => (
        <div
          key={step.id}
          style={{
            border:
              step.status === "pending_approval"
                ? "1.5px solid #d83b01"
                : step.status === "failed"
                  ? "1px solid #d13438"
                  : "1px solid var(--colorNeutralStroke2, #e1dfdd)",
            borderRadius: 8,
            padding: "8px 10px",
            background: "var(--colorNeutralBackground2, rgba(0,0,0,0.03))",
            fontSize: 12,
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 6,
            }}
          >
            <span style={{ fontWeight: 600 }}>
              #{step.stepNumber} • {step.toolName}
            </span>
            <Badge size="small" appearance="filled" color={statusColor(step.status)}>
              {step.status.replace("_", " ")}
            </Badge>
          </div>

          <div style={{ marginTop: 4, lineHeight: 1.4 }}>{step.description}</div>

          {step.status === "running" && (
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6 }}>
              <Spinner size="extra-tiny" />
              <span style={{ fontSize: 11 }}>Executing in PowerPoint...</span>
            </div>
          )}

          {step.resultSummary && step.status !== "failed" && (
            <div
              style={{
                marginTop: 4,
                fontSize: 11,
                opacity: 0.85,
                maxHeight: 68,
                overflowY: "auto",
              }}
            >
              ✓ {step.resultSummary.slice(0, 220)}
            </div>
          )}

          {/* Failed steps are kept visible with the error so the user can retry */}
          {step.status === "failed" && (
            <div
              style={{
                marginTop: 6,
                padding: "6px 8px",
                borderRadius: 6,
                background: "rgba(209, 52, 56, 0.12)",
                color: "#d13438",
                fontSize: 11,
              }}
            >
              <div>⚠️ Error: {step.errorMessage || "Step failed"}</div>
              <div style={{ marginTop: 4 }}>
                <Button
                  size="small"
                  appearance="secondary"
                  onClick={() => onRetryStep(step)}
                  style={{ fontSize: 11, padding: "2px 8px" }}
                >
                  ↻ Retry Step
                </Button>
              </div>
            </div>
          )}

          {/* Approval gate for write tools */}
          {step.status === "pending_approval" && (
            <div
              style={{
                marginTop: 8,
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              <Button
                size="small"
                appearance="primary"
                onClick={() => onApproveStep(step.id)}
                style={{ backgroundColor: "#b73618", fontSize: 11 }}
              >
                ✓ Approve & Apply
              </Button>
              <Button
                size="small"
                appearance="outline"
                onClick={() => onRejectStep(step.id)}
                style={{ fontSize: 11 }}
              >
                ✕ Skip
              </Button>
            </div>
          )}

          {/* Multi-step Undo button */}
          {step.status === "completed" && step.undoSnapshot && (
            <div style={{ marginTop: 6, display: "flex", justifyContent: "flex-end" }}>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => onUndoStep(step)}
                style={{ fontSize: 11, padding: "1px 6px" }}
              >
                ↶ Undo Step
              </Button>
            </div>
          )}
        </div>
      ))}

      {awaitingContinue && onContinueAgent && (
        <div
          style={{
            marginTop: 4,
            padding: "8px 10px",
            borderRadius: 8,
            border: "1px dashed #b73618",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 8,
          }}
        >
          <span style={{ fontSize: 11 }}>
            Step limit reached to prevent runaway loops. Continue agent execution?
          </span>
          <Button
            size="small"
            appearance="primary"
            onClick={onContinueAgent}
            style={{ backgroundColor: "#b73618", fontSize: 11 }}
          >
            ▶ Continue
          </Button>
        </div>
      )}
    </div>
  );
};
