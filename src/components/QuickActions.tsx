import React from "react";
import {
  Button,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Tooltip,
  Badge,
} from "@fluentui/react-components";
import { QuickActionType } from "../types/chat";
import { PowerPointApiRequirements } from "../types/ppt";
import { ZeeSettings } from "../types/settings";

interface QuickActionsProps {
  requirements: PowerPointApiRequirements;
  settings: ZeeSettings;
  disabled?: boolean;
  onRunQuickAction: (action: QuickActionType) => void;
  onOpenReportToDeck: () => void;
  onOpenExcelToSlide: () => void;
  onOpenConsistencyChecker: () => void;
  onOpenTextFitFixer: () => void;
  onOpenDeckReviewer: () => void;
  onOpenBeautifier: () => void;
}

export const QuickActions: React.FC<QuickActionsProps> = ({
  requirements,
  settings,
  disabled,
  onRunQuickAction,
  onOpenReportToDeck,
  onOpenExcelToSlide,
  onOpenConsistencyChecker,
  onOpenTextFitFixer,
  onOpenDeckReviewer,
  onOpenBeautifier,
}) => {
  // Hide features that require unsupported PowerPointApi requirement sets instead of failing
  const canCreateSlides = !requirements.isPowerPointHost || requirements.pptApi1_1;
  const canInspectShapes = !requirements.isPowerPointHost || requirements.pptApi1_4;
  const canWriteNotes =
    !requirements.isPowerPointHost || requirements.pptApi1_4 || requirements.pptApi1_7;

  return (
    <div
      className="zee-compact-toolbar"
      style={{ flexDirection: "column", alignItems: "stretch", gap: 5 }}
    >
      {/* Row 1: Text & Slide Quick Actions */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center" }}>
        <Tooltip
          content="Rewrite selected text or active slide for executive clarity"
          relationship="description"
        >
          <Button
            size="small"
            appearance="outline"
            disabled={disabled}
            onClick={() => onRunQuickAction("rewrite")}
            style={{ fontSize: 11, padding: "2px 7px" }}
          >
            ✨ Rewrite
          </Button>
        </Tooltip>

        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button
              size="small"
              appearance="outline"
              disabled={disabled}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              🌐 Translate ▾
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem onClick={() => onRunQuickAction("translate_ar")}>
                Translate to Arabic (RTL + Protected Terms)
              </MenuItem>
              <MenuItem onClick={() => onRunQuickAction("translate_en")}>
                Translate to English (LTR + Protected Terms)
              </MenuItem>
              <MenuItem onClick={() => onRunQuickAction("translate_bilingual")}>
                Bilingual (English + Arabic Side-by-Side)
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        <Tooltip
          content="Summarize selected text or current slide into crisp bullets"
          relationship="description"
        >
          <Button
            size="small"
            appearance="outline"
            disabled={disabled}
            onClick={() => onRunQuickAction("summarize")}
            style={{ fontSize: 11, padding: "2px 7px" }}
          >
            📝 Summarize
          </Button>
        </Tooltip>

        {canWriteNotes && (
          <Tooltip
            content="Generate speaker notes for the active slide and write to notes pane"
            relationship="description"
          >
            <Button
              size="small"
              appearance="outline"
              disabled={disabled}
              onClick={() => onRunQuickAction("speaker_notes")}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              🎙️ Notes
            </Button>
          </Tooltip>
        )}

        <Tooltip
          content="Convert oilfield units in selected text or slide (bbl/d, psi, ft, °F, scf/stb)"
          relationship="description"
        >
          <Button
            size="small"
            appearance="outline"
            disabled={disabled}
            onClick={() => onRunQuickAction("convert_units")}
            style={{ fontSize: 11, padding: "2px 7px" }}
          >
            🛢️ Units
          </Button>
        </Tooltip>
      </div>

      {/* Row 2: PowerPoint-Specific Deck, Petroleum & Review Tools */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center" }}>
        {canCreateSlides && (
          <Tooltip
            content="Turn a Word report or PDF into a slide outline first, then build slides after approval"
            relationship="description"
          >
            <Button
              size="small"
              appearance="secondary"
              disabled={disabled}
              onClick={onOpenReportToDeck}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              📄 Report → Deck
            </Button>
          </Tooltip>
        )}

        {canCreateSlides && (
          <Tooltip
            content="Excel → Slide (Auto-Plot Decline/ESP Charts & KPIs) + Multi-Well Batch Well-Book Generator"
            relationship="description"
          >
            <Button
              size="small"
              appearance="secondary"
              disabled={disabled}
              onClick={onOpenExcelToSlide}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              📊 Excel / Well-Book
            </Button>
          </Tooltip>
        )}

        {canInspectShapes && (
          <Tooltip
            content="Management Q&A & Deck Reviewer: spot numerical contradictions across slides + build backup Q&A slides"
            relationship="description"
          >
            <Button
              size="small"
              appearance="secondary"
              disabled={disabled}
              onClick={onOpenDeckReviewer}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              🎯 Q&amp;A Review
            </Button>
          </Tooltip>
        )}

        {canInspectShapes && (
          <Tooltip
            content="Smart Slide Beautifier: convert dense slides into 2-Column splits, 3-Card callouts, or aligned grids"
            relationship="description"
          >
            <Button
              size="small"
              appearance="secondary"
              disabled={disabled}
              onClick={onOpenBeautifier}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              🎨 Beautify
            </Button>
          </Tooltip>
        )}

        {canInspectShapes && (
          <Tooltip
            content="Check fonts, sizes, title capitalization, units, and protected terms across all slides"
            relationship="description"
          >
            <Button
              size="small"
              appearance="secondary"
              disabled={disabled}
              onClick={onOpenConsistencyChecker}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              ✅ Consistency
            </Button>
          </Tooltip>
        )}

        {canInspectShapes && (
          <Tooltip
            content="Find and shorten bullets that overflow slide placeholders"
            relationship="description"
          >
            <Button
              size="small"
              appearance="secondary"
              disabled={disabled}
              onClick={onOpenTextFitFixer}
              style={{ fontSize: 11, padding: "2px 7px" }}
            >
              📐 Text-Fit
            </Button>
          </Tooltip>
        )}

        {settings.companyTemplate.enabled && (
          <Badge size="small" appearance="tint" color="brand" style={{ marginLeft: "auto" }}>
            Template: {settings.companyTemplate.layoutName}
          </Badge>
        )}
      </div>
    </div>
  );
};
