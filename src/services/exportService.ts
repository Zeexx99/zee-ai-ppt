import { ChatMessage } from "../types/chat";
import { APP_VERSION, BUILD_DATE, ZeeSettings } from "../types/settings";
import { addSlide, getActiveSlideIndex, insertTextBox } from "./pptService";

export function formatChatHistoryForExport(
  messages: ChatMessage[],
  deckTitle: string = "PowerPoint Presentation"
): string {
  const header = [
    `# Zee AI for PowerPoint — Chat Export`,
    `- **Presentation:** ${deckTitle}`,
    `- **Exported:** ${new Date().toISOString()}`,
    `- **Version:** v${APP_VERSION} (${BUILD_DATE})`,
    "",
    "---",
    "",
  ].join("\n");

  const body = messages
    .map((m) => {
      const roleLabel =
        m.role === "user"
          ? "### 👤 User"
          : m.role === "assistant"
            ? "### 🤖 Zee AI"
            : "### ℹ️ System";
      const timeStr = new Date(m.timestamp).toLocaleTimeString();
      const lines = [`${roleLabel} _(${timeStr})_`, "", m.content.trim()];

      if (m.agentSteps && m.agentSteps.length > 0) {
        lines.push("", "**Agent Slide Steps:**");
        for (const st of m.agentSteps) {
          lines.push(
            `- [${st.status.toUpperCase()}] Step ${st.stepNumber}: ${st.description}${
              st.resultSummary ? ` — ${st.resultSummary}` : ""
            }${st.errorMessage ? ` (Error: ${st.errorMessage})` : ""}`
          );
        }
      }

      if (m.citations && m.citations.length > 0) {
        lines.push("", "**Sources:**");
        for (const c of m.citations) {
          lines.push(`- [${c.title}](${c.uri})`);
        }
      }

      return lines.join("\n");
    })
    .join("\n\n---\n\n");

  return `${header}${body}\n`;
}

export async function copyTextToClipboard(text: string): Promise<boolean> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // Fallback below
  }

  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(textarea);
    return ok;
  } catch {
    return false;
  }
}

export function downloadChatTranscript(
  messages: ChatMessage[],
  filename: string = `zee-ai-ppt-chat-${new Date().toISOString().slice(0, 10)}.md`
): void {
  const markdown = formatChatHistoryForExport(messages);
  const blob = new Blob([markdown], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export async function exportTextToSlideElement(
  text: string,
  mode: "current_slide_box" | "new_slide",
  settings: ZeeSettings
): Promise<string> {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.replace(/^#+\s*/, "").trim())
    .filter(Boolean);

  if (mode === "new_slide") {
    const title = lines[0]?.slice(0, 90) || "Zee AI Summary";
    const body = lines.length > 1 ? lines.slice(1).join("\n") : text;
    const res = await addSlide({
      title,
      bodyText: body,
      settings,
    });
    return res.summary;
  }

  const activeSlide = await getActiveSlideIndex();
  return await insertTextBox({
    slideNumber: activeSlide,
    text,
    settings,
  });
}
