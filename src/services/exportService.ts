import { ChatMessage } from '../types/chat';
import { pptService } from './pptService';

export const exportService = {
  /**
   * Formats chat messages into Markdown string
   */
  formatAsMarkdown(messages: ChatMessage[]): string {
    const lines: string[] = [
      '# Zee AI for PowerPoint - Chat Transcript',
      `Exported: ${new Date().toLocaleString()}`,
      '---\n'
    ];

    for (const msg of messages) {
      const roleName = msg.role === 'user' ? '**You**' : '**Zee AI**';
      const time = new Date(msg.timestamp).toLocaleTimeString();
      lines.push(`### ${roleName} (${time})\n`);
      lines.push(`${msg.content}\n`);

      if (msg.attachments && msg.attachments.length > 0) {
        lines.push(`*Attachments: ${msg.attachments.map(a => a.name).join(', ')}*\n`);
      }

      if (msg.agentSteps && msg.agentSteps.length > 0) {
        lines.push('**Agent Actions:**');
        for (const step of msg.agentSteps) {
          lines.push(`- [${step.status.toUpperCase()}] ${step.description}`);
        }
        lines.push('');
      }

      lines.push('---\n');
    }

    return lines.join('\n');
  },

  /**
   * Copies formatted transcript to clipboard
   */
  async copyToClipboard(messages: ChatMessage[]): Promise<boolean> {
    try {
      const text = this.formatAsMarkdown(messages);
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      return false;
    }
  },

  /**
   * Downloads formatted transcript as markdown file
   */
  downloadAsFile(messages: ChatMessage[], filename = 'zee-ai-chat-export.md'): void {
    const text = this.formatAsMarkdown(messages);
    const blob = new Blob([text], { type: 'text/markdown;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  },

  /**
   * Inserts chat summary directly as a new slide in the presentation
   */
  async exportToNewSlide(messages: ChatMessage[]): Promise<{ slideIndex: number }> {
    const recent = messages.filter(m => m.role === 'assistant' && m.content.trim()).slice(-3);
    const summaryBullets = recent.map(m => m.content.substring(0, 150).replace(/\n+/g, ' '));

    return await pptService.addSlide({
      title: 'Zee AI Conversation Summary',
      bullets: summaryBullets.length > 0 ? summaryBullets : ['No recent assistant messages to summarize.']
    });
  }
};
