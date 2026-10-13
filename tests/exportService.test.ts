import { describe, it, expect } from 'vitest';
import { exportService } from '../src/services/exportService';
import { ChatMessage } from '../src/types/chat';

describe('exportService', () => {
  it('formats conversation transcript as markdown', () => {
    const messages: ChatMessage[] = [
      {
        id: '1',
        role: 'user',
        content: 'Create a slide about ESP optimization.',
        timestamp: 1700000000000
      },
      {
        id: '2',
        role: 'assistant',
        content: 'I have added the slide with key metrics.',
        timestamp: 1700000005000,
        agentSteps: [
          {
            id: 's1',
            stepNumber: 1,
            toolName: 'add_slide',
            description: 'Add ESP slide',
            status: 'success',
            timestamp: 1700000002000
          }
        ]
      }
    ];

    const md = exportService.formatAsMarkdown(messages);
    expect(md).toContain('# Zee AI for PowerPoint - Chat Transcript');
    expect(md).toContain('**You**');
    expect(md).toContain('Create a slide about ESP optimization.');
    expect(md).toContain('**Zee AI**');
    expect(md).toContain('[SUCCESS] Add ESP slide');
  });
});
