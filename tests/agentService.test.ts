import { describe, it, expect } from 'vitest';
import { agentService, AGENT_TOOLS } from '../src/services/agentService';
import { DEFAULT_SETTINGS } from '../src/types/settings';

describe('agentService', () => {
  it('correctly distinguishes read-only vs mutating tools', () => {
    expect(agentService.isReadOnlyTool('list_slides')).toBe(true);
    expect(agentService.isReadOnlyTool('read_slide')).toBe(true);
    expect(agentService.isReadOnlyTool('read_selected_text')).toBe(true);
    expect(agentService.isReadOnlyTool('read_speaker_notes')).toBe(true);

    // Mutating write tools
    expect(agentService.isReadOnlyTool('add_slide')).toBe(false);
    expect(agentService.isReadOnlyTool('delete_slide')).toBe(false);
    expect(agentService.isReadOnlyTool('insert_text_box')).toBe(false);
    expect(agentService.isReadOnlyTool('insert_table')).toBe(false);
    expect(agentService.isReadOnlyTool('write_speaker_notes')).toBe(false);
    expect(agentService.isReadOnlyTool('replace_text')).toBe(false);
  });

  it('formats readable approval descriptions', () => {
    const addDesc = agentService.formatApprovalDescription('add_slide', {
      title: 'Reservoir Simulation',
      bullets: ['Point 1', 'Point 2']
    });
    expect(addDesc).toContain('Add new slide: "Reservoir Simulation"');
    expect(addDesc).toContain('2 bullet points');

    const delDesc = agentService.formatApprovalDescription('delete_slide', { slideIndex: 4 });
    expect(delDesc).toBe('Delete slide 4');

    const tableDesc = agentService.formatApprovalDescription('insert_table', {
      slideIndex: 2,
      rows: 4,
      columns: 3
    });
    expect(tableDesc).toBe('Insert 4x3 table on slide 2');
  });

  it('builds system prompt with presentation summary and protected terms', () => {
    const overview = {
      slideCount: 3,
      slides: [
        { index: 1, title: 'Title Slide' },
        { index: 2, title: 'ESP Performance' },
        { index: 3, title: 'Conclusion' }
      ]
    };

    const prompt = agentService.buildSystemPrompt(overview, DEFAULT_SETTINGS);
    expect(prompt).toContain('Total slides: 3');
    expect(prompt).toContain('Slide 2: "ESP Performance"');
    expect(prompt).toContain('CRITICAL INSTRUCTION - PROTECTED TERMS');
    expect(prompt).toContain('"ESP"');
    expect(prompt).toContain('"TDH"');
  });

  it('registers all required PowerPoint agent tools', () => {
    const names = AGENT_TOOLS.map(t => t.name);
    expect(names).toContain('list_slides');
    expect(names).toContain('read_slide');
    expect(names).toContain('add_slide');
    expect(names).toContain('delete_slide');
    expect(names).toContain('insert_text_box');
    expect(names).toContain('insert_table');
    expect(names).toContain('write_speaker_notes');
    expect(names).toContain('replace_text');
  });
});
