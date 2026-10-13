import React, { useState } from 'react';
import {
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Button,
  Tooltip,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Input,
  Textarea,
  Field
} from '@fluentui/react-components';
import { BookStar24Regular, Add24Regular } from '@fluentui/react-icons';
import { SavedPrompt } from '../types/chat';

interface SavedPromptsMenuProps {
  prompts: SavedPrompt[];
  onSelectPrompt: (promptText: string) => void;
  onAddPrompt: (newPrompt: SavedPrompt) => void;
}

export const SavedPromptsMenu: React.FC<SavedPromptsMenuProps> = ({
  prompts,
  onSelectPrompt,
  onAddPrompt
}) => {
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [title, setTitle] = useState('');
  const [promptText, setPromptText] = useState('');

  const handleSave = () => {
    if (title.trim() && promptText.trim()) {
      onAddPrompt({
        id: `custom-${Date.now()}`,
        title: title.trim(),
        prompt: promptText.trim(),
        category: 'generate'
      });
      setTitle('');
      setPromptText('');
      setIsDialogOpen(false);
    }
  };

  return (
    <>
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <Tooltip content="Saved Prompts" relationship="label">
            <Button
              appearance="subtle"
              size="small"
              icon={<BookStar24Regular style={{ fontSize: '16px' }} />}
              aria-label="Open saved prompts menu"
            />
          </Tooltip>
        </MenuTrigger>
        <MenuPopover>
          <MenuList style={{ minWidth: '220px', maxWidth: '300px' }}>
            {prompts.map(p => (
              <MenuItem
                key={p.id}
                onClick={() => onSelectPrompt(p.prompt)}
                secondaryContent={p.category}
              >
                {p.title}
              </MenuItem>
            ))}
            <MenuItem
              icon={<Add24Regular />}
              onClick={() => setIsDialogOpen(true)}
            >
              Save New Prompt...
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Dialog open={isDialogOpen} onOpenChange={(_, d) => setIsDialogOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Save Custom Prompt</DialogTitle>
            <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <Field label="Prompt Title">
                <Input
                  value={title}
                  onChange={(_, d) => setTitle(d.value)}
                  placeholder="e.g. Executive Summary Generator"
                />
              </Field>
              <Field label="Prompt Instructions">
                <Textarea
                  value={promptText}
                  onChange={(_, d) => setPromptText(d.value)}
                  rows={4}
                  placeholder="Enter instructions for Zee AI..."
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setIsDialogOpen(false)}>
                Cancel
              </Button>
              <Button appearance="primary" onClick={handleSave}>
                Save Prompt
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
};
