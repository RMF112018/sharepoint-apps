import React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import DemoEmbed from '../DemoEmbed';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

jest.mock('@hbi/sp-client', () => ({
  createSharePointClient: () => ({}),
  getCurrentUser: async () => ({ id: 1, loginName: 'user', displayName: 'User' }),
  ensureListExists: async () => true,
  createListItem: async () => ({})
}));

function ctx(): BaseComponentContext { return {} as any; }

function renderComp() {
  return render(
    <DemoEmbed
      description="test"
      isDarkTheme={false}
      environmentMessage="env"
      hasTeamsContext={false}
      userDisplayName="User"
      context={ctx()}
      targetListName="CursorDemo"
      showHero={true}
      compactMode={false}
    />
  );
}

test('title is required and disables submit', async () => {
  renderComp();
  const submit = await screen.findByRole('button', { name: /submit to sharepoint/i });
  expect(submit).toBeDisabled();

  const title = screen.getByLabelText(/title/i);
  fireEvent.change(title, { target: { value: 'Hello' } });
  expect(submit).not.toBeDisabled();
});

test('notes character limit shows error and disables submit', async () => {
  renderComp();
  const title = await screen.findByLabelText(/title/i);
  fireEvent.change(title, { target: { value: 'Hello' } });
  const notes = screen.getByLabelText(/notes/i);
  fireEvent.change(notes, { target: { value: 'x'.repeat(600) } });

  const submit = screen.getByRole('button', { name: /submit to sharepoint/i });
  expect(submit).toBeDisabled();
  expect(screen.getByText(/reduce to 500 characters or fewer/i)).toBeInTheDocument();
});
