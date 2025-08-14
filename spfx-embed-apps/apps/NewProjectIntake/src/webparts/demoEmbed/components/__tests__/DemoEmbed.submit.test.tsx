import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import DemoEmbed from '../DemoEmbed';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

const sp = {
  createSharePointClient: () => ({}),
  getCurrentUser: async () => ({ id: 1, loginName: 'user', displayName: 'User' }),
};

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
      targetListName="NewProjectIntake"
      showHero={true}
      compactMode={false}
    />
  );
}

test('successful submit shows success toast', async () => {
  renderComp();
  fireEvent.change(await screen.findByLabelText(/title/i), { target: { value: 'Hello' } });
  fireEvent.click(screen.getByRole('button', { name: /save entry/i }));

  await waitFor(() => expect(screen.getByText(/saved to newprojectintake/i)).toBeInTheDocument());
});

test('keyboard-only: Enter submits when valid (simulated form submit)', async () => {
  renderComp();
  const title = await screen.findByLabelText(/title/i);
  fireEvent.change(title, { target: { value: 'Hello' } });
  const form = title.closest('form');
  if (!form) throw new Error('Form not found');
  fireEvent.submit(form);

  await waitFor(() => expect(screen.getByText(/saved to newprojectintake/i)).toBeInTheDocument());
});
