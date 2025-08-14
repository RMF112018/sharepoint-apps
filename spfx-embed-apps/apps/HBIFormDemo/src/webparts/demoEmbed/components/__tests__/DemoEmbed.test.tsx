import React from 'react';
import { render, screen } from '@testing-library/react';
import DemoEmbed from '../DemoEmbed';
import type { BaseComponentContext } from '@microsoft/sp-component-base';

jest.mock('@hbi/sp-client', () => ({
  createSharePointClient: () => ({}),
  getCurrentUser: async () => ({ id: 1, loginName: 'john', displayName: 'John' }),
  ensureListExists: async () => true,
  createListItem: async () => ({})
}));

function ctx(): BaseComponentContext { return {} as any; }

test('renders hero and form', async () => {
  render(
    <DemoEmbed
      description="test"
      isDarkTheme={false}
      environmentMessage="local"
      hasTeamsContext={false}
      userDisplayName="User"
      context={ctx()}
      targetListName="CursorDemo"
      showHero={true}
      compactMode={false}
    />
  );

  expect(await screen.findByText(/HBI Form Demo/i)).toBeInTheDocument();
  expect(screen.getByLabelText(/Title/i)).toBeInTheDocument();
});


