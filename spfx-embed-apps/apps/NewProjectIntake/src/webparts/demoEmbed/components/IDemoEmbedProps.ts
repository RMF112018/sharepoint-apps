import type { BaseComponentContext } from '@microsoft/sp-component-base';
export interface IDemoEmbedProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: BaseComponentContext;
  targetListName: string;
  showHero: boolean;
  compactMode: boolean;
}
