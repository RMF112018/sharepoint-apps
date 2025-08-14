export interface IDemoEmbedProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: import("@microsoft/sp-webpart-base").BaseComponentContext;
}
