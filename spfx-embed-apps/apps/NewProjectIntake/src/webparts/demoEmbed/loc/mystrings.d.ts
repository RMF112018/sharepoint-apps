declare interface IDemoEmbedWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TargetListNameLabel: string;
  ShowHeroLabel: string;
  CompactModeLabel: string;
  ToggleOnText: string;
  ToggleOffText: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'DemoEmbedWebPartStrings' {
  const strings: IDemoEmbedWebPartStrings;
  export = strings;
}
