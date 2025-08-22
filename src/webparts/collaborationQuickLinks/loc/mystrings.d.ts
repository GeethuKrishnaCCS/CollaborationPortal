declare interface ICollaborationQuickLinksWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  AddNewLinkTitle: string;
  AddButtonLabel: string;
  NolinksMessage: string;
  UseDefaultIcon: string;
  SelectFromFiles: string;
  SelectFromIcons: string;
  TextOnly: string;
}

declare module 'CollaborationQuickLinksWebPartStrings' {
  const strings: ICollaborationQuickLinksWebPartStrings;
  export = strings;
}
