declare interface ICollaborationDirectoryWebPartStrings {
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

  DropDownPlaceLabelMessage: string;
  DropDownPlaceHolderMessage: string;
  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  DirectoryMessage: string;
  LoadingText: string;
  SearchPropsLabel: string;
  SearchPropsDesc: string;
  ClearTextSearchPropsLabel: string;
  ClearTextSearchPropsDesc: string;
  PagingLabel: string;
}

declare module 'CollaborationDirectoryWebPartStrings' {
  const strings: ICollaborationDirectoryWebPartStrings;
  export = strings;
}
