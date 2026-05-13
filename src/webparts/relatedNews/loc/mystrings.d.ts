declare interface IRelatedNewsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'RelatedNewsWebPartStrings' {
  const strings: IRelatedNewsWebPartStrings;
  export = strings;
}
