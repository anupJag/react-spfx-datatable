declare interface IMarsDatatableWebPartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListSelectionLabel : string;
  ColumnSelectionLabel : string;
  SearchTextBoxLabel : string;

  //Placeholder Components
  noTilesIconText: string;
  noTilesConfigured: string;
  noTilesConfiguredClassicSP : string;
  noTilesBtn: string;

  //Error Component
  errorOccured : string;
  ErrorOnItemsFetch : string;
  ErrorOnPermissions : string;
}

declare module 'MarsDatatableWebPartWebPartStrings' {
  const strings: IMarsDatatableWebPartWebPartStrings;
  export = strings;
}
