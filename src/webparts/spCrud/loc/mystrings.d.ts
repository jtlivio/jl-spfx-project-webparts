declare interface ISpCrudStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListsFieldLabel: string;
  ListsOperationsFieldLabel: string;
}

declare module 'spCrudStrings' {
  const strings: ISpCrudStrings;
  export = strings;
}
