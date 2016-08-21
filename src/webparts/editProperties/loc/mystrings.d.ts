declare interface IStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  UserProfilePropertyFieldLabel: string;
}

declare module 'mystrings' {
  const strings: IStrings;
  export = strings;
}
