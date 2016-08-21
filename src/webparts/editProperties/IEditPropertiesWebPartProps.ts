import { IWebPartContext } from '@microsoft/sp-client-preview';

export interface IEditPropertiesWebPartProps {
  description: string;
  userprofileproperty: string;
  context: IWebPartContext;
  webAbsoluteUrl: string;
  userLoginName: string;
  propertyName: string;
}
