import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import * as strings from 'mystrings';
import EditProperties, { IEditPropertiesProps } from './components/EditProperties';
import { IEditPropertiesWebPartProps } from './IEditPropertiesWebPartProps';

export default class EditPropertiesWebPart extends BaseClientSideWebPart<IEditPropertiesWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<IEditPropertiesProps> = React.createElement(EditProperties, {
      description: this.properties.description,
      userprofileproperty: this.properties.userprofileproperty,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      userLoginName: encodeURIComponent(_spClientSidePageContext.user.LoginName),
      propertyName: this.properties.userprofileproperty,
      context: this.context
    });

    ReactDom.render(element, this.domElement);


  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('userprofileproperty', {
                  label: strings.UserProfilePropertyFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
		return true;
	}
}
