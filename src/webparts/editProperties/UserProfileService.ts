import { IUserDetails } from './IUserDetails.ts';
import { IWebPartContext } from '@microsoft/sp-client-preview';
import { IEditPropertiesProps } from './components/EditProperties';

interface IUserProfileService {
  getUserProperties: Promise<IUserDetails>;
  setUserProperties: void;
  webAbsoluteUrl: string;
  propertyName: string;
  userLoginName: string;
  context: IWebPartContext;
}

export interface IResult{
  status: number;
  statusText: string;
}

export class UserProfileService {
  private context: IWebPartContext;
  private props: IEditPropertiesProps;

  constructor(_props: IEditPropertiesProps){
      this.props = _props;
      this.context = _props.context;
  }

  public getUserProperties(): Promise<IUserDetails> {
    return this.context.httpClient.get(
    `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/
    GetUserProfilePropertyFor(accountName=@v,propertyName='${this.props.propertyName}')?@v='${this.props.userLoginName}'`)
    .then((response: Response) => {
        return response.json();
    });
  }

  public setUserProperties(propertyValue: string): Promise<Response> {
    const postBody: Object = {
        'accountName': decodeURIComponent(this.props.userLoginName),
        'propertyName': this.props.propertyName,
        'propertyValue': propertyValue
    };

    //Explicitly add the odata v3 header to work with SharePoint REST API
    const reqHeaders: Headers = new Headers();
    reqHeaders.append('odata-version', '3.0');

    return this.context.httpClient.post(
    `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`,
      {
        body: JSON.stringify(postBody),
        headers: reqHeaders
      })
    .then((response: Response) => {
          return response;
     });
  }
}