import { IUserDetails } from './IUserDetails.ts';
import { IWebPartContext } from '@microsoft/sp-client-preview';
import { IEditPropertiesProps } from './components/EditProperties';

interface IUserProfileService {
  getUserProperties: Promise<IUserDetails>;
  setUserProperties: Promise<void>;
  webAbsoluteUrl: string;
  propertyName: string;
  userLoginName: string;
  context: IWebPartContext;
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

  public setUserProperties(propertyValue: string): void{

    this.context.httpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`,
    { body: `{"accountName":"${decodeURIComponent(this.props.userLoginName)}","propertyName":"${this.props.propertyName}","propertyValue":"${propertyValue}"}`  })

    .then((response: any) => {
        console.log(response);
      },
      (response: any) => {
        console.log(response);
      });

  }

  // public setUserProperties(propertyValue: string): Promis<any>{
  //   return this.context.httpClient.post(
  //   `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`,
  //   {
  //     body: `{"accountName":"${decodeURIComponent(this.props.userLoginName)}","propertyName":"${this.props.propertyName}","propertyValue":"${propertyValue}"}`
  //   })
  //   .then((response: any) => {
  //       return response.json();
  //   });
  // }
}