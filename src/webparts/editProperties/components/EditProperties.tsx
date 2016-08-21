import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import styles from '../EditProperties.module.scss';
import { IEditPropertiesWebPartProps } from '../IEditPropertiesWebPartProps';
import { UserProfileService } from '../UserProfileService';

export interface IEditPropertiesProps extends IEditPropertiesWebPartProps {
}

export interface IEditPropertiesWebPartState {
  userprofileproperty: string;
  result?: string;
}

export default class EditProperties extends React.Component<IEditPropertiesProps, IEditPropertiesWebPartState> {

  constructor(props: IEditPropertiesProps) {
    super(props);
    this.state = {
      userprofileproperty: "",
      result: ""
    };
  }

  public render(): JSX.Element {
    return (
      <div className={styles.editProperties}>
        <div className={styles.container}>
          <div className={css('ms-Grid-row ms-bgColor-themeSecondary ms-fontColor-white', styles.row) }>
            <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
              <span className='ms-font-xl ms-fontColor-white'>
                Welcome to the SharePoint Framework!
              </span>
              <p className='ms-font-l ms-fontColor-white'>
                {this.props.description}
              </p>
              <div className={css('ms-TextField')}>
                <label className={css('ms-Label ms-fontColor-white')}>{this.props.userprofileproperty}</label>
                <input className={css('ms-TextField-field')} value={this.state.userprofileproperty} onChange={this.handleChange.bind(this)}></input>
              </div>
              <a className={css('ms-Button', styles.button) } href='#' onClick={this._setProperties.bind(this)}>
                <span className='ms-Button-label'>Update</span>
              </a>
              <div>
                <label className="ms-Label">{this.state.result}</label>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }


  public handleChange (event: any): void {
    this.setState({ userprofileproperty: event.target.value});
  }

  public componentDidMount(): void {
    this._getProperties();
  }

  private _setProperties(): void {
    const userProfileService: UserProfileService = new UserProfileService(this.props);
    userProfileService.setUserProperties(this.state.userprofileproperty);
    // userProfileService.setUserProperties(this.state.userprofileproperty).then((response) => {
    //   this.setState({
    //     userprofileproperty: response.value,
    //     result: response.error.message
    //   });
    // });
  }

  private _getProperties(): void {
    const userProfileService: UserProfileService = new UserProfileService(this.props);
    userProfileService.getUserProperties().then((response) => {
      this.setState({ userprofileproperty: response.value });
    });
  }
}
