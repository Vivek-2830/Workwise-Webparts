import * as React from 'react';
import styles from './EmployeehubPageBanner.module.scss';
import { IEmployeehubPageBannerProps } from './IEmployeehubPageBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IEmployeehubPageBannerState {

}

require('../assets/style.css');

export default class EmployeehubPageBanner extends React.Component<IEmployeehubPageBannerProps, IEmployeehubPageBannerState> {

  constructor(props: IEmployeehubPageBannerProps, state: IEmployeehubPageBannerState) {
  
    super(props);

    this.state = {

    };

  
  }

  public render(): React.ReactElement<IEmployeehubPageBannerProps> {
    const {
      description,
      Title,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const BannerImageLink = this.props.filePickerResult == undefined ? require('../assets/Rectangle26.png') : this.props.filePickerResult.fileAbsoluteUrl;

    return (
      <section className="employeehubPageBanner">
        
        <div className="hero-section" style={{ backgroundImage: " url(" + BannerImageLink + ")" }}>
          <div className="hero-overlay"></div>

          <div className="hero-content">
            <h1>{Title}</h1>
            <p>
              {description}
            </p>
          </div>
        </div>


      </section>
    );
  }
}
