import * as React from 'react';
import styles from './BusinessResourcesBanner.module.scss';
import { IBusinessResourcesBannerProps } from './IBusinessResourcesBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

require('../assets/style.css');

export default class BusinessResourcesBanner extends React.Component<IBusinessResourcesBannerProps, {}> {
  public render(): React.ReactElement<IBusinessResourcesBannerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      BusinessTitle,
      BusinessDescription
    } = this.props;

    const BannerImageLink = this.props.filePickerResult == undefined ? require('../assets/Rectangle26.png') : this.props.filePickerResult.fileAbsoluteUrl;

    return (
      <section className="businessResourcesBanner">
        
        <div className="Business-section" style={{ backgroundImage: " url(" + BannerImageLink + ")" }}>
          <div className="Business-overlay"></div>

          <div className="Business-content">
            <h1>{BusinessTitle}</h1>
            <p>
              {BusinessDescription}
            </p>
          </div>
        </div>
       
      </section>
    );
  }
}
