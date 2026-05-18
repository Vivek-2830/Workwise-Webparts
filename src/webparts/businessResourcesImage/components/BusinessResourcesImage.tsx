import * as React from 'react';
import styles from './BusinessResourcesImage.module.scss';
import { IBusinessResourcesImageProps } from './IBusinessResourcesImageProps';
import { escape } from '@microsoft/sp-lodash-subset';

require('../assets/style.css');


export default class BusinessResourcesImage extends React.Component<IBusinessResourcesImageProps, {}> {
  public render(): React.ReactElement<IBusinessResourcesImageProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const ImageSectionLink = this.props.BusinessImage == undefined ? require('../assets/Image.png') : this.props.BusinessImage.fileAbsoluteUrl;

    return (
      <section className="businessResourcesImage">

            <div className="business-image">
              <img className='Responsive-Img' src={ImageSectionLink} alt="FAQ Image" />
            </div>
        
      </section>
    );
  }
}
