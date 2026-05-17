import * as React from 'react';
import styles from './EmployeeHubAxisClcGallery.module.scss';
import { IEmployeeHubAxisClcGalleryProps } from './IEmployeeHubAxisClcGalleryProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IEmployeeHubAxisClcGalleryState {

}

require('../assets/style.css');

export default class EmployeeHubAxisClcGallery extends React.Component<IEmployeeHubAxisClcGalleryProps, IEmployeeHubAxisClcGalleryState> {

  constructor(props: IEmployeeHubAxisClcGalleryProps, state: IEmployeeHubAxisClcGalleryState) {

    super(props);

    this.state = {

    };

  }

  public render(): React.ReactElement<IEmployeeHubAxisClcGalleryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubAxisClcGallery">

        <div className="gallery-section">


          <div className="gallery-header">
            Life at Axis CLC (Photo Gallery)
          </div>


          <div className="gallery-grid">
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 1" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 2" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 3" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 4" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 5" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 6" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 7" /></div>
            <div className="gallery-item"><img src={require('../assets/Frame14.png')} alt="Gallery Image 8" /></div>
          </div>

          <button className="submit-btn">
            Submit an Image
            <svg viewBox="0 0 24 24">
              <path d="M5 20h14v-2H5v2zm7-18l-5.5 5.5 1.42 1.42L11 6.84V16h2V6.84l3.08 3.08 1.42-1.42L12 2z" />
            </svg>
          </button>

        </div>

      </section>
    );
  }
}
