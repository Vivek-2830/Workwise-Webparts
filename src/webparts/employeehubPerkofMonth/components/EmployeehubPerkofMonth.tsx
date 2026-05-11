import * as React from 'react';
import styles from './EmployeehubPerkofMonth.module.scss';
import { IEmployeehubPerkofMonthProps } from './IEmployeehubPerkofMonthProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IEmployeehubPerkofMonthState {

}

require('../assets/style.css');

export default class EmployeehubPerkofMonth extends React.Component<IEmployeehubPerkofMonthProps, IEmployeehubPerkofMonthState> {

  constructor(props: IEmployeehubPerkofMonthProps, state:IEmployeehubPerkofMonthState) {

    super(props);

    this.state = {

    };

  }


  public render(): React.ReactElement<IEmployeehubPerkofMonthProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      PerkMonthDescription,
      LinkButton
    } = this.props;

    const ImageSectionLink = this.props.PerkMonthImage == undefined ? require('../assets/Frame14.png') : this.props.PerkMonthImage.fileAbsoluteUrl;

    return (
      <section className="employeehubPerkofMonth">

        <div className="perk-card">

          <div className="perk-header">
            {/* <h3>Perk of the Month</h3> */}
            {/* <span className="arrow">⌃</span> */}
          </div>


          <div className="image-container">
            <a href={LinkButton} style={{ textDecoration: 'none', color: 'inherit' }}>
              <img src={ImageSectionLink} alt="Mixo 8 Machine" className='ImageCon' />
            </a>
            <p className='perkdes' style={{ textAlign: 'center' }}>{PerkMonthDescription}</p>
          </div>

        </div>


      </section>
    );
  }
}
