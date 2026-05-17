import * as React from 'react';
import styles from './EmployeeHubFeedBackSection.module.scss';
import { IEmployeeHubFeedBackSectionProps } from './IEmployeeHubFeedBackSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';


export interface IEmployeeHubFeedBackSectionState {

}

require('../assets/style.css');

export default class EmployeeHubFeedBackSection extends React.Component<IEmployeeHubFeedBackSectionProps, IEmployeeHubFeedBackSectionState> {

  constructor(props: IEmployeeHubFeedBackSectionProps, state: IEmployeeHubFeedBackSectionState) {

    super(props);

    this.state = {

    };

  }


  public render(): React.ReactElement<IEmployeeHubFeedBackSectionProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubFeedBackSection">

        <h1 className="section-feed">Hearts & Minds - Our Culture, Our People</h1>
        <p className="section-subtitle">
          Join us in shaping a more engaged, connected, and vibrant workplace ready for our Investors in People survey!
        </p>

        <div className="feedback-panel">
          <div className="feedback-grid">

            <div>
              <h3 className="column-title">Your Feedback (You Said)</h3>
              <div className="column-items">
                <div className="feedback-item">Need clearer communication on teams roles</div>
                <div className="feedback-item">More structured mentorship opportunities</div>
                <div className="feedback-item">More structured team interest in the employees</div>
                <div className="feedback-item">More clear and communicative mentorship opportunities</div>
              </div>
            </div>

            <div>
              <h3 className="column-title">Our Response (We Did)</h3>
              <div className="column-items">
                <div className="feedback-item">Implementing a weekly department-wide email digest</div>
                <div className="feedback-item">Launched a new peer mentoring program for all staff</div>
                <div className="feedback-item">Launched a new peer mentorship for career opportunities</div>
                <div className="feedback-item">Launched a new peer mentorship program for training and success</div>
              </div>
            </div>

          </div>
        </div>

      </section>
    );
  }
}
