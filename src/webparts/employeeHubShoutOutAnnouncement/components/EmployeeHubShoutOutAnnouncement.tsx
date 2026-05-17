import * as React from 'react';
import styles from './EmployeeHubShoutOutAnnouncement.module.scss';
import { IEmployeeHubShoutOutAnnouncementProps } from './IEmployeeHubShoutOutAnnouncementProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IEmployeeHubShoutOutAnnouncementState {

}

require('../assets/style.css');

export default class EmployeeHubShoutOutAnnouncement extends React.Component<IEmployeeHubShoutOutAnnouncementProps, IEmployeeHubShoutOutAnnouncementState> {

  constructor(props: IEmployeeHubShoutOutAnnouncementProps, state: IEmployeeHubShoutOutAnnouncementState) {

    super(props);

    this.state = {

    };

  }

  public render(): React.ReactElement<IEmployeeHubShoutOutAnnouncementProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubShoutOutAnnouncement">

        <div className="shoutouts-section">
          <h2 className="shoutouts-title">Shout Out's</h2>

          <div className="shout-card">
            <div className="shout-icon">📣</div>
            <div className="shout-content">
              <div className="shout-name">John Smith</div>
              <div className="shout-message">
                John has been an amazing help finalising the end of year accounts
              </div>
            </div>
          </div>

          <div className="shout-card">
            <div className="shout-icon">📣</div>
            <div className="shout-content">
              <div className="shout-name">Ella Ferguson</div>
              <div className="shout-message">
                Ella really stepped up and has taken on extra workload since a team member left
              </div>
            </div>
          </div>

          <div className="shout-card">
            <div className="shout-icon">📣</div>
            <div className="shout-content">
              <div className="shout-name">Mark Bronson</div>
              <div className="shout-message">
                Mark has been a great addition to our team and has picked things up really quickly
              </div>
            </div>
          </div>

          <div className="shout-card">
            <div className="shout-icon">📣</div>
            <div className="shout-content">
              <div className="shout-name">Andrew Grange</div>
              <div className="shout-message">
                Andy recently started helping with the training of new staff and is doing great
              </div>
            </div>
          </div>

          <div className="shout-card">
            <div className="shout-icon">📣</div>
            <div className="shout-content">
              <div className="shout-name">Amanda Brogue</div>
              <div className="shout-message">
                Amanda is really holding together the administration of the Exeter office until we get a new hire in to assist
              </div>
            </div>
          </div>

          <div className="shout-card">
            <div className="shout-icon">📣</div>
            <div className="shout-content">
              <div className="shout-name">Oliver Ridgeley</div>
              <div className="shout-message">
                Ollie has been a hard working member of the team since he joined in June
              </div>
            </div>
          </div>
        </div>

      </section>
    );
  }
}
