import * as React from 'react';
import styles from './EmployeeHubShoutOutAcknowledge.module.scss';
import { IEmployeeHubShoutOutAcknowledgeProps } from './IEmployeeHubShoutOutAcknowledgeProps';
import { escape } from '@microsoft/sp-lodash-subset';


export interface IEmployeeHubShoutOutAcknowledgeState {

}

require('../assets/style.css');

export default class EmployeeHubShoutOutAcknowledge extends React.Component<IEmployeeHubShoutOutAcknowledgeProps, IEmployeeHubShoutOutAcknowledgeState> {

  constructor(props: IEmployeeHubShoutOutAcknowledgeProps, state: IEmployeeHubShoutOutAcknowledgeState) {

    super(props);

    this.state = {

    };

  }


  public render(): React.ReactElement<IEmployeeHubShoutOutAcknowledgeProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      ShoutOutTitle,
      ShoutOutDescription,
      Link
    } = this.props;

    return (
      <section className="employeeHubShoutOutAcknowledge">

        <div className="shoutout-card">
          <div className="card-top">
            <img className='sharedImage' src={require('../assets/sharedimage.jpg')}/>
          </div>
          <div className="card-body">
            <h3 className="card-title">
              {ShoutOutTitle}
              {/* Shout Out: Acknowledge a Colleague! */}
            </h3>
            <p className="card-text">
              {ShoutOutDescription}
              {/* Recognise a fellow team member for their hard work, support,
              or great attitude. Leave your message here! */}
            </p>
            <a href={Link} className="submitbtn">Submit a Shout Out</a>
          </div>
        </div>

      </section>
    );
  }
}
