import * as React from 'react';
import styles from './EmployeeHubPeopleStories.module.scss';
import { IEmployeeHubPeopleStoriesProps } from './IEmployeeHubPeopleStoriesProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IEmployeeHubPeopleStoriesState {

}


export default class EmployeeHubPeopleStories extends React.Component<IEmployeeHubPeopleStoriesProps, IEmployeeHubPeopleStoriesState> {

  constructor(props: IEmployeeHubPeopleStoriesProps, state: IEmployeeHubPeopleStoriesState) {

    super(props);

    this.state = {

    };

  }


  public render(): React.ReactElement<IEmployeeHubPeopleStoriesProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubPeopleStories">

        <section className="stories-section">
          <h2 className="stories-title">Investing in People: Stories of Growth</h2>

          <div className="stories-grid">

            <div className="story-card">
              <div className="story-thumbnail">
                <video autoPlay muted loop playsInline controls style={{ width: "100%", height: "149px", objectFit: "cover" }}>
                  <source src="https://cdn.pixabay.com/video/2026/03/02/337459_large.mp4" type="video/mp4" />
                  Your browser does not support the video tag.
                </video>
              </div>
              <div className="story-content">
                <div className="story-title">
                  Sarah’s Development - New Learning Program Participation
                </div>
              </div>
            </div>


            <div className="story-card">
              <div className="story-thumbnail">
                <video autoPlay muted loop playsInline controls style={{ width: "100%", height: "149px", objectFit: "cover" }}>
                  <source src="https://cdn.pixabay.com/video/2026/03/02/337459_large.mp4" type="video/mp4" />
                  Your browser does not support the video tag.
                </video>
              </div>
              <div className="story-content">
                <div className="story-title">
                  Joanne’s Promotion - Mentorship &amp; Career Path
                </div>
              </div>
            </div>


            <div className="story-card">
              <div className="story-thumbnail">
                <video autoPlay muted loop playsInline controls style={{ width: "100%", height: "149px", objectFit: "cover" }}>
                  <source src="https://cdn.pixabay.com/video/2026/03/02/337459_large.mp4" type="video/mp4" />
                  Your browser does not support the video tag.
                </video>

              </div>
              <div className="story-content">
                <div className="story-title">
                  David’s Retirement - What did Axis CLC do for his career?
                </div>
              </div>
            </div>


            <div className="story-card">
              <div className="story-thumbnail">
                <video autoPlay muted loop playsInline controls style={{ width: "100%", height: "149px", objectFit: "cover" }}>
                  <source src="https://cdn.pixabay.com/video/2026/03/02/337459_large.mp4" type="video/mp4" />
                  Your browser does not support the video tag.
                </video>

              </div>
              <div className="story-content">
                <div className="story-title">
                  Maxine’s Promotion - Mentorship &amp; Career Path
                </div>
              </div>
            </div>
          </div>

          <div className="accent-line"></div>
        </section>

      </section>
    );
  }
}
