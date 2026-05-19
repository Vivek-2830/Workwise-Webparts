import * as React from 'react';
import styles from './HomePageEvents.module.scss';
import { IHomePageEventsProps } from './IHomePageEventsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import { sp } from '@pnp/sp/presets/all';
import { PrimaryButton } from 'office-ui-fabric-react';

export interface IHomePageEventsState {
  EventsAllDate: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

export default class HomePageEvents extends React.Component<IHomePageEventsProps, IHomePageEventsState> {

  constructor(props: IHomePageEventsProps, state: IHomePageEventsState) {
    super(props);
    
    this.state = {
      EventsAllDate: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };
  }

  public render(): React.ReactElement<IHomePageEventsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageEvents">

        <div className="events-panel">

          <h2 className="section-title">Events</h2>
          <div className="title-underline"></div>

          {
            this.state.IsAdmin ?
            <>
                <div>
                  <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Lists/Events/calendar.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
                    <PrimaryButton className='Adddoc' text="Add Event" />
                  </a>
                </div>
            </>
            :
            <></>
          }

          <div className='events-scroll'>

            {/* filter(item => moment(item.EventDate).isSameOrAfter(moment(), "day")) */}

            {
              this.state.EventsAllDate.length > 0 &&
              this.state.EventsAllDate.map((item) => {
                return (
                  <>

                    <a href={"https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/_layouts/15/Event.aspx?ListGuid=66c1121d-07ad-4acc-9eeb-ff8af6ba70da&ItemId=" + item.Id} style={{ textDecoration: "none", cursor: "pointer", color: "inherit" }}>
                      <div className="event-Meeting">

                        <div className="event-date">
                          <h3>{moment(item.EventDate).format("DD")}</h3>
                          <span>{moment(item.EventDate).format("MMM").toUpperCase()}</span>
                        </div>

                        <div className="event-info">
                          <p>{item.Title}</p>
                          <p className='event-time'>{item.EventTime}</p>

                        </div>
                        {item.Category ? <><span className="event-tag week">{item.Category}</span></> : <></>}

                      </div>
                    </a>
                    {/* {
            item.EventCategory == "Knowledge Exchange" ?
              <>
                <a href={item.Link.Url} style={{ textDecoration: "none", cursor: "pointer", color: "inherit" }}>
                  <div className="event-Meeting">

                    <div className="event-date">
                      <h3>{moment(item.EventDate).format("DD")}</h3>
                      <span>{moment(item.EventDate).format("MMM").toUpperCase()}</span>
                    </div>

                    <div className="event-info">
                      <p>{item.Title}</p>
                      <p className='event-time'>{item.EventTime}</p>

                    </div>

                    <span className="event-tag week">{item.EventCategory}</span>
                  </div>
                </a>
              </>
              :
              <>
                {
                  item.EventCategory == "Exhibitions & Sponsorships" ?
                    <>
                      <a href={item.Link.Url} style={{ textDecoration: "none", cursor: "pointer", color: "inherit" }}>
                        <div className="event-Business">

                          <div className="event-date">
                            <h3>{moment(item.EventDate).format("DD")}</h3>
                            <span>{moment(item.EventDate).format("MMM").toUpperCase()}</span>
                          </div>

                          <div className="event-info">
                            <p>{item.EventTitle}</p>
                            <p className='event-time'>{item.EventTime}</p>

                          </div>

                          <span className="event-tag Business">{item.EventCategory}</span>
                        </div>
                      </a>
                    </>
                    :
                    <>
                      {
                        item.EventCategory == "Awards" ?
                          <>
                            <a href={item.Link.Url} style={{ textDecoration: "none", cursor: "pointer", color: "inherit" }}>
                              <div className="event-Training">

                                <div className="event-date">
                                  <h3>{moment(item.EventDate).format("DD")}</h3>
                                  <span>{moment(item.EventDate).format("MMM").toUpperCase()}</span>
                                </div>

                                <div className="event-info">
                                  <p>{item.EventTitle}</p>
                                  <p className='event-time'>{item.EventTime}</p>

                                </div>

                                <span className="event-tag Training">{item.EventCategory}</span>

                              </div>
                            </a>
                          </>
                          :
                          <>
                          </>
                      }
                    </>
                }
              </>
          } */}

                  </>

                );
              })
            }

            {/* <div className="event-card active">
          <div className="event-date">
            <h3>27</h3>
            <span>FEB</span>
          </div>

          <div className="event-info">
            <p>Meet &amp; Greet</p>
            <p className='event-time'>4:00 PM - 5:00 PM</p>
            <a href="#">+ RSVP</a>
          </div>

          <span className="event-tag week">Week Two</span>
        </div>

        <div className="event-card">
          <div className="event-date">
            <h3>27</h3>
            <span>FEB</span>
          </div>

          <div className="event-info">
            <p>Lunch &amp; Learn</p>
            <p className='event-time'>4:00 PM - 5:00 PM</p>
            <a href="#">+ RSVP</a>
          </div>

          <span className="event-tag meeting">Meeting</span>
        </div>

        <div className="event-card">
          <div className="event-date">
            <h3>27</h3>
            <span>FEB</span>
          </div>

          <div className="event-info">
            <p>New Hire Orientation</p>
            <p className='event-time'>4:00 PM - 5:00 PM</p>
            <a href="#">+ RSVP</a>
          </div>

          <span className="event-tag business">Business</span>
        </div>

        <div className="event-card">
          <div className="event-date">
            <h3>27</h3>
            <span>FEB</span>
          </div>

          <div className="event-info">
            <p>New Hire Orientation</p>
            <p className='event-time'>4:00 PM - 5:00 PM</p>
            <a href="#">+ RSVP</a>
          </div>

          <span className="event-tag business">Business</span>
        </div> */}

          </div>

        </div>
        
      </section>
    );
  }

  public async componentDidMount() {
    // this.getEvents();
    this.getEventsDetails();
    this.GetCurrentUser();
  }

  public async GetCurrentUser() {
    try {
      const currentUser = await sp.web.currentUser.get();
      const userEmail = currentUser.Email.toLowerCase().trim();
      const ownerGroup = await sp.web.associatedOwnerGroup();
      const groupUsers = await sp.web.siteGroups.getById(ownerGroup.Id).users.get();

      const isAdmin = groupUsers.some(user =>
        user.LoginName.toLowerCase() === currentUser.LoginName.toLowerCase()
      );
      this.setState({ IsAdmin: true });
      this.setState({ IsAdmin: isAdmin, CurrentUserEmail: userEmail });
    } catch (error) {
      console.error("Error checking admin status:", error);
    }
  }

  // public async getEvents() {
  //   const today = new Date().toISOString();
  //   const event = await sp.web.lists.getByTitle("Company Events").items.select(
  //     "ID",
  //     "EventTitle",
  //     "EventTime",
  //     "EventDate",
  //     "EventCategory",
  //     "Link"
  //   )
  //   .filter(`EventDate ge datetime'${today}'`).orderBy("EventDate", true).top(5).get().then((data) => {
  //     let AllData = [];
  //     console.log(event);
  //     console.log(data);
  //     if (data.length > 0) {
  //       data.forEach((item) => {
  //         AllData.push({
  //           ID: item.ID ? item.ID : "",
  //           EventTitle: item.EventTitle ? item.EventTitle : "",
  //           EventTime: item.EventTime ? item.EventTime : "",
  //           EventDate: item.EventDate ? item.EventDate : "",
  //           EventCategory: item.EventCategory ? item.EventCategory : "",
  //           Link: item.Link ? item.Link : ""
  //         });
  //       });
  //       this.setState({ EventsAllDate: AllData });
  //     }
  //   }).catch((error) => {
  //     console.log("Error Fetching Events data: ", error);
  //   });
  // }

  private getEventsDetails = async () => {
    try {
      const EventsDetails = await sp.web.lists.getByTitle("Events").items.select("*").filter(`EventDate ge '${moment.parseZone().format('YYYY-MM-DD')}'`).top(5).get();

      if (EventsDetails.length > 0) {
        this.setState({
          EventsAllDate: EventsDetails.sort((o1, o2) => {
            if (moment(moment.parseZone(o1.EventDate).format("YYYY-MM-DD")).isBefore(moment.parseZone(o2.EventDate).format("YYYY-MM-DD"))) {
              return -1;
            }
            else {
              return 1;
            }
          })
        });
      }
    }
    catch (error) {
      console.log(error);
    }
  }


}
