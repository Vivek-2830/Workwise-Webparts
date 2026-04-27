import * as React from 'react';
import styles from './HomePageEvents.module.scss';
import { IHomePageEventsProps } from './IHomePageEventsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import { sp } from '@pnp/sp/presets/all';

export interface IHomePageEventsState {
  EventsAllDate: any;
}

require('../assets/style.css');

export default class HomePageEvents extends React.Component<IHomePageEventsProps, IHomePageEventsState> {

  constructor(props: IHomePageEventsProps, state: IHomePageEventsState) {
    super(props);
    
    this.state = {
      EventsAllDate: ""
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

              <div className='events-scroll'>

                {/* filter(item => moment(item.EventDate).isSameOrAfter(moment(), "day")) */}

                {
                  this.state.EventsAllDate.length > 0 &&
                  this.state.EventsAllDate.map((item) => {
                    return (
                      <>
                        {
                          item.EventCategory == "Knowledge Exchange" ?
                            <>
                              <a href={item.Link.Url} style={{ textDecoration: "none", cursor: "pointer", color: "inherit" }}>
                                <div className="event-Meeting">

                                  <div className="event-date">
                                    <h3>{moment(item.EventDate).format("DD")}</h3>
                                    <span>{moment(item.EventDate).format("MMM").toUpperCase()}</span>
                                  </div>

                                  <div className="event-info">
                                    <p>{item.EventTitle}</p>
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
                                    <a href={item.Link.Url} style={{ textDecoration: "none", cursor: "pointer" , color: "inherit" }}>
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
                                          <a href={item.Link.Url} style={{ textDecoration: "none", cursor: "pointer" , color: "inherit"}}>
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
                        }

                      </>

                    );
                  })
                }

                

              </div>

            </div>
        
      </section>
    );
  }

  public async componentDidMount() {
    this.getEvents();
  }

   public async getEvents() {
    const today = new Date().toISOString();
    const event = await sp.web.lists.getByTitle("Company Events").items.select(
      "ID",
      "EventTitle",
      "EventTime",
      "EventDate",
      "EventCategory",
      "Link"
    )
    .filter(`EventDate ge datetime'${today}'`).orderBy("EventDate", true).top(5).get().then((data) => {
      let AllData = [];
      console.log(event);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            EventTitle: item.EventTitle ? item.EventTitle : "",
            EventTime: item.EventTime ? item.EventTime : "",
            EventDate: item.EventDate ? item.EventDate : "",
            EventCategory: item.EventCategory ? item.EventCategory : "",
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ EventsAllDate: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Events data: ", error);
    });
  }


}
