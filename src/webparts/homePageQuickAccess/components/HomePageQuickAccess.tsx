import * as React from 'react';
import styles from './HomePageQuickAccess.module.scss';
import { IHomePageQuickAccessProps } from './IHomePageQuickAccessProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

export interface IHomePageQuickAccessState {
  QuickAccessData : any;
}

require('../assets/style.css');

export default class HomePageQuickAccess extends React.Component<IHomePageQuickAccessProps, IHomePageQuickAccessState> {

  constructor(props: IHomePageQuickAccessProps, state: IHomePageQuickAccessState) {

    super(props);

    this.state = {
      QuickAccessData: ""  
    };

  }


  public render(): React.ReactElement<IHomePageQuickAccessProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageQuickAccess">

        <div className="quick-access-wrapper">

          <h2>Quick Access</h2>
          <p className="subtitle">
            Find the applications and resources you need to get work done efficiently
          </p>

          <div className="qa-cards">

            {/* ================= Productivity ================= */}
            {
              this.state.QuickAccessData.length > 0 &&
              this.state.QuickAccessData.filter(i => i.QuickAccessCategories === "Tools").length > 0 &&

              <div className="qa-card">
                <h3>
                  <span className="card-icon">
                    <img src={require('../assets/baggage.png')} />
                  </span>
                  Tools
                </h3>

                <ul>
                  {
                    this.state.QuickAccessData.map((item) => {

                      if (item.QuickAccessCategories !== "Tools") return null;

                      let imagePath = "";
                      let ImageInfo = JSON.parse(item.Icon);
                      if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                        imagePath = ImageInfo["serverRelativeUrl"];
                      } else {
                        imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Access/Attachments/${item.ID}/${ImageInfo.fileName}`;
                      }

                      return (
                        <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                          <li key={item.ID}>
                            <span className="item-icon">
                              <img src={imagePath} />
                            </span>
                            <div>
                              <strong>{item.AccessTitle}</strong>
                              <p>{item.AccessDescription}</p>
                            </div>
                          </li>
                        </a>
                      );
                    })
                  }
                </ul>
              </div>
            }


            {/* ================= Human Resources ================= */}
            {
              this.state.QuickAccessData.length > 0 &&
              this.state.QuickAccessData.filter(i => i.QuickAccessCategories === "Support").length > 0 &&

              <div className="qa-card">
                <h3>
                  <span className="card-icon">
                    <img src={require('../assets/group.png')} />
                  </span>
                  Support
                </h3>

                <ul>
                  {
                    this.state.QuickAccessData.map((item) => {

                      if (item.QuickAccessCategories !== "Support") return null;

                      let imagePath = "";
                      let ImageInfo = JSON.parse(item.Icon);
                      if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                        imagePath = ImageInfo["serverRelativeUrl"];
                      } else {
                        imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Access/Attachments/${item.ID}/${ImageInfo.fileName}`;
                      }

                      return (
                        <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                          <li key={item.ID}>
                            <span className="item-icon">
                              <img src={imagePath} />
                            </span>
                            <div>
                              <strong>{item.AccessTitle}</strong>
                              <p>{item.AccessDescription}</p>
                            </div>
                          </li>
                        </a>
                      );
                    })
                  }
                </ul>
              </div>
            }


            {/* ================= Business Applications ================= */}
            {
              this.state.QuickAccessData.length > 0 &&
              this.state.QuickAccessData.filter(i => i.QuickAccessCategories === "Resources").length > 0 &&

              <div className="qa-card">
                <h3>
                  <span className="card-icon">
                    <img src={require('../assets/phone.png')} />
                  </span>
                  Resources
                </h3>

                <ul>
                  {
                    this.state.QuickAccessData.map((item) => {

                      if (item.QuickAccessCategories !== "Resources") return null;

                      let imagePath = "";
                      let ImageInfo = JSON.parse(item.Icon);
                      if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                        imagePath = ImageInfo["serverRelativeUrl"];
                      } else {
                        imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Access/Attachments/${item.ID}/${ImageInfo.fileName}`;
                      }

                      return (
                        <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                          <li key={item.ID}>
                            <span className="item-icon">
                              <img src={imagePath} />
                            </span>
                            <div>
                              <strong>{item.AccessTitle}</strong>
                              <p>{item.AccessDescription}</p>
                            </div>
                          </li>
                        </a>
                      );
                    })
                  }
                </ul>
              </div>
            }

          </div>

        </div>

      </section>
    );
  }

  public async componentDidMount() {
    this.getQuickAccessData();
  }

  public async getQuickAccessData() {
    const quickdata = await sp.web.lists.getByTitle("Quick Access").items.select(
      "ID",
      "QuickAccessCategories",
      "AccessTitle",
      "AccessDescription",
      "Icon",
      "Link"
    ).get().then((data) => {
      let AllData = [];
      console.log(quickdata);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            QuickAccessCategories: item.QuickAccessCategories ? item.QuickAccessCategories : "",
            AccessTitle: item.AccessTitle ? item.AccessTitle : "",
            AccessDescription: item.AccessDescription ? item.AccessDescription : "",
            Icon: item.Icon,
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ QuickAccessData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Quick Access Data:", error);
    });
  }



}
