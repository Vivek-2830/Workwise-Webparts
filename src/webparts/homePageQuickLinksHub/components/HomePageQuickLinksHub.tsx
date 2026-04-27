import * as React from 'react';
import styles from './HomePageQuickLinksHub.module.scss';
import { IHomePageQuickLinksHubProps } from './IHomePageQuickLinksHubProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';


export interface IHomePageQuickLinksHubState {
  QuickLinkAllData: any;
}

require('../assets/style.css');

export default class HomePageQuickLinksHub extends React.Component<IHomePageQuickLinksHubProps, IHomePageQuickLinksHubState> {

  constructor(props: IHomePageQuickLinksHubProps, state: IHomePageQuickLinksHubState) {

    super(props);

    this.state = {
      QuickLinkAllData: ""
    };

  }

  public render(): React.ReactElement<IHomePageQuickLinksHubProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;



    return (
      <section className="homePageQuickLinksHub">

        <div className="quick-links">

          {
            this.state.QuickLinkAllData.length > 0 &&
            this.state.QuickLinkAllData.map((item) => {
              let imagePath = "";
              let ImageInfo = JSON.parse(item.Icons);
              if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                imagePath = ImageInfo["serverRelativeUrl"];
              }
              else {
                imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Links Hub/Attachments/${item.ID}/${ImageInfo.fileName}`;
              }

              return (
                <div className="link-card">
                  <a href={item.Link.Url} style={{ textDecoration: "none" }}>
                    <img src={imagePath} />
                    <p>{item.Title}</p>
                  </a>
                </div>
              );
            })
          }

        </div>

      </section>
    );
  }

  public async componentDidMount() {
    this.getquicklinksData();
  }

  public async getquicklinksData() {
    const links = await sp.web.lists.getByTitle("Quick Links Hub").items.select(
      "ID",
      "Title",
      "Icons",
      "Link"
    ).get().then((data) => {
      let AllData = [];
      console.log(links);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Icons: item.Icons ? item.Icons : "",
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ QuickLinkAllData: AllData });
      }
    }).catch((error) => {
      console.log("Error fetching Quick Links Hub data: ", error);
    });
  }


}
