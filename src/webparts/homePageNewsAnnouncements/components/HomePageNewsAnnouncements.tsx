import * as React from 'react';
import styles from './HomePageNewsAnnouncements.module.scss';
import { IHomePageNewsAnnouncementsProps } from './IHomePageNewsAnnouncementsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { Pivot, PivotItem } from 'office-ui-fabric-react';
import * as moment from 'moment';

export interface IHomePageNewsAnnouncementsState {
  NewsAnnouncementsData: any;
  NewsFilterdData: any;
 
}

require('../assets/style.css');

export default class HomePageNewsAnnouncements extends React.Component<IHomePageNewsAnnouncementsProps, IHomePageNewsAnnouncementsState> {

  constructor(props: IHomePageNewsAnnouncementsProps, state: IHomePageNewsAnnouncementsState) {
  
    super(props);
    
    this.state = {
      NewsAnnouncementsData: "",
      NewsFilterdData: "",
      
    };

  }


  public render(): React.ReactElement<IHomePageNewsAnnouncementsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.homePageNewsAnnouncements} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className="news-panel">

          <div className="news-header">
            <h2 className="section-title">News &amp; Announcements</h2>
            <a href='https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/SitePages/News%20&%20Announcements%20Page.aspx' style={{ textDecoration: "none", color: "black" }} target="_blank" rel="noopener noreferrer">
              <button className="view-news">View all</button>
            </a>
          </div>

          <div className="title-underline"></div>

          <div className="news-filters">
            <Pivot onLinkClick={this._onPivotChange}>
              <PivotItem headerText="All News" itemKey="all" />
              <PivotItem headerText="Company" itemKey="company" />
              <PivotItem headerText="Community" itemKey="community" />
              <PivotItem headerText="Charity" itemKey="charity" />
              <PivotItem headerText="Colleagues" itemKey="colleagues" />
              <PivotItem headerText="Contracts" itemKey="contracts" />
            </Pivot>
          </div>

          <div className='news-scroll'>

            <div className="news-list">

              {
                this.state.NewsFilterdData.length > 0 &&
                this.state.NewsFilterdData.map((item) => {
                  // let imagePath = "";
                  // let ImageInfo = JSON.parse(item.NewsPhoto);
                  // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                  //   imagePath = ImageInfo["serverRelativeUrl"];
                  // }
                  // else {
                  //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/News Announcement/Attachments/${item.ID}/${ImageInfo.fileName}`;
                  // }

                  return (
                    <div className="news-card">
                      <img src={item.NewsPhoto} />

                      <div className="news-content">
                        <p className="news-tag">{item.NewsCategory}</p>
                        <h4>{item.NewsTitle}</h4>
                        <p>{moment(item.NewsDate).format("Do MMMM,YYYY")}</p>
                        <a href={item.Link ? item.Link.Url : ""} style={{ textDecoration: "none", color: "black" }}>View more →</a>
                      </div>
                    </div>
                  );
                })
              }

            </div>

          </div>

        </div>
      </section>
    );
  }

  public async componentDidMount() {
    this.getNewsAnnouncementsData();
  }

   public async getNewsAnnouncementsData() {

    const items = await sp.web.lists
      .getByTitle("News Announcements")
      .items.select(
        "ID",
        "NewsTitle",
        "NewsPhoto",
        "NewsCategory",
        "NewsDate",
        "Link",
        "AttachmentFiles"
      )
      .expand("AttachmentFiles")
      .orderBy("NewsDate", false)
      .get();

    let formattedData: any[] = [];

    if (items.length > 0) {

      items.forEach((news) => {
        formattedData.push({
          ID: news.ID || "",
          NewsTitle: news.NewsTitle || "",
          NewsPhoto:
            news.AttachmentFiles.length > 0
              ? news.AttachmentFiles[0].ServerRelativeUrl
              : require("../assets/Rectangle1.png"),
          NewsCategory: news.NewsCategory || "",
          NewsDate: news.NewsDate || "",
          Link: news.Link || ""
        });
      });

      // GROUP BY CATEGORY
      const grouped = formattedData.reduce((acc, item) => {
        if (!acc[item.NewsCategory]) {
          acc[item.NewsCategory] = [];
        }
        acc[item.NewsCategory].push(item);
        return acc;
      }, {});

      // TAKE TOP 4 FROM EACH CATEGORY
      let topFourPerCategory: any[] = [];

      Object.keys(grouped).forEach((category) => {
        const top4 = grouped[category].slice(0, 6);
        topFourPerCategory = [...topFourPerCategory, ...top4];
      });

      const reduced = formattedData.reduce((acc: any, item: any) => {
        const category = item.NewsCategory;
      
        if (
          !acc[category] ||
          new Date(item.Created) > new Date(acc[category].Created)
        ) {
          acc[category] = item;
        }
      
        return acc;
      }, {});
      
      // // Convert object → array (ES5 safe)
      // const latestPerCategory: any[] = [];
      
      // for (let key in reduced) {
      //   if (reduced.hasOwnProperty(key)) {
      //     latestPerCategory.push(reduced[key]);
      //   }
      // }

      this.setState({
        NewsAnnouncementsData: topFourPerCategory,
        NewsFilterdData: topFourPerCategory
      });

    }
  }

  private _onPivotChange = (item?: PivotItem): void => {
    if (!item) return;

    let filterdata = this.state.NewsAnnouncementsData;

    

    switch (item.props.itemKey) {

      case "company":
        filterdata = filterdata.filter(t => t.NewsCategory === "Company");
        break;

      case "community":
        filterdata = filterdata.filter(t => t.NewsCategory === "Community");
        break;

      case "charity":
        filterdata = filterdata.filter(t => t.NewsCategory === "Charity");
        break;

      case "colleagues":
        filterdata = filterdata.filter(t => t.NewsCategory === "Colleagues");
        break;

      case "contracts":
        filterdata = filterdata.filter(t => t.NewsCategory === "Contracts");
        break;

      case "all":
      default:
        filterdata = this.state.NewsFilterdData;
    }

    this.setState({ NewsFilterdData: filterdata });
  }

}
