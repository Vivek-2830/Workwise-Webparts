import * as React from 'react';
import styles from './HomePageAllNewsAnnouncementPage.module.scss';
import { IHomePageAllNewsAnnouncementPageProps } from './IHomePageAllNewsAnnouncementPageProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Pivot, PivotItem } from 'office-ui-fabric-react';
import * as moment from 'moment';
import { sp } from '@pnp/sp/rest';

require('../assets/style.css');

export interface IHomePageAllNewsAnnouncementPageState {
  NewsAnnouncementsData: any;
  NewsFilterdData: any;
}


export default class HomePageAllNewsAnnouncementPage extends React.Component<IHomePageAllNewsAnnouncementPageProps, IHomePageAllNewsAnnouncementPageState> {

  constructor(props: IHomePageAllNewsAnnouncementPageProps, state: IHomePageAllNewsAnnouncementPageState) {

    super(props);
    this.state = {
      NewsAnnouncementsData: "",
      NewsFilterdData: ""
    };

  }


  public render(): React.ReactElement<IHomePageAllNewsAnnouncementPageProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageAllNewsAnnouncementPage">

        <div className="news-panel">

          <div className="news-header">
            <h2 className="section-title">News &amp; Announcements</h2>
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
              <PivotItem headerText="Case Studies" itemKey="case studies" />
            </Pivot>
          </div>

          <div className='news-scroll'>

            <div className="news-list">

              {
                this.state.NewsFilterdData.length > 0 &&
                this.state.NewsFilterdData.map((item) => {

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
        NewsAnnouncementsData: formattedData,
        NewsFilterdData: formattedData
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

      case "case studies":
        filterdata = filterdata.filter(t => t.NewsCategory === "Case Studies");
        break;

      case "all":
      default:
        filterdata = this.state.NewsAnnouncementsData;
    }

    this.setState({ NewsFilterdData: filterdata });
  }

}
