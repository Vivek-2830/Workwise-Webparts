import * as React from 'react';
import styles from './RelatedNews.module.scss';
import { IRelatedNewsProps } from './IRelatedNewsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';

require('../assets/style.css');

export interface IRelatedNewsState {
  NewsAnnouncementsData: any;
  NewsFilterdData: any;
}

export default class RelatedNews extends React.Component<IRelatedNewsProps, IRelatedNewsState> {

  constructor(props: IRelatedNewsProps, state: IRelatedNewsState) {

    super(props);

    this.state = {
      NewsAnnouncementsData: "",
      NewsFilterdData: ""
    };

  }


  public render(): React.ReactElement<IRelatedNewsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="relatedNews">

        <div className="relatedNews-panel">

          <div className="relatedNews-underline"></div>

          <div className="relatedNews-list">

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
                  <div className="relatedNews-card">
                    <img src={item.NewsPhoto} />

                    <div className="relatedNews-content">
                      <p className="relatedNews-tag">{item.NewsCategory}</p>
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


      </section>
    );
  }

  public async componentDidMount() {
    this.getNewsAnnouncementsData();
  }

  public async getNewsAnnouncementsData() {

    let filterQuery = "";

    if (this.props.category && this.props.category !== "") {
      filterQuery = `NewsCategory eq '${this.props.category}'`;
    }

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
      .top(3)
      .filter(`NewsCategory eq '${this.props.category}' and NewsDate le '${new Date().toISOString()}'` )
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
        const top4 = grouped[category].slice(0, 4);
        topFourPerCategory = [...topFourPerCategory, ...top4];
      });

      this.setState({
        NewsAnnouncementsData: topFourPerCategory,
        NewsFilterdData: topFourPerCategory
      });

    }
  }

}
