import * as React from 'react';
import styles from './HomePageAnnouncementPart.module.scss';
import { IHomePageAnnouncementPartProps } from './IHomePageAnnouncementPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { sp } from '@pnp/sp/presets/all';

export interface IHomePageAnnouncementPartState {
  AnnouncementsData: any;
}

require('../assets/style.css');


export default class HomePageAnnouncementPart extends React.Component<IHomePageAnnouncementPartProps, IHomePageAnnouncementPartState> {

  constructor(props: IHomePageAnnouncementPartProps, state: IHomePageAnnouncementPartState) {

    super(props);

    this.state = {
      AnnouncementsData: ""
    };

  }

  public render(): React.ReactElement<IHomePageAnnouncementPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    var settings = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 1,
      slidesToScroll: 1,
      autoplaySpeed: 5000,
      autoplay: true,
      cssEase: "linear",
      fade: true,
      // nextArrow: <SampleNextArrow />,
      // prevArrow: <SamplePrevArrow />
    };

    return (
      <section className="homePageAnnouncementPart">

        <Slider {...settings}>

          {
            this.state.AnnouncementsData.length > 0 &&
            this.state.AnnouncementsData.map((item) => {
              // let imagePath = "";
              // let ImageInfo = JSON.parse(item.Images);
              // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
              //   imagePath = ImageInfo["serverRelativeUrl"];
              // }
              // else {
              //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Announcements/Attachments/${item.ID}/${ImageInfo.fileName}`;
              // }
              return (

                <div className="welcome-container">

                  <div className="welcome-left">
                    <p className="welcome-user">{item.Title}</p>

                    <h1>
                      {item.Description}
                    </h1>

                    <p className="welcome-desc">
                      {item.Source}
                    </p>

                    {item.Link && (
                      <div className='announcement-read'>
                        <a href={item.Link} className='anno-read' target="_blank" rel="noopener noreferrer">
                          Read more...
                        </a>
                      </div>
                    )}

                  </div>

                  <div className="welcome-right">
                    {
                      item.Images ? (
                        <img src={item.Images} alt="announcement" />
                      ) : item.Videos ? (
                        <video autoPlay muted loop playsInline controls style={{ width: "420px", borderRadius: "18px", objectFit: "cover", height: "300px" }} >
                          <source src={item.Videos} type="video/mp4" />
                          Your browser does not support the video tag.
                        </video>
                      ) : (
                        <img src={require("../assets/Rectangle1.png")} alt="default" />
                      )
                    }
                  </div>

                </div>

              );
            })
          }

        </Slider>

      </section>
    );
  }

  public async componentDidMount() {
    this.getannouncement();
  }

  public async getannouncement(): Promise<void> {
    try {

      const items: any[] = await sp.web.lists
        .getByTitle("Announcements")
        .items
        .select(
          "ID",
          "Title",
          "Description",
          "Source",
          "Link",
          "Videos",
          "SlideOrder",
          "AttachmentFiles"
        ).orderBy("SlideOrder", true)
        .expand("AttachmentFiles")
        .get();

      let AllData: any[] = [];

      if (items && items.length > 0) {

        items.forEach((item: any) => {

          let imageUrl: string = "";
          let videoUrl: string = "";

          /* ===========================
             CHECK ATTACHMENTS FIRST
          ============================ */

          if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {

            const file = item.AttachmentFiles[0];
            const fileName = file.FileName.toLowerCase();

            if (fileName.match(/\.(jpg|jpeg|png|gif)$/)) {
              imageUrl = file.ServerRelativeUrl;
            }
            else if (fileName.match(/\.(mp4|webm|ogg|mov|avi|m4v)$/)) {
              videoUrl = file.ServerRelativeUrl;
            }
          }

          /* ===========================
             CHECK HYPERLINK VIDEO FIELD
          ============================ */

          let videoColumnUrl: string = "";

          if (item.Videos) {

            // Case 1: Hyperlink field object
            if (typeof item.Videos === "object" && item.Videos.Url) {
              videoColumnUrl = item.Videos.Url;
            }

            // Case 2: Direct string
            else if (typeof item.Videos === "string") {
              videoColumnUrl = item.Videos;
            }
          }

          /* ===========================
             PUSH CLEAN DATA
          ============================ */

          AllData.push({
            ID: item.ID || "",
            Title: item.Title || "",
            Description: item.Description || "",
            Source: item.Source || "",
            Images: imageUrl,
            Videos: videoUrl || videoColumnUrl,
            Link: item.Link ? (item.Link.Url ? item.Link.Url : item.Link) : ""
          });

        });

        this.setState({
          AnnouncementsData: AllData
        });
      }

    } catch (error) {
      console.log("Error Fetching details :", error);
    }
  }


}
