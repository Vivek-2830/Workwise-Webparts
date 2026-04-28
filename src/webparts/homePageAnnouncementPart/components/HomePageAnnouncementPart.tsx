import * as React from 'react';
import styles from './HomePageAnnouncementPart.module.scss';
import { IHomePageAnnouncementPartProps } from './IHomePageAnnouncementPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { sp } from '@pnp/sp/presets/all';
import { Announced, Dialog, PrimaryButton } from 'office-ui-fabric-react';

export interface IHomePageAnnouncementPartState {
  AnnouncementsData: any;
  AddAnnouncementDialog: boolean;
  Title: any;
  Description: any;
  Source: any;
  Images: any;
  Link: any;
  Videos: any;
  UploadImages: any;
  UploadVideo: any;
  AllAnnouncementDocuments: any;
}

require('../assets/style.css');

const AddAnnouncementDetailsDialogContentProps = {
  title: "Add Announcement Details",
};

const addmodelProps = {
  className: "Add-Dialog"
};

export default class HomePageAnnouncementPart extends React.Component<IHomePageAnnouncementPartProps, IHomePageAnnouncementPartState> {

  constructor(props: IHomePageAnnouncementPartProps, state: IHomePageAnnouncementPartState) {

    super(props);

    this.state = {
      AnnouncementsData: "",
      AddAnnouncementDialog: true,
      Title: "",
      Description: "",
      Source: "",
      Images: [],
      Link: "",
      Videos: [],
      UploadImages: [],
      UploadVideo: [],
      AllAnnouncementDocuments: []
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

        <div className='AddAnnouncemt'>
          <PrimaryButton text='Add Announcements' onClick={() => this.setState({ AddAnnouncementDialog: false })} />
        </div>

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
                        <video autoPlay muted loop playsInline controls style={{ width: "400px", borderRadius: "18px", objectFit: "cover", height: "203px" }} >
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

  public async AddAnnouncementData() {
    try {
  
      if (this.state.Title.length === 0) {
        alert("Title is required");
        return;
      }
  
      let imageColumnValue: any = null;
  
      if (this.state.UploadImages && this.state.UploadImages.length > 0) {
  
        const fileObj = this.state.UploadImages[0];
        const file = fileObj.content; // ✅ FIX
  
        const uploadResult = await sp.web
          .getFolderByServerRelativeUrl("SiteAssets")
          .files.add(file.name, file, true);
  
        const fileUrl = uploadResult.data.ServerRelativeUrl;
  
        // ✅ FIX (NO stringify)
        imageColumnValue = {
          fileName: file.name,
          serverRelativeUrl: fileUrl
        };
      }
  
      const itemAddResult = await sp.web.lists
        .getByTitle("Announcements")
        .items.add({
          Title: this.state.Title,
          Description: this.state.Description,
          Source: this.state.Source,
          Link: this.state.Link,
  
          Images: imageColumnValue,
  
          Videos: this.state.UploadVideo
            ? {
                Url: this.state.UploadVideo,
                Description: "Video"
              }
            : null
        });
  
      const itemId = itemAddResult.data.Id;
  
      // Attachments (optional)
      if (this.state.UploadImages && this.state.UploadImages.length > 0) {
        for (const fileObj of this.state.UploadImages) {
          await sp.web.lists
            .getByTitle("Announcements")
            .items.getById(itemId)
            .attachmentFiles.add(fileObj.name, fileObj.content);
        }
      }
  
      alert("Announcement added successfully!");
  
      this.setState({ AddAnnouncementDialog: true });
      this.getannouncement();
  
    } catch (error) {
      console.error("Error adding announcement:", error);
    }
  }
  
}
