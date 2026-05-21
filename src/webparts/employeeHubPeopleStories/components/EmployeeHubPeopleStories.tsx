import * as React from 'react';
import styles from './EmployeeHubPeopleStories.module.scss';
import { IEmployeeHubPeopleStoriesProps } from './IEmployeeHubPeopleStoriesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";

export interface IEmployeeHubPeopleStoriesState {
  Title: any;
  Video: any;
  PeopleStoriesData: any;
  PeopleStoriesDetailsDialog: boolean;
  AddPeopleStoriesDataDialog: boolean;
  EditTitle: any;
  EditVideo: any;
  EditPeopleStoriesDialog: boolean;
  CurrentPeopleStoriesItemID: any;
  DeletePeopleStoriesItemID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

export default class EmployeeHubPeopleStories extends React.Component<IEmployeeHubPeopleStoriesProps, IEmployeeHubPeopleStoriesState> {

  constructor(props: IEmployeeHubPeopleStoriesProps, state: IEmployeeHubPeopleStoriesState) {

    super(props);

    this.state = {
      Title: "",
      Video: "",
      PeopleStoriesData: "",
      PeopleStoriesDetailsDialog: true,
      AddPeopleStoriesDataDialog: true,
      EditTitle: "",
      EditVideo: "",
      EditPeopleStoriesDialog: true,
      CurrentPeopleStoriesItemID: "",
      DeletePeopleStoriesItemID: "",
      IsAdmin: false,
      CurrentUserEmail: "",
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
      <section className="employeeHubPeopleStories">

        <section className="stories-section">
          <h2 className="stories-title">Investing in People: Stories of Growth</h2>


          <Slider {...settings}>
            {
              this.state.PeopleStoriesData.length > 0 &&
              this.state.PeopleStoriesData.map((item) => {
                return (
                  <div className="stories-grid">

                    <div className="story-card">
                      <div className="story-thumbnail">
                        {
                          item.Video ? (
                            this.getYouTubeEmbedUrl(item.Video) ? (
                              // ✅ YouTube iframe
                              <iframe
                                style={{
                                  width: "400px",
                                  borderRadius: "18px",
                                  objectFit: "cover",
                                  height: "203px"
                                }}
                                src={this.getYouTubeEmbedUrl(item.Video)!}
                                title="YouTube video player"
                                frameBorder="0"
                                allow="autoplay; encrypted-media"
                                allowFullScreen
                                loading="lazy"
                              />
                            ) : (
                              // ✅ Normal video file (mp4 etc.)
                              <video
                                autoPlay
                                muted
                                loop
                                playsInline
                                controls
                                style={{
                                  width: "400px",
                                  borderRadius: "18px",
                                  objectFit: "cover",
                                  height: "203px"
                                }}
                              >
                                <source src={item.Video} type="video/mp4" />
                              </video>
                            )
                          ) :
                            (
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
                            )
                        }

                      </div>
                      <div className="story-content">
                        <div className="story-title">
                         {item.Title}
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })
            }
          </Slider>

          {/* <div className="stories-grid">

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
          </div> */}

          <div className="accent-line"></div>
        </section>

      </section>
    );
  }

  public async componentDidMount() {
    this.getPeopleStoriesData();
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


  public async getPeopleStoriesData(): Promise<void> {
    try {
  
      const items: any[] = await sp.web.lists
        .getByTitle("People Stories")
        .items
        .select(
          "ID",
          "Title",
          "Video",
          "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .get();
  
      let AllData: any[] = [];
  
      if (items && items.length > 0) {
  
        items.forEach((item: any) => {
  
          let videoUrl: string = "";
  
          /* ===========================
             CHECK ATTACHMENT VIDEO
          ============================ */
  
          if (
            item.AttachmentFiles &&
            item.AttachmentFiles.length > 0
          ) {
  
            const file = item.AttachmentFiles[0];
            const fileName = file.FileName.toLowerCase();
  
            if (fileName.match(/\.(mp4|webm|ogg|mov|avi|m4v)$/)) {
              videoUrl = file.ServerRelativeUrl;
            }
          }
  
          /* ===========================
             CHECK HYPERLINK VIDEO FIELD
          ============================ */
  
          let videoColumnUrl: string = "";
  
          if (item.Video) {
  
            // Hyperlink field object
            if (
              typeof item.Video === "object" &&
              item.Video.Url
            ) {
              videoColumnUrl = item.Video.Url;
            }
  
            // Direct string
            else if (typeof item.Video === "string") {
              videoColumnUrl = item.Video;
            }
          }
  
          /* ===========================
             PUSH CLEAN DATA
          ============================ */
  
          AllData.push({
            ID: item.ID || "",
            Title: item.Title || "",
            Video: videoUrl || videoColumnUrl
          });
  
        });
  
        this.setState({
          PeopleStoriesData: AllData
        });
  
        console.log(AllData);
      }
  
    } catch (error) {
      console.log("Error Fetching details :", error);
    }
  }

  private getYouTubeEmbedUrl = (url: any): string | null => {
    if (!url) return null;

    // Handle SharePoint Hyperlink field object:
    // { Url: "...", Description: "..." }
    if (typeof url === "object" && url.Url) {
      url = url.Url;
    }

    // Ensure url is a string
    if (typeof url !== "string") {
      return null;
    }

    const regExp =
      /(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/shorts\/)([^&\n?#]+)/;

    const match = url.match(regExp);

    if (match && match[1]) {
      return `https://www.youtube.com/embed/${match[1]}?autoplay=1&mute=1&controls=0&rel=0&modestbranding=1&loop=1&playlist=${match[1]}`;
    }

    return null;
  }

  private handleVideoChange = (e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
    this.setState({ Video: value || "" });
  }

}
