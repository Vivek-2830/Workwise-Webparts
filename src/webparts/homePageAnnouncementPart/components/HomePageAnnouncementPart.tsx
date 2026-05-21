import * as React from 'react';
import styles from './HomePageAnnouncementPart.module.scss';
import { IHomePageAnnouncementPartProps } from './IHomePageAnnouncementPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { sp } from '@pnp/sp/presets/all';
import { Announced, DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePageAnnouncementPartState {
  AnnouncementsData: any;
  AddAnnouncementDialog: boolean;
  AddAnnouncementDataDiaolg: boolean;
  Title: any;
  Description: any;
  Source: any;
  Images: any;
  Link: any;
  Videos: any;
  UploadImages: any;
  UploadVideo: any;
  AllAnnouncementDocuments: any;
  file: any;
  previewImage: any;
  EditTitle: any;
  EditDescription: any;
  EditSource: any;
  EditImages: any;
  EditLink: any;
  EditVideos: any;
  EditUploadImages: any;
  EditAnnouncementDataDialog: boolean;
  CurrentAnnouncementDetailsID: any;
  DeleteAnnouncementID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');
require("../assets/fabric.min.css");

const AddAnnouncementDetailsDialogContentProps = {
  title: "Add Announcement Details",
};

const AddAnnouncementDataDialogContentProps = {
  title: "Add Announcements"
};

const UpdateAnnouncementDetailsDialogContentProps = {
  title: "Update Announcement Details"
};

const updatemodelProps = {
  className: "Update-Dialog"
};

const addmodelProps = {
  className: "Add-Dialog"
};

const addmodelProps2 = {
  className: "Add-Data-Dialog"
};

export default class HomePageAnnouncementPart extends React.Component<IHomePageAnnouncementPartProps, IHomePageAnnouncementPartState> {

  constructor(props: IHomePageAnnouncementPartProps, state: IHomePageAnnouncementPartState) {

    super(props);

    this.state = {
      AnnouncementsData: "",
      AddAnnouncementDialog: true,
      AddAnnouncementDataDiaolg: true,
      Title: "",
      Description: "",
      Source: "",
      Images: [],
      Link: "",
      Videos: "",
      UploadImages: [],
      UploadVideo: [],
      previewImage: "",
      AllAnnouncementDocuments: [],
      file: "",
      EditTitle: "",
      EditDescription: "",
      EditSource: "",
      EditImages: [],
      EditLink: "",
      EditVideos: "",
      EditUploadImages: [],
      EditAnnouncementDataDialog: true,
      CurrentAnnouncementDetailsID: "",
      DeleteAnnouncementID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IHomePageAnnouncementPartProps> {
    const {
      description,
      isDarkTheme,
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

        {
          this.state.IsAdmin ?
          <>
            <div className='AddAnnouncemt'>
              <PrimaryButton text='Add Announcements' onClick={() => this.setState({ AddAnnouncementDialog: false })} />
            </div>
          </>
          :
          <>
          </>
        }

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
                        <a href={item.Link.Url} className='anno-read' target="_blank" rel="noopener noreferrer">
                          Read more...
                        </a>
                      </div>
                    )}

                  </div>

                  {/* <div className="welcome-right">
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
                  </div> */}

                  <div className="welcome-right">
                    {
                      item.Images ? (
                        <img src={item.Images} alt="announcement" />
                      ) : item.Videos ? (

                        this.getYouTubeEmbedUrl(item.Videos) ? (
                          // ✅ YouTube iframe
                          <iframe
                            style={{
                              width: "400px",
                              borderRadius: "18px",
                              objectFit: "cover",
                              height: "203px"
                            }}
                            src={this.getYouTubeEmbedUrl(item.Videos)!}
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
                            <source src={item.Videos} type="video/mp4" />
                          </video>
                        )

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

        <Dialog
          hidden={this.state.AddAnnouncementDialog}
          onDismiss={() =>
            this.setState({
              AddAnnouncementDialog: true,
            })
          }
          dialogContentProps={AddAnnouncementDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddAnnouncmentData'>
            <PrimaryButton className='AddAnnounInfo' text='Add Data' onClick={() => this.setState({ AddAnnouncementDataDiaolg: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>Title</th>
                  <th style={{ width: '30%' }}>Description</th>
                  <th style={{ width: '30%' }}>Source</th>
                  <th style={{ width: '15%' }}>Images</th>
                  <th style={{ width: '15%' }}>Link</th>
                  <th style={{ width: '15%' }}>Videos</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.AnnouncementsData.length > 0 &&
                  this.state.AnnouncementsData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Title}</td>
                        <td>{item.Description}</td>
                        <td>{item.Source}</td>
                        <td>
                          {
                            item.Images ? (
                              <img src={item.Images} alt="announcement" style={{ width: "120px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Image"
                            )
                          }
                        </td>
                        <td>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>
                        {/* <td>
                          {
                            item.Videos ? (
                              
                              <a
                                href={item.Videos.Url || item.Videos}
                                target="_blank"
                                rel="noopener noreferrer"
                              >
                                Watch Video
                              </a>
                            ) : (
                              "No Video"
                            )
                          }
                        </td> */}

                        

                       
                          <td>
                            {
                              item.Videos ? (
                                this.getYouTubeEmbedUrl(item.Videos) ? (
                                  <span></span>
                                ) : (
                                  <a
                                    href={item.Videos}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                  >
                                    Watch Video
                                  </a>
                                )
                              ) : (
                                "No Video"
                              )
                            }
                          </td>
                          <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditAnnouncementDataDialog: false, CurrentAnnouncementDetailsID: item.ID }, () => this.EditAnnouncementInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteAnnouncementInfo(item.ID)}
                            />

                          </div>
                        </td>

                      </tr>
                    );
                  })
                }

              </tbody>
            </table>
          </div>

        </Dialog>

        <Dialog
          hidden={this.state.AddAnnouncementDataDiaolg}
          onDismiss={() =>
            this.setState({
              AddAnnouncementDataDiaolg: true,
              Title: "",
              Description: "",
              Source: "",
              // Images: [],
              Link: "",
              Videos: "",
              //UploadImages: [],
              //UploadVideo: ""
            })
          }
          dialogContentProps={AddAnnouncementDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={1100}
        >
          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Announcement Title'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Title: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Description'
                  type='text'
                  multiline rows={3}
                  onChange={(value) =>
                    this.setState({ Description: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Source'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Source: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <label><b>Upload Image</b></label><br />

                <input
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleImageChange(e)}
                />

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Link'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Link: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label="Video"
                  type="text"
                  value={this.state.Videos}
                  onChange={this.handleVideoChange}
                  placeholder="Enter YouTube URL or text"
                />

                {/* 🔹 Inline Preview */}
                {this.state.Videos && (
                  this.getYouTubeEmbedUrl(this.state.Videos) ? (
                    <iframe
                      width="100%"
                      height="200"
                      src={this.getYouTubeEmbedUrl(this.state.Videos)!}
                      title="YouTube video player"
                      frameBorder="0"
                      allow="autoplay; encrypted-media"
                      allowFullScreen
                      loading="lazy"
                      style={{ marginTop: "10px" }}
                    />
                  ) : (
                    <p style={{ marginTop: "10px" }}>
                      {this.state.Videos}
                    </p>
                  )
                )}
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddAnnouncementInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddAnnouncementDataDiaolg: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditAnnouncementDataDialog}
          onDismiss={() =>
            this.setState({
              EditAnnouncementDataDialog: true,
              EditTitle: "",
              EditDescription: "",
              EditSource: "",
              EditLink: "",
              EditVideos: "",
              EditImages: [],
              EditUploadImages: []
            })
          }
          dialogContentProps={UpdateAnnouncementDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={1100}
        >
          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Announcement Title'
                  type='text'
                  value={this.state.EditTitle}
                  onChange={(value) =>
                    this.setState({ EditTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Description'
                  type='text'
                  multiline rows={3}
                  value={this.state.EditDescription}
                  onChange={(value) =>
                    this.setState({ EditDescription: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Source'
                  type='text'
                  value={this.state.EditSource}
                  onChange={(value) =>
                    this.setState({ EditSource: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <label><b>Upload Image</b></label><br />

                <input
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleUpdateImageChange(e)}
                />

                {
                  this.state.EditUploadImages && (
                    <div className="Attached-img">

                      {/* ✅ Handle BOTH string + file */}
                      <p>
                        {
                          typeof this.state.EditUploadImages === "string"
                            ? this.state.EditUploadImages.split('/').pop()
                            : this.state.EditUploadImages[0]?.name
                        }
                      </p>

                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadImages: "" })}
                      />
                    </div>
                  )
                }

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div>
                <TextField
                  label='Link'
                  type='text'
                  value={this.state.EditLink}
                  onChange={(value) =>
                    this.setState({ EditLink: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div>
                <TextField
                  label='Video'
                  type='text'
                  value={this.state.EditVideos}
                  onChange={(value) =>
                    this.setState({ EditVideos: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateAnnouncementDetails(this.state.CurrentAnnouncementDetailsID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditAnnouncementDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>


          </div>

        </Dialog>

      </section>
    );
  }

  public async componentDidMount() {
    this.getannouncement();
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
            Link: item.Link
          });

        });

        this.setState({
          AnnouncementsData: AllData
        });
        console.log(this.state.AnnouncementsData);
      }

    } catch (error) {
      console.log("Error Fetching details :", error);
    }
  }

  public async AddAnnouncementInfo() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const announcement = await sp.web.lists.getByTitle("Announcements").items.add({
        Title: this.state.Title,
        Description: this.state.Description,
        Source: this.state.Source,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null,

        Videos: this.state.Videos
          ? {
            Url: this.state.Videos,
            Description: "Video"
          }
          : null
      });

      if (this.state.UploadImages && this.state.UploadImages.length > 0) {

        const file = this.state.UploadImages[0];

        await sp.web.lists
          .getByTitle("Announcements")
          .items.getById(announcement.data.Id)
          .attachmentFiles.add(file.name, file);
      }


      // this.setState({ AnnouncementsData: announcement });
      this.setState({ AddAnnouncementDataDiaolg: true });
      this.getannouncement();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadImages: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  }

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadImages: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }
  
  }

  public async EditAnnouncementInfo(ID) {
    let EditAnnouncement = this.state.AnnouncementsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditAnnouncement);
    this.setState({
      EditTitle: EditAnnouncement[0].Title,
      EditDescription: EditAnnouncement[0].Description,
      EditSource: EditAnnouncement[0].Source,
      EditLink: EditAnnouncement[0].Link.Url,
      EditVideos: EditAnnouncement[0].Videos,
      EditUploadImages: EditAnnouncement[0].Images,

    });
  }

  public async UpdateAnnouncementDetails(CurrentAnnouncementDetailsID) {
    try {
      const updateannouncement: any = {
        Title: this.state.EditTitle,
        Description: this.state.EditDescription,
        Source: this.state.EditSource,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null,
        Videos: this.state.EditVideos ? {
          Url: this.state.EditVideos,
          Description: "Video"
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Announcements").items.getById(CurrentAnnouncementDetailsID).update(updateannouncement);

      // if (this.state.EditUploadImages && this.state.EditUploadImages.length > 0) {
      //   const file = this.state.EditUploadImages[0];

      //   const itemRef = sp.web.lists
      //     .getByTitle("Announcements")
      //     .items.getById(CurrentAnnouncementDetailsID);

      //   const attachments = await itemRef.attachmentFiles();

      //   for (let att of attachments) {
      //     await itemRef.attachmentFiles.getByName(att.FileName).delete();
      //   }

      //   await itemRef.attachmentFiles.add(file.name, file);
      // }

      if (Array.isArray(this.state.EditUploadImages) && this.state.EditUploadImages.length > 0) {

        const file = this.state.EditUploadImages[0];
  
        const itemRef = sp.web.lists
          .getByTitle("Announcements")
          .items.getById(CurrentAnnouncementDetailsID);
  
        // delete old attachments
        const attachments = await itemRef.attachmentFiles();
  
        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }
  
        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }
  

      this.setState({ EditAnnouncementDataDialog: true });
      this.getannouncement();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteAnnouncementInfo(DeleteAnnouncementID) {
    const deleteinfo = await sp.web.lists.getByTitle("Announcements").items.getById(DeleteAnnouncementID).delete();
    this.setState({ AnnouncementsData: deleteinfo });
    this.getannouncement();
  }

  // private getYouTubeEmbedUrl = (url: string): string | null => {

  //   if (!url) return null;
  
  //   const regExp =
  //     /(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/shorts\/)([^&\n?#]+)/;
  
  //   const match = url.match(regExp);
  
  //   if (match && match[1]) {
  //     return `https://www.youtube.com/embed/${match[1]}?autoplay=1&mute=1&controls=0&rel=0&modestbranding=1`;
  //   }
  
  //   return null;
  // };

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
    this.setState({ Videos: value || "" });
  }

}
