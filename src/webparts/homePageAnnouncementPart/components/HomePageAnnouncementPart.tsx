import * as React from 'react';
import styles from './HomePageAnnouncementPart.module.scss';
import { IHomePageAnnouncementPartProps } from './IHomePageAnnouncementPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { sp } from '@pnp/sp/presets/all';
import { Announced, DefaultButton, Dialog, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

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
  previewImage: any
}

require('../assets/style.css');
require("../assets/fabric.min.css");

const AddAnnouncementDetailsDialogContentProps = {
  title: "Add Announcement Details",
};

const AddAnnouncementDataDialogContentProps = {
  title: "Add Announcements"
}

const addmodelProps = {
  className: "Add-Dialog"
};

const addmodelProps2 = {
  className: "Add-Data-Dialog"
}

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
      file: ""
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
                        <a href={item.Link.Url} className='anno-read' target="_blank" rel="noopener noreferrer">
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

        <Dialog
          hidden={this.state.AddAnnouncementDialog}
          onDismiss={() =>
            this.setState({
              AddAnnouncementDialog: true,
            })
          }
          dialogContentProps={AddAnnouncementDetailsDialogContentProps}
          modalProps={addmodelProps}
          maxWidth={1500}
        >

          <div className='AddAnnouncmentData'>
            <PrimaryButton className='AddAnnounInfo' text='Add Data' onClick={() => this.setState({ AddAnnouncementDataDiaolg : false })}/>
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Description</th>
                  <th>Source</th>
                  <th>Images</th>
                  <th>Link</th>
                  <th>Videos</th>
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
                          <td>
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
                          </td>

                          <td>
                            <div style={{ display: "flex", gap: "8px" }}>

                              <IconButton
                                iconProps={{ iconName: "Edit" }}
                                title="Edit"
                                ariaLabel="Edit"
                                
                              />
                             
                              <IconButton
                                iconProps={{ iconName: "Delete" }}
                                title="Delete"
                                ariaLabel="Delete"
                               
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
              Videos:"",
              //UploadImages: [],
              //UploadVideo: ""
            })
          }
          dialogContentProps={AddAnnouncementDataDialogContentProps}
          modalProps={addmodelProps2}
          maxWidth={1500}
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
                  label='Video'
                  type='text'
                  onChange={(value) => 
                    this.setState({ Videos: value.target["value"] })
                  }
                />
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

  // public async AddAnnouncementData() {
  //   try {

  //     if (this.state.Title.length === 0) {
  //       alert("Title is required");
  //       return;
  //     }

  //     let imageColumnValue: any = null;

  //     if (this.state.UploadImages && this.state.UploadImages.length > 0) {

  //       const fileObj = this.state.UploadImages[0];
  //       const file = fileObj.content; // ✅ FIX

  //       const uploadResult = await sp.web
  //         .getFolderByServerRelativeUrl("SiteAssets")
  //         .files.add(file.name, file, true);

  //       const fileUrl = uploadResult.data.ServerRelativeUrl;

  //       // ✅ FIX (NO stringify)
  //       imageColumnValue = {
  //         fileName: file.name,
  //         serverRelativeUrl: fileUrl
  //       };
  //     }

  //     const itemAddResult = await sp.web.lists
  //       .getByTitle("Announcements")
  //       .items.add({
  //         Title: this.state.Title,
  //         Description: this.state.Description,
  //         Source: this.state.Source,
  //         Link: this.state.Link,

  //         Images: imageColumnValue,

  //         Videos: this.state.Videos
  //           ? {
  //             Url: this.state.Videos,
  //             Description: "Video"
  //           }
  //           : null
  //       });

  //     const itemId = itemAddResult.data.Id;

  //     // Attachments (optional)
  //     if (this.state.UploadImages && this.state.UploadImages.length > 0) {
  //       for (const fileObj of this.state.UploadImages) {
  //         await sp.web.lists
  //           .getByTitle("Announcements")
  //           .items.getById(itemId)
  //           .attachmentFiles.add(fileObj.name, fileObj.content);
  //       }
  //     }

  //     alert("Announcement added successfully!");

  //     this.setState({ AddAnnouncementDataDiaolg: true });
  //     this.getannouncement();

  //   } catch (error) {
  //     console.error("Error adding announcement:", error);
  //   }
  // }

  public async AddAnnouncementInfo() {
    if(this.state.Title.length == 0) {
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
  };

  // public async AddAnnouncementData() {
  //   try {
  
  //     if (!this.state.Title || this.state.Title.trim().length === 0) {
  //       alert("Please enter Title");
  //       return;
  //     }
  
  //     let imageJson: any = null;
  
  //     if (this.state.Images) {
  
  //       const file = this.state.Images;
  
  //       const uploadResult = await sp.web
  //         .getFolderByServerRelativePath("SiteAssets")
  //         .files.addUsingPath(file.name, file, { Overwrite: true });
  
  //       const fileUrl = uploadResult.data.ServerRelativeUrl;
  
  //       // ✅ NO JSON.stringify HERE
  //       imageJson = {
  //         type: "thumbnail",
  //         fileName: file.name,
  //         serverUrl: window.location.origin,
  //         serverRelativeUrl: fileUrl
  //       };
  //     }
  
  //    const annodata =  await sp.web.lists.getByTitle("Announcements").items.add({
  //       Title: this.state.Title,
  //       Description: this.state.Description,
  //       Source: this.state.Source,
  //       Link: this.state.Link,
  
  //       // ⚠️ FIX for Hyperlink column
  //       Videos: this.state.Videos
  //         ? {
  //             Url: this.state.Videos,
  //             Description: "Video"
  //           }
  //         : null,
  
  //       Images: imageJson
  //     });
      
  //     this.setState({ AnnouncementsData : annodata });
  //     this.setState({ AddAnnouncementDataDiaolg: true });
  //     this.getannouncement();
  
  //   } catch (error) {
  //     console.error("Error:", error);
  //     alert("Something went wrong");
  //   }
  // }

}
