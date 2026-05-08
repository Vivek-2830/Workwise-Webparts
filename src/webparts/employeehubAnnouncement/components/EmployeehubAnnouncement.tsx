import * as React from 'react';
import styles from './EmployeehubAnnouncement.module.scss';
import { IEmployeehubAnnouncementProps } from './IEmployeehubAnnouncementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IEmployeehubAnnouncementState {
  EmployeeAnnouncementsData: any;
  AddEmpAnnouncementDialog: boolean;
  AddEmpAnnouncementDataDialog: boolean;
  Title: any;
  Description: any;
  Source: any;
  Images: any;
  link: any;
  Videos: any;
  EmpVideos:any;
  UploadImage: any;
  previewImage: any;
  EditTitle: any;
  EditDescription: any;
  EditSource: any;
  EditImage: any;
  Editlink: any;
  EditVideos: any;
  EditUploadImage: any;
  EditEmpAnnouncementDataDialog: boolean;
  CurrentEmpAnnouncementDataID: any;
  DeleteEmpAnnouncementDataID: any;
}

require('../assets/style.css');

const AddEmpAnnouncementDetailsDialogContentProps = {
  title: "Add Announcement Details",
};

const AddEmpAnnouncementDataDialogContentProps = {
  title: "Add Announcements"
}

const UpdateEmpAnnouncementDetailsDialogContentProps = {
  title: "Update Announcement Details"
}

const updatemodelProps = {
  className: "Update-Dialog"
};

const addmodelProps = {
  className: "Add-Dialog"
};

const addmodelProps2 = {
  className: "Add-Data-Dialog"
}


export default class EmployeehubAnnouncement extends React.Component<IEmployeehubAnnouncementProps, IEmployeehubAnnouncementState> {

  constructor(props: IEmployeehubAnnouncementProps, state: IEmployeehubAnnouncementState) {

    super(props);

    this.state = {
      EmployeeAnnouncementsData: "",
      AddEmpAnnouncementDialog: true,
      AddEmpAnnouncementDataDialog: true,
      Title: "",
      Description: "",
      Source: "",
      Images: [],
      link: "",
      Videos: [],
      UploadImage: [],
      previewImage: "",
      EditTitle: "",
      EditDescription: "",
      EditSource: "",
      EditImage: [],
      Editlink: "",
      EditVideos: "",
      EditUploadImage: [],
      EditEmpAnnouncementDataDialog: true,
      CurrentEmpAnnouncementDataID: "",
      DeleteEmpAnnouncementDataID: "",
      EmpVideos: ""
    }

  }


  public render(): React.ReactElement<IEmployeehubAnnouncementProps> {
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
      <section className="employeehubAnnouncement">

        <div className='AddAnnouncemt'>
          <PrimaryButton text='Add Announcements' onClick={() => this.setState({ AddEmpAnnouncementDialog: false })} />
        </div>

        <Slider {...settings}>

          {
            this.state.EmployeeAnnouncementsData.length > 0 &&
            this.state.EmployeeAnnouncementsData.map((item) => {

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

                    <div className='announcement-read'>
                      <a href={item.link} className='anno-read'>Read more...</a>
                    </div>

                  </div>

                  <div className="welcome-right">
                    {
                      item.Images ? (
                        <img src={item.Images} alt="announcement" />
                      ) : item.Videos ? (

                       item.Videos ? (
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
          hidden={this.state.AddEmpAnnouncementDialog}
          onDismiss={() =>
            this.setState({
              AddEmpAnnouncementDialog: true,
            })
          }
          dialogContentProps={AddEmpAnnouncementDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddAnnouncmentData'>
            <PrimaryButton className='AddAnnounInfo' text='Add Data' onClick={() => this.setState({ AddEmpAnnouncementDataDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>Title</th>
                  <th style={{ width: '30%' }}>Description</th>
                  <th style={{ width: '30%' }}>Source</th>
                  <th style={{ width: '15%' }}>Images</th>
                  <th style={{ width: '15%' }}>link</th>
                  <th style={{ width: '15%' }}>Videos</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.EmployeeAnnouncementsData.length > 0 &&
                  this.state.EmployeeAnnouncementsData.map((item) => {
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
                          <a href={item.link} target="_blank" rel="noopener noreferrer">{item.link.Description}</a>
                        </td>

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
                              onClick={() => this.setState({ EditEmpAnnouncementDataDialog: false, CurrentEmpAnnouncementDataID: item.ID }, () => this.EditEmpAnnouncement(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteEmpAnnouncementInfo(item.ID)}
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
          hidden={this.state.AddEmpAnnouncementDataDialog}
          onDismiss={() =>
            this.setState({
              AddEmpAnnouncementDataDialog: true,
              Title: "",
              Description: "",
              Source: "",
              Images: [],
              link: "",
              Videos: [],
              UploadImage: []
            })
          }
          dialogContentProps={AddEmpAnnouncementDataDialogContentProps}
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
                    this.setState({ link: value.target["value"] })
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
                    onClick={() => this.AddEmpAnnouncementInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddEmpAnnouncementDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditEmpAnnouncementDataDialog}
          onDismiss={() =>
            this.setState({
              EditEmpAnnouncementDataDialog: true,
              EditTitle: "",
              EditDescription: "",
              EditSource: "",
              Editlink: "",
              EditVideos: "",
              EditImage: [],
              EditUploadImage: []
            })
          }
          dialogContentProps={UpdateEmpAnnouncementDetailsDialogContentProps}
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
                  this.state.EditUploadImage && (
                    <div className="Attached-img">

                      {/* ✅ Handle BOTH string + file */}
                      <p>
                        {
                          typeof this.state.EditUploadImage === "string"
                            ? this.state.EditUploadImage.split('/').pop()
                            : this.state.EditUploadImage[0]?.name
                        }
                      </p>

                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadImage: "" })}
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
                  value={this.state.Editlink}
                  onChange={(value) =>
                    this.setState({ Editlink: value.target["value"] })
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
                    onClick={() => this.UpdateEmpAnnouncementDetails(this.state.CurrentEmpAnnouncementDataID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditEmpAnnouncementDataDialog: true })
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
    this.getEmpannouncement();
  }

  public async getEmpannouncement(): Promise<void> {
    try {

      const items: any[] = await sp.web.lists
        .getByTitle("Employee Announcements")
        .items
        .select(
          "ID",
          "Title",
          "Description",
          "Source",
          "link",
          "Videos",
          "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .get();

      let AllData: any[] = [];

      if (items && items.length > 0) {

        items.forEach((item: any) => {

          let imageUrl: string = "";
          let videoUrl: string = "";

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
            link: item.link ? item.link.Url : ""
          });

        });

        this.setState({
          EmployeeAnnouncementsData: AllData
        });
      }

    } catch (error) {
      console.log("Error Fetching details :", error);
    }
  }

  public async AddEmpAnnouncementInfo() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const empannouncement = await sp.web.lists.getByTitle("Employee Announcements").items.add({
        Title: this.state.Title,
        Description: this.state.Description,
        Source: this.state.Source,
        link: this.state.link
          ? {
            Url: this.state.link,
            Description: this.state.link
          }
          : null,

        Videos: this.state.Videos
          ? {
            Url: this.state.Videos,
            Description: "Video"
          }
          : null
      });

      if (this.state.UploadImage && this.state.UploadImage.length > 0) {

        const file = this.state.UploadImage[0];

        await sp.web.lists
          .getByTitle("Employee Announcements")
          .items.getById(empannouncement.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddEmpAnnouncementDataDialog: true });
      this.getEmpannouncement();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadImage: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  };

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadImage: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditEmpAnnouncement(ID) {
    let EditEmpAnnouncement = this.state.EmployeeAnnouncementsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditEmpAnnouncement);
    this.setState({
      EditTitle: EditEmpAnnouncement[0].Title,
      EditDescription: EditEmpAnnouncement[0].Description,
      EditSource: EditEmpAnnouncement[0].Source,
      Editlink: EditEmpAnnouncement[0].link.Url,
      EditVideos: EditEmpAnnouncement[0].Videos,
      EditUploadImage: EditEmpAnnouncement[0].Images,

    });
  }

  public async UpdateEmpAnnouncementDetails(CurrentEmpAnnouncementDataID) {
    try {
      const updateempannouncement: any = {
        Title: this.state.EditTitle,
        Description: this.state.EditDescription,
        Source: this.state.EditSource,
        link: this.state.Editlink ? {
          Url: this.state.Editlink,
          Description: this.state.Editlink
        } : null,
        Videos: this.state.EditVideos ? {
          Url: this.state.EditVideos,
          Description: "Video"
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Employee Announcements").items.getById(CurrentEmpAnnouncementDataID).update(updateempannouncement);

      // if (this.state.EditUploadImages && this.state.EditUploadImages.length > 0) {
      //   const file = this.state.EditUploadImages[0];

      //   const itemRef = sp.web.lists
      //     .getByTitle("Announcements")
      //     .items.getById(CurrentEmpAnnouncementDataID);

      //   const attachments = await itemRef.attachmentFiles();

      //   for (let att of attachments) {
      //     await itemRef.attachmentFiles.getByName(att.FileName).delete();
      //   }

      //   await itemRef.attachmentFiles.add(file.name, file);
      // }

      if (Array.isArray(this.state.EditUploadImage) && this.state.EditUploadImage.length > 0) {

        const file = this.state.EditUploadImage[0];

        const itemRef = sp.web.lists
          .getByTitle("Employee Announcements")
          .items.getById(CurrentEmpAnnouncementDataID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditEmpAnnouncementDataDialog: true });
      this.getEmpannouncement();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteEmpAnnouncementInfo(DeleteEmpAnnouncementDataID) {
    const deleteinfo = await sp.web.lists.getByTitle("Employee Announcements").items.getById(DeleteEmpAnnouncementDataID).delete();
    this.setState({ EmployeeAnnouncementsData: deleteinfo });
    this.getEmpannouncement();
  }

  private getYouTubeEmbedUrl = (url: string): string | null => {
    if (!url) return null;

    const regExp =
      /(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/shorts\/)([^&\n?#]+)/;

    const match = url.match(regExp);

    if (match && match[1]) {
      return `https://www.youtube.com/embed/${match[1]}?autoplay=1&mute=1&controls=0&rel=0&modestbranding=1`;
    }

    return null;
  };

  private handleVideoChange = (e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string) => {
    this.setState({ EmpVideos: value || "" });
  };

}
