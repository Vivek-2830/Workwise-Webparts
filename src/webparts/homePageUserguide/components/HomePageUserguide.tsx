import * as React from 'react';
import styles from './HomePageUserguide.module.scss';
import { IHomePageUserguideProps } from './IHomePageUserguideProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { Announced, DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePageUserguideState {
  EssentialLearningsData: any;
  AddUserGuideDialog: boolean;
  AddUserGuideDataDialog: boolean;
  Title: any;
  EssentialDescription: any;
  Images: any;
  link: any;
  UploadImages: any;
  EditUserGuideDataDialog: boolean;
  EditTitle: any;
  EditEssentialDescription: any;
  EditImages: any;
  Editlink: any;
  EditUploadImages: any;
  CurrentUserguideDetailsID: any;
  DeleteUserguideDetailsID: any;
  previewImage: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddUserguideDetailsDialogContentProps = {
  title: "Add Userguide Details",
};

const AddAnnouncementDataDialogContentProps = {
  title: "Add Userguide"
}

const UpdateuserguideDataDialogContentProps = {
  title: "Update Userguide Details"
}

const updatemodelProps = {
  className: "Update-Dialog"
};

const addmodelProps = {
  className: "Add-Dialog"
};

const addmodelProps2 = {
  className: "Add-Data-Dialog"
};


export default class HomePageUserguide extends React.Component<IHomePageUserguideProps, IHomePageUserguideState> {

  constructor(props: IHomePageUserguideProps, state: IHomePageUserguideState) {

    super(props);

    this.state = {
      EssentialLearningsData: "",
      AddUserGuideDialog: true,
      AddUserGuideDataDialog: true,
      Title: "",
      EssentialDescription: "",
      Images: [],
      link: "",
      UploadImages: [],
      EditUserGuideDataDialog: true,
      EditTitle: "",
      EditEssentialDescription: "",
      EditImages: [],
      Editlink: "",
      EditUploadImages: [],
      CurrentUserguideDetailsID: "",
      DeleteUserguideDetailsID: "",
      previewImage: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }


  public render(): React.ReactElement<IHomePageUserguideProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const userguide = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 4,
      slidesToScroll: 4,
      arrows: true,
      autoplay: true,
      cssEase: "linear"
    };

    return (
      <section className="homePageUserguide">

        <div className="essential-section">
          <div className="essential-header">
            <h2 className="section-title">User guide</h2>
            <a href='https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/SitePages/User-Guides1.aspx' style={{ textDecoration: "none", color: "inherit" }} target="_blank" rel="noopener noreferrer">
              <button className="view-userguide">View all</button>
            </a>
          </div>

          {
            this.state.IsAdmin ?
            <>
              <div className='AddAnnouncemt'>
                <PrimaryButton text='Add UserGuides' onClick={() => this.setState({ AddUserGuideDialog: false })} />
              </div>
            </> 
            :
            <>
            </>
          }

          <Slider {...userguide}>

            {
              this.state.EssentialLearningsData.length > 0 &&
              this.state.EssentialLearningsData.map((item) => {
                return (
                  <div className="learning-card">
                    <img src={item.Images} alt="Training" className="learning-image" />
                    <div className="learning-content">
                      <h3>{item.Title}</h3>
                      <p>
                        {item.EssentialDescription}
                      </p>
                      <a href={item.link.Url} style={{ cursor: "pointer" }} className="read-more">
                        Read more →
                      </a>
                    </div>
                  </div>
                );
              })
            }

          </Slider>

          <Dialog
            hidden={this.state.AddUserGuideDialog}
            onDismiss={() =>
              this.setState({
                AddUserGuideDialog: true,
              })
            }
            dialogContentProps={AddUserguideDetailsDialogContentProps}
            modalProps={addmodelProps}
            minWidth={1500}
          >

            <div className='AddUserData'>
              <PrimaryButton className='Add Userguide' text='Add UserInfo' onClick={() => this.setState({ AddUserGuideDataDialog: false })} />
            </div>

            <div className="news-container">
              <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
                <thead>
                  <tr>
                    <th style={{ width: '20%' }}>Title</th>
                    <th style={{ width: '30%' }}>Essential Description</th>
                    <th style={{ width: '15%' }}>Images</th>
                    <th style={{ width: '15%' }}>link</th>
                    <th style={{ width: '15%' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>

                  {
                    this.state.EssentialLearningsData.length > 0 &&
                    this.state.EssentialLearningsData.map((item) => {
                      return (
                        <tr key={item.ID}>
                          <td className="title">{item.Title}</td>
                          <td>{item.EssentialDescription}</td>
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
                            <a href={item.link.Url} target="_blank" rel="noopener noreferrer">{item.link.Description}</a>
                          </td>

                          <td>
                            <div style={{ display: "flex", gap: "8px" }}>

                              <IconButton
                                iconProps={{ iconName: "Edit" }}
                                title="Edit"
                                ariaLabel="Edit"
                                onClick={() => this.setState({ EditUserGuideDataDialog: false, CurrentUserguideDetailsID: item.ID }, () => this.EditUserGuideInfo(item.ID))}
                              />

                              <IconButton
                                iconProps={{ iconName: "Delete" }}
                                title="Delete"
                                ariaLabel="Delete"
                                onClick={() => this.DeleteUserGuideInfo(item.ID)}
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
            hidden={this.state.AddUserGuideDataDialog}
            onDismiss={() =>
              this.setState({
                AddUserGuideDataDialog: true,
                Title: "",
                EssentialDescription: "",
                link: "",
            
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
                    label='Title'
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
                    label='Essential Description'
                    type='text'
                    multiline rows={3}
                    onChange={(value) =>
                      this.setState({ EssentialDescription: value.target["value"] })
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
                        this.setState({ AddUserGuideDataDialog: true })
                      }
                    />
                  </div>

                </div>
              </div>

            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EditUserGuideDataDialog}
            onDismiss={() =>
              this.setState({
                EditUserGuideDataDialog: true,
                EditTitle: "",
                EditEssentialDescription: "",
                Editlink: "",
                EditImages: [],
                EditUploadImages: []
              })
            }
            dialogContentProps={UpdateuserguideDataDialogContentProps}
            modalProps={updatemodelProps}
            minWidth={1100}
          >
            <div className='ms-Grid-row'>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <TextField
                    label='Title'
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
                    label='Essential Description'
                    type='text'
                    multiline rows={3}
                    value={this.state.EditEssentialDescription}
                    onChange={(value) =>
                      this.setState({ EditEssentialDescription: value.target["value"] })
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
                    value={this.state.Editlink}
                    onChange={(value) =>
                      this.setState({ Editlink: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Announcement-Submit'>
                  <div className='Submit-Button'>
                    <PrimaryButton
                      text='Update'
                      onClick={() => this.UpdateAnnouncementDetails(this.state.CurrentUserguideDetailsID)}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ EditUserGuideDataDialog: true })
                      }
                    />
                  </div>

                </div>
              </div>


            </div>

          </Dialog>

        </div>

      </section>
    );
  }

  public async componentDidMount() {
    this.getEssentiallearnings();
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

  public async getEssentiallearnings() {
    const roadmap = await sp.web.lists.getByTitle("User guides").items.select(
      "ID",
      "Title",
      "EssentialDescription",
      "Images",
      "link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(roadmap);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            EssentialDescription: item.EssentialDescription ? item.EssentialDescription : "",
            Images: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Image ? JSON.parse(item.Image).serverRelativeUrl : require(`../assets/Rectangle1.png`),
            link: item.link ? item.link : ""
          });
        });
        this.setState({ EssentialLearningsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details: ", error);
    });
  }

  public async AddAnnouncementInfo() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const announcement = await sp.web.lists.getByTitle("User guides").items.add({
        Title: this.state.Title,
        EssentialDescription: this.state.EssentialDescription,
        link: this.state.link
          ? {
            Url: this.state.link,
            Description: this.state.link
          }
          : null
      });

      if (this.state.UploadImages && this.state.UploadImages.length > 0) {

        const file = this.state.UploadImages[0];

        await sp.web.lists
          .getByTitle("User guides")
          .items.getById(announcement.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddUserGuideDataDialog: true });
      this.getEssentiallearnings();

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

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadImages: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }
  
  }

  public async EditUserGuideInfo(ID) {
    let Edituserguide = this.state.EssentialLearningsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(Edituserguide);
    this.setState({
      EditTitle: Edituserguide[0].Title,
      EditEssentialDescription: Edituserguide[0].EssentialDescription,
      Editlink: Edituserguide[0].link.Url,
      EditUploadImages: Edituserguide[0].Images,
    });
  }

  public async UpdateAnnouncementDetails(CurrentUserguideDetailsID) {
    try {
      const updateuserguides: any = {
        Title: this.state.EditTitle,
        EssentialDescription: this.state.EditEssentialDescription,
        link: this.state.Editlink ? {
          Url: this.state.Editlink,
          Description: this.state.Editlink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("User guides").items.getById(CurrentUserguideDetailsID).update(updateuserguides);

      if (Array.isArray(this.state.EditUploadImages) && this.state.EditUploadImages.length > 0) {

        const file = this.state.EditUploadImages[0];
  
        const itemRef = sp.web.lists
          .getByTitle("User guides")
          .items.getById(CurrentUserguideDetailsID);
  
        // delete old attachments
        const attachments = await itemRef.attachmentFiles();
  
        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }
  
        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }

      this.setState({ EditUserGuideDataDialog: true });
      this.getEssentiallearnings();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteUserGuideInfo(DeleteUserguideDetailsID) {
    const deleteinfo = await sp.web.lists.getByTitle("User guides").items.getById(DeleteUserguideDetailsID).delete();
    this.setState({ EssentialLearningsData: deleteinfo });
    this.getEssentiallearnings();
  }

}
