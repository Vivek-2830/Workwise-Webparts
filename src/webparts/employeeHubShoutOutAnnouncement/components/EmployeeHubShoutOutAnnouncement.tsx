import * as React from 'react';
import styles from './EmployeeHubShoutOutAnnouncement.module.scss';
import { IEmployeeHubShoutOutAnnouncementProps } from './IEmployeeHubShoutOutAnnouncementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IEmployeeHubShoutOutAnnouncementState {
  ShoutOutAnnouncementData: any;
  Title: any;
  Description: any;
  Icon: any;
  UploadIcon: any;
  ShoutOutAnnouncementDialog: boolean;
  AddShoutoutAnnouncementDialog: boolean;
  EditTitle: any;
  EditDescription: any;
  EditIcon: any;
  EditUploadIcon: any;
  EditShoutOutAnnouncementDialog: boolean;
  CurrentShoutOutItemID: any;
  DeleteShoutOutItemID: any;
  previewImage: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const ShoutOutAnnouncementDialogContentProps = {
  title: "Add  ShoutOutAnnouncement Details",
};

const AddShoutOutAnnouncementDataDialogContentProps = {
  title: "Add  ShoutOutAnnouncement"
};

const UpdateShoutOutAnnouncementDataDialogContentProps = {
  title: "Update  ShoutOutAnnouncement Details"
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

export default class EmployeeHubShoutOutAnnouncement extends React.Component<IEmployeeHubShoutOutAnnouncementProps, IEmployeeHubShoutOutAnnouncementState> {

  constructor(props: IEmployeeHubShoutOutAnnouncementProps, state: IEmployeeHubShoutOutAnnouncementState) {

    super(props);

    this.state = {
      ShoutOutAnnouncementData: "",
      Title: "",
      Description: "",
      Icon: [],
      UploadIcon: [],
      ShoutOutAnnouncementDialog: true,
      AddShoutoutAnnouncementDialog: true,
      EditTitle: "",
      EditDescription: "",
      EditIcon: [],
      EditUploadIcon: [],
      EditShoutOutAnnouncementDialog: true,
      CurrentShoutOutItemID: "",
      DeleteShoutOutItemID: "",
      previewImage: "",
      IsAdmin: false,
      CurrentUserEmail: "",
    };

  }

  public render(): React.ReactElement<IEmployeeHubShoutOutAnnouncementProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubShoutOutAnnouncement">

        <div className="shoutouts-section">
          
          <div className='shoutoutsAnnouncement'>
            <div>
              <h2 className="shoutouts-title">Shout Out's</h2>
            </div>
              {
              this.state.IsAdmin ?
                <>
                  <div className='Addshouout'>
                    <PrimaryButton text='Add ShoutOut' onClick={() => this.setState({ ShoutOutAnnouncementDialog: false })} />
                  </div>
                </>
                :
                <>
                </>
              }
          </div>

          <div className='shoutout-scroll'>

            {
              this.state.ShoutOutAnnouncementData.length > 0 &&
              this.state.ShoutOutAnnouncementData.map((item) => {
                return (
                  <div className="shout-card">
                    <div className="shout-icon">📣</div>
                    <div className="shout-content">
                      <div className="shout-name">{item.Title}</div>
                      <div className="shout-message">
                        {item.Description}
                      </div>
                    </div>
                  </div>
                );
              })
            }

          </div>

        </div>

        <Dialog
          hidden={this.state.ShoutOutAnnouncementDialog}
          onDismiss={() =>
            this.setState({
              ShoutOutAnnouncementDialog: true,
            })
          }
          dialogContentProps={ShoutOutAnnouncementDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1200}
        >

          <div className='shoutoutbox'>
            <div>
              <h2>ShoutOut Announcement Information</h2>
            </div>
            <div className='AddshoutData'>
              <PrimaryButton className='AddshoutInfo' text='Add Data' onClick={() => this.setState({ AddShoutoutAnnouncementDialog: false })} />
            </div>
          </div>  

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Description</th>
                  <th>Icon</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.ShoutOutAnnouncementData.length > 0 &&
                  this.state.ShoutOutAnnouncementData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Title}</td>
                        <td>{item.Description}</td>
                        <td>
                          {
                            item.Icon ? (
                              <img src={item.Icon} alt="announcement" style={{ width: "120px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Icon"
                            )
                          }
                        </td>
                        
                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditShoutOutAnnouncementDialog: false, CurrentShoutOutItemID: item.ID }, () => this.EditShoutoutInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteShoutOutInfo(item.ID)}
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
          hidden={this.state.AddShoutoutAnnouncementDialog}
          onDismiss={() =>
            this.setState({
              AddShoutoutAnnouncementDialog: true,
              Title: "",
              Description: "",
              Icon: [],
              UploadIcon : []
            })
          }
          dialogContentProps={AddShoutOutAnnouncementDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={900}
        >

          <div>
            <h2>Add ShoutOut Announcement</h2>
          </div>

          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 ShoutOutSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 ShoutOutSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 ShoutOutSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icon</b></label>

                <input className='shoutouticon'
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleImageChange(e)}
                />

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 ShoutOutSection'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddShoutoutInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddShoutoutAnnouncementDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditShoutOutAnnouncementDialog}
          onDismiss={() =>
            this.setState({
              EditShoutOutAnnouncementDialog: true,
              EditTitle: "",
              EditDescription: "",
              EditIcon: [],
              EditUploadIcon : []
            })
          }
          dialogContentProps={UpdateShoutOutAnnouncementDataDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={900}
        >
          <div>
            <h2>Update ShoutOut Announcement</h2>
          </div>

          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 ShoutOutSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 ShoutOutSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 ShoutOutSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Image</b></label>

                <input className='shoutouticon'
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleUpdateImageChange(e)}
                />

                {
                  this.state.EditUploadIcon && (
                    <div className="Attached-img">

                      {/* ✅ Handle BOTH string + file */}
                      <p>
                        {
                          typeof this.state.EditUploadIcon === "string"
                            ? this.state.EditUploadIcon.split('/').pop()
                            : this.state.EditUploadIcon[0]?.name
                        }
                      </p>

                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadIcon: "" })}
                      />
                    </div>
                  )
                }

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 '>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateShoutOutDetails(this.state.CurrentShoutOutItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditShoutOutAnnouncementDialog: true })
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
    this.getShoutoutannouncementInfo();
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

  public async getShoutoutannouncementInfo() {
    const info = await sp.web.lists.getByTitle("ShoutOut Announcement").items.select(
      "ID",
      "Title",
      "Description",
      "Icon"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      // console.log(info);
      // console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Description: item.Description ? item.Description : "",
            Icon: item.Icon ? item.Icon : ""
          });
        });
        this.setState({ ShoutOutAnnouncementData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details", error);
    });
  }

  public async AddShoutoutInfo() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const shoutout = await sp.web.lists.getByTitle("ShoutOut Announcement").items.add({
        Title: this.state.Title,
        Description: this.state.Description,
      });

      if (this.state.UploadIcon && this.state.UploadIcon.length > 0) {

        const file = this.state.UploadIcon[0];

        await sp.web.lists
          .getByTitle("ShoutOut Announcement")
          .items.getById(shoutout.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddShoutoutAnnouncementDialog: true });
      this.getShoutoutannouncementInfo();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadIcon: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  }

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadIcon: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditShoutoutInfo(ID) {
    let EditAnnouncement = this.state.ShoutOutAnnouncementData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditAnnouncement);
    this.setState({
      EditTitle: EditAnnouncement[0].Title,
      EditDescription: EditAnnouncement[0].Description,
      EditUploadIcon: EditAnnouncement[0].Icon,

    });
  }

  public async UpdateShoutOutDetails(CurrentShoutOutItemID) {
    try {
      const updateannouncement: any = {
        Title: this.state.EditTitle,
        Description: this.state.EditDescription,
      };

      const updateItem = await sp.web.lists.getByTitle("ShoutOut Announcement").items.getById(CurrentShoutOutItemID).update(updateannouncement);

      if (Array.isArray(this.state.EditUploadIcon) && this.state.EditUploadIcon.length > 0) {

        const file = this.state.EditUploadIcon[0];

        const itemRef = sp.web.lists
          .getByTitle("ShoutOut Announcement")
          .items.getById(CurrentShoutOutItemID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditShoutOutAnnouncementDialog: true });
      this.getShoutoutannouncementInfo();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteShoutOutInfo(DeleteShoutOutItemID) {
    const deleteinfo = await sp.web.lists.getByTitle("ShoutOut Announcement").items.getById(DeleteShoutOutItemID).delete();
    this.setState({ ShoutOutAnnouncementData: deleteinfo });
    this.getShoutoutannouncementInfo();
  }

}
