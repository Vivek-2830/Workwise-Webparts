import * as React from 'react';
import styles from './EmployeehubHrTeamContact.module.scss';
import { IEmployeehubHrTeamContactProps } from './IEmployeehubHrTeamContactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IEmployeehubHrTeamContactState {
  HRDetailsData: any;
  Name: any;
  JobTitle: any;
  Phone: any;
  Email: any;
  Photo: any;
  UploadPhoto: any;
  AddHRTemaDialog: boolean;
  AddHRTeamDataDialog: boolean;
  EditName: any;
  EditJobTitle: any;
  EditPhone: any;
  EditEmail: any;
  EditPhoto: any;
  EditUploadPhoto: any;
  EditHRTeamDataDialog: boolean;
  CurrentHrTeamDetailsID: any;
  DeleteHrTeamDetailsID: any;
  previewImage: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddHRTeamContactDetailsDialogContentProps = {
  title: "Add HRTeamContact Details",
};

const AddHRTeamContactDataDialogContentProps = {
  title: "Add HRTeamContact"
};

const UpdateHRTeamContactDetailsDialogContentProps = {
  title: "Update HRTeamContact Details"
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


export default class EmployeehubHrTeamContact extends React.Component<IEmployeehubHrTeamContactProps, IEmployeehubHrTeamContactState> {

  constructor(props: IEmployeehubHrTeamContactProps, state: IEmployeehubHrTeamContactState) {

    super(props);

    this.state = {
      HRDetailsData: "",
      Name: "",
      JobTitle: "",
      Phone: "",
      Email: "",
      Photo: [],
      UploadPhoto: [],
      AddHRTemaDialog: true,
      AddHRTeamDataDialog: true,
      EditName: "",
      EditJobTitle: "",
      EditPhone: "",
      EditEmail: "",
      EditPhoto: [],
      EditUploadPhoto: [],
      EditHRTeamDataDialog: true,
      CurrentHrTeamDetailsID: "",
      DeleteHrTeamDetailsID: "",
      previewImage: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }


  public render(): React.ReactElement<IEmployeehubHrTeamContactProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubHrTeamContact">

        <div className="hr-wrapper">

          <div className="hr-title">
            <div>
              <h3>HR Team Contact Details</h3>
              <span className="underline"></span>
            </div>
            {
              this.state.IsAdmin ?
              <>
                <div className='Addteam'>
                  <PrimaryButton text='Add TeamContact' onClick={() => this.setState({ AddHRTemaDialog: false })} />
                </div>
              </>
              :
              <></>
            }

          </div>

          <div className="hr-grid">

            {
              this.state.HRDetailsData.length > 0 &&
              this.state.HRDetailsData.map((item, index) => {

                return (
                  <div className="hr-card">

                    <div className='main-card'>
                      <img src={item.Photo} className="avatar" />
                      <div className="hr-info">
                        <h4>{item.Name}</h4>
                        <span className="role">{item.JobTitle}</span>
                      </div>
                    </div>

                    {
                      !!item.Phone && (
                        <div className="contact">
                          <img src={require('../assets/phone.png')} alt="phone" />
                          <p>{item.Phone}</p>
                        </div>
                      )
                    }

                    <div className='contact-email'>
                      <img src={require('../assets/mail01.png')} /> <p>{item.Email}</p>
                    </div>

                  </div>
                );
              })
            }

          </div>
        </div>

        <Dialog
          hidden={this.state.AddHRTemaDialog}
          onDismiss={() =>
            this.setState({
              AddHRTemaDialog: true,
            })
          }
          dialogContentProps={AddHRTeamContactDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1200}
        >

          <div className='hrteambox'>
            <div>
              <h2>HR Team Contact Details</h2>
            </div>
            <div className='AddHRData'>
              <PrimaryButton className='AddHRInfo' text='Add Data' onClick={() => this.setState({ AddHRTeamDataDialog: false })} />
            </div>
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>JobTitle</th>
                  <th>Phone</th>
                  <th>Email</th>
                  <th>Photo</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.HRDetailsData.length > 0 &&
                  this.state.HRDetailsData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Name}</td>
                        <td>{item.JobTitle}</td>
                        <td>{item.Phone}</td>
                        <td>{item.Email}</td>
                        <td>
                          {
                            item.Photo ? (
                              <img src={item.Photo} alt="Photo" style={{ width: "120px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Photo"
                            )
                          }
                        </td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditHRTeamDataDialog: false, CurrentHrTeamDetailsID: item.ID }, () => this.EditHRContact(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteHRTeamInfo(item.ID)}
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
          hidden={this.state.AddHRTeamDataDialog}
          onDismiss={() =>
            this.setState({
              AddHRTeamDataDialog: true,
              Name: "",
              JobTitle: "",
              Phone: "",
              Email: "",
              Photo: [],
              UploadPhoto: []
            })
          }
          dialogContentProps={AddHRTeamContactDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={900}
        >

          <div>
            <h2>Add HR Team Contact Details</h2>
          </div>

          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label="Name"
                  type='text'
                  onChange={(value) =>
                    this.setState({ Name: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label='Job Title'
                  type='text'
                  onChange={(value) =>
                    this.setState({ JobTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label='Phone'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Phone: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label='Email'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Email: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Photo</b></label>

                <input className='teamicon'
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleImageChange(e)}
                />

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddHRContactInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddHRTeamDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditHRTeamDataDialog}
          onDismiss={() =>
            this.setState({
              EditHRTeamDataDialog: true,
              EditName: "",
              EditJobTitle: "",
              EditPhone: "",
              EditEmail: "",
              EditPhoto: [],
              EditUploadPhoto: []
            })
          }
          dialogContentProps={UpdateHRTeamContactDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={900}
        >
          <div>
            <h2>Update HR Team Contact Details</h2>
          </div>

          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label='Name'
                  type='text'
                  value={this.state.EditName}
                  onChange={(value) =>
                    this.setState({ EditName: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label='Job Title'
                  type='text'
                  value={this.state.EditJobTitle}
                  onChange={(value) =>
                    this.setState({ EditJobTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <TextField
                  label='Phone'
                  type='text'
                  value={this.state.EditPhone}
                  onChange={(value) =>
                    this.setState({ EditPhone: value.target["value"] })
                  }
                />
              </div>
            </div>
            
            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div>
                <TextField
                  label='Email'
                  type='text'
                  value={this.state.EditEmail}
                  onChange={(value) =>
                    this.setState({ EditEmail: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 HRTeamSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Photo</b></label>

                <input className='teamicon'
                  type="file" 
                  accept="image/*"
                  onChange={(e: any) => this.handleUpdateImageChange(e)}
                />

                {
                  this.state.EditUploadPhoto && (
                    <div className="Attached-img">

                      {/* ✅ Handle BOTH string + file */}
                      <p>
                        {
                          typeof this.state.EditUploadPhoto === "string"
                            ? this.state.EditUploadPhoto.split('/').pop()
                            : this.state.EditUploadPhoto[0]?.name
                        }
                      </p>

                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadPhoto: "" })}
                      />
                    </div>
                  )
                }

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateHRContactDetails(this.state.CurrentHrTeamDetailsID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditHRTeamDataDialog: true })
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
    this.getHRTeamDetails();
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

  public async getHRTeamDetails() {
    const details = await sp.web.lists.getByTitle("HR Team Contact Details").items.select(
      "ID",
      "Name",
      "JobTitle",
      "Phone",
      "Email",
      "Photo"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(details);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : [],
            Name: item.Name ? item.Name : "",
            JobTitle: item.JobTitle ? item.JobTitle : "",
            Phone: item.Phone ? item.Phone : "",
            Email: item.Email ? item.Email : "",
            Photo: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Photo ? JSON.parse(item.Photo).serverRelativeUrl : require(`../assets/avatar3.png`)
          });
        });
        this.setState({ HRDetailsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Detail in HR Team Contact Details:", error);
    });
  }

  public async AddHRContactInfo() {
    if (this.state.Name.length == 0) {
      alert("Please Enter Details");
    } else {
      const hrcontact = await sp.web.lists.getByTitle("HR Team Contact Details").items.add({
        Name: this.state.Name,
        JobTitle: this.state.JobTitle,
        Phone: this.state.Phone,
        Email: this.state.Email
      });

      if (this.state.UploadPhoto && this.state.UploadPhoto.length > 0) {

        const file = this.state.UploadPhoto[0];

        await sp.web.lists
          .getByTitle("HR Team Contact Details")
          .items.getById(hrcontact.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddHRTeamDataDialog: true });
      this.getHRTeamDetails();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadPhoto: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  }

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadPhoto: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditHRContact(ID) {
    let EditHRInfo = this.state.HRDetailsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditHRInfo);
    this.setState({
      EditName: EditHRInfo[0].Name,
      EditJobTitle: EditHRInfo[0].JobTitle,
      EditPhone: EditHRInfo[0].Phone,
      EditEmail: EditHRInfo[0].Email,
      EditUploadPhoto: EditHRInfo[0].Photo,
    });
  }

  public async UpdateHRContactDetails(CurrentHrTeamDetailsID) {
    try {
      const updateempannouncement: any = {
        Name: this.state.EditName,
        JobTitle: this.state.EditJobTitle,
        Phone: this.state.EditPhone,
        Email: this.state.EditEmail
      };

      const updateItem = await sp.web.lists.getByTitle("HR Team Contact Details").items.getById(CurrentHrTeamDetailsID).update(updateempannouncement);

      if (Array.isArray(this.state.EditUploadPhoto) && this.state.EditUploadPhoto.length > 0) {

        const file = this.state.EditUploadPhoto[0];

        const itemRef = sp.web.lists
          .getByTitle("HR Team Contact Details")
          .items.getById(CurrentHrTeamDetailsID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditHRTeamDataDialog: true });
      this.getHRTeamDetails();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteHRTeamInfo(DeleteHrTeamDetailsID) {
    const deleteinfo = await sp.web.lists.getByTitle("HR Team Contact Details").items.getById(DeleteHrTeamDetailsID).delete();
    this.setState({ HRDetailsData: deleteinfo });
    this.getHRTeamDetails();
  }

}
