import * as React from 'react';
import styles from './EmployeehubOurValues.module.scss';
import { IEmployeehubOurValuesProps } from './IEmployeehubOurValuesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';


export interface IEmployeehubOurValuesState {
  OurValuesData: any;
  Title: any;
  Tag: any;
  Description: any;
  Icon: any;
  UploadIcon: any;
  AddOurvalueDialog: boolean;
  AddOurvalueDataDialog: boolean;
  EditTitle: any;
  EditTag: any;
  EditDescription: any;
  EditIcon: any;
  EditUploadIcon: any;
  EditOurValuesDataDialog: boolean;
  CurrentOurValueItemID: any;
  DeleteOurvalueItemID: any;
  previewImage: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddOurvalueDetailsDialogContentProps = {
  title: "Add OurValue Details",
};

const AddOurValuesDataDialogContentProps = {
  title: "Add Value"
};

const UpdateOurValuesDetailsDialogContentProps = {
  title: "Update OurValue Details"
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

export default class EmployeehubOurValues extends React.Component<IEmployeehubOurValuesProps, IEmployeehubOurValuesState> {

  constructor(props: IEmployeehubOurValuesProps, state: IEmployeehubOurValuesState) {

    super(props);

    this.state = {
      OurValuesData: "",
      Title: "",
      Tag: "",
      Description: "",
      Icon: [],
      UploadIcon: [],
      AddOurvalueDialog: true,
      AddOurvalueDataDialog: true,
      EditTitle: "",
      EditTag: "",
      EditDescription: "",
      EditIcon: [],
      EditUploadIcon: [],
      EditOurValuesDataDialog: true,
      CurrentOurValueItemID: "",
      DeleteOurvalueItemID: "",
      previewImage: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }


  public render(): React.ReactElement<IEmployeehubOurValuesProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubOurValues">

        <div className='Trust-sec'>

          <div className='OurValuePart'>
            <div>
              <h2 className='Value-Title'>Our Values</h2>
            </div>
            {
              this.state.IsAdmin ?
              <>
                <div className='Addvalue'>
                  <PrimaryButton text='Add OurValue' onClick={() => this.setState({ AddOurvalueDialog: false })} />
                </div>
              </>
              :
              <>
              </>
            }
          </div>

          <div className="trust-wrapper">

            {
              this.state.OurValuesData.length > 0 &&
              this.state.OurValuesData.map((item) => {
                return (
                  <div className="trust-section">
                    <div className="trust-container">
                      <div className="trust-left">
                        <div className="trust-badge">{item.Tag}</div>
                        <h2>{item.Title}</h2>
                        <p>
                          {item.Description}
                        </p>
                      </div>
                      <div className="trust-right">
                        <img src={item.Icon} alt="Safety" />
                      </div>
                    </div>
                  </div>
                );
              })
            }

          </div>
        </div>

        <Dialog
          hidden={this.state.AddOurvalueDialog}
          onDismiss={() =>
            this.setState({
              AddOurvalueDialog: true,
            })
          }
          dialogContentProps={AddOurvalueDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1200}
        >

          <div className='ourvaluebox'>
            <div>
              <h2>Employee Our Value Details</h2>
            </div>
            <div className='AddvalueData'>
              <PrimaryButton className='AddourValue' text='Add Data' onClick={() => this.setState({ AddOurvalueDataDialog: false })} />
            </div>
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Tag</th>
                  <th>Description</th>
                  <th>Icon</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.OurValuesData.length > 0 &&
                  this.state.OurValuesData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Title}</td>
                        <td>{item.Tag}</td>
                        <td>{item.Description}</td>
                        <td>
                          {
                            item.Icon ? (
                              <img src={item.Icon} alt="announcement" style={{ width: "80px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Image"
                            )
                          }
                        </td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditOurValuesDataDialog: false, CurrentOurValueItemID: item.ID }, () => this.EditOurValueInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteOurvalueItems(item.ID)}
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
          hidden={this.state.AddOurvalueDataDialog}
          onDismiss={() =>
            this.setState({
              AddOurvalueDataDialog: true,
              Title: "",
              Tag: "",
              Description: "",
              Icon: [],
              UploadIcon: []
            })
          }
          dialogContentProps={AddOurValuesDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={900}
        >

          <div>
            <h2>Add Employee Our Value Details</h2>
          </div>

          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
              <div className='Add-Form'>
                <TextField
                  label='Tag'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Tag: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icon</b></label>

                <input className='ouricon'
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
                    onClick={() => this.AddOurValueDetails()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddOurvalueDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditOurValuesDataDialog}
          onDismiss={() =>
            this.setState({
              EditOurValuesDataDialog: true,
              EditTitle: "",
              EditTag: "",
              EditDescription: "",
              EditIcon: [],
              EditUploadIcon: []
            })
          }
          dialogContentProps={UpdateOurValuesDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={900}
        >
          <div>
            <h2>Update Employee Our Value Details</h2>
          </div>

          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
              <div className='Add-Form'>
                <TextField
                  label='Tag'
                  type='text'
                  value={this.state.EditTag}
                  onChange={(value) =>
                    this.setState({ EditTag: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurValueSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icon</b></label>

                <input className='ouricon'
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

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateourvalueDetails(this.state.CurrentOurValueItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditOurValuesDataDialog: true })
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
    this.getOurvalues();
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

  public async getOurvalues() {
    const value = await sp.web.lists.getByTitle("Our Values").items.select(
      "ID",
      "Title",
      "Tag",
      "Description",
      "Icon"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(value);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Tag: item.Tag ? item.Tag : "",
            Description: item.Description ? item.Description : "",
            Icon: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Icon ? JSON.parse(item.Icon).serverRelativeUrl : require(`../assets/Safety.jpg`)
          });
        });
        this.setState({ OurValuesData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching details: ", error);
    });
  }

  public async AddOurValueDetails() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const ourvalueinfo = await sp.web.lists.getByTitle("Our Values").items.add({
        Title: this.state.Title,
        Tag: this.state.Tag,
        Description: this.state.Description
      });

      if (this.state.UploadIcon && this.state.UploadIcon.length > 0) {

        const file = this.state.UploadIcon[0];

        await sp.web.lists
          .getByTitle("Our Values")
          .items.getById(ourvalueinfo.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddOurvalueDataDialog: true });
      this.getOurvalues();
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

  public async EditOurValueInfo(ID) {
    let EditourValues = this.state.OurValuesData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditourValues);
    this.setState({
      EditTitle: EditourValues[0].Title,
      EditDescription: EditourValues[0].Description,
      EditTag: EditourValues[0].Tag,
      EditUploadIcon: EditourValues[0].Icon,
    });
  }

  public async UpdateourvalueDetails(CurrentOurValueItemID) {
    try {
      const updateourvalues: any = {
        Title: this.state.EditTitle,
        Description: this.state.EditDescription,
        Tag: this.state.EditTag
      };

      const updateItem = await sp.web.lists.getByTitle("Our Values").items.getById(CurrentOurValueItemID).update(updateourvalues);

      if (Array.isArray(this.state.EditUploadIcon) && this.state.EditUploadIcon.length > 0) {

        const file = this.state.EditUploadIcon[0];

        const itemRef = sp.web.lists
          .getByTitle("Our Values")
          .items.getById(CurrentOurValueItemID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditOurValuesDataDialog: true });
      this.getOurvalues();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteOurvalueItems(DeleteOurvalueItemID) {
    const deleteinfo = await sp.web.lists.getByTitle("Our Values").items.getById(DeleteOurvalueItemID).delete();
    this.setState({ OurValuesData: deleteinfo });
    this.getOurvalues();
  }

}
