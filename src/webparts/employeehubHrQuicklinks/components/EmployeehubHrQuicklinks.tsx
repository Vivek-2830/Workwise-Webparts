import * as React from 'react';
import styles from './EmployeehubHrQuicklinks.module.scss';
import { IEmployeehubHrQuicklinksProps } from './IEmployeehubHrQuicklinksProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DatePicker, DefaultButton, Dialog, Dropdown, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IEmployeehubHrQuicklinksState {
  RecruitmentToolsData: any;
  RecruitmentTitle: any;
  RecruitmentImage: any;
  UploadRecruitmentImage: any;
  Link: any;
  EmployeeHRQuickDialog: boolean;
  AddEmployeeHRQuicklinkDialog: boolean;
  EditRecruitmentTitle: any;
  EditRecruitmentImage: any;
  EditUploadRecruitmentImage: any;
  EditLink: any;
  EditEmployeeHRQuicklinkDialog: boolean;
  previewImage: any;
  CurrentEmployeeHRQuickItemID: any;
  DeleteEmployeeHRQuickItemID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const HRQuicklinkDialogContentProps = {
  title: "Add HRQuicklink Details",
};

const AddHRQuicklinkDataDialogContentProps = {
  title: "Add HRQuicklink"
};

const UpdateHRQuicklinkDetailsDialogContentProps = {
  title: "Update HRQuicklink Details"
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

export default class EmployeehubHrQuicklinks extends React.Component<IEmployeehubHrQuicklinksProps, IEmployeehubHrQuicklinksState> {

  constructor(props: IEmployeehubHrQuicklinksProps, state: IEmployeehubHrQuicklinksState) {

    super(props);

    this.state = {
      RecruitmentToolsData: "",
      RecruitmentTitle: "",
      RecruitmentImage: [],
      UploadRecruitmentImage: [],
      Link: "",
      EmployeeHRQuickDialog: true,
      AddEmployeeHRQuicklinkDialog: true,
      EditRecruitmentTitle: "",
      EditRecruitmentImage: [],
      EditUploadRecruitmentImage: [],
      EditLink: "",
      EditEmployeeHRQuicklinkDialog: true,
      previewImage: "",
      CurrentEmployeeHRQuickItemID: "",
      DeleteEmployeeHRQuickItemID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IEmployeehubHrQuicklinksProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubHrQuicklinks">

        <div className="tools-wrapper">

          <h2 className="tools-title">
            HR Quick Links
            <span className="underline"></span>

            {
              this.state.IsAdmin ?
              <>
                <div className='Addquicklink'>
                  <PrimaryButton text='Add HRQuicklink' onClick={() => this.setState({ EmployeeHRQuickDialog: false })} />
                </div>
              </>
              :
              <>
              </>
            }

          </h2>

          <div className="tools-grid">

            {
              this.state.RecruitmentToolsData.length > 0 &&
              this.state.RecruitmentToolsData.map((item) => {

                return (
                  <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                    <div className="tool-card">
                      <img src={item.RecruitmentImage} />
                      <p>{item.RecruitmentTitle}</p>
                    </div>
                  </a>

                );
              })
            }

          </div>

        </div>

        <Dialog
          hidden={this.state.EmployeeHRQuickDialog}
          onDismiss={() =>
            this.setState({
              EmployeeHRQuickDialog: true,
            })
          }
          dialogContentProps={HRQuicklinkDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddHrInfo'>
            <PrimaryButton className='AddQuickInfo' text='Add Data' onClick={() => this.setState({ AddEmployeeHRQuicklinkDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>RecruitmentTitle</th>
                  <th style={{ width: '30%' }}>RecruitmentImage</th>
                  <th style={{ width: '30%' }}>Link</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.RecruitmentToolsData.length > 0 &&
                  this.state.RecruitmentToolsData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.RecruitmentTitle}</td>
                        <td>
                          {
                            item.RecruitmentImage ? (
                              <img src={item.RecruitmentImage} alt="announcement" style={{ width: "120px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Image"
                            )
                          }
                        </td>
                        <td>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditEmployeeHRQuicklinkDialog: false, CurrentEmployeeHRQuickItemID: item.ID }, () => this.EditQuickLinksInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteHRLinkInfo(item.ID)}
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
          hidden={this.state.AddEmployeeHRQuicklinkDialog}
          onDismiss={() =>
            this.setState({
              AddEmployeeHRQuicklinkDialog: true,
              RecruitmentTitle: "",
              RecruitmentImage: [],
              UploadRecruitmentImage: [],
              Link: ""
            })
          }
          dialogContentProps={AddHRQuicklinkDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={1100}
        >
          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='RecruitmentTitle'
                  type='text'
                  onChange={(value) =>
                    this.setState({ RecruitmentTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <label><b>Upload RecruitmentImage</b></label><br />

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

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddHRQuickInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddEmployeeHRQuicklinkDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditEmployeeHRQuicklinkDialog}
          onDismiss={() =>
            this.setState({
              EditEmployeeHRQuicklinkDialog: true,
              EditRecruitmentTitle: "",
              EditRecruitmentImage: [],
              EditUploadRecruitmentImage: [],
              EditLink: ""
            })
          }
          dialogContentProps={UpdateHRQuicklinkDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={1100}
        >
          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='RecruitmentTitle'
                  type='text'
                  value={this.state.EditRecruitmentTitle}
                  onChange={(value) =>
                    this.setState({ EditRecruitmentTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <label><b>Upload RecruitmentImage</b></label><br />

                <input
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleUpdateImageChange(e)}
                />

                {
                  this.state.EditUploadRecruitmentImage && (
                    <div className="Attached-img">

                      {/* ✅ Handle BOTH string + file */}
                      <p>
                        {
                          typeof this.state.EditUploadRecruitmentImage === "string"
                            ? this.state.EditUploadRecruitmentImage.split('/').pop()
                            : this.state.EditUploadRecruitmentImage[0]?.name
                        }
                      </p>

                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadRecruitmentImage: "" })}
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

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateHRQuickDetails(this.state.CurrentEmployeeHRQuickItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditEmployeeHRQuicklinkDialog: true })
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
    this.getRecruitmentData();
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

  public async getRecruitmentData() {
    const tools = await sp.web.lists.getByTitle("HR Quick Links").items.select(
      "ID",
      "RecruitmentTitle",
      "RecruitmentImage",
      "Link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(tools);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            RecruitmentTitle: item.RecruitmentTitle ? item.RecruitmentTitle : "",
            RecruitmentImage: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.RecruitmentImage ? JSON.parse(item.RecruitmentImage).serverRelativeUrl : require(`../assets/fi3194787.png`),
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ RecruitmentToolsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details in Recruitment Tools:", error);
    });
  }

  public async AddHRQuickInfo() {
    if (this.state.RecruitmentTitle.length == 0) {
      alert("Please Enter Details");
    } else {
      const hrquick = await sp.web.lists.getByTitle("HR Quick Links").items.add({
        RecruitmentTitle: this.state.RecruitmentTitle,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null
      });

      if (this.state.UploadRecruitmentImage && this.state.UploadRecruitmentImage.length > 0) {

        const file = this.state.UploadRecruitmentImage[0];

        await sp.web.lists
          .getByTitle("HR Quick Links")
          .items.getById(hrquick.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddEmployeeHRQuicklinkDialog: true });
      this.getRecruitmentData();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadRecruitmentImage: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  }

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadRecruitmentImage: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditQuickLinksInfo(ID) {
    let EditquickLink = this.state.RecruitmentToolsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditquickLink);
    this.setState({
      EditRecruitmentTitle: EditquickLink[0].RecruitmentTitle,
      EditLink: EditquickLink[0].Link.Url,
      EditUploadRecruitmentImage: EditquickLink[0].RecruitmentImage
    });
  }

  public async UpdateHRQuickDetails(CurrentEmployeeHRQuickItemID) {
    try {
      const updatequicklinks: any = {
        RecruitmentTitle: this.state.EditRecruitmentTitle,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("HR Quick Links").items.getById(CurrentEmployeeHRQuickItemID).update(updatequicklinks);

      if (Array.isArray(this.state.EditUploadRecruitmentImage) && this.state.EditUploadRecruitmentImage.length > 0) {

        const file = this.state.EditUploadRecruitmentImage[0];

        const itemRef = sp.web.lists
          .getByTitle("HR Quick Links")
          .items.getById(CurrentEmployeeHRQuickItemID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditEmployeeHRQuicklinkDialog: true });
      this.getRecruitmentData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteHRLinkInfo(DeleteEmployeeHRQuickItemID) {
    const deleteinfo = await sp.web.lists.getByTitle("HR Quick Links").items.getById(DeleteEmployeeHRQuickItemID).delete();
    this.setState({ RecruitmentToolsData: deleteinfo });
    this.getRecruitmentData();
  }

}
