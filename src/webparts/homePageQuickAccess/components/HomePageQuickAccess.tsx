import * as React from 'react';
import styles from './HomePageQuickAccess.module.scss';
import { IHomePageQuickAccessProps } from './IHomePageQuickAccessProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, Dropdown, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePageQuickAccessState {
  QuickAccessData: any;
  QuickAccessCategories: any;
  AccessTitle: any;
  AccessDescription: any;
  Icon: any;
  Link: any;
  UploadIcon: any;
  AddQuickaccessDialog: boolean;
  AddQuickAccessDataDialog: boolean;
  EditQuickAccessDataDialog: boolean;
  EditQuickAccessCategories: any;
  EditAccessTitle: any;
  EditAccessDescription: any;
  EditIcon: any;
  EditUploadIcon: any;
  EditLink: any;
  CurrentQuickAccessDataID: any;
  DeleteQuickAccessDataID: any;
  previewImage: any;
  QuickAccessCategorieslist: any; 
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddQuickaccessDetailsDialogContentProps = {
  title: "Add Quickaccess Details",
};

const AddQuickaccessDataDialogContentProps = {
  title: "Add Quickaccess"
};

const UpdateQuickaccessDetailsDialogContentProps = {
  title: "Update Quickaccess Details"
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

export default class HomePageQuickAccess extends React.Component<IHomePageQuickAccessProps, IHomePageQuickAccessState> {

  constructor(props: IHomePageQuickAccessProps, state: IHomePageQuickAccessState) {

    super(props);

    this.state = {
      QuickAccessData: "",
      QuickAccessCategories: [],
      AccessTitle: "",
      AccessDescription: "",
      Icon: [],
      Link: "",
      UploadIcon: [],
      AddQuickaccessDialog: true,
      AddQuickAccessDataDialog: true,
      EditQuickAccessDataDialog: true,
      EditQuickAccessCategories: [],
      EditAccessTitle: "",
      EditAccessDescription: "",
      EditIcon: [],
      EditUploadIcon: [],
      EditLink: "",
      CurrentQuickAccessDataID: "",
      DeleteQuickAccessDataID: "",
      previewImage: "",
      QuickAccessCategorieslist : [],
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }


  public render(): React.ReactElement<IHomePageQuickAccessProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageQuickAccess">

        {
          this.state.IsAdmin ?
            <>
              <div className='AddAnnouncemt'>
                <PrimaryButton text='Add Access Details' onClick={() => this.setState({ AddQuickaccessDialog: false })} />
              </div>
            </>
            :
            <>
            </>
        }

        <div className="quick-access-wrapper">

          <h2>Quick Access</h2>
          <p className="subtitle">
            Find the applications and resources you need to get work done efficiently
          </p>

          <div className="qa-cards">

            {/* ================= Productivity ================= */}
            {
              this.state.QuickAccessData.length > 0 &&
              this.state.QuickAccessData.filter(i => i.QuickAccessCategories === "Tools").length > 0 &&

              <div className="qa-card">
                <h3>
                  <span className="card-icon">
                    <img src={require('../assets/baggage.png')} />
                  </span>
                  Tools
                </h3>

                <ul>
                  {
                    this.state.QuickAccessData.map((item) => {

                      if (item.QuickAccessCategories !== "Tools") return null;

                      // let imagePath = "";
                      // let ImageInfo = JSON.parse(item.Icon);
                      // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                      //   imagePath = ImageInfo["serverRelativeUrl"];
                      // } else {
                      //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Access/Attachments/${item.ID}/${ImageInfo.fileName}`;
                      // }

                      return (
                        <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                          <li key={item.ID}>
                            <span className="item-icon">
                              <img src={item.Icon} />
                            </span>
                            <div>
                              <strong>{item.AccessTitle}</strong>
                              <p>{item.AccessDescription}</p>
                            </div>
                          </li>
                        </a>
                      );
                    })
                  }
                </ul>
              </div>
            }


            {/* ================= Human Resources ================= */}
            {
              this.state.QuickAccessData.length > 0 &&
              this.state.QuickAccessData.filter(i => i.QuickAccessCategories === "Support").length > 0 &&

              <div className="qa-card">
                <h3>
                  <span className="card-icon">
                    <img src={require('../assets/group.png')} />
                  </span>
                  Support
                </h3>

                <ul>
                  {
                    this.state.QuickAccessData.map((item) => {

                      if (item.QuickAccessCategories !== "Support") return null;

                      // let imagePath = "";
                      // let ImageInfo = JSON.parse(item.Icon);
                      // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                      //   imagePath = ImageInfo["serverRelativeUrl"];
                      // } else {
                      //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Access/Attachments/${item.ID}/${ImageInfo.fileName}`;
                      // }

                      return (
                        <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                          <li key={item.ID}>
                            <span className="item-icon">
                              <img src={item.Icon} />
                            </span>
                            <div>
                              <strong>{item.AccessTitle}</strong>
                              <p>{item.AccessDescription}</p>
                            </div>
                          </li>
                        </a>
                      );
                    })
                  }
                </ul>
              </div>
            }


            {/* ================= Business Applications ================= */}
            {
              this.state.QuickAccessData.length > 0 &&
              this.state.QuickAccessData.filter(i => i.QuickAccessCategories === "Resources").length > 0 &&

              <div className="qa-card">
                <h3>
                  <span className="card-icon">
                    <img src={require('../assets/phone.png')} />
                  </span>
                  Resources
                </h3>

                <ul>
                  {
                    this.state.QuickAccessData.map((item) => {

                      if (item.QuickAccessCategories !== "Resources") return null;

                      // let imagePath = "";
                      // let ImageInfo = JSON.parse(item.Icon);
                      // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                      //   imagePath = ImageInfo["serverRelativeUrl"];
                      // } else {
                      //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Access/Attachments/${item.ID}/${ImageInfo.fileName}`;
                      // }

                      return (
                        <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                          <li key={item.ID}>
                            <span className="item-icon">
                              <img src={item.Icon} />
                            </span>
                            <div>
                              <strong>{item.AccessTitle}</strong>
                              <p>{item.AccessDescription}</p>
                            </div>
                          </li>
                        </a>
                      );
                    })
                  }
                </ul>
              </div>
            }

          </div>

        </div>

        <Dialog
          hidden={this.state.AddQuickaccessDialog}
          onDismiss={() =>
            this.setState({
              AddQuickaccessDialog: true,
            })
          }
          dialogContentProps={AddQuickaccessDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1200}
        >

          <div className='Quickbox'>
            <div>
              <h2>Quick Access Information</h2>
            </div>
            <div className='AddquickData'>
              <PrimaryButton className='AddquickInfo' text='Add Data' onClick={() => this.setState({ AddQuickAccessDataDialog: false })} />
            </div>
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>QuickAccess Categories</th>
                  <th>Access Title</th>
                  <th>Access Description</th>
                  <th>Icon</th>
                  <th>Link</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.QuickAccessData.length > 0 &&
                  this.state.QuickAccessData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.QuickAccessCategories}</td>
                        <td>{item.AccessTitle}</td>
                        <td>{item.AccessDescription}</td>
                        <td>
                          {
                            item.Icon ? (
                              <img src={item.Icon} alt="announcement" style={{ width: "80px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Icon"
                            )
                          }
                        </td>
                        <td style={{ wordBreak: "break-all" }}>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditQuickAccessDataDialog: false, CurrentQuickAccessDataID: item.ID }, () => this.EditQuickAccessInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteQuickAccessInfo(item.ID)}
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
          hidden={this.state.AddQuickAccessDataDialog}
          onDismiss={() =>
            this.setState({
              AddQuickAccessDataDialog: true,
              QuickAccessCategories: [],
              AccessTitle: "",
              AccessDescription: "",
              Icon: [],
              Link: "",
              UploadIcon: []
            })
          }
          dialogContentProps={AddQuickaccessDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={900}
        >
          <div>
            <h2>Add QuickAccess Details</h2>
          </div>

          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                  <Dropdown
                    options={this.state.QuickAccessCategorieslist}
                    label='QuickAccess Category'
                    required
                    onChange={(e, option, text) =>
                      this.setState({ QuickAccessCategories: option.text })
                    }
                  />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <TextField
                  label='Access Title'
                  type='text'
                  onChange={(value) =>
                    this.setState({ AccessTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icon</b></label>

                <input className='quickaccessicon'
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleImageChange(e)}
                />

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <TextField
                  label='Access Description'
                  type='text'
                  onChange={(value) =>
                    this.setState({ AccessDescription: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddQuickAccessInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddQuickAccessDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditQuickAccessDataDialog}
          onDismiss={() =>
            this.setState({
              EditQuickAccessDataDialog: true,
              EditQuickAccessCategories: [],
              EditAccessTitle: "",
              EditAccessDescription: "",
              EditIcon: [],
              EditUploadIcon: [],
              EditLink: "",
            })
          }
          dialogContentProps={UpdateQuickaccessDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={900}
        >

          <div>
            <h2>Update QuickAccess Details</h2>
          </div>

          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <Dropdown
                  options={this.state.QuickAccessCategorieslist}
                  label="QuickAccess Category"
                  required
                  defaultSelectedKey={this.state.EditQuickAccessCategories}
                  onChange={(e, option, text) =>
                    this.setState({ EditQuickAccessCategories: option.text })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <TextField
                  label='Access Title'
                  type='text'
                  value={this.state.EditAccessTitle}
                  onChange={(value) =>
                    this.setState({ EditAccessTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icon</b></label>

                <input className='quickaccessicon'
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuickAccessSection'>
              <div className='Add-Form'>
                <TextField
                  label='Access Description'
                  type='text'
                  value={this.state.EditAccessDescription}
                  onChange={(value) =>
                    this.setState({ EditAccessDescription: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateQuickAccessDetails(this.state.CurrentQuickAccessDataID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditQuickAccessDataDialog: true })
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
    this.getQuickAccessData();
    this.GetTicketsChoicesItems();
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

  public async getQuickAccessData() {
    const quickdata = await sp.web.lists.getByTitle("Quick Access").items.select(
      "ID",
      "QuickAccessCategories",
      "AccessTitle",
      "AccessDescription",
      "Icon",
      "Link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(quickdata);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            QuickAccessCategories: item.QuickAccessCategories ? item.QuickAccessCategories : "",
            AccessTitle: item.AccessTitle ? item.AccessTitle : "",
            AccessDescription: item.AccessDescription ? item.AccessDescription : "",
            Icon: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Icon ? JSON.parse(item.Icon).serverRelativeUrl : require(`../assets/baggage.png`),
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ QuickAccessData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Quick Access Data:", error);
    });
  }

 public async AddQuickAccessInfo() {
    if (this.state.AccessTitle.length == 0) {
      alert("Please Enter Details");
    } else {
      const quickaccess = await sp.web.lists.getByTitle("Quick Access").items.add({
        QuickAccessCategories: this.state.QuickAccessCategories,
        AccessTitle: this.state.AccessTitle,
        AccessDescription: this.state.AccessDescription,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null
      });

      if (this.state.UploadIcon && this.state.UploadIcon.length > 0) {

        const file = this.state.UploadIcon[0];

        await sp.web.lists
          .getByTitle("Quick Access")
          .items.getById(quickaccess.data.Id)
          .attachmentFiles.add(file.name, file);
      }


      // this.setState({ AnnouncementsData: announcement });
      this.setState({ AddQuickAccessDataDialog: true });
      this.getQuickAccessData();

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

  public async EditQuickAccessInfo(ID) {
    let EditQuickaccess = this.state.QuickAccessData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditQuickaccess);
    this.setState({
      EditQuickAccessCategories: EditQuickaccess[0].QuickAccessCategories,
      EditAccessTitle: EditQuickaccess[0].AccessTitle,
      EditAccessDescription: EditQuickaccess[0].AccessDescription,
      EditLink: EditQuickaccess[0].Link.Url,
      EditUploadIcon: EditQuickaccess[0].Icon,

    });
  }

  public async UpdateQuickAccessDetails(CurrentQuickAccessDataID) {
    try {
      const updatequickaccess: any = {
        QuickAccessCategories: this.state.EditQuickAccessCategories,
        AccessTitle: this.state.EditAccessTitle,
        AccessDescription: this.state.EditAccessDescription,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Quick Access").items.getById(CurrentQuickAccessDataID).update(updatequickaccess);

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

      if (Array.isArray(this.state.EditUploadIcon) && this.state.EditUploadIcon.length > 0) {

        const file = this.state.EditUploadIcon[0];
  
        const itemRef = sp.web.lists
          .getByTitle("Announcements")
          .items.getById(CurrentQuickAccessDataID);
  
        // delete old attachments
        const attachments = await itemRef.attachmentFiles();
  
        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }
  
        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }
  

      this.setState({ EditQuickAccessDataDialog: true });
      this.getQuickAccessData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteQuickAccessInfo(DeleteQuickAccessDataID) {
    const deleteinfo = await sp.web.lists.getByTitle("Quick Access").items.getById(DeleteQuickAccessDataID).delete();
    this.setState({ QuickAccessData: deleteinfo });
    this.getQuickAccessData();
  }

  public async GetTicketsChoicesItems() {
    const choiceFieldName1 = "QuickAccess Categories";
    const field1 = await sp.web.lists.getByTitle("Quick Access").fields.getByInternalNameOrTitle(choiceFieldName1)();
    let quickaccesscategories = [];
    field1["Choices"].forEach(function (dname, i) {
      quickaccesscategories.push({ key: dname, text: dname });
    });
    this.setState({ QuickAccessCategorieslist: quickaccesscategories });
  }

}
