import * as React from 'react';
import styles from './HomePageQuickLinksHub.module.scss';
import { IHomePageQuickLinksHubProps } from './IHomePageQuickLinksHubProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { Announced, DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePageQuickLinksHubState {
  QuickLinkAllData: any;
  Title: any;
  Icons: any;
  Link: any;
  UploadIcons: any;
  AddQuicklinkDialog: boolean;
  AddQuicklinksDataDialog: boolean;
  previewImage: any;
  EditTitle: any;
  EditIcons: any;
  EditLink: any;
  EditUploadIcons: any;
  EditQuicklinksDataDialog: boolean;
  CurrentQuickLinkDetailsID: any;
  DeleteQuickLinkDataID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddQuickLinkDetailsDialogContentProps = {
  title: "Add Quick Link Details",
};

const AddQuickLinksDataDialogContentProps = {
  title: "Add Quick Links"
};

const UpdateQuickLinkDetailsDialogContentProps = {
  title: "Update Quick Link Details"
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

export default class HomePageQuickLinksHub extends React.Component<IHomePageQuickLinksHubProps, IHomePageQuickLinksHubState> {

  constructor(props: IHomePageQuickLinksHubProps, state: IHomePageQuickLinksHubState) {

    super(props);

    this.state = {
      QuickLinkAllData: "",
      Title: "",
      Icons: [],
      Link: "",
      UploadIcons: [],
      AddQuicklinkDialog: true,
      AddQuicklinksDataDialog: true,
      previewImage: "",
      EditTitle: "",
      EditIcons: [],
      EditLink: "",
      EditUploadIcons: [],
      EditQuicklinksDataDialog: true,
      CurrentQuickLinkDetailsID: "",
      DeleteQuickLinkDataID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IHomePageQuickLinksHubProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageQuickLinksHub">

        {
          this.state.IsAdmin ?
            <>
              <div className='Addquick'>
                <PrimaryButton text='Add Quicklinks' onClick={() => this.setState({ AddQuicklinkDialog: false })} />
              </div>
            </>
            :
            <></>
        }

        <div className="quick-links">

          {
            this.state.QuickLinkAllData.length > 0 &&
            this.state.QuickLinkAllData.map((item) => {
              // let imagePath = "";
              // let ImageInfo = JSON.parse(item.Icons);
              // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
              //   imagePath = ImageInfo["serverRelativeUrl"];
              // }
              // else {
              //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Quick Links Hub/Attachments/${item.ID}/${ImageInfo.fileName}`;
              // }

              return (
                <div className="link-card">
                  <a href={item.Link.Url} style={{ textDecoration: "none" }}>
                    <img src={item.Icons} />
                    <p>{item.Title}</p>
                  </a>
                </div>
              );
            })
          }

        </div>

        <Dialog
          hidden={this.state.AddQuicklinkDialog}
          onDismiss={() =>
            this.setState({
              AddQuicklinkDialog: true,
            })
          }
          dialogContentProps={AddQuickLinkDetailsDialogContentProps}
          modalProps={addmodelProps}
          maxWidth={1200}
        >
          
          <div className='linkhub'>
            <h2>Quick Links Hub Details</h2>
              <div className='AddQuickdata'>
                <PrimaryButton className='AddQuicklnfo' text='Add Data' onClick={() => this.setState({ AddQuicklinksDataDialog: false })} />
              </div>
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Icons</th>
                  <th>Links</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.QuickLinkAllData.length > 0 &&
                  this.state.QuickLinkAllData.map((item) => {

                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Title}</td>
                        <td><img src={item.Icons} /></td>
                        <td style={{ wordBreak: "break-all" }}>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>

                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditQuicklinksDataDialog: false, CurrentQuickLinkDetailsID: item.ID }, () => this.EditAnnouncementInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteQuicklinkinfo(item.ID)}
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
          hidden={this.state.AddQuicklinksDataDialog}
          onDismiss={() =>
            this.setState({
              AddQuicklinksDataDialog: true,
              Title: "",
              Icons: "",
              Link: "",
            })
          }
          dialogContentProps={AddQuickLinksDataDialogContentProps}
          modalProps={addmodelProps2}
          maxWidth={900}
        >

          <div>
            <h2>Add QuickLinks</h2>
          </div>

          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuicklinksInfo'>
              <div className='Add-Form'>
                <TextField
                  label='Quick Title'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Title: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuicklinksInfo'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuicklinksInfo'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icon</b></label>

                <input className='quicklinkicon'
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
                    onClick={() => this.AddQuicklinks()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddQuicklinksDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditQuicklinksDataDialog}
          onDismiss={() =>
            this.setState({
              EditQuicklinksDataDialog: true,
              EditTitle: "",
              EditLink: "",
              EditIcons: [],
              EditUploadIcons: [],
            })
          }
          dialogContentProps={UpdateQuickLinkDetailsDialogContentProps}
          modalProps={updatemodelProps}
          maxWidth={900}
        >
          <div>
            <h2>Update QuickLinks</h2>
          </div>

          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuicklinksInfo'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuicklinksInfo'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 QuicklinksInfo'>
              <div className='Add-Form'>
                <label style={{ display: 'flex' }}><b style={{ fontWeight : '600'}}>Upload Icons</b></label>
                
                <input className='quicklinkicon'
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleUpdateImageChange(e)}
                />

                {
                  this.state.EditUploadIcons && (
                    <div className='Attached-img'>
                      <p>
                        {
                          typeof this.state.EditUploadIcons === "string"
                            ? this.state.EditUploadIcons.split('/').pop()
                            : this.state.EditUploadIcons[0]?.name
                        }
                      </p>
                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadIcons: "" })}
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
                    onClick={() => this.UpdateQuicklinksDetails(this.state.CurrentQuickLinkDetailsID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditQuicklinksDataDialog: true })
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
    this.getquicklinksData();
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

  public async getquicklinksData() {
    const links = await sp.web.lists.getByTitle("Quick Links Hub").items.select(
      "ID",
      "Title",
      "Icons",
      "Link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(links);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Icons: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Icons ? JSON.parse(item.Icons).serverRelativeUrl : require(`../assets/Access.png`),
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ QuickLinkAllData: AllData });
      }
    }).catch((error) => {
      console.log("Error fetching Quick Links Hub data: ", error);
    });
  }

  public async AddQuicklinks() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const quicklink = await sp.web.lists.getByTitle("Quick Links Hub").items.add({
        Title: this.state.Title,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null
      });

      if (this.state.UploadIcons && this.state.UploadIcons.length > 0) {

        const file = this.state.UploadIcons[0];

        await sp.web.lists
          .getByTitle("Quick Links Hub")
          .items.getById(quicklink.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddQuicklinksDataDialog: true });
      this.getquicklinksData();
    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadIcons: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  }

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadIcons: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditAnnouncementInfo(ID) {
    let EditQuicklinksdata = this.state.QuickLinkAllData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditQuicklinksdata);
    this.setState({
      EditTitle: EditQuicklinksdata[0].Title,
      EditLink: EditQuicklinksdata[0].Link.Url,
      EditUploadIcons: EditQuicklinksdata[0].Icons,

    });
  }

  public async UpdateQuicklinksDetails(CurrentQuickLinkDetailsID) {
    try {
      const updateannouncement: any = {
        Title: this.state.EditTitle,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Quick Links Hub").items.getById(CurrentQuickLinkDetailsID).update(updateannouncement);

      if (Array.isArray(this.state.EditUploadIcons) && this.state.EditUploadIcons.length > 0) {

        const file = this.state.EditUploadIcons[0];

        const itemRef = sp.web.lists
          .getByTitle("Quick Links Hub")
          .items.getById(CurrentQuickLinkDetailsID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditQuicklinksDataDialog: true });
      this.getquicklinksData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteQuicklinkinfo(DeleteQuickLinkDataID) {
    const deleteinfo = await sp.web.lists.getByTitle("Quick Links Hub").items.getById(DeleteQuickLinkDataID).delete();
    this.setState({ QuickLinkAllData: deleteinfo });
    this.getquicklinksData();
  }

}
