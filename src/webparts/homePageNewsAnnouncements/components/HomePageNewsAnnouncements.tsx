import * as React from 'react';
import styles from './HomePageNewsAnnouncements.module.scss';
import { IHomePageNewsAnnouncementsProps } from './IHomePageNewsAnnouncementsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DatePicker, Dropdown, Pivot, PivotItem } from 'office-ui-fabric-react';
import * as moment from 'moment';
import { Announced, DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePageNewsAnnouncementsState {
  NewsAnnouncementsData: any;
  NewsFilterdData: any;
  NewsTitle: any;
  NewsPhoto: any;
  NewsCategory: any;
  Categorylist: any;
  NewsDate: any;
  Link: any;
  AddNewsDialog: boolean;
  AddNewsDataDialog: boolean;
  UploadNewsPhoto: any;
  previewImage: any;
  EditNewsTitle: any;
  EditNewsCategory: any;
  EditNewsPhoto: any;
  EditNewsDate: any;
  EditLink: any;
  EditNewsDataDialog: boolean;
  EditUploadNewsPhoto: any;
  EditNewsAnnouncementDataDialog: boolean;
  CurrentNewsAnnouncementID: any;
  DeleteNewsAnnouncementID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddNewsDetailsDialogContentProps = {
  title: "Add News Details",
};

const AddNewsDataDialogContentProps = {
  title: "Add News "
};

const UpdateNewsDetailsDialogContentProps = {
  title: "Update News Details"
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

export default class HomePageNewsAnnouncements extends React.Component<IHomePageNewsAnnouncementsProps, IHomePageNewsAnnouncementsState> {

  constructor(props: IHomePageNewsAnnouncementsProps, state: IHomePageNewsAnnouncementsState) {

    super(props);

    this.state = {
      NewsAnnouncementsData: "",
      NewsFilterdData: "",
      NewsTitle: "",
      NewsPhoto: [],
      NewsCategory: "",
      Categorylist: [],
      NewsDate: "",
      Link: "",
      AddNewsDialog: true,
      AddNewsDataDialog: true,
      UploadNewsPhoto: [],
      previewImage: "",
      EditNewsTitle: "",
      EditNewsCategory: "",
      EditNewsPhoto: [],
      EditNewsDate: "",
      EditLink: "",
      EditNewsDataDialog: true,
      EditUploadNewsPhoto: [],
      EditNewsAnnouncementDataDialog: true,
      CurrentNewsAnnouncementID: "",
      DeleteNewsAnnouncementID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }


  public render(): React.ReactElement<IHomePageNewsAnnouncementsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.homePageNewsAnnouncements} ${hasTeamsContext ? styles.teams : ''}`}>

        <div className="news-panel">

          <div className="news-header">
            <h2 className="section-title">News &amp; Announcements</h2>

            {
              this.state.IsAdmin ?
              <>
                <div className='AddNews'>
                  <PrimaryButton text='Add News' onClick={() => this.setState({ AddNewsDialog: false })} />
                </div>
              </>
              :
              <>
              </>
            }

            <a href='https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/SitePages/News%20&%20Announcements%20Page.aspx' style={{ textDecoration: "none", color: "black" }} target="_blank" rel="noopener noreferrer">
              <button className="view-news">View all</button>
            </a>
          </div>

          <div className="title-underline"></div>

          <div className="news-filters">
            <Pivot onLinkClick={this._onPivotChange}>
              <PivotItem headerText="All News" itemKey="all" />
              <PivotItem headerText="Company" itemKey="company" />
              <PivotItem headerText="Community" itemKey="community" />
              <PivotItem headerText="Charity" itemKey="charity" />
              <PivotItem headerText="Colleagues" itemKey="colleagues" />
              <PivotItem headerText="Contracts" itemKey="contracts" />
              <PivotItem headerText="Case Studies" itemKey="case studies" />
            </Pivot>
          </div>

          <div className='news-scroll'>

            <div className="news-list">

              {
                this.state.NewsFilterdData.length > 0 &&
                this.state.NewsFilterdData.map((item) => {
                  // let imagePath = "";
                  // let ImageInfo = JSON.parse(item.NewsPhoto);
                  // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                  //   imagePath = ImageInfo["serverRelativeUrl"];
                  // }
                  // else {
                  //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/News Announcement/Attachments/${item.ID}/${ImageInfo.fileName}`;
                  // }

                  return (
                    <div className="news-card">
                      <img src={item.NewsPhoto} />

                      <div className="news-content">
                        <p className="news-tag">{item.NewsCategory}</p>
                        <h4>{item.NewsTitle}</h4>
                        <p>{moment(item.NewsDate).format("Do MMMM,YYYY")}</p>
                        <a href={item.Link ? item.Link.Url : ""} style={{ textDecoration: "none", color: "black" }}>View more →</a>
                      </div>
                    </div>
                  );
                })
              }

            </div>

          </div>

          <Dialog
            hidden={this.state.AddNewsDialog}
            onDismiss={() =>
              this.setState({
                AddNewsDialog: true,
              })
            }
            dialogContentProps={AddNewsDetailsDialogContentProps}
            modalProps={addmodelProps}
            minWidth={1500}
          >

            <div className='AddnewsInfo'>
              <PrimaryButton className='AddNewsData' text='Add News' onClick={() => this.setState({ AddNewsDataDialog: false })} />
            </div>

            <div className="news-container">
              <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
                <thead>
                  <tr>
                    <th style={{ width: '20%' }}>NewsTitle</th>
                    <th style={{ width: '30%' }}>NewsCategory</th>
                    <th style={{ width: '15%' }}>NewsPhoto</th>
                    <th style={{ width: '15%' }}>NewsDate</th>
                    <th style={{ width: '15%' }}>Link</th>
                    <th style={{ width: '15%' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>

                  {
                    this.state.NewsAnnouncementsData.length > 0 &&
                    this.state.NewsAnnouncementsData.map((item) => {
                      return (
                        <tr key={item.ID}>
                          <td className="title">{item.NewsTitle}</td>
                          <td>{item.NewsCategory}</td>
                          <td>
                            {
                              item.NewsPhoto ? (
                                <img src={item.NewsPhoto} alt="announcement" style={{ width: "120px", height: "80px", objectFit: "cover" }} />
                              ) : (
                                "No Image"
                              )
                            }
                          </td>
                          <td>{moment(item.NewsDate).format("DD-MM-YYYY")}</td>
                          <td>
                            <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                          </td>

                          <td>
                            <div style={{ display: "flex", gap: "8px" }}>

                              <IconButton
                                iconProps={{ iconName: "Edit" }}
                                title="Edit"
                                ariaLabel="Edit"
                                onClick={() => this.setState({ EditNewsAnnouncementDataDialog: false, CurrentNewsAnnouncementID: item.ID }, () => this.EditNewsAnnouncementInfo(item.ID))}
                              />

                              <IconButton
                                iconProps={{ iconName: "Delete" }}
                                title="Delete"
                                ariaLabel="Delete"
                                onClick={() => this.DeleteNewsAnnouncementInfo(item.ID)}
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
            hidden={this.state.AddNewsDataDialog}
            onDismiss={() =>
              this.setState({
                AddNewsDataDialog: true,
                NewsTitle: "",
                NewsPhoto: [],
                NewsCategory: "",
                NewsDate: "",
                Link: ""
              })
            }
            dialogContentProps={AddNewsDataDialogContentProps}
            modalProps={addmodelProps2}
            minWidth={1100}
          >
            <div className="ms-Grid-row">

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <TextField
                    label='News Title'
                    type='text'
                    onChange={(value) =>
                      this.setState({ NewsTitle: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <Dropdown
                    options={this.state.Categorylist}
                    label='Category'
                    required
                    onChange={(e, option, text) =>
                      this.setState({ NewsCategory: option.text })
                    }
                  />
                </div>
              </div>


              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <label><b>Upload NewsPhoto</b></label><br />

                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e: any) => this.handleImageChange(e)}
                  />

                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <DatePicker
                    label='News Date'
                    allowTextInput={false}
                    value={this.state.NewsDate ? this.state.NewsDate : null}
                    onSelectDate={(date: any) => this.setState({ NewsDate: date })}
                    aria-label="Select a Date" placeholder='Select a News Date' isRequired
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
                      onClick={() => this.AddNewsAnnouncementInfo()}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ AddNewsDataDialog: true })
                      }
                    />
                  </div>

                </div>
              </div>

            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EditNewsAnnouncementDataDialog}
            onDismiss={() =>
              this.setState({
                EditNewsAnnouncementDataDialog: true,
                EditNewsTitle: "",
                EditNewsCategory: "",
                EditNewsPhoto: [],
                EditNewsDate: "",
                EditLink: "",
                EditUploadNewsPhoto: []
              })
            }
            dialogContentProps={UpdateNewsDetailsDialogContentProps}
            modalProps={updatemodelProps}
            minWidth={1100}
          >
            <div className='ms-Grid-row'>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <TextField
                    label='Announcement Title'
                    type='text'
                    value={this.state.EditNewsTitle}
                    onChange={(value) =>
                      this.setState({ EditNewsTitle: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <Dropdown
                    options={this.state.Categorylist}
                    label="News Category"
                    required
                    placeholder="Select News Category"
                    defaultSelectedKey={this.state.EditNewsCategory}
                    onChange={(e, option, text) =>
                      this.setState({ EditNewsCategory: option.text })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <DatePicker
                    label='News Date'
                    allowTextInput={false}
                    value={this.state.EditNewsDate ? this.state.EditNewsDate : null}
                    onSelectDate={(date: any) => this.setState({ EditNewsDate: date })}
                    aria-label="Select a Date" placeholder='Select a News Date' isRequired
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <label><b>Upload NewsPhoto</b></label><br />

                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e: any) => this.handleUpdateImageChange(e)}
                  />

                  {

                    this.state.EditUploadNewsPhoto && (
                      <div className="Attached-img">

                        <p>
                          {
                            typeof this.state.EditUploadNewsPhoto === "string"
                              ? this.state.EditUploadNewsPhoto.split('/').pop()
                              : this.state.EditUploadNewsPhoto[0]?.name
                          }
                        </p>

                        <Icon
                          iconName="Cancel"
                          onClick={() => this.setState({ EditUploadNewsPhoto: "" })}
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
                      onClick={() => this.UpdateNewsAnnouncementDetails(this.state.CurrentNewsAnnouncementID)}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ EditNewsAnnouncementDataDialog: true })
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
    this.getNewsAnnouncementsData();
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

  public async getNewsAnnouncementsData() {

    const items = await sp.web.lists
      .getByTitle("News Announcements")
      .items.select(
        "ID",
        "NewsTitle",
        "NewsPhoto",
        "NewsCategory",
        "NewsDate",
        "Link"
      )
      .expand("AttachmentFiles")
      .orderBy("NewsDate", false)
      .get();

    let formattedData: any[] = [];

    if (items.length > 0) {

      items.forEach((news) => {
        formattedData.push({
          ID: news.ID || "",
          NewsTitle: news.NewsTitle || "",
          NewsPhoto:
            news.AttachmentFiles.length > 0
              ? news.AttachmentFiles[0].ServerRelativeUrl
              : require("../assets/Rectangle1.png"),
          NewsCategory: news.NewsCategory || "",
          NewsDate: news.NewsDate || "",
          Link: news.Link || ""
        });
      });

      // GROUP BY CATEGORY
      const grouped = formattedData.reduce((acc, item) => {
        if (!acc[item.NewsCategory]) {
          acc[item.NewsCategory] = [];
        }
        acc[item.NewsCategory].push(item);
        return acc;
      }, {});

      // TAKE TOP 4 FROM EACH CATEGORY
      let topFourPerCategory: any[] = [];

      Object.keys(grouped).forEach((category) => {
        const top4 = grouped[category].slice(0, 6);
        topFourPerCategory = [...topFourPerCategory, ...top4];
      });

      const reduced = formattedData.reduce((acc: any, item: any) => {
        const category = item.NewsCategory;

        if (
          !acc[category] ||
          new Date(item.Created) > new Date(acc[category].Created)
        ) {
          acc[category] = item;
        }

        return acc;
      }, {});

      // // Convert object → array (ES5 safe)
      // const latestPerCategory: any[] = [];

      // for (let key in reduced) {
      //   if (reduced.hasOwnProperty(key)) {
      //     latestPerCategory.push(reduced[key]);
      //   }
      // }

      this.setState({
        NewsAnnouncementsData: formattedData,
        NewsFilterdData: formattedData
      });

    }
  }

  private _onPivotChange = (item?: PivotItem): void => {
    if (!item) return;

    let filterdata = this.state.NewsAnnouncementsData;

    switch (item.props.itemKey) {

      case "company":
        filterdata = filterdata.filter(t => t.NewsCategory === "Company");
        break;

      case "community":
        filterdata = filterdata.filter(t => t.NewsCategory === "Community");
        break;

      case "charity":
        filterdata = filterdata.filter(t => t.NewsCategory === "Charity");
        break;

      case "colleagues":
        filterdata = filterdata.filter(t => t.NewsCategory === "Colleagues");
        break;

      case "contracts":
        filterdata = filterdata.filter(t => t.NewsCategory === "Contracts");
        break;

      case "case studies":
        filterdata = filterdata.filter(t => t.NewsCategory === "Case Studies");
        break;

      case "all":
      default:
        filterdata = this.state.NewsAnnouncementsData;
    }

    this.setState({ NewsFilterdData: filterdata });
  }

  public async AddNewsAnnouncementInfo() {
    if (this.state.NewsTitle.length == 0) {
      alert("Please Enter Details");
    } else {
      const announcement = await sp.web.lists.getByTitle("News Announcements").items.add({
        NewsTitle: this.state.NewsTitle,
        NewsCategory: this.state.NewsCategory,
        NewsDate: this.state.NewsDate,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null
      });

      if (this.state.UploadNewsPhoto && this.state.UploadNewsPhoto.length > 0) {

        const file = this.state.UploadNewsPhoto[0];

        await sp.web.lists
          .getByTitle("News Announcements")
          .items.getById(announcement.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddNewsDataDialog: true });
      this.getNewsAnnouncementsData();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadNewsPhoto: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  }

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadNewsPhoto: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditNewsAnnouncementInfo(ID) {
    let EditNewsannouncement = this.state.NewsAnnouncementsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditNewsannouncement);
    this.setState({
      EditNewsTitle: EditNewsannouncement[0].NewsTitle,
      EditNewsCategory: EditNewsannouncement[0].NewsCategory,
      EditNewsDate: new Date(EditNewsannouncement[0].NewsDate),
      EditLink: EditNewsannouncement[0].Link.Url,
      EditUploadNewsPhoto: EditNewsannouncement[0].NewsPhoto,
    });
  }

  public async UpdateNewsAnnouncementDetails(CurrentNewsAnnouncementID) {
    try {
      const updateannouncement: any = {
        NewsTitle: this.state.EditNewsTitle,
        NewsCategory: this.state.EditNewsCategory,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null,
        NewsDate: this.state.EditNewsDate
      };

      const updateItem = await sp.web.lists.getByTitle("News Announcements").items.getById(CurrentNewsAnnouncementID).update(updateannouncement);

      if (Array.isArray(this.state.EditNewsPhoto) && this.state.EditNewsPhoto.length > 0) {
        const file = this.state.EditNewsPhoto[0];

        const itemRef = sp.web.lists
          .getByTitle("News Announcements")
          .items.getById(CurrentNewsAnnouncementID);

        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        await itemRef.attachmentFiles.add(file.name, file);
      }

      this.setState({ EditNewsAnnouncementDataDialog: true });
      this.getNewsAnnouncementsData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteNewsAnnouncementInfo(DeleteNewsAnnouncementID) {
    const deleteinfo = await sp.web.lists.getByTitle("News Announcements").items.getById(DeleteNewsAnnouncementID).delete();
    this.setState({ NewsAnnouncementsData: deleteinfo });
    this.getNewsAnnouncementsData();
  }

  public async GetTicketsChoicesItems() {
    const choiceFieldName1 = "News Category";
    const field1 = await sp.web.lists.getByTitle("News Announcements").fields.getByInternalNameOrTitle(choiceFieldName1)();
    let categorylist = [];
    field1["Choices"].forEach(function (dname, i) {
      categorylist.push({ key: dname, text: dname });
    });
    this.setState({ Categorylist: categorylist });
  }

}
