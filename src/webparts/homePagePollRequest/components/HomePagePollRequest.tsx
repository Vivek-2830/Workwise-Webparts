import * as React from 'react';
import styles from './HomePagePollRequest.module.scss';
import { IHomePagePollRequestProps } from './IHomePagePollRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { Chart } from 'chart.js';
import { DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePagePollRequestState {
  options: any;
  question: any;
  hasVoted: boolean;
  selectedOption: any;
  TotalResponses: any;
  counts: any;
  SurveyData: any;
  SurveyResponseData: any;
  UserWidgetData: any;
  UseFullAppsData: any
  Title: any;
  Icon: any;
  Link: any;
  UploadIcon: any;
  EditTitle: any;
  EditIcon: any;
  EditLink: any;
  EditUploadIcon: any;
  AddUserwidgetDialog: boolean;
  AdduserwidgetDataDialog: boolean;
  EdituserwidgetDataDialog: boolean;
  AppName: any;
  AppIcon: any;
  Applink: any;
  UploadAppIcon: any;
  EditAppName: any;
  EditAppIcon: any;
  EditApplink: any;
  EditUploadAppIcon: any;
  AddUsefullDialog: any;
  AddUsefullDataDialog: any;
  EditUserfullDataDialog: boolean;
  IsAdmin: boolean;
  CurrentUserEmail: any;
  previewIcon: any;
  previewAppIcon: any;
  CurrentUserWidgetDataID: any;
  DeleteuserWidgetDataID: any;
  CurrentUsefullappDataID: any;
  DeleteUsefullappDataID: any;
}

require('../assets/style.css');

const AddUserwidgetDataDialogContentProps = {
  title: "Add Details",
};

const AddUserWidgetDetailsContentProps = {
  title: "Add User Widget Details"
}

const AddUseFullDataDialogContentProps = {
  title: "Add UseFullApp Details"
}

const UpdateUserWidgetDetailsDialogContentProps = {
  title: "Update User Widget Details"
}

const UpdateUsefullappDetailsDialogContentProps = {
  title: "Update UseFullApp Details"
}

const updatemodelProps = {
  className: "Update-Dialog"
};

const updatemodelProps2 = {
  className: "Update-Data-Dialog"
}

const addmodelProps = {
  className: "Add-Dialog"
};

const addmodelProps2 = {
  className: "Add-Data-Dialog"
}

const addmodelProps3 =  {
  className: "Add-UseFull-Dialog"
}


export default class HomePagePollRequest extends React.Component<IHomePagePollRequestProps, IHomePagePollRequestState> {

  constructor(props: IHomePagePollRequestProps, state: IHomePagePollRequestState) {
    super(props);

    this.state = {
      options: [],
      question: "",
      hasVoted: false,
      selectedOption: "",
      TotalResponses: "",
      counts: "",
      SurveyData: "",
      SurveyResponseData: "",
      UserWidgetData: "",
      UseFullAppsData: "",
      Title: "",
      Icon: [],
      Link: "",
      UploadIcon: [],
      EditTitle: "",
      EditIcon: [],
      EditLink: "",
      EditUploadIcon: [],
      AddUserwidgetDialog: true,
      AdduserwidgetDataDialog: true,
      EdituserwidgetDataDialog: true,
      AppName: "",
      AppIcon: [],
      Applink: "",
      UploadAppIcon: [],
      EditAppName: "",
      EditAppIcon: [],
      EditApplink: "",
      EditUploadAppIcon: [],
      AddUsefullDialog: true,
      AddUsefullDataDialog: true,
      EditUserfullDataDialog: true,
      IsAdmin: false,
      CurrentUserEmail: "",
      previewAppIcon: "",
      previewIcon: "",
      CurrentUserWidgetDataID: "",
      DeleteuserWidgetDataID: "",
      CurrentUsefullappDataID: "",
      DeleteUsefullappDataID: ""
    };

  }

  public render(): React.ReactElement<IHomePagePollRequestProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePagePollRequest">

        <div className="Poll-card">
          <div className="Poll-header">
            <h3>Poll</h3>
            <span className="arrow">⌃</span>
          </div>


          {
            this.state.hasVoted ?
              <>
                <h3>{this.state.question}</h3>

                <canvas id="pollChart" height="250"></canvas>

              </> : <>

                <h3>{this.state.question}</h3>

                {
                  this.state.options.map((opt, index) => {

                    return (
                      <div key={index} style={{ marginBottom: 8 }}>
                        <input
                          type="radio"
                          name="poll"
                          value={opt}
                          onChange={(e) => this.setState({ selectedOption: e.target.value })}
                        />
                        <span style={{ marginLeft: 8 }}>{opt}</span>
                      </div>
                    );

                  })
                }

                <button
                  style={{
                    marginTop: 12,
                    padding: "6px 16px",
                    cursor: "pointer"
                  }}
                  className='PollButton'
                  onClick={() => this.submitVote()}
                >
                  Submit
                </button>

              </>
          }

        </div>

        {/* ---------------------------------------------------------------- */}

        <div className="user-widget">

          <h3 className="hello">Hello {userDisplayName}</h3>
          <div className="hello-underline"></div>

          {
            this.state.IsAdmin ?
            <>
                <div>
                  <PrimaryButton text='Add' onClick={() => this.setState({ AddUserwidgetDialog: false })} />
                </div>
            </>
            :
            <>
            </>
          }

          {
            this.state.UserWidgetData.length > 0 &&
            this.state.UserWidgetData.map((item) => {
              return (
                <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                  <div className="stat-card">
                    <div className='context'>
                      <img src={item.Icon} /><span> {item.Title}</span>
                    </div>
                    {/* <strong>3</strong> */}
                  </div>
                </a>
              );
            })
          }

          <Dialog
            hidden={this.state.AddUserwidgetDialog}
            onDismiss={() =>
              this.setState({
                AddUserwidgetDialog: true,
              })
            }
            dialogContentProps={AddUserwidgetDataDialogContentProps}
            modalProps={addmodelProps}
            minWidth={1500}
          >

            {/* -----------------------------Useful Widget-------------------------------- */}

            <div className='AddAnnouncmentData'>
              <PrimaryButton className='AddAnnounInfo' text='Add User Widget' onClick={() => this.setState({ AdduserwidgetDataDialog: false })} />
            </div>

            <div className="news-container">
              <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
                <thead>
                  <tr>
                    <th style={{ width: '20%' }}>Title</th>
                    <th style={{ width: '30%' }}>Icon</th>
                    <th style={{ width: '30%' }}>Link</th>
                    <th style={{ width: '15%' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>

                  {
                    this.state.UserWidgetData.length > 0 &&
                    this.state.UserWidgetData.map((item) => {
                      return (
                        <tr key={item.ID}>
                          <td className="title">{item.Title}</td>
                          <td>
                            {
                              item.Icon ? (
                                <img src={item.Icon} alt="announcement" style={{ width: "30px", height: "30px", objectFit: "cover" }} />
                              ) : (
                                "No Icon"
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
                                onClick={() => this.setState({ EdituserwidgetDataDialog: false, CurrentUserWidgetDataID: item.ID }, () => this.EditUserWidgetsInfo(item.ID))}
                              />

                              <IconButton
                                iconProps={{ iconName: "Delete" }}
                                title="Delete"
                                ariaLabel="Delete"
                                onClick={() => this.DeleteUserWidgetsDetail(item.ID)}
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

            {/* -----------------------------Useful Apps-------------------------------- */}
              
            <div className='AddAnnouncmentData'>
                
              <h2>Useful Apps</h2>

              <PrimaryButton className='AddAnnounInfo' text='Add Apps' onClick={() => this.setState({ AddUsefullDataDialog: false })} />
            </div>

            <div className="news-container">
              <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
                <thead>
                  <tr>
                    <th style={{ width: '20%' }}>AppName</th>
                    <th style={{ width: '30%' }}>AppIcon</th>
                    <th style={{ width: '30%' }}>Applink</th>
                    <th style={{ width: '15%' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>

                  {
                    this.state.UseFullAppsData.length > 0 &&
                    this.state.UseFullAppsData.map((item) => {
                      return (
                        <tr key={item.ID}>
                          <td className="title">{item.AppName}</td>
                          <td>
                            {
                              item.AppIcon ? (
                                <img src={item.AppIcon} alt="announcement" style={{ width: "30px", height: "30px", objectFit: "cover" }} />
                              ) : (
                                "No AppIcon"
                              )
                            }
                          </td>
                          <td>
                            <a href={item.Applink.Url} target="_blank" rel="noopener noreferrer">{item.Applink.Description}</a>
                          </td>

                          <td>
                            <div style={{ display: "flex", gap: "8px" }}>

                              <IconButton
                                iconProps={{ iconName: "Edit" }}
                                title="Edit"
                                ariaLabel="Edit"
                                onClick={() => this.setState({ EditUserfullDataDialog: false, CurrentUsefullappDataID: item.ID }, () => this.EditUsefullappInfo(item.ID))}
                              />

                              <IconButton
                                iconProps={{ iconName: "Delete" }}
                                title="Delete"
                                ariaLabel="Delete"
                                onClick={() => this.DeleteusefullappDetail(item.ID)}
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

          {/* -----------------------------Useful Widget-------------------------------- */}
          <Dialog
            hidden={this.state.AdduserwidgetDataDialog}
            onDismiss={() =>
              this.setState({
                AdduserwidgetDataDialog: true,
                Title: "",
                Icon: "",
                Link: ""
              })
            }
            dialogContentProps={AddUserWidgetDetailsContentProps}
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
                  <label><b>Upload Image</b></label><br />
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e: any) => this.handleUserWidgetIcomChange(e)}
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
                      onClick={() => this.AddUserwidgetInfo()}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ AdduserwidgetDataDialog: true })
                      }
                    />
                  </div>

                </div>
              </div>

            </div>
          </Dialog>

          {/* -----------------------------Useful Widget-------------------------------- */}
          <Dialog
            hidden={this.state.EdituserwidgetDataDialog}
            onDismiss={() =>
              this.setState({
                EdituserwidgetDataDialog: true,
                EditTitle: "",
                EditLink: "",
                EditIcon: "",
                EditUploadIcon: [],
              })
            }
            dialogContentProps={UpdateUserWidgetDetailsDialogContentProps}
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
                  <label><b>Upload Image</b></label><br />

                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e: any) => this.handleUpdateUserWidgetIconChange(e)}
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
                      onClick={() => this.UpdateUserWidgetstDetails(this.state.CurrentUserWidgetDataID)}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ EdituserwidgetDataDialog: true })
                      }
                    />
                  </div>

                </div>
              </div>


            </div>

          </Dialog>

          <h4>Useful Apps</h4>

          <div className="apps-grid">

            {
              this.state.UseFullAppsData.length > 0 &&
              this.state.UseFullAppsData.map((item) => {
                return (
                  <a href={item.Applink.Url} style={{ textDecoration: "none" }}>
                    <div className="app-card">{item.AppName} <img className='next-i' src={item.AppIcon} /></div>
                  </a>
                );
              })
            }

            {/* <a href='https://axiseurope.crm4.dynamics.com/main.aspx?appid=9fa6e94b-63a5-4a31-89d6-6298402f0d3e&pagetype=dashboard&type=system&_canOverride=true' style={{ textDecoration: "none" }}><div className="app-card">Dynamics CE <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://uk.sheassure.net/clc' style={{ textDecoration: "none" }}><div className="app-card">Evotix <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://bit.ly/4l6gNQc' style={{ textDecoration: "none" }}><div className="app-card" >Outlook <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://go.accessacloud.com/o/repbp/workspaces/98d34671c16d4d2e9e1429a2fd965ec2/Access.PeopleXDEmpMain/2f1f1cba97924b2b891fc2f51a13677a?location=https%3A%2F%2Fmy.xd.accessacloud.com%2Fpls%2Fcoreportal_repbp%2Fi%23EmpMain%2Fmytime' style={{ textDecoration: "none" }}><div className="app-card">PeopleXD <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://teams.cloud.microsoft/' style={{ textDecoration: "none" }}><div className="app-card">Teams <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://servicedesk.axisclc.com/' style={{ textDecoration: "none" }}><div className="app-card">Halo <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://go.accessacloud.com/o/repbp/workspaces/28609fedc58441bdbf7a8a4cbe52b1c7/Access.Product.Learning/f899dafaf580404cb513eeab0849d751?location=https%3A%2F%2Faxisclcgroup.lms.accessacloud.com%2Fw%2Fhome' style={{ textDecoration: "none" }}><div className="app-card">Training (LMS) <img className='next-i' src={require("../assets/icon.png")} /></div></a> */}
          </div>

          <Dialog
            hidden={this.state.AddUsefullDataDialog}
            onDismiss={() =>
              this.setState({
                AddUsefullDataDialog: true,
                Title: "",
                Icon: "",
                Link: ""
              })
            }
            dialogContentProps={AddUseFullDataDialogContentProps}
            modalProps={addmodelProps3}
            minWidth={1100}
          >
            <div className="ms-Grid-row">

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <TextField
                    label='AppName'
                    type='text'
                    onChange={(value) =>
                      this.setState({ AppName: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <label><b>Upload AppIcon</b></label><br />
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e: any) => this.handleUseFullappIconChange(e)}
                  />

                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <TextField
                    label='App Link'
                    type='text'
                    onChange={(value) =>
                      this.setState({ Applink: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Announcement-Submit'>
                  <div className='Submit-Button'>
                    <PrimaryButton
                      text='Submit'
                      onClick={() => this.AddUseFullAppInfo()}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ AddUsefullDataDialog: true })
                      }
                    />
                  </div>

                </div>
              </div>

            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EditUserfullDataDialog}
            onDismiss={() =>
              this.setState({
                EditUserfullDataDialog: true,
                EditAppName: "",
                EditApplink: "",
                EditAppIcon: "",
                EditUploadAppIcon: [],
              })
            }
            dialogContentProps={UpdateUsefullappDetailsDialogContentProps}
            modalProps={updatemodelProps2}
            minWidth={1100}
          >
            <div className='ms-Grid-row'>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <TextField
                    label='Title'
                    type='text'
                    value={this.state.EditAppName}
                    onChange={(value) =>
                      this.setState({ EditAppName: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-Form'>
                  <label><b>Upload AppIcon</b></label><br />

                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e: any) => this.handleUpdateUserWidgetIconChange(e)}
                  />

                  {
                    this.state.EditUploadAppIcon && (
                      <div className="Attached-img">

                        {/* ✅ Handle BOTH string + file */}
                        <p>
                          {
                            typeof this.state.EditUploadAppIcon === "string"
                              ? this.state.EditUploadAppIcon.split('/').pop()
                              : this.state.EditUploadAppIcon[0]?.name
                          }
                        </p>

                        <Icon
                          iconName="Cancel"
                          onClick={() => this.setState({ EditUploadAppIcon: "" })}
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
                    value={this.state.EditApplink}
                    onChange={(value) =>
                      this.setState({ EditApplink: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Announcement-Submit'>
                  <div className='Submit-Button'>
                    <PrimaryButton
                      text='Update'
                      onClick={() => this.UpdateusefullappDetails(this.state.CurrentUsefullappDataID)}
                    />
                  </div>

                  <div className='Cancel-Button'>
                    <DefaultButton
                      text='Cancel'
                      onClick={() =>
                        this.setState({ EditUserfullDataDialog: true })
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
    this.getUserFullAppsData();
    this.getUserwidgetInfo();
    this.GetCurrentUser();

    await this.loadSurvey();
    await this.checkUserVote();
    await this.loadResults();
    this.getSurveyInfo();


    Chart.pluginService.register({
      beforeDraw: function (chart) {
        if (chart.config.options.centerText) {
          const ctx = chart.chart.ctx;
          const txt = chart.config.options.centerText;

          ctx.save();
          ctx.font = "bold 18px Arial";
          ctx.fillStyle = "#ffffff";
          ctx.textAlign = "center";
          ctx.textBaseline = "middle";

          const centerX = chart.chart.width / 2;
          const centerY = chart.chart.height / 1.8;

          ctx.fillText(txt, centerX, centerY);
          ctx.restore();
        }
      }
    });

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

  private async loadSurvey() {

    const items = await sp.web.lists
      .getByTitle("Poll")
      .items
      .select("Id", "Question", "Options")
      .top(1)
      .get();

    if (items.length === 0) return;

    const item = items[0];

    /* If Options column is Choice (single) with semicolon values */
    let optionsArray: string[] = [];

    if (item.Options) {
      optionsArray = item.Options;
    }

    this.setState({
      question: item.Question,
      options: optionsArray
    });

  }

  /* Check if user already voted */
  private async checkUserVote() {

    const userId = this.props.context.pageContext.legacyPageContext.userId;

    const items = await sp.web.lists
      .getByTitle("Poll Response")
      .items
      .select("Id", "Author/Id", "Title")
      .expand("Author")
      .filter(`Author/Id eq ${userId} and Title eq '${this.state.question}'`)
      .get();

    if (items.length > 0) {
      this.setState({ hasVoted: true });
    }

  }

  private async submitVote() {

    if (!this.state.selectedOption) {
      alert("Please select option");
      return;
    }

    const email = this.props.context.pageContext.user.email;

    await sp.web.lists.getByTitle("Poll Response").items.add({
      Title: this.state.question,
      // UserEmail: email,
      Option: this.state.selectedOption
    });

    this.setState({ hasVoted: true });

    await this.loadResults();
  }

  /* Load aggregated results */
  // private async loadResults() {

  //   const items = await sp.web.lists
  //     .getByTitle("Survey Response")
  //     .items
  //     .select("Option")
  //     .get();

  //   let counts = [0, 0, 0, 0];

  //   items.forEach((item: any) => {

  //     if (item.Option === "test 1") counts[0]++;
  //     if (item.Option === "test 2") counts[1]++;
  //     if (item.Option === "Option 3") counts[2]++;
  //     if (item.Option === "Option 4") counts[3]++;

  //   });

  //   this.setState({ counts: counts }, () => {

  //     if (this.state.hasVoted) {
  //       this.renderChart();
  //     }

  //   });
  // }

  private async loadResults() {

    const items = await sp.web.lists
      .getByTitle("Poll Response")
      .items
      .select("Option")
      .get();

    let counts: number[] = [];

    for (let i = 0; i < this.state.options.length; i++) {
      counts.push(0);
    }
    items.forEach((item: any) => {

      this.state.options.forEach((opt, index) => {
        if (item.Option === opt) {
          counts[index]++;
        }
      });

    });

    this.setState({ counts: counts, TotalResponses: items.length }, () => {

      if (this.state.hasVoted && this.state.options.length > 0) {
        this.renderChart();
      }

    });

  }

  /* Render Pie Chart */
  private renderChart() {

    const backgroundColors = [
      'rgba(95, 255, 214, 0.6)',
      'rgba(217, 198, 255, 0.6)',
      'rgba(255, 210, 168, 0.6)',
      'rgba(41, 199, 217, 0.6)',
    ];

    setTimeout(() => {

      const canvas = document.getElementById("pollChart") as HTMLCanvasElement;
      if (!canvas) return;

      new Chart(canvas, {
        type: 'doughnut',
        data: {
          labels: this.state.options,
          datasets: [{
            data: this.state.counts,
            backgroundColor: backgroundColors.slice(0, this.state.options.length),
          }]
        },
        options: {
          cutoutPercentage: 80,
          legend: {
            labels: {
              fontColor: "#ffffff"
            }
          },
          centerText: this.state.TotalResponses + " Responses"
        }
      });

    }, 300);
  }

  public async getSurveyInfo() {
    const poll = await sp.web.lists.getByTitle("Poll").items.select(
      "ID",
      "Question",
      "Answer",
      "Options"
    ).get().then((data) => {
      let AllData = [];
      console.log(poll);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Question: item.Question ? item.Question : "",
            Answer: item.Answer ? item.Answer : "",
            Options: item.Options ? item.Options : ""
          });
        });
        this.setState({ SurveyData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching details ", error);
    });
  }

  public async getSurveyresponse() {
    const response = await sp.web.lists.getByTitle("Poll Response").items.select(
      "ID",
      "Title",
      "Option",
      "PersonName"
    ).get().then((data) => {
      let AllData = [];
      console.log(response);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Option: item.Option ? item.Option : "",
            PersonName: item.PersonName ? item.PersonName : ""
          });
        });
        this.setState({ SurveyResponseData: AllData });
      }
    });
  }

  public async getUserwidgetInfo() {
    const widgetdata = await sp.web.lists.getByTitle("User widget").items.select(
      "ID",
      "Title",
      "Icon",
      "Link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(widgetdata);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Icon: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Icon ? JSON.parse(item.Icon).serverRelativeUrl : require(`../assets/ticket01.png`),
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ UserWidgetData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching details ", error);
    });
  }

  public async getUserFullAppsData() {
    const apps = await sp.web.lists.getByTitle("Useful Apps").items.select(
      "ID",
      "AppName",
      "AppIcon",
      "Applink"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(apps);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            AppName: item.AppName ? item.AppName : "",
            AppIcon: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.AppIcon ? JSON.parse(item.AppIcon).serverRelativeUrl : require(`../assets/Icon.png`),
            Applink: item.Applink ? item.Applink : ""
          });
        });
        this.setState({ UseFullAppsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching details ", error);
    });
  }

  public async AddUserwidgetInfo() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const userwidget = await sp.web.lists.getByTitle("User widget").items.add({
        Title: this.state.Title,
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
          .getByTitle("User widget")
          .items.getById(userwidget.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AdduserwidgetDataDialog: true });
      this.getUserwidgetInfo();

    }
  }
  
  handleUserWidgetIcomChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadIcon: [file],
        previewIcon: URL.createObjectURL(file)
      });
    }
  };

  handleUpdateUserWidgetIconChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadIcon: [file],
      });
    }
  
  }

  public async EditUserWidgetsInfo(ID) {
    let Edituserwidgets = this.state.UserWidgetData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(Edituserwidgets);
    this.setState({
      EditTitle: Edituserwidgets[0].Title,
      EditLink: Edituserwidgets[0].Link.Url,
      EditUploadIcon: Edituserwidgets[0].Icon
    });
  }

  public async UpdateUserWidgetstDetails(CurrentUserWidgetDataID) {
    try {
      const updateuserWidgets: any = {
        Title: this.state.EditTitle,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("User widget").items.getById(CurrentUserWidgetDataID).update(updateuserWidgets);

      if (Array.isArray(this.state.EditUploadIcon) && this.state.EditUploadIcon.length > 0) {

        const file = this.state.EditUploadIcon[0];
  
        const itemRef = sp.web.lists
          .getByTitle("User widget")
          .items.getById(CurrentUserWidgetDataID);
  
        // delete old attachments
        const attachments = await itemRef.attachmentFiles();
  
        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }
  
        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }
  

      this.setState({ EdituserwidgetDataDialog: true });
      this.getUserwidgetInfo();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteUserWidgetsDetail(DeleteuserWidgetDataID) {
    const deleteinfo = await sp.web.lists.getByTitle("User widget").items.getById(DeleteuserWidgetDataID).delete();
    this.setState({ UserWidgetData: deleteinfo });
    this.getUserwidgetInfo()
  }

  public async AddUseFullAppInfo() {
    if (this.state.Title.length == 0) {
      alert("Please Enter Details");
    } else {
      const useapp = await sp.web.lists.getByTitle("Useful Apps").items.add({
        AppName: this.state.AppName,
        Applink: this.state.Applink
          ? {
            Url: this.state.Applink,
            Description: this.state.Applink
          }
          : null
      });

      if (this.state.UploadAppIcon && this.state.UploadAppIcon.length > 0) {

        const file = this.state.UploadAppIcon[0];

        await sp.web.lists
          .getByTitle("Useful Apps")
          .items.getById(useapp.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddUsefullDataDialog: true });
      this.getUserFullAppsData();

    }
  }
  
  handleUseFullappIconChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadAppIcon: [file],
        previewAppIcon: URL.createObjectURL(file)
      });
    }
  };

  handleUpdateUseFullappIconChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        EditUploadAppIcon: [file],
      });
    }
  
  }

  public async EditUsefullappInfo(ID) {
    let Editusefullapps = this.state.UseFullAppsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(Editusefullapps);
    this.setState({
      EditAppName: Editusefullapps[0].AppName,
      EditApplink: Editusefullapps[0].Applink.Url,
      EditUploadAppIcon: Editusefullapps[0].AppIcon
    });
  }

  public async UpdateusefullappDetails(CurrentUsefullappDataID) {
    try {
      const updateuseFullapps: any = {
        AppName: this.state.EditAppName,
        Applink: this.state.EditApplink ? {
          Url: this.state.EditApplink,
          Description: this.state.EditApplink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Useful Apps").items.getById(CurrentUsefullappDataID).update(updateuseFullapps);

      if (Array.isArray(this.state.EditUploadAppIcon) && this.state.EditUploadAppIcon.length > 0) {

        const file = this.state.EditUploadAppIcon[0];
  
        const itemRef = sp.web.lists
          .getByTitle("Useful Apps")
          .items.getById(CurrentUsefullappDataID);
  
        // delete old attachments
        const attachments = await itemRef.attachmentFiles();
  
        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }
  
        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }
  

      this.setState({ EditUserfullDataDialog: true });
      this.getUserFullAppsData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteusefullappDetail(DeleteUsefullappDataID) {
    const deletapps = await sp.web.lists.getByTitle("Useful Apps").items.getById(DeleteUsefullappDataID).delete();
    this.setState({ UserWidgetData: deletapps });
    this.getUserFullAppsData()
  }

}





{/* <a href="https://axiseuropeplc.sharepoint.com/sites/AxisLMS/SitePages/My-Training-Dashboard.aspx?isSPOFile=1" style={{ textDecoration: "none", color: "black" }}>
            <div className="stat-card">
              <div className='context'>
                <img src={require("../assets/checkcirclebroken.png")} /><span> My Training</span>
              </div>
              <strong>3</strong>
            </div>
              </a> 

              <div className="stat-card">
                <div className='context'>
                  <img src={require("../assets/icon2.png")} /><span> My Approvels</span>
                </div>
                <strong>4</strong>
              </div>

              <a href='https://servicedesk.axisclc.com/portal/tickets?btn=60&viewid=1' style={{ textDecoration: "none", color: "black" }}>
              <div className="stat-card">
                <div className='context'>
                  <img src={require("../assets/ticket01.png")} /><span> My IT Tickets</span>
                </div>
                <strong>5</strong>
              </div>
              </a> 

              <h4>My Favorite Articles</h4>

              <ul className="fav-list">
                <li>Better Understanding your patients needs</li>
                <li>401k Updates fpr 2020</li>
                <li className="active">Covid Frequently Asked Questions</li>
                <li>HR Polices and Procedures Guidelines</li>
              </ul>

              <a className="views-all" href="#">View all</a> 
          */}