import * as React from 'react';
import styles from './EmployeehubInternalVacancies.module.scss';
import { IEmployeehubInternalVacanciesProps } from './IEmployeehubInternalVacanciesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import { sp } from '@pnp/sp/presets/all';
import { Dialog, IconButton, PrimaryButton } from 'office-ui-fabric-react';

export interface IEmployeehubInternalVacanciesState {
  InternalVacanciesData: any;
  JobTitle: any;
  Department: any;
  LocationType: any;
  EmploymentType: any;
  Salary: any;
  KeyRequirements: any;
  ApplicationDeadline: any;
  Link: any;
  AddInternalVacancieDialog: boolean;
  AddInternalvacancieDataDialog: boolean;
  EditJobTitle: any;
  EditDepartment: any;
  EditLocationType: any;
  EditEmploymentType: any;
  EditSalary: any;
  EditKeyRequirements: any;
  EditApplicationDeadline: any;
  EditLink: any;
  EditInternalVacanciesDialog: boolean;
  CurrentInternalVacanciesItemID: any;
  DeleteVacanciesItemID: any;
}

require('../assets/style.css');

const AddInternalVacanciesDetailsDialogContentProps = {
  title: "Add InternalVacancies Details",
};

const AddEInternalVacanciesDataDialogContentProps = {
  title: "Add InternalVacancies"
}

const UpdateInternalVacanciesDetailsDialogContentProps = {
  title: "Update InternalVacancies Details"
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

export default class EmployeehubInternalVacancies extends React.Component<IEmployeehubInternalVacanciesProps, IEmployeehubInternalVacanciesState> {

  constructor(props: IEmployeehubInternalVacanciesProps, state: IEmployeehubInternalVacanciesState) {

    super(props);

    this.state = {
      InternalVacanciesData: "",
      JobTitle: "",
      Department: "",
      LocationType: "",
      EmploymentType: "",
      Salary: "",
      KeyRequirements: "",
      ApplicationDeadline: "",
      Link: "",
      AddInternalVacancieDialog: true,
      AddInternalvacancieDataDialog: true,
      EditJobTitle: "",
      EditDepartment: "",
      EditLocationType: "",
      EditEmploymentType: "",
      EditSalary: "",
      EditKeyRequirements: "",
      EditApplicationDeadline: "",
      EditLink: "",
      EditInternalVacanciesDialog: true,
      CurrentInternalVacanciesItemID: "",
      DeleteVacanciesItemID: "",
    };

  }


  public render(): React.ReactElement<IEmployeehubInternalVacanciesProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubInternalVacancies">

        <div className="jobs-wrapper">

          <div className="jobs-header">
            <div>
              <h3>Internal Vacancies</h3>
              <span className="underline"></span>
            </div>
            <a href='https://www.axisclc.com/work-with-us/' style={{ textDecoration: 'none', color: 'black' }}><button className="view-all">View all</button></a>
          </div>

          <div className="jobs-grid">

            {
              this.state.InternalVacanciesData.length > 0 &&
              this.state.InternalVacanciesData.map((item) => {
                return (
                  <div className="job-card">

                    <h4>{item.JobTitle}</h4>
                    <p className="dept">{item.Department}</p>

                    <div className="job-info">
                      <div className="info-row">
                        <img src={require('../assets/markerpin01.png')} />
                        <span>Location: {item.LocationType}</span>
                      </div>
                      <div className="info-row">
                        <img src={require('../assets/currencydollarcircle.png')} />
                        <span>Salary: {item.Salary}</span>
                      </div>
                      <div className="info-row">
                        <img src={require('../assets/alarmclock.png')} />
                        <span>Type: {item.EmploymentType}</span>
                      </div>
                    </div>

                    <h5>Key Requirements</h5>

                    <div className="req">
                      <div className="req-row">
                        {/* <img src={require('../assets/check.png')} /> */}
                        <span dangerouslySetInnerHTML={{ __html: item.KeyRequirements }} />
                      </div>

                    </div>

                    <p className="deadline">Application Deadline: {moment(item.ApplicationDeadline).format("MMM DD, YYYY")}</p>

                    <a href={item.Link.Url} style={{ textDecoration: "none" }}><button className="apply-btn">Apply Now</button></a>

                  </div>
                );
              })
            }

          </div>
        </div>

        <Dialog
          hidden={this.state.AddInternalVacancieDialog}
          onDismiss={() =>
            this.setState({
              AddInternalVacancieDialog: true,
            })
          }
          dialogContentProps={AddInternalVacanciesDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddvacancieData'>
            <PrimaryButton className='AddInternalInfo' text='Add Data' onClick={() => this.setState({ AddInternalvacancieDataDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>JobTitle</th>
                  <th style={{ width: '30%' }}>Department</th>
                  <th style={{ width: '30%' }}>LocationType</th>
                  <th style={{ width: '15%' }}>EmploymentType</th>
                  <th style={{ width: '15%' }}>Salary</th>
                  <th style={{ width: '15%' }}>KeyRequirements</th>
                  <th style={{ width: '15%' }}>ApplicationDeadline</th>
                  <th style={{ width: '15%' }}>Link</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.InternalVacanciesData.length > 0 &&
                  this.state.InternalVacanciesData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.JobTitle}</td>
                        <td>{item.Department}</td>
                        <td>{item.LocationType}</td>
                        <td>{item.EmploymentType}</td>
                        <td>{item.Salary}</td>
                        <td>{item.KeyRequirements}</td>
                        <td>{item.ApplicationDeadline}</td>
                        <td>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>
                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            {/* <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditInternalVacanciesDialog: false, CurrentInternalVacanciesItemID: item.ID }, () => this.EditEmpAnnouncement(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteEmpAnnouncementInfo(item.ID)}
                            /> */}

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

        {/* <Dialog
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
        </Dialog> */}

        {/* <Dialog
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

        </Dialog> */}

      </section>
    );
  }

  public async componentDidMount() {
    this.getInternalVacancies();
  }

  public async getInternalVacancies() {
    const vacancies = await sp.web.lists.getByTitle("Internal Vacancies").items.select(
      "ID",
      "JobTitle",
      "Department",
      "LocationType",
      "EmploymentType",
      "ReportsTo",
      "Salary",
      "KeyRequirements",
      "ApplicationDeadline",
      "Link"
    ).get().then((data) => {
      let AllData = [];
      console.log(vacancies);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            JobTitle: item.JobTitle ? item.JobTitle : "",
            Department: item.Department ? item.Department : "",
            LocationType: item.LocationType ? item.LocationType : "",
            EmploymentType: item.EmploymentType ? item.EmploymentType : "",
            ReportsTo: item.ReportsTo ? item.ReportsTo : "",
            Salary: item.Salary ? item.Salary : "",
            KeyRequirements: item.KeyRequirements ? item.KeyRequirements : "",
            ApplicationDeadline: item.ApplicationDeadline ? item.ApplicationDeadline : "",
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ InternalVacanciesData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details in Internal Vacancies:", error);
    });
  }

}
