import * as React from 'react';
import styles from './EmployeehubInternalVacancies.module.scss';
import { IEmployeehubInternalVacanciesProps } from './IEmployeehubInternalVacanciesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import { sp } from '@pnp/sp/presets/all';
import { DatePicker, DefaultButton, Dialog, Dropdown, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";

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
  LocationTypelist: any;
  EmploymentTypelist: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddInternalVacanciesDetailsDialogContentProps = {
  title: "Add InternalVacancies Details",
};

const AddEInternalVacanciesDataDialogContentProps = {
  title: "Add InternalVacancies"
};

const UpdateInternalVacanciesDetailsDialogContentProps = {
  title: "Update InternalVacancies Details"
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
      LocationTypelist: [],
      EmploymentTypelist: [],
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }


  public render(): React.ReactElement<IEmployeehubInternalVacanciesProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      InternalVacanciesLink
    } = this.props;

    return (
      <section className="employeehubInternalVacancies">

        <div className="jobs-wrapper">

          <div className="jobs-header">
            <div>
              <h3>Internal Vacancies</h3>
              <span className="underline"></span>
            </div>

            {
              this.state.IsAdmin ?
              <>
                  <div className='Addvacancieinfo'>
                    <PrimaryButton text='Add Vacancies' onClick={() => this.setState({ AddInternalVacancieDialog: false })} />
                  </div>
              </>
              :
              <></>
            }

            <a href={InternalVacanciesLink} style={{ textDecoration: 'none', color: 'black' }}><button className="view-all">View all</button></a>
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
          minWidth={1200}
        >

          <div className='Internalbox'>
            <div>
              <h2>InternalVacancies Details</h2>
            </div>
            <div className='AddvacancieData'>
              <PrimaryButton className='AddInternalInfo' text='Add Data' onClick={() => this.setState({ AddInternalvacancieDataDialog: false })} />
            </div>
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th>JobTitle</th>
                  <th>Department</th>
                  <th>LocationType</th>
                  <th>EmploymentType</th>
                  <th>Salary</th>
                  <th>KeyRequirements</th>
                  <th>ApplicationDeadline</th>
                  <th>Link</th>
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
                        <td  dangerouslySetInnerHTML={{ __html: item.KeyRequirements }}></td>
                        <td>{moment(item.ApplicationDeadline).format("MMM DD, YYYY")}</td>
                        <td style={{ wordBreak: "break-all" }}>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>
                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditInternalVacanciesDialog: false, CurrentInternalVacanciesItemID: item.ID }, () => this.EditVacancies(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteInternalvacanciesInfo(item.ID)}
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
          hidden={this.state.AddInternalvacancieDataDialog}
          onDismiss={() =>
            this.setState({
              AddInternalvacancieDataDialog: true,
              JobTitle: "",
              Department: "",
              LocationType: "",
              EmploymentType: "",
              Salary: "",
              KeyRequirements: "",
              ApplicationDeadline: "",
              Link: ""
            })
          }
          dialogContentProps={AddEInternalVacanciesDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={900}
        >
          <div>
            <h2>Add InternalVacancies Details</h2>
          </div>

          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <TextField
                  label='JobTitle'
                  type='text'
                  onChange={(value) =>
                    this.setState({ JobTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <Dropdown
                  options={this.state.LocationTypelist}
                  label='Location Type'
                  required
                  onChange={(e, option, text) =>
                    this.setState({ LocationType: option.text })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <Dropdown
                  options={this.state.EmploymentTypelist}
                  label='Employment Type'
                  required
                  onChange={(e, option, text) =>
                    this.setState({ EmploymentType: option.text })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <TextField
                  label='Salary'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Salary: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <DatePicker
                  label='ApplicationDeadline'
                  allowTextInput={false}
                  value={this.state.ApplicationDeadline ? this.state.ApplicationDeadline : null}
                  onSelectDate={(date: any) => this.setState({ ApplicationDeadline: date })}
                  aria-label="Select Date" placeholder='Select Date' isRequired
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Department'
                  type='text'
                  multiline rows={3}
                  onChange={(value) =>
                    this.setState({ Department: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg12 KeySection'>
              <div className='Add-Form'>
                <label><b style={{ fontWeight: '600' }}>KeyRequirements</b></label>
                <RichText
                  value={this.state.KeyRequirements}
                  onChange={(text: string) => {
                    this.setState({ KeyRequirements: text });
                    return text;
                  }}
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddInternalVacanciesInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddInternalvacancieDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditInternalVacanciesDialog}
          onDismiss={() =>
            this.setState({
              EditInternalVacanciesDialog: true,
              EditJobTitle: "",
              EditDepartment: "",
              EditLocationType: "",
              EditEmploymentType: "",
              EditSalary: "",
              EditKeyRequirements: "",
              EditApplicationDeadline: "",
              EditLink: ""
            })
          }
          dialogContentProps={UpdateInternalVacanciesDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={900}
        >
          <div>
            <h2>Update InternalVacancies Details</h2>
          </div>

          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <TextField
                  label='JobTitle'
                  type='text'
                  value={this.state.EditJobTitle}
                  onChange={(value) =>
                    this.setState({ EditJobTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <Dropdown
                  options={this.state.LocationTypelist}
                  label="Location Type"
                  required
                  defaultSelectedKey={this.state.EditLocationType}
                  onChange={(e, option, text) =>
                    this.setState({ EditLocationType: option.text })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <Dropdown
                  options={this.state.EmploymentTypelist}
                  label="Employment Type"
                  required
                  defaultSelectedKey={this.state.EditEmploymentType}
                  onChange={(e, option, text) =>
                    this.setState({ EditEmploymentType: option.text })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <TextField
                  label='Salary'
                  type='text'
                  value={this.state.EditSalary}
                  onChange={(value) =>
                    this.setState({ EditSalary: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
              <div className='Add-Form'>
                <DatePicker
                  label='ApplicationDeadline'
                  allowTextInput={false}
                  value={this.state.EditApplicationDeadline ? this.state.EditApplicationDeadline : null}
                  onSelectDate={(date: any) => this.setState({ EditApplicationDeadline: date })}
                  aria-label="Select a Date" isRequired
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 OurInternalVacancies'>
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

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Department'
                  type='text'
                  value={this.state.EditDepartment}
                  multiline rows={3}
                  onChange={(value) =>
                    this.setState({ EditDepartment: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg12 KeySection'>
              <div className='Add-Form'>
                <label><b style={{ fontWeight: '600' }}>KeyRequirements</b></label>
                <RichText
                  value={this.state.EditKeyRequirements}
                  onChange={(text: string) => {
                    this.setState({ EditKeyRequirements: text });
                    return text;
                  }}
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateInternalvacanciesInfo(this.state.CurrentInternalVacanciesItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditInternalVacanciesDialog: true })
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
    this.getInternalVacancies();
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

  public async AddInternalVacanciesInfo() {
    if (this.state.JobTitle.length == 0) {
      alert("Please Enter Details");
    } else {
      const vacancies = await sp.web.lists.getByTitle("Internal Vacancies").items.add({
        JobTitle: this.state.JobTitle,
        Department: this.state.Department,
        LocationType: this.state.LocationType,
        Salary: this.state.Salary,
        EmploymentType: this.state.EmploymentType,
        KeyRequirements: this.state.KeyRequirements,
        ApplicationDeadline: this.state.ApplicationDeadline,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null
      });

      this.setState({ AddInternalvacancieDataDialog: true });
      this.getInternalVacancies();

    }
  }

  public async EditVacancies(ID) {
    let EditInternalvacancies = this.state.InternalVacanciesData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditInternalvacancies);
    this.setState({
      EditJobTitle: EditInternalvacancies[0].JobTitle,
      EditDepartment: EditInternalvacancies[0].Department,
      EditLocationType: EditInternalvacancies[0].LocationType,
      EditEmploymentType: EditInternalvacancies[0].EmploymentType,
      EditSalary: EditInternalvacancies[0].Salary,
      EditKeyRequirements: EditInternalvacancies[0].KeyRequirements,
      EditApplicationDeadline: new Date(EditInternalvacancies[0].ApplicationDeadline),
      EditLink: EditInternalvacancies[0].Link.Url
    });
  }

  public async UpdateInternalvacanciesInfo(CurrentInternalVacanciesItemID) {
    try {
      const updateempannouncement: any = {
        JobTitle: this.state.EditJobTitle,
        Department: this.state.EditDepartment,
        LocationType: this.state.EditLocationType,
        EmploymentType: this.state.EditEmploymentType,
        Salary: this.state.EditSalary,
        KeyRequirements: this.state.EditKeyRequirements,
        ApplicationDeadline: this.state.EditApplicationDeadline,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Internal Vacancies").items.getById(CurrentInternalVacanciesItemID).update(updateempannouncement);


      this.setState({ EditInternalVacanciesDialog: true });
      this.getInternalVacancies();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteInternalvacanciesInfo(DeleteVacanciesItemID) {
    const deleteinfo = await sp.web.lists.getByTitle("Internal Vacancies").items.getById(DeleteVacanciesItemID).delete();
    this.setState({ InternalVacanciesData: deleteinfo });
    this.getInternalVacancies();
  }

  public async GetTicketsChoicesItems() {
    const choiceFieldName1 = "LocationType";
    const field1 = await sp.web.lists.getByTitle("Internal Vacancies").fields.getByInternalNameOrTitle(choiceFieldName1)();
    let locationtypelist = [];
    field1["Choices"].forEach(function (dname, i) {
      locationtypelist.push({ key: dname, text: dname });
    });
    this.setState({ LocationTypelist: locationtypelist });

    const choiceFieldName2 = "Employment Type";
    const field2 = await sp.web.lists.getByTitle("Internal Vacancies").fields.getByInternalNameOrTitle(choiceFieldName2)();
    let employmenttypelist = [];
    field2["Choices"].forEach(function (dname, i) {
      employmenttypelist.push({ key: dname, text: dname });
    });
    this.setState({ EmploymentTypelist: employmenttypelist });

  }


}
