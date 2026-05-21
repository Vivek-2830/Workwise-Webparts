import * as React from 'react';
import styles from './EmployeeHubFeedBackSection.module.scss';
import { IEmployeeHubFeedBackSectionProps } from './IEmployeeHubFeedBackSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';


export interface IEmployeeHubFeedBackSectionState {
  FeedBackResponseData: any;
  YourFeedback: any;
  OurResponse: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
  FeedbackResponseDialog: boolean;
  AddFeedbackResponseDialog: boolean;
  EditYourFeedback: any;
  EditOurResponse: any;
  EditFeedbackResponseDialog: boolean;
  CurrentFeedbackItemID: any;
  DeleteFeedbackItemID: any;
}

require('../assets/style.css');

const FeedbackDetailsDialogContentProps = {
  title: "Add Feedback Details",
};

const AddAFeedbackDataDialogContentProps = {
  title: "Add Feedback"
};

const UpdateFeedbackDataDialogContentProps = {
  title: "Update Feedback Details"
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


export default class EmployeeHubFeedBackSection extends React.Component<IEmployeeHubFeedBackSectionProps, IEmployeeHubFeedBackSectionState> {

  constructor(props: IEmployeeHubFeedBackSectionProps, state: IEmployeeHubFeedBackSectionState) {

    super(props);

    this.state = {
      FeedBackResponseData: "",
      IsAdmin: false,
      CurrentUserEmail: "",
      YourFeedback: "",
      OurResponse: "",
      FeedbackResponseDialog: true,
      AddFeedbackResponseDialog: true,
      EditYourFeedback: "",
      EditOurResponse: "",
      EditFeedbackResponseDialog: true,
      CurrentFeedbackItemID: "",
      DeleteFeedbackItemID: "",
    };

  }


  public render(): React.ReactElement<IEmployeeHubFeedBackSectionProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubFeedBackSection">

        <h1 className="section-feed">Hearts & Minds - Our Culture, Our People</h1>
        <p className="section-subtitle">
          Join us in shaping a more engaged, connected, and vibrant workplace ready for our Investors in People survey!
        </p>

        {
          this.state.IsAdmin ?
            <>
              <div className='Addfeed'>
                <PrimaryButton text='Add Feedback' onClick={() => this.setState({ FeedbackResponseDialog: false })} />
              </div>
            </>
            :
            <>
            </>
        }

        <div className="feedback-panel">
          <div className="feedback-grid">

            <div>
              <h3 className="column-title">Your Feedback (You Said)</h3>
              <div className="column-items">
                {
                  this.state.FeedBackResponseData.length > 0 &&
                  this.state.FeedBackResponseData.map((item) => {
                    return (
                      <div className="feedback-item">{item.YourFeedback}</div>
                    );
                  })
                }
              </div>
            </div>

            <div>
              <h3 className="column-title">Our Response (We Did)</h3>
              <div className="column-items">
                {
                  this.state.FeedBackResponseData.length > 0 &&
                  this.state.FeedBackResponseData.map((item) => {
                    return (
                      <div className="feedback-item">{item.OurResponse}</div>
                    );
                  })
                }
              </div>
            </div>

          </div>
        </div>


        <Dialog
          hidden={this.state.FeedbackResponseDialog}
          onDismiss={() =>
            this.setState({
              FeedbackResponseDialog: true,
            })
          }
          dialogContentProps={FeedbackDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddUserData'>
            <PrimaryButton className='Add Userguide' text='Add Feedbacks' onClick={() => this.setState({ AddFeedbackResponseDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>YourFeedback</th>
                  <th style={{ width: '30%' }}>OurResponse</th>
                  <th style={{ width: '15%' }}>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.FeedBackResponseData.length > 0 &&
                  this.state.FeedBackResponseData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.YourFeedback}</td>
                        <td>{item.OurResponse}</td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>

                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditFeedbackResponseDialog: false, CurrentFeedbackItemID: item.ID }, () => this.EditFeedbackInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteFeedbackInfo(item.ID)}
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
          hidden={this.state.AddFeedbackResponseDialog}
          onDismiss={() =>
            this.setState({
              AddFeedbackResponseDialog: true,
              YourFeedback: "",
              OurResponse: ""
            })
          }
          dialogContentProps={AddAFeedbackDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={1100}
        >
          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Title'
                  type='text'
                  onChange={(value) =>
                    this.setState({ YourFeedback: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Essential Description'
                  type='text'
                  multiline rows={3}
                  onChange={(value) =>
                    this.setState({ OurResponse: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddFeedbackInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddFeedbackResponseDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditFeedbackResponseDialog}
          onDismiss={() =>
            this.setState({
              EditFeedbackResponseDialog: true,
              EditYourFeedback: "",
              EditOurResponse: "",
            })
          }
          dialogContentProps={UpdateFeedbackDataDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={1100}
        >
          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Title'
                  type='text'
                  value={this.state.EditYourFeedback}
                  onChange={(value) =>
                    this.setState({ EditYourFeedback: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Essential Description'
                  type='text'
                  multiline rows={3}
                  value={this.state.EditOurResponse}
                  onChange={(value) =>
                    this.setState({ EditOurResponse: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdatefeedbackDetails(this.state.CurrentFeedbackItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditFeedbackResponseDialog: true })
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
    this.getFeedbackData();
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

  public async getFeedbackData() {
    const response = await sp.web.lists.getByTitle("Feedback and Response").items.select(
      "ID",
      "YourFeedback",
      "OurResponse"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(response);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            YourFeedback: item.YourFeedback ? item.YourFeedback : "",
            OurResponse: item.OurResponse ? item.OurResponse : ""
          });
        });
        this.setState({ FeedBackResponseData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details", error);
    });
  }

  public async AddFeedbackInfo() {
    if (this.state.YourFeedback.length == 0) {
      alert("Please Enter Details");
    } else {
      const announcement = await sp.web.lists.getByTitle("Feedback and Response").items.add({
        YourFeedback: this.state.YourFeedback,
        OurResponse: this.state.OurResponse,
      });

      this.setState({ AddFeedbackResponseDialog: true });
      this.getFeedbackData();

    }
  }

  public async EditFeedbackInfo(ID) {
    let Edituserguide = this.state.FeedBackResponseData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(Edituserguide);
    this.setState({
      EditYourFeedback: Edituserguide[0].YourFeedback,
      EditOurResponse: Edituserguide[0].OurResponse
    });
  }

  public async UpdatefeedbackDetails(CurrentFeedbackItemID) {
    try {
      const updatfeedback: any = {
        EditYourFeedback: this.state.YourFeedback,
        EditOurResponse: this.state.OurResponse
      };

      const updateItem = await sp.web.lists.getByTitle("Feedback and Response").items.getById(CurrentFeedbackItemID).update(updatfeedback);

      this.setState({ EditFeedbackResponseDialog: true });
      this.getFeedbackData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteFeedbackInfo(DeleteFeedbackItemID) {
    const deleteinfo = await sp.web.lists.getByTitle("Feedback and Response").items.getById(DeleteFeedbackItemID).delete();
    this.setState({ FeedBackResponseData: deleteinfo });
    this.getFeedbackData();
  }

}
