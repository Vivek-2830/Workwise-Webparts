import * as React from 'react';
import styles from './EmployeehubHrfaQs.module.scss';
import { IEmployeehubHrfaQsProps } from './IEmployeehubHrfaQsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DatePicker, DefaultButton, Dialog, Dropdown, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';
import 'react-accessible-accordion/dist/fancy-example.css';

export interface IEmployeehubHrfaQsState {
  EmployeeFaQData: any;
  Question: any;
  Answers: any;
  EmployeeHRFAQsDialog: boolean;
  AddEmployeeHRFAQsDialog: boolean;
  EditQuestion: any;
  EditAnswers: any;
  EditEmployeeHRFAQsDialog: boolean;
  CurrentEmployeeFAQsItemID: any;
  DeleteEmployeeFAQsItemID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const EmployeeHubFAQsDetailsDialogContentProps = {
  title: "Add EmployeeHubFAQs Details",
};

const AddEmployeeHubFAQsDataDialogContentProps = {
  title: "Add EmployeeHubFAQs"
};

const UpdateEmployeeHubFAQsDetailsDialogContentProps = {
  title: "Update EmployeeHubFAQs Details"
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

export default class EmployeehubHrfaQs extends React.Component<IEmployeehubHrfaQsProps, IEmployeehubHrfaQsState> {

  constructor(props: IEmployeehubHrfaQsProps, state: IEmployeehubHrfaQsState) {

    super(props);

    this.state = {
      EmployeeFaQData: "",
      Question: "",
      Answers: "",
      EmployeeHRFAQsDialog: true,
      AddEmployeeHRFAQsDialog: true,
      EditQuestion: "",
      EditAnswers: "",
      EditEmployeeHRFAQsDialog: true,
      CurrentEmployeeFAQsItemID: "",
      DeleteEmployeeFAQsItemID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IEmployeehubHrfaQsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubHrfaQs">

        <div className="faq-wrapper">

          <h2 className="faq-title">
            HR FAQs
            <span className="underline"></span>
          </h2>

          {
            this.state.IsAdmin ?
            <>
              <div className='AddHRInfo'>
                <PrimaryButton text='Add HRFAQs' onClick={() => this.setState({ EmployeeHRFAQsDialog: false })} />
              </div>
            </>
            :
            <>
            </>
          }

          {
            this.state.EmployeeFaQData.length > 0 &&
            this.state.EmployeeFaQData.map((item) => {
              return (
                <div className="faq-card">
                  <div className="faq-item">
                    <Accordion allowZeroExpanded>
                      <AccordionItem>
                        <AccordionItemHeading>
                          <AccordionItemButton>
                            {item.Question}
                          </AccordionItemButton>
                        </AccordionItemHeading>
                        <AccordionItemPanel>
                          <p className="faq-answer" dangerouslySetInnerHTML={{ __html: item.Answers }}>
                          </p>
                        </AccordionItemPanel>
                      </AccordionItem>
                    </Accordion>
                  </div>
                </div>
              );
            })
          }

        </div>

        <Dialog
          hidden={this.state.EmployeeHRFAQsDialog}
          onDismiss={() =>
            this.setState({
              EmployeeHRFAQsDialog: true,
            })
          }
          dialogContentProps={EmployeeHubFAQsDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddfaqsData'>
            <PrimaryButton className='Addhrfaqs' text='Add Data' onClick={() => this.setState({ AddEmployeeHRFAQsDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>Question</th>
                  <th style={{ width: '30%' }}>Answers</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.EmployeeFaQData.length > 0 &&
                  this.state.EmployeeFaQData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Question}</td>
                        <td dangerouslySetInnerHTML={{ __html: item.Answers }}></td>
                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditEmployeeHRFAQsDialog: false, CurrentEmployeeFAQsItemID: item.ID }, () => this.EditEmpFAQsInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteEmpHRFaqsInfo(item.ID)}
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
          hidden={this.state.AddEmployeeHRFAQsDialog}
          onDismiss={() =>
            this.setState({
              AddEmployeeHRFAQsDialog: true,
              Question: "",
              Answers: ""
            })
          }
          dialogContentProps={AddEmployeeHubFAQsDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={1100}
        >
          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Question'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Question: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Answers'
                  type='text'
                  multiline rows={3}
                  onChange={(value) =>
                    this.setState({ Answers: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddEmpFAQsInfo()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddEmployeeHRFAQsDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditEmployeeHRFAQsDialog}
          onDismiss={() =>
            this.setState({
              EditEmployeeHRFAQsDialog: true,
              EditQuestion: "",
              EditAnswers: ""
            })
          }
          dialogContentProps={UpdateEmployeeHubFAQsDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={1100}
        >
          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Question'
                  type='text'
                  value={this.state.EditQuestion}
                  onChange={(value) =>
                    this.setState({ EditQuestion: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Answers'
                  type='text'
                  multiline rows={3}
                  value={this.state.EditAnswers}
                  onChange={(value) =>
                    this.setState({ EditAnswers: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateEmpFAQsDetails(this.state.CurrentEmployeeFAQsItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditEmployeeHRFAQsDialog: true })
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
    this.getEmployeeFaQs();
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

  public async getEmployeeFaQs() {
    const faq = await sp.web.lists.getByTitle(this.props.ListName).items.select(
      "ID",
      "Question",
      "Answers"
    ).get().then((data) => {
      let AllData = [];
      console.log(faq);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Question: item.Question ? item.Question : "",
            Answers: item.Answers ? item.Answers : ""
          });
        });
        this.setState({ EmployeeFaQData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details in Employee FaQs:", error);
    });
  }

  public async AddEmpFAQsInfo() {
    if (this.state.Question.length == 0) {
      alert("Please Enter Details");
    } else {
      const empannouncement = await sp.web.lists.getByTitle(this.props.ListName).items.add({
        Question: this.state.Question,
        Answers: this.state.Answers
      });

      this.setState({ AddEmployeeHRFAQsDialog: true });
      this.getEmployeeFaQs();

    }
  }

  public async EditEmpFAQsInfo(ID) {
    let EditFaqs = this.state.EmployeeFaQData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditFaqs);
    this.setState({
      EditQuestion: EditFaqs[0].Question,
      EditAnswers: EditFaqs[0].Answers
    });
  }

  public async UpdateEmpFAQsDetails(CurrentEmployeeFAQsItemID) {
    try {
      const updateHRFaqs: any = {
        Question: this.state.EditQuestion,
        Answers: this.state.EditAnswers
      };

      const updateItem = await sp.web.lists.getByTitle(this.props.ListName).items.getById(CurrentEmployeeFAQsItemID).update(updateHRFaqs);

      this.setState({ EditEmployeeHRFAQsDialog: true });
      this.getEmployeeFaQs();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteEmpHRFaqsInfo(DeleteEmployeeFAQsItemID) {
    const deleteinfo = await sp.web.lists.getByTitle(this.props.ListName).items.getById(DeleteEmployeeFAQsItemID).delete();
    this.setState({ EmployeeFaQData: deleteinfo });
    this.getEmployeeFaQs();
  }

}
