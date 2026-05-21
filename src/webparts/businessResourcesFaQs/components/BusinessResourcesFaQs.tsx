import * as React from 'react';
import styles from './BusinessResourcesFaQs.module.scss';
import { IBusinessResourcesFaQsProps } from './IBusinessResourcesFaQsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import 'react-accessible-accordion/dist/fancy-example.css';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';
import { DefaultButton, Dialog, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IBusinessResourcesFaQsState {
  BusinessFaQsData: any;
  Question: any;
  Answer: any;
  BusinessResourceFaQDialog: boolean;
  AddBusinessResourceDataFaqDialog: boolean;
  EditQuestion: any;
  EditAnswer: any;
  EditBusinessResourceFaQDialog: boolean;
  CurrenFaQItemID: any;
  DeleteFaQItemID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const BusinessResourceFAQsDetailsDialogContentProps = {
  title: "Add BusinessResourceFAQs Details",
};

const AddBusinessResourceFAQsDataDialogContentProps = {
  title: "Add BusinessResourceFAQs"
};

const UpdateBusinessResourceFAQsDetailsDialogContentProps = {
  title: "Update BusinessResource Details"
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

export default class BusinessResourcesFaQs extends React.Component<IBusinessResourcesFaQsProps, IBusinessResourcesFaQsState> {

  constructor(props: IBusinessResourcesFaQsProps, state: IBusinessResourcesFaQsState) {

    super(props);

    this.state = {
      BusinessFaQsData: "",
      Question: "",
      Answer: "",
      BusinessResourceFaQDialog: true,
      AddBusinessResourceDataFaqDialog: true,
      EditQuestion: "",
      EditAnswer: "",
      EditBusinessResourceFaQDialog: true,
      CurrenFaQItemID: "",
      DeleteFaQItemID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IBusinessResourcesFaQsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="businessResourcesFaQs">

        <div className="Business-contents">
          <h2>Frequently Asked Questions</h2>
          <div className='Faq-underline'></div>

          {
            this.state.IsAdmin ?
              <>
                <div className='AddHRInfo'>
                  <PrimaryButton text='Add ResourceFAQs' onClick={() => this.setState({ BusinessResourceFaQDialog: false })} />
                </div>
              </>
              :
              <>
              </>
          }

          {
            this.state.BusinessFaQsData.length > 0 &&
            this.state.BusinessFaQsData.map((item) => {
              return (
                <div className="Business-list">
                  <div className="Business-row">
                    <Accordion allowZeroExpanded>
                      <AccordionItem>
                        <AccordionItemHeading>
                          <AccordionItemButton>
                            {item.Question}
                          </AccordionItemButton>
                        </AccordionItemHeading>
                        <AccordionItemPanel>
                          <p className="Business-answer" dangerouslySetInnerHTML={{ __html: item.Answer }}>
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
          hidden={this.state.BusinessResourceFaQDialog}
          onDismiss={() =>
            this.setState({
              BusinessResourceFaQDialog: true,
            })
          }
          dialogContentProps={BusinessResourceFAQsDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddResourceData'>
            <PrimaryButton className='AddResourceFAQ' text='Add Data' onClick={() => this.setState({ AddBusinessResourceDataFaqDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>Question</th>
                  <th style={{ width: '30%' }}>Answer</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.BusinessFaQsData.length > 0 &&
                  this.state.BusinessFaQsData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Question}</td>
                        <td dangerouslySetInnerHTML={{ __html: item.Answer }}></td>
                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditBusinessResourceFaQDialog: false, CurrenFaQItemID: item.ID }, () => this.EditResourceFAQsItem(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteResourceFAQsItemInfo(item.ID)}
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
          hidden={this.state.AddBusinessResourceDataFaqDialog}
          onDismiss={() =>
            this.setState({
              AddBusinessResourceDataFaqDialog: true,
              Question: "",
              Answer: ""
            })
          }
          dialogContentProps={AddBusinessResourceFAQsDataDialogContentProps}
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
                    this.setState({ Answer: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddResourceFAQsItem()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddBusinessResourceDataFaqDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditBusinessResourceFaQDialog}
          onDismiss={() =>
            this.setState({
              EditBusinessResourceFaQDialog: true,
              EditQuestion: "",
              EditAnswer: ""
            })
          }
          dialogContentProps={UpdateBusinessResourceFAQsDetailsDialogContentProps}
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
                  value={this.state.EditAnswer}
                  onChange={(value) =>
                    this.setState({ EditAnswer: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateEmpFAQsDetails(this.state.CurrenFaQItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditBusinessResourceFaQDialog: true })
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
    this.getBusinessFaQsData();
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

  public async getBusinessFaQsData() {
    const faqs = await sp.web.lists.getByTitle(this.props.ListName).items.select(
      "ID",
      "Question",
      "Answer"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID,
            Question: item.Question,
            Answer: item.Answer
          });
        });
        this.setState({ BusinessFaQsData: AllData });
      }
    }).catch((error) => {
      console.log("Error fetching Business FaQs data: ", error);
    });
  }


  public async AddResourceFAQsItem() {
    if (this.state.Question.length == 0) {
      alert("Please Enter Details");
    } else {
      const empannouncement = await sp.web.lists.getByTitle(this.props.ListName).items.add({
        Question: this.state.Question,
        Answer: this.state.Answer
      });

      this.setState({ BusinessFaQsData: true });
      this.getBusinessFaQsData();

    }
  }

  public async EditResourceFAQsItem(ID) {
    let EditFaqs = this.state.BusinessFaQsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditFaqs);
    this.setState({
      EditQuestion: EditFaqs[0].Question,
      EditAnswer: EditFaqs[0].Answer
    });
  }

  public async UpdateEmpFAQsDetails(CurrenFaQItemID) {
    try {
      const updateFaqs: any = {
        Question: this.state.EditQuestion,
        Answers: this.state.EditAnswer
      };

      const updateItem = await sp.web.lists.getByTitle(this.props.ListName).items.getById(CurrenFaQItemID).update(updateFaqs);

      this.setState({ EditBusinessResourceFaQDialog: true });
      this.getBusinessFaQsData();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteResourceFAQsItemInfo(DeleteFaQItemID) {
    const deleteinfo = await sp.web.lists.getByTitle(this.props.ListName).items.getById(DeleteFaQItemID).delete();
    this.setState({ BusinessFaQsData: deleteinfo });
    this.getBusinessFaQsData();
  }

}
