import * as React from 'react';
import styles from './HomePageIntranetFaQs.module.scss';
import { IHomePageIntranetFaQsProps } from './IHomePageIntranetFaQsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';
import 'react-accessible-accordion/dist/fancy-example.css';
import { DefaultButton, Dialog, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IHomePageIntranetFaQsState {
  FaqsAnswersData: any;
  Questions: any;
  Answers: any;
  AddIntranetFaqDialog: boolean;
  AddIntranetFaqDataDialog: boolean;
  EditQuestions: any;
  EditAnswers: any;
  EditIntranetFaqDataDiaolg: boolean;
  CurrentIntranetFaqID: any;
  DeleteIntranetFaqID: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
}

require('../assets/style.css');

const AddIntranetFaqDetailsDialogContentProps = {
  title: "Add Intranet Details",
};

const AddAIntranetfaqDataDialogContentProps = {
  title: "Add Intranet"
};

const UpdateIntranetFaqDetailsDialogContentProps = {
  title: "Update Intranet Details"
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

export default class HomePageIntranetFaQs extends React.Component<IHomePageIntranetFaQsProps, IHomePageIntranetFaQsState> {

  constructor(props: IHomePageIntranetFaQsProps, state: IHomePageIntranetFaQsState) {

    super(props);

    this.state = {
      FaqsAnswersData: "",
      Questions: "",
      Answers: "",
      AddIntranetFaqDialog: true,
      AddIntranetFaqDataDialog: true,
      EditQuestions: "",
      EditAnswers: "",
      EditIntranetFaqDataDiaolg: true,
      CurrentIntranetFaqID: "",
      DeleteIntranetFaqID: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IHomePageIntranetFaQsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageIntranetFaQs">

        <div className="faq-panel">

          <h2>Intranet FAQs</h2>

          {
            this.state.IsAdmin ?
              <>
                <div className='AddAnnouncemt'>
                  <PrimaryButton text='Add FAQs' onClick={() => this.setState({ AddIntranetFaqDialog: false })} />
                </div>
              </>
              :
              <>
              </>
          }

          {
            this.state.FaqsAnswersData.length > 0 &&
            this.state.FaqsAnswersData.map((item) => {
              return (
                <div className="faq-item open">
                  <div className="faq-question">
                    <Accordion allowZeroExpanded>
                      <AccordionItem>
                        <AccordionItemHeading>
                          <AccordionItemButton>
                            {item.Questions}
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
          hidden={this.state.AddIntranetFaqDialog}
          onDismiss={() =>
            this.setState({
              AddIntranetFaqDialog: true,
            })
          }
          dialogContentProps={AddIntranetFaqDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddAnnouncmentData'>
            <PrimaryButton className='AddInfo' text='Add IntranetFaq Info' onClick={() => this.setState({ AddIntranetFaqDataDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>Questions</th>
                  <th style={{ width: '30%' }}>Answers</th>
                  <th style={{ width: '15%' }}>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.FaqsAnswersData.length > 0 &&
                  this.state.FaqsAnswersData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Questions}</td>
                        <td  dangerouslySetInnerHTML={{ __html: item.Answers }}></td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>

                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditIntranetFaqDataDiaolg: false, CurrentIntranetFaqID: item.ID }, () => this.EditIntranetFaqDetails(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteIntranetFaqItems(item.ID)}
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
          hidden={this.state.AddIntranetFaqDataDialog}
          onDismiss={() =>
            this.setState({
              AddIntranetFaqDataDialog: true,
              Questions: "",
              Answers: ""
            })
          }
          dialogContentProps={AddAIntranetfaqDataDialogContentProps}
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
                    this.setState({ Answers: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddIntranetFaqItems()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddIntranetFaqDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditIntranetFaqDataDiaolg}
          onDismiss={() =>
            this.setState({
              EditIntranetFaqDataDiaolg: true,
              EditQuestions: "",
              EditAnswers: ""
            })
          }
          dialogContentProps={UpdateIntranetFaqDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={1100}
        >
          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Announcement Title'
                  type='text'
                  value={this.state.EditQuestions}
                  onChange={(value) =>
                    this.setState({ EditQuestions: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdateAnnouncementDetails(this.state.CurrentIntranetFaqID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditIntranetFaqDataDiaolg: true })
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
    this.getFAQs();
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

  public async getFAQs() {
    const faqs = await sp.web.lists.getByTitle(this.props.ListName).items.select(
      "ID",
      "Questions",
      "Answers"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(faqs);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID,
            Questions: item.Questions,
            Answers: item.Answers
          });
        });
        this.setState({ FaqsAnswersData: AllData });
      }
    }).catch((error) => {
      console.log("Error fetching FAQs items: ", error);
    });
  }

  public async AddIntranetFaqItems() {
    if (this.state.Questions.length == 0) {
      alert("Please Enter Details");
    } else {
      const announcement = await sp.web.lists.getByTitle(this.props.ListName).items.add({
        Questions: this.state.Questions,
        Answers: this.state.Answers
      });

      this.setState({ AddIntranetFaqDataDialog: true });
      this.getFAQs();

    }
  }

  public async EditIntranetFaqDetails(ID) {
    let Editfaq = this.state.FaqsAnswersData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(Editfaq);
    this.setState({
      EditQuestions: Editfaq[0].Questions,
      EditAnswers: Editfaq[0].Answers
    });
  }

  public async UpdateAnnouncementDetails(CurrentIntranetFaqID) {
    try {
      const updatefaq: any = {
        Questions: this.state.EditQuestions,
        Answers: this.state.EditAnswers,
      };

      const updateItem = await sp.web.lists.getByTitle(this.props.ListName).items.getById(CurrentIntranetFaqID).update(updatefaq);

      this.setState({ EditIntranetFaqDataDiaolg: true });
      this.getFAQs();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteIntranetFaqItems(DeleteIntranetFaqID) {
    const deleteinfo = await sp.web.lists.getByTitle(this.props.ListName).items.getById(DeleteIntranetFaqID).delete();
    this.setState({ FaqsAnswersData: deleteinfo });
    this.getFAQs();
  }

}
