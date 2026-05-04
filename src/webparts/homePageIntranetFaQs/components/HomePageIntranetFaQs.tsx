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
import { Dialog, IconButton, PrimaryButton } from 'office-ui-fabric-react';

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
}

require('../assets/style.css');

const AddIntranetFaqDetailsDialogContentProps = {
  title: "Add Intranet Details",
};

const AddAIntranetfaqDataDialogContentProps = {
  title: "Add Intranet"
}

const UpdateAnnouncementDetailsDialogContentProps = {
  title: "Update Intranet Details"
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
          maxWidth={1500}
        >

          <div className='AddAnnouncmentData'>
            <PrimaryButton className='AddInfo' text='Add IntranetFaq Info' onClick={() => this.setState({ AddIntranetFaqDataDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>Title</th>
                  <th style={{ width: '30%' }}>Description</th>
                  <th style={{ width: '30%' }}>Source</th>
                  <th style={{ width: '15%' }}>Images</th>
                  <th style={{ width: '15%' }}>Link</th>
                  <th style={{ width: '15%' }}>Videos</th>
                  <th style={{ width: '15%' }}>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.FaqsAnswersData.length > 0 &&
                  this.state.FaqsAnswersData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.Title}</td>
                        <td>{item.Description}</td>
                        <td>{item.Source}</td>
                        <td>
                          {
                            item.Images ? (
                              <img src={item.Images} alt="announcement" style={{ width: "120px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Image"
                            )
                          }
                        </td>
                        <td>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>
                        <td>
                          {
                            item.Videos ? (
                              <a
                                href={item.Videos.Url || item.Videos}
                                target="_blank"
                                rel="noopener noreferrer"
                              >
                                Watch Video
                              </a>
                            ) : (
                              "No Video"
                            )
                          }
                        </td>

                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>

                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              // onClick={() => this.setState({ EditIntranetFaqDataDiaolg: false, CurrentIntranetFaqID: item.ID }, () => this.EditAnnouncementInfo(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              // onClick={() => this.DeleteAnnouncementInfo(item.ID)}
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

      </section>
    );
  }

  public async componentDidMount() {
    this.getFAQs();
  }

  public async getFAQs() {
    const faqs = await sp.web.lists.getByTitle("Intranet FAQ").items.select(
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

}
