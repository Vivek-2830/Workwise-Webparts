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

export interface IHomePageIntranetFaQsState {
  FaqsAnswersData: any;
}

require('../assets/style.css');

export default class HomePageIntranetFaQs extends React.Component<IHomePageIntranetFaQsProps, IHomePageIntranetFaQsState> {

  constructor(props: IHomePageIntranetFaQsProps, state: IHomePageIntranetFaQsState) {
    
    super(props);

    this.state = {
      FaqsAnswersData : ""
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
