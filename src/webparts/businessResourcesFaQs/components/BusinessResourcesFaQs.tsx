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

export interface IBusinessResourcesFaQsState {
  BusinessFaQsData: any;
}

require('../assets/style.css');

export default class BusinessResourcesFaQs extends React.Component<IBusinessResourcesFaQsProps, IBusinessResourcesFaQsState> {

  constructor(props:IBusinessResourcesFaQsProps, state:IBusinessResourcesFaQsState) {

    super(props);

    this.state = {
      BusinessFaQsData : ""
    }

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
        
      </section >
    );
  }

  public async componentDidMount() {
    this.getBusinessFaQsData();
  }

  public async getBusinessFaQsData() {
    const faqs = await sp.web.lists.getByTitle("Business FaQ").items.select(
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

}
