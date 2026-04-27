import * as React from 'react';
import styles from './HomePageUserguide.module.scss';
import { IHomePageUserguideProps } from './IHomePageUserguideProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";

export interface IHomePageUserguideState {
  EssentialLearningsData: any;
}

require('../assets/style.css');

export default class HomePageUserguide extends React.Component<IHomePageUserguideProps, IHomePageUserguideState> {

  constructor(props: IHomePageUserguideProps, state: IHomePageUserguideState) {
  
    super(props);

    this.state = {
      EssentialLearningsData : ""
    };
  
  }


  public render(): React.ReactElement<IHomePageUserguideProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

     const userguide = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 4,
      slidesToScroll: 4,
      arrows: true,
      autoplay: true,
      cssEase: "linear"
    };

    return (
      <section className="homePageUserguide">
        
          <div className="essential-section"> 
          <h2 className="essential-title">User guides</h2>

            <Slider {...userguide}>
          
            

              {
                this.state.EssentialLearningsData.length > 0 &&
                this.state.EssentialLearningsData.map((item) => {
                  return (
                    <div className="learning-card">
                      <img src={item.Images} alt="Training" className="learning-image" />
                      <div className="learning-content">
                        <h3>{item.Title}</h3>
                        <p>
                          {item.EssentialDescription}
                        </p>
                        <a href={item.link.Url} style={{ cursor: "pointer" }} className="read-more">
                          Read more →
                        </a>
                      </div>
                    </div>
                  );
                })
              }

              
            
          </Slider>
         
        </div>


      </section>
    );
  }

  public async componentDidMount() {
    this.getEssentiallearnings();
  }

  public async getEssentiallearnings() {
    const roadmap = await sp.web.lists.getByTitle("User guides").items.select(
      "ID",
      "Title",
      "EssentialDescription",
      "Images",
      "link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(roadmap);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            EssentialDescription: item.EssentialDescription ? item.EssentialDescription : "",
            Images: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Image ? JSON.parse(item.Image).serverRelativeUrl : require(`../assets/Rectangle1.png`),
            link: item.link ? item.link : ""
          });
        });
        this.setState({ EssentialLearningsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Details: ", error);
    });
  }

}
