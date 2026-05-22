import * as React from 'react';
import styles from './HomePageVideoSection.module.scss';
import { IHomePageVideoSectionProps } from './IHomePageVideoSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { sp } from '@pnp/sp/presets/all';
import { PrimaryButton } from 'office-ui-fabric-react';

export interface IHomePageVideoSectionState {
  videos: any;
}

require('../assets/style.css');

export default class HomePageVideoSection extends React.Component<IHomePageVideoSectionProps, IHomePageVideoSectionState> {

  constructor(props: IHomePageVideoSectionProps, state: IHomePageVideoSectionState) {
    super(props);

    this.state = {
      videos: []
    };
  }


  public render(): React.ReactElement<IHomePageVideoSectionProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const videos = {
      dots: true,
      infinite: true,
      speed: 1000,
      slidesToShow: 3,
      slidesToScroll: 1,
      arrows: true,
      autoplay: true,
      cssEase: "linear"
    };

    return (
      <section className="homePageVideoSection">

        <div>
          <div className="news-header">
            <h2 className="section-video">Case Study Videos</h2>

            <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Videos/Forms/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
              <PrimaryButton className='Adddoc' text="Add Video" />
            </a>

            <a href='https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/SitePages/Company%20videos.aspx' style={{ textDecoration: "none", color: "black" }} target="_blank" rel="noopener noreferrer">
              <button className="view-news">View all</button>
            </a>
          </div>
          <br />
          <Slider {...videos}>

            {this.state.videos.map((item: any, index: number) => (
              <div key={index} style={{ marginBottom: "20px" }}>

                {/* <h3>{item.Name}</h3> */}

                <video className='GalleryVideos' autoPlay width="400" muted controls>
                  <source src={item.ServerRelativeUrl} type="video/mp4" />
                </video>

              </div>
            ))}

          </Slider>
        </div>

      </section>
    );
  }

  public async componentDidMount() {
    this.getVideos();
  }

  public getVideos = async () => {
    try {

      const files = await sp.web.lists
        .getByTitle("Videos")
        .rootFolder
        .files
        .select("Name", "ServerRelativeUrl")
        .get();

      const videoFiles = files.filter((item: any) =>
        item.Name.endsWith(".mp4") ||
        item.Name.endsWith(".mov") ||
        item.Name.endsWith(".webm") ||
        item.Name.endsWith(".avi")
      );

      this.setState({
        videos: videoFiles
      });

    } catch (error) {
      console.log(error);
    }
  }


}
