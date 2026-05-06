import * as React from 'react';
import styles from './HomePageAllVideoPage.module.scss';
import { IHomePageAllVideoPageProps } from './IHomePageAllVideoPageProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

export interface IHomePageAllVideoPageState {
  videos: any;
}

require('../assets/style.css');

export default class HomePageAllVideoPage extends React.Component<IHomePageAllVideoPageProps, IHomePageAllVideoPageState> {

  constructor(props: IHomePageAllVideoPageProps, state: IHomePageAllVideoPageState) {

    super(props);

    this.state = {
      videos: []
    };


  }



  public render(): React.ReactElement<IHomePageAllVideoPageProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePageAllVideoPage">

        <section id="homePageAllVideoPage">
          <div>
            <div className="news-header">
              <h2 className="section-title">Videos</h2>
            </div>
            <br />

            <div className="VideosHeader">

              {this.state.videos.map((item: any, index: number) => (

                <div className="video-container" key={index} style={{ marginBottom: "20px" }}>

                  {/* <h3>{item.Name}</h3> */}

                  <video className="video-bg" autoPlay width="400" muted controls>
                    <source src={item.ServerRelativeUrl} type="video/mp4" />
                  </video>

                </div>

              ))}

            </div>

          </div>
        </section>



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
