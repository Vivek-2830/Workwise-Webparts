import * as React from 'react';
import styles from './HomePagePoliciesAndDoc.module.scss';
import { IHomePagePoliciesAndDocProps } from './IHomePagePoliciesAndDocProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

export interface IHomePagePoliciesAndDocState {
  PolicesData : any;
}

require('../assets/style.css');

export default class HomePagePoliciesAndDoc extends React.Component<IHomePagePoliciesAndDocProps, IHomePagePoliciesAndDocState> {

  constructor(props: IHomePagePoliciesAndDocProps, state: IHomePagePoliciesAndDocState) {

    super(props);

    this.state = {
      PolicesData : ""
    };

  }


  public render(): React.ReactElement<IHomePagePoliciesAndDocProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePagePoliciesAndDoc">

        <div className="policy-panel">

          <div className="policy-header">
            <h2>
              Policies &amp; Documents
              <div className="policy-underline"></div>
            </h2>

            {/* <button className="view-all">View all</button> */}
          </div>

          <div className="policy-grid">

            {/* {
                this.state.PolicesData.length > 0 &&
                this.state.PolicesData.map((item) => {
                  // let imagePath = "";
                  // let ImageInfo = JSON.parse(item.PolicesImage);
                  // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                  //   imagePath = ImageInfo["serverRelativeUrl"];
                  // }
                  // else {
                  //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Policies Documents/Attachments/${item.ID}/${ImageInfo.fileName}`;
                  // }
                  return (
                    <div className="policy-card">
                      <a href={item.FileRef} target="_blank" data-interception="off" style={{ textDecoration: "none" }} >
                        <img src={item.PolicesImage} />
                        <span>{item.FileLeafRef}</span>
                      </a>
                    </div>
                  );
                })
              } */}

            {
              this.state.PolicesData.length > 0 &&
              this.state.PolicesData.map((item: any) => {

                let imagePath = "";

                if (item.PolicesImage) {
                  let imageInfo: any = item.PolicesImage;

                  // Convert string to JSON if needed
                  if (typeof item.PolicesImage === "string") {
                    try {
                      imageInfo = JSON.parse(item.PolicesImage);
                    } catch (error) {
                      console.log("Image JSON parse error", error);
                    }
                  }

                  if (imageInfo && imageInfo.serverRelativeUrl) {
                    imagePath = this.props.context.pageContext.web.absoluteUrl + imageInfo.serverRelativeUrl;
                  }
                }

                return (
                  <div className="policy-card" key={item.Id}>
                    <a href={item.FileRef} target="_blank" data-interception="off" style={{ textDecoration: "none", color: "black" }}>
                      {imagePath ? (
                        <img src={imagePath} />
                      ) : (
                        <img
                          src={require("../assets/Rectangle1.png")}
                          alt="default"
                        />
                      )}
                      <span>{item.FileLeafRef}</span>
                    </a>
                  </div>
                );
              })
            }

          </div>

        </div>


      </section>
    );
  }

  public async componentDidMount() {
    this.getPoliciesData();
  }

  public async getPoliciesData() {
    try {

      const libraryRoot = "/sites/GroupIntranet/Policies Documents"; // update site path

      const doc = await sp.web.lists
        .getByTitle("Policies Documents")
        .items
        .select(
          "Id",
          "FileLeafRef",
          "FileRef",
          "Modified",
          "Editor/Title",
          "File_x0020_Type",
          "PolicesImage",
          "FSObjType",
          "FileDirRef"
        )
        .expand("Editor")
        .filter(`FSObjType eq 1 and FileDirRef eq '${libraryRoot}'`)
        .get();

      this.setState({ PolicesData: doc });

    } catch (error) {
      console.log("Error fetching Company Documents data: ", error);
    }
  }


}
