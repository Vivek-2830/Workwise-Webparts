import * as React from 'react';
import styles from './EmployeeHubAxisClcGallery.module.scss';
import { IEmployeeHubAxisClcGalleryProps } from './IEmployeeHubAxisClcGalleryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

export interface IEmployeeHubAxisClcGalleryState {
  AxisGalleryData: any;
  IsAdmin: boolean;
  CurrentUserEmail: any;
  
}

require('../assets/style.css');

export default class EmployeeHubAxisClcGallery extends React.Component<IEmployeeHubAxisClcGalleryProps, IEmployeeHubAxisClcGalleryState> {

  constructor(props: IEmployeeHubAxisClcGalleryProps, state: IEmployeeHubAxisClcGalleryState) {

    super(props);

    this.state = {
      AxisGalleryData: "",
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IEmployeeHubAxisClcGalleryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeHubAxisClcGallery">

        <div className="gallery-section">


          <div className="gallery-header">
            Life at Axis CLC (Photo Gallery)
          </div>

          <div className="gallery-grid">
            {
              this.state.AxisGalleryData.length > 0 &&
                this.state.AxisGalleryData.map((item) => {
                  return (
                    <div className="gallery-item" key={item.Id}>
                      <img src={item.EncodedAbsUrl} alt={item.FileLeafRef} />
                    </div>
                  );     
                })

            }
          </div>

          {
            this.state.IsAdmin ?
            <>
              <a href='https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Axis%20CLC%20Gallery/Forms/AllItems.aspx' style={{ textDecoration : 'none', color: 'inherit' }}>
                <button className="submit-btn">
                  Submit an Image
                  <svg viewBox="0 0 24 24">
                    <path d="M5 20h14v-2H5v2zm7-18l-5.5 5.5 1.42 1.42L11 6.84V16h2V6.84l3.08 3.08 1.42-1.42L12 2z" />
                  </svg>
                </button>
              </a>
            </>
            :
            <>
            </>
          }

        </div>

      </section>
    );
  }

  public async componentDidMount() {
    this.getGalleryImages();
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
      console.error("Error checking details:", error);
    }
  }

  public async getGalleryImages() {
    try {
      const imageExtensions = [
        "png",
        "jpg",
        "jpeg",
        "gif",
        "webp",
        "svg",
        "bmp"
      ];
  
      const items = await sp.web.lists
        .getByTitle("Axis CLC Gallery")
        .items
        .select(
          "Id",
          "Title",
          "FileLeafRef",
          "FileRef",
          "EncodedAbsUrl",
          "File_x0020_Type",
          "UniqueId",
          "Modified"
        )
        .filter("FSObjType eq 0") // Only files, no folders
        .orderBy("Modified", false)
        .getAll();
  
      // Filter only image files
      const galleryImages = items.filter((item: any) => {
        const extension = item.File_x0020_Type
          ? item.File_x0020_Type.toLowerCase()
          : "";
  
        return imageExtensions.indexOf(extension) > -1;
      });
  
      // Format image URLs
      const formattedImages = galleryImages.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        FileLeafRef: item.FileLeafRef,
        FileRef: item.FileRef,
        EncodedAbsUrl:
          item.EncodedAbsUrl ||
          `${window.location.origin}${item.FileRef}`,
        File_x0020_Type: item.File_x0020_Type,
        UniqueId: item.UniqueId
      }));
  
      console.log("Gallery Images:", formattedImages);
  
      this.setState({
        AxisGalleryData: formattedImages
      });
    } catch (error) {
      console.error("Error fetching details:", error);
    }
  }


}
