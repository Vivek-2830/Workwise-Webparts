import * as React from 'react';
import styles from './EmployeehubHrTeamContact.module.scss';
import { IEmployeehubHrTeamContactProps } from './IEmployeehubHrTeamContactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

export interface IEmployeehubHrTeamContactState {
  HRDetailsData: any;
  Name: any;
  JobTitle: any;
  Phone: any;
  Email: any;
  Photo: any;
  UploadPhoto: any;
  AddHRTemaDialog: boolean;
  AddHRTeamDataDialog: boolean;
  EditName: any;
  EditJobTitle: any;
  EditPhone: any;
  EditEmail: any;
  EditPhoto: any;
  EditUploadPhoto: any;
  EditHRTeamDataDialog: boolean;
  CurrentHrTeamDetailsID: any;
  DeleteHrTeamDetailsID: any;
  previewImage: any;
}

require('../assets/style.css');

export default class EmployeehubHrTeamContact extends React.Component<IEmployeehubHrTeamContactProps, IEmployeehubHrTeamContactState> {

  constructor(props: IEmployeehubHrTeamContactProps, state:IEmployeehubHrTeamContactState) {
    
      super(props);

      this.state = {
        HRDetailsData: "",
        Name: "",
        JobTitle: "",
        Phone: "",
        Email: "",
        Photo: [],
        UploadPhoto: [],
        AddHRTemaDialog: true,
        AddHRTeamDataDialog: true,
        EditName: "",
        EditJobTitle: "",
        EditPhone: "",
        EditEmail: "",
        EditPhoto: [],
        EditUploadPhoto: [],
        EditHRTeamDataDialog: true,
        CurrentHrTeamDetailsID: "",
        DeleteHrTeamDetailsID: "",
        previewImage: "",
      };

  }


  public render(): React.ReactElement<IEmployeehubHrTeamContactProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubHrTeamContact">

        <div className="hr-wrapper">

          <div className="hr-title">
            <h3>HR Team Contact Details</h3>
            <span className="underline"></span>
          </div>

          <div className="hr-grid">

            {
              this.state.HRDetailsData.length > 0 &&
              this.state.HRDetailsData.map((item, index) => {

                return (
                  <div className="hr-card">

                    <div className='main-card'>
                      <img src={item.Photo} className="avatar" />
                      <div className="hr-info">
                        <h4>{item.Name}</h4>
                        <span className="role">{item.JobTitle}</span>
                      </div>
                    </div>

                    {
                      !!item.Phone && (
                        <div className="contact">
                          <img src={require('../assets/phone.png')} alt="phone" />
                          <p>{item.Phone}</p>
                        </div>
                      )
                    }

                    <div className='contact-email'>
                      <img src={require('../assets/mail01.png')} /> <p>{item.Email}</p>
                    </div>

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
    this.getHRTeamDetails();
  }

  public async getHRTeamDetails() {
    const details = await sp.web.lists.getByTitle("HR Team Contact Details").items.select(
      "ID",
      "Name",
      "JobTitle",
      "Phone",
      "Email",
      "Photo"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(details);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : [],
            Name: item.Name ? item.Name : "",
            JobTitle: item.JobTitle ? item.JobTitle : "",
            Phone: item.Phone ? item.Phone : "",
            Email: item.Email ? item.Email : "",
            Photo: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Photo ? JSON.parse(item.Photo).serverRelativeUrl : require(`../assets/avatar3.png`)
          });
        });
        this.setState({ HRDetailsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching Detail in HR Team Contact Details:", error);
    });
  }

}
