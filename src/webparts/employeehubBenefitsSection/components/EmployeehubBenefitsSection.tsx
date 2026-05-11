import * as React from 'react';
import styles from './EmployeehubBenefitsSection.module.scss';
import { IEmployeehubBenefitsSectionProps } from './IEmployeehubBenefitsSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { DefaultButton, Dialog, Icon, IconButton, PrimaryButton, TextField } from 'office-ui-fabric-react';


export interface IEmployeehubBenefitsSectionState {
  BenefitsData: any;
  BenefitsTitle: any;
  BenefitsDescription: any;
  BenefitsIcon: any;
  Link: any;
  UploadBenefitsIcon: any;
  AddBenefitsDialog: boolean;
  AddBenefitsDataDialog: boolean;
  EditBenefitsDataDialog: boolean;
  EditBenefitsTitle: any;
  EditBenefitsDescrition: any;
  EditBenefitsIcon: any;
  EditLink: any;
  EditUploadBenefitsIcon: any;
  CurrentBenefitsItemID: any;
  DeleteBenefitsItemID: any;
  previewImage: any;
}

require('../assets/style.css');

const AddBenefitsDetailsDialogContentProps = {
  title: "Add Benefits Details",
};

const AddBenefitsDataDialogContentProps = {
  title: "Add BenefitsItem"
}

const UpdateBenefitsDetailsDialogContentProps = {
  title: "Update Benefits Details"
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

export default class EmployeehubBenefitsSection extends React.Component<IEmployeehubBenefitsSectionProps, IEmployeehubBenefitsSectionState> {

  constructor(props: IEmployeehubBenefitsSectionProps, state: IEmployeehubBenefitsSectionState) {

    super(props);

    this.state = {
      BenefitsData: "",
      BenefitsTitle: "",
      BenefitsDescription: "",
      BenefitsIcon: [],
      Link: "",
      UploadBenefitsIcon: [],
      AddBenefitsDialog: true,
      AddBenefitsDataDialog: true,
      EditBenefitsDataDialog: true,
      EditBenefitsTitle: "",
      EditBenefitsDescrition: "",
      EditBenefitsIcon: [],
      EditLink: "",
      EditUploadBenefitsIcon: [],
      CurrentBenefitsItemID: "",
      DeleteBenefitsItemID: "",
      previewImage: "",
    };

  }

  public render(): React.ReactElement<IEmployeehubBenefitsSectionProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="employeehubBenefitsSection">

        <div className="benefits-wrapper">

          <div className="benefits-title">
            <h3>Benefits Section</h3>
            <span className="underline"></span>

            <div className='AddAnnouncemt'>
              <PrimaryButton text='Add Benefits Item' onClick={() => this.setState({ AddBenefitsDialog: false })} />
            </div>

          </div>

          <div className="benefits-grid">

            {
              this.state.BenefitsData.length > 0 &&
              this.state.BenefitsData.map((item, index) => {
                // let imagePath = "";
                // let ImageInfo = JSON.parse(item.BenefitsIcon);
                // if (ImageInfo && ImageInfo["serverRelativeUrl"]) {
                //   imagePath = ImageInfo["serverRelativeUrl"];
                // }
                // else {
                //   imagePath = `${this.props.context.pageContext.site.absoluteUrl}/Lists/Benefits Section/Attachments/${item.ID}/${ImageInfo.fileName}`;
                // }
                return (
                  <a href={item.Link.Url} style={{ textDecoration: "none", color: "black" }}>
                    <div className="benefit-card" key={index}>
                      <div className="icons"><img src={item.BenefitsIcon} /></div>
                      <h4>{item.BenefitsTitle}</h4>
                      <p>{item.BenefitsDescription}</p>
                    </div>
                  </a>
                );
              })
            }

          </div>
        </div>

        <Dialog
          hidden={this.state.AddBenefitsDialog}
          onDismiss={() =>
            this.setState({
              AddBenefitsDialog: true,
            })
          }
          dialogContentProps={AddBenefitsDetailsDialogContentProps}
          modalProps={addmodelProps}
          minWidth={1500}
        >

          <div className='AddAnnouncmentData'>
            <PrimaryButton className='AddBenefits' text='Add Data' onClick={() => this.setState({ AddBenefitsDataDialog: false })} />
          </div>

          <div className="news-container">
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }} className="news-table">
              <thead>
                <tr>
                  <th style={{ width: '20%' }}>BenefitsTitle</th>
                  <th style={{ width: '30%' }}>BenefitsDescription</th>
                  <th style={{ width: '30%' }}>BenefitsIcon</th>
                  <th style={{ width: '15%' }}>Link</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>

                {
                  this.state.BenefitsData.length > 0 &&
                  this.state.BenefitsData.map((item) => {
                    return (
                      <tr key={item.ID}>
                        <td className="title">{item.BenefitsTitle}</td>
                        <td>{item.BenefitsDescription}</td>
                        <td>
                          {
                            item.BenefitsIcon ? (
                              <img src={item.BenefitsIcon} alt="announcement" style={{ width: "80px", height: "80px", objectFit: "cover" }} />
                            ) : (
                              "No Icon"
                            )
                          }
                        </td>
                        <td>
                          <a href={item.Link.Url} target="_blank" rel="noopener noreferrer">{item.Link.Description}</a>
                        </td>

                       
                        <td>
                          <div style={{ display: "flex", gap: "8px" }}>
                            <IconButton
                              iconProps={{ iconName: "Edit" }}
                              title="Edit"
                              ariaLabel="Edit"
                              onClick={() => this.setState({ EditBenefitsDataDialog: false, CurrentBenefitsItemID: item.ID }, () => this.EditBenefits(item.ID))}
                            />

                            <IconButton
                              iconProps={{ iconName: "Delete" }}
                              title="Delete"
                              ariaLabel="Delete"
                              onClick={() => this.DeleteBenefitsInfo(item.ID)}
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
          hidden={this.state.AddBenefitsDataDialog}
          onDismiss={() =>
            this.setState({
              AddBenefitsDataDialog: true,
              BenefitsTitle: "",
              BenefitsDescription: "",
              BenefitsIcon: [],
              Link: "",
              UploadBenefitsIcon: []
            })
          }
          dialogContentProps={AddBenefitsDataDialogContentProps}
          modalProps={addmodelProps2}
          minWidth={1100}
        >
          <div className="ms-Grid-row">

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Benefits Title'
                  type='text'
                  onChange={(value) =>
                    this.setState({ BenefitsTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Benefits Description'
                  type='text'
                  multiline rows={3}
                  onChange={(value) =>
                    this.setState({ BenefitsDescription: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <label><b>Upload BenefitsIcon</b></label><br />

                <input
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleImageChange(e)}
                />

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Link'
                  type='text'
                  onChange={(value) =>
                    this.setState({ Link: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Submit'
                    onClick={() => this.AddBenefitsItem()}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ AddBenefitsDataDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </div>
        </Dialog>

        <Dialog
          hidden={this.state.EditBenefitsDataDialog}
          onDismiss={() =>
            this.setState({
              EditBenefitsDataDialog: true,
              EditBenefitsTitle: "",
              EditBenefitsDescrition: "",
              EditBenefitsIcon: [],
              EditLink: "",
              EditUploadBenefitsIcon: []
            })
          }
          dialogContentProps={UpdateBenefitsDetailsDialogContentProps}
          modalProps={updatemodelProps}
          minWidth={1100}
        >
          <div className='ms-Grid-row'>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Benefits Title'
                  type='text'
                  value={this.state.EditBenefitsTitle}
                  onChange={(value) =>
                    this.setState({ EditBenefitsTitle: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <TextField
                  label='Benefits Description'
                  type='text'
                  multiline rows={3}
                  value={this.state.EditBenefitsDescrition}
                  onChange={(value) =>
                    this.setState({ EditBenefitsDescrition: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div className='Add-Form'>
                <label><b>Upload BenefitsIcon</b></label><br />

                <input
                  type="file"
                  accept="image/*"
                  onChange={(e: any) => this.handleUpdateImageChange(e)}
                />

                {
                  this.state.EditUploadBenefitsIcon && (
                    <div className="Attached-img">

                      {/* ✅ Handle BOTH string + file */}
                      <p>
                        {
                          typeof this.state.EditUploadBenefitsIcon === "string"
                            ? this.state.EditUploadBenefitsIcon.split('/').pop()
                            : this.state.EditUploadBenefitsIcon[0]?.name
                        }
                      </p>

                      <Icon
                        iconName="Cancel"
                        onClick={() => this.setState({ EditUploadBenefitsIcon: "" })}
                      />
                    </div>
                  )
                }

              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
              <div>
                <TextField
                  label='Link'
                  type='text'
                  value={this.state.EditLink}
                  onChange={(value) =>
                    this.setState({ EditLink: value.target["value"] })
                  }
                />
              </div>
            </div>

            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Announcement-Submit'>
                <div className='Submit-Button'>
                  <PrimaryButton
                    text='Update'
                    onClick={() => this.UpdatBenefitsItemDetails(this.state.CurrentBenefitsItemID)}
                  />
                </div>

                <div className='Cancel-Button'>
                  <DefaultButton
                    text='Cancel'
                    onClick={() =>
                      this.setState({ EditBenefitsDataDialog: true })
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
    this.getBenefitsItems();
  }

  public async getBenefitsItems() {
    const benefit = await sp.web.lists.getByTitle("Benefits Section").items.select(
      "ID",
      "BenefitsIcon",
      "BenefitsTitle",
      "BenefitsDescription",
      "Link"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(benefit);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            BenefitsIcon: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.BenefitsIcon ? JSON.parse(item.BenefitsIcon).serverRelativeUrl : require(`../assets/fi3198344.png`),
            BenefitsTitle: item.BenefitsTitle ? item.BenefitsTitle : "",
            BenefitsDescription: item.BenefitsDescription ? item.BenefitsDescription : "",
            Link: item.Link ? item.Link : ""
          });
        });
        this.setState({ BenefitsData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching details: ", error);
    });
  }

  public async AddBenefitsItem() {
    if (this.state.BenefitsTitle.length == 0) {
      alert("Please Enter Details");
    } else {
      const empannouncement = await sp.web.lists.getByTitle("Benefits Section").items.add({
        BenefitsTitle: this.state.BenefitsTitle,
        BenefitsDescription: this.state.BenefitsDescription,
        Link: this.state.Link
          ? {
            Url: this.state.Link,
            Description: this.state.Link
          }
          : null
      });

      if (this.state.UploadBenefitsIcon && this.state.UploadBenefitsIcon.length > 0) {

        const file = this.state.UploadBenefitsIcon[0];

        await sp.web.lists
          .getByTitle("Benefits Section")
          .items.getById(empannouncement.data.Id)
          .attachmentFiles.add(file.name, file);
      }

      this.setState({ AddBenefitsDataDialog: true });
      this.getBenefitsItems();

    }
  }

  handleImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadBenefitsIcon: [file],
        previewImage: URL.createObjectURL(file)
      });
    }
  };

  handleUpdateImageChange = (e: any) => {
    const file = e.target.files[0];

    if (file) {
      this.setState({
        UploadBenefitsIcon: [file],
        // previewImage: URL.createObjectURL(file)
      });
    }

  }

  public async EditBenefits(ID) {
    let EditbenefitsItem = this.state.BenefitsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditbenefitsItem);
    this.setState({
      EditBenefitsTitle: EditbenefitsItem[0].BenefitsTitle,
      EditBenefitsDescrition: EditbenefitsItem[0].BenefitsDescription,
      EditLink: EditbenefitsItem[0].Link.Url,
      EditUploadBenefitsIcon: EditbenefitsItem[0].BenefitsIcon,
    });
  }

  public async UpdatBenefitsItemDetails(CurrentBenefitsItemID) {
    try {
      const updatebenefitsItems: any = {
        BenefitsTitle: this.state.EditBenefitsTitle,
        BenefitsDescription: this.state.EditBenefitsDescrition,
        Link: this.state.EditLink ? {
          Url: this.state.EditLink,
          Description: this.state.EditLink
        } : null
      };

      const updateItem = await sp.web.lists.getByTitle("Benefits Section").items.getById(CurrentBenefitsItemID).update(updatebenefitsItems);

      // if (this.state.EditUploadImages && this.state.EditUploadImages.length > 0) {
      //   const file = this.state.EditUploadImages[0];

      //   const itemRef = sp.web.lists
      //     .getByTitle("Announcements")
      //     .items.getById(CurrentEmpAnnouncementDataID);

      //   const attachments = await itemRef.attachmentFiles();

      //   for (let att of attachments) {
      //     await itemRef.attachmentFiles.getByName(att.FileName).delete();
      //   }

      //   await itemRef.attachmentFiles.add(file.name, file);
      // }

      if (Array.isArray(this.state.EditUploadBenefitsIcon) && this.state.EditUploadBenefitsIcon.length > 0) {

        const file = this.state.EditUploadBenefitsIcon[0];

        const itemRef = sp.web.lists
          .getByTitle("Benefits Section")
          .items.getById(CurrentBenefitsItemID);

        // delete old attachments
        const attachments = await itemRef.attachmentFiles();

        for (let att of attachments) {
          await itemRef.attachmentFiles.getByName(att.FileName).delete();
        }

        // add new file
        await itemRef.attachmentFiles.add(file.name, file);
      }


      this.setState({ EditBenefitsDataDialog: true });
      this.getBenefitsItems();

    } catch (error) {
      console.log("Error Updating details :", error);
    }
  }

  public async DeleteBenefitsInfo(DeleteBenefitsItemID) {
    const deleteinfo = await sp.web.lists.getByTitle("Benefits Section").items.getById(DeleteBenefitsItemID).delete();
    this.setState({ BenefitsData: deleteinfo });
    this.getBenefitsItems();
  }

}
