import * as React from 'react';
import styles from './BusinessResourcesAllDocument.module.scss';
import { IBusinessResourcesAllDocumentProps } from './IBusinessResourcesAllDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { Pivot, PivotItem, Icon, IIconProps, SearchBox, PrimaryButton } from 'office-ui-fabric-react';

export interface IBusinessResourcesAllDocumentState {
  BusinessAnnouncementsData: any;
  CompanyDocumentsData: any;
  DocumentsAndPoliciesData: any;
  DocumentsAndPoliciesFilterdData: any;
  TemplatesandBrandingData: any;
  TemplatesandBrandingFilterdData: any;
  InsuranceandCertificateData: any;
  InsuranceandCertificateFilterdData: any;
  HowtoGuidesData: any;
  HowtoGuidesFilterdData: any;
  BusinessApplicationData: any;
  BusinessApplicationFilterdData: any;
  BrochuresFilterdData: any;
  BrochuresData: any;
  BusinessFaQsData: any;
  SelectedDocsPivot: string;
  SelectedTemplatesPivot: string;
  SelectInsuranceCertificatePivot: string;
  SelectHowtoGuidePivot: string;
  SelectBusinessApplicationPivot: string;
  SelectBrochurePivot: string;
  results: any;
  searchText: any;
  AddDocumentsandPoliciesDialog : boolean;
  AddTemplatesandBrandingDialog : boolean;
  AddInsuranceandCertificatesDialog: boolean;
  AddHSDocumentsDialog: boolean;
  AddBrochuresDialog: boolean;
  IsAdmin: boolean;
  CurrentUserEmail: any;
} 

const SearchIconProps: IIconProps = { iconName: 'Search' };

require('../assets/style.css');

export default class BusinessResourcesAllDocument extends React.Component<IBusinessResourcesAllDocumentProps, IBusinessResourcesAllDocumentState> {

  constructor(props: IBusinessResourcesAllDocumentProps, state:IBusinessResourcesAllDocumentState) {

    super(props);

    this.state = {
      BusinessAnnouncementsData: "",
      CompanyDocumentsData: "",
      DocumentsAndPoliciesData: [],
      DocumentsAndPoliciesFilterdData: "",
      TemplatesandBrandingData: [],
      TemplatesandBrandingFilterdData: [],
      InsuranceandCertificateData: [],
      InsuranceandCertificateFilterdData: "",
      HowtoGuidesData: [],
      HowtoGuidesFilterdData: "",
      BusinessApplicationData: [],
      BusinessApplicationFilterdData: "",
      BrochuresData: [],
      BrochuresFilterdData: "",
      BusinessFaQsData: "",
      SelectedDocsPivot: "all",
      SelectedTemplatesPivot: "all",
      SelectInsuranceCertificatePivot: "all",
      SelectHowtoGuidePivot: "all",
      SelectBusinessApplicationPivot: "all",
      SelectBrochurePivot: "all",
      searchText: [],
      results: [],
      AddDocumentsandPoliciesDialog : true,
      AddTemplatesandBrandingDialog : true,
      AddInsuranceandCertificatesDialog: true,
      AddHSDocumentsDialog: true,
      AddBrochuresDialog: true,
      IsAdmin: false,
      CurrentUserEmail: ""
    };

  }

  public render(): React.ReactElement<IBusinessResourcesAllDocumentProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="businessResourcesAllDocument">

        <div className="quick-accesswrapper">

          <SearchBox iconProps={SearchIconProps} placeholder="What are you looking for?" className='new-search' value={this.state.searchText}
            onChange={(e) => {
              const value = e.target.value;

              this.setState({ searchText: value }, () => {
                this.searchDocuments(); 
              });
            }}
          />
          
          <ul className="doclist searchitems">
            {this.state.results.map((item, i) => (

              <li key={item.Id} className="docs-items">
                <div className="docfile-info">
                  <span className="doc-icon">
                    <img src={require('../assets/Vector.png')} />
                  </span>
                  <div className="doc-name">{item.FileLeafRef}</div>
                </div>
                <a href={item.FileRef} target="_blank" data-interception="off" className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
              </li>
            ))}
          </ul>

          <div className="quick-links">

            <a href='#BusinessCatalog' style={{ textDecoration: 'none', color: 'inherit' }}>
              <div className="quick-item">
                <div className="icon-circle">
                  <img src={require('../assets/Document.png')} />
                </div>
                <p>Documents<br />and Policies</p>
              </div>
            </a>

            <a href='#TempCatalog' style={{ textDecoration: 'none', color: 'inherit' }}>
              <div className="quick-item">
                <div className="icon-circle">
                  <img src={require('../assets/TemplateBranding.png')} />
                </div>
                <p>Templates<br />and Branding</p>
              </div>
            </a>

            <a href='#InsuranceTab' style={{ textDecoration: 'none', color: 'inherit' }}>
              <div className="quick-item">
                <div className="icon-circle">
                  <img src={require('../assets/Insurance.png')} />
                </div>
                <p>Insurance and<br />Certificates</p>
              </div>
            </a>

            <a href='#HSTab' style={{ textDecoration: 'none', color: 'inherit' }}>
              <div className="quick-item">
                <div className="icon-circle">
                  <img src={require('../assets/BusinessApp.png')} />
                </div>
                <p>H&S Documents</p>
              </div>
            </a>

          </div>

        </div>

        {/* -------------------------------Documents and Policies------------------------------------- */}
        <div id="BusinessCatalog" className="docs-wrapper">

          <h2 className="page-title">Documents and Policies</h2>

          {
            this.state.IsAdmin ?
            <>
              <div className='DocPolicieButton'>
                <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Documents%20and%20Policies/Forms/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
                  <PrimaryButton className='AddDocPolicie' text="Add Documents and Policies" />
                </a>
              </div>
            </>
            :
            <>
            </>
          }

          <div className="category-tabs">
           
            <Pivot
              selectedKey={this.state.SelectedDocsPivot}
              onLinkClick={(item) => this._onPivotChange("docs", item)}
            >
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="HR" itemKey="hr" />
              <PivotItem headerText="Finance" itemKey="finance" />
              <PivotItem headerText="Safety" itemKey="safety" />
              <PivotItem headerText="IT" itemKey="it" />
              <PivotItem headerText="Fleet" itemKey="fleet" />
              <PivotItem headerText="Procurement" itemKey="procurement" />
              <PivotItem headerText="Axis" itemKey="axis" />
              <PivotItem headerText="CLC" itemKey="clc" />
              <PivotItem headerText="Concept" itemKey="concept" />
            </Pivot>

          </div>

          <div className="content-grid">

            <div className="doc-card">
              <div className="doc-header">Latest</div>

              <ul className="doclist">
                {
                  Array.isArray(this.state.DocumentsAndPoliciesData)
                    ? [...this.state.DocumentsAndPoliciesData]   // clone array
                      .sort((a, b) => {
                        const dateA = a.Modified ? new Date(a.Modified).getTime() : 0;
                        const dateB = b.Modified ? new Date(b.Modified).getTime() : 0;
                        return dateB - dateA;
                      })
                      .slice(0, 5)
                      .map((doc) => (
                        <li key={doc.Id} className="docs-items">
                          <div className="docfile-info">
                            <span className="doc-icon">
                              <img src={require('../assets/Vector.png')} />
                            </span>
                            <div className="doc-name">{doc.FileLeafRef}</div>
                          </div>
                          <a href={doc.FileRef} target="_blank" data-interception="off" download className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
                        </li>
                      ))
                    : null
                }
              </ul>
            </div>

            <div className="documents-card">
              <ul className="doc-list scroll">
                {
                  this.state.DocumentsAndPoliciesFilterdData.length > 0 &&
                  this.state.DocumentsAndPoliciesFilterdData.map((doc) => {
                    return (
                      <li key={doc.Id} className="docsPolice-items">
                        <div className="docsPolice-info">
                          <span className="docsPolice-icon">
                            <img src={require('../assets/Vector.png')} />
                          </span>
                          <div className="docsPolice-name">{doc.FileLeafRef}</div>
                        </div>
                        <a href={doc.FileRef} target="_blank" data-interception="off" download className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
                      </li>
                    );
                  })
                }

              </ul>
            </div>

          </div>
        </div>

        {/* -------------------------------Templates and Branding------------------------------------- */}
        <div id="TempCatalog" className="temps-wrapper">

          <h2 className="page-title">Templates and Branding</h2>
          
          {
            this.state.IsAdmin ?
            <>
              <div className='TemplateButton'>
                <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Templates%20and%20Branding/Forms/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
                  <PrimaryButton className='AddTemplates' text="Add Templates and Branding" />
                </a>
              </div>
            </>
            :
            <>
            </>
          }

          <div className="category-tabs">
            
            <Pivot selectedKey={this.state.SelectedTemplatesPivot} onLinkClick={(item) => this._onPivotChange("templates", item)}>
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="Logos" itemKey="logos" />
              <PivotItem headerText="Logo Lockups" itemKey="logo lockups" />
              <PivotItem headerText="Branding Guidelines" itemKey="Branding Guidelines" />
              <PivotItem headerText="Presentations" itemKey="Presentations" />
              <PivotItem headerText="Teams Backgrounds" itemKey="Teams Backgrounds" />
              <PivotItem headerText="LinkedIn Banners" itemKey="LinkedIn Banners" />
            </Pivot>
          </div>

          <div className="content-grid">

            {/* ================= LEFT: LATEST ================= */}
            <div className="category-card">
              <div className="doc-header">Latest</div>

              <ul className="doclist">
                {
                  Array.isArray(this.state.TemplatesandBrandingData)
                    ? [...this.state.TemplatesandBrandingData]
                      .sort((a, b) => {
                        var dateA = a.Modified ? new Date(a.Modified).getTime() : 0;
                        var dateB = b.Modified ? new Date(b.Modified).getTime() : 0;
                        return dateB - dateA;
                      })
                      .slice(0, 5)
                      .map((doc) => (
                        <li key={doc.Id} className="docs-items">

                          <div className="docfile-info">
                            <span className="doc-icon">
                              <img src={require('../assets/Vector.png')} />
                            </span>
                            <div className="doc-name">{doc.FileLeafRef}</div>
                          </div>

                          <a
                            href={doc.FileRef}
                            target="_blank"
                            rel="noopener noreferrer"
                            data-interception="off"
                            download
                            className="download"
                          >
                            <img src={require('../assets/DownloadIcon.png')} />
                          </a>

                        </li>
                      ))
                    : null
                }
              </ul>
            </div>


            {/* ================= RIGHT: PREVIEW ================= */}
            <div className="categorys-card">
              <ul className="doc-list scroll TemplatesandBranding">

                {
                  this.state.TemplatesandBrandingFilterdData.length > 0 &&
                  this.state.TemplatesandBrandingFilterdData.map((doc) => {

                    var type = doc.File_x0020_Type
                      ? doc.File_x0020_Type.toLowerCase()
                      : "";

                    var videoTypes = ["mp4", "mov", "mkv", "wmv", "flv"];
                    var imageTypes = ["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg"];
                    var officeTypes = ["ppt", "pptx", "doc", "docx", "xls", "xlsx"];
                    var templateTypes = ["pot", "potx"];

                    var fileUrl = doc.EncodedAbsUrl;
                    var previewUrl = fileUrl;

                    // ✅ OFFICE FILES
                    if (officeTypes.indexOf(type) > -1 && doc.UniqueId) {
                      previewUrl =
                        this.props.context.pageContext.web.absoluteUrl +
                        "/_layouts/15/Doc.aspx?sourcedoc=" + doc.UniqueId +
                        "&file=" + encodeURIComponent(doc.FileLeafRef) +
                        "&action=embedview&wdStartOn=1";
                    }

                    // ✅ TEMPLATE FILES (POTX FIX)
                    else if (templateTypes.indexOf(type) > -1 && doc.UniqueId) {
                      previewUrl =
                        this.props.context.pageContext.web.absoluteUrl +
                        "/_layouts/15/WopiFrame.aspx?sourcedoc=" + doc.UniqueId +
                        "&action=interactivepreview";
                    }

                    // ✅ PDF
                    else if (type === "pdf") {
                      previewUrl = fileUrl;
                    }

                    // fallback
                    else {
                      previewUrl = fileUrl.replace("interactivepreview", "embedview");
                    }

                    return (
                      <li key={doc.Id} className="docsPolice-items">

                        <div className="docsPolice-info">

                          <span className="docsPolice-icon">

                            {/* VIDEO */}
                            {
                              videoTypes.indexOf(type) > -1 ? (
                                <video width="100%" height="580" controls>
                                  <source src={fileUrl} />
                                </video>
                              )

                                /* IMAGE */
                                : imageTypes.indexOf(type) > -1 ? (
                                  <img
                                    className="fileimage"
                                    src={fileUrl}
                                    style={{
                                      width: "100%",
                                      maxHeight: "580px",
                                      objectFit: "contain"
                                    }}
                                    alt={doc.FileLeafRef}
                                  />
                                )

                                  /* OFFICE + TEMPLATE PREVIEW */
                                  : (
                                    <iframe
                                      src={previewUrl}
                                      style={{ width: "100%", height: "580px" }}
                                      frameBorder="0"
                                      allowFullScreen
                                      title={doc.FileLeafRef}
                                    ></iframe>
                                  )
                            }

                          </span>

                          <div className="docsPolice-name">
                            {doc.FileLeafRef}
                          </div>

                        </div>

                        {/* DOWNLOAD */}
                        <a
                          href={doc.FileRef}
                          target="_blank"
                          rel="noopener noreferrer"
                          data-interception="off"
                          download
                          className="download"
                        >
                          <img src={require('../assets/DownloadIcon.png')} />
                        </a>

                      </li>
                    );
                  })
                }

              </ul>
            </div>

          </div>


        </div>

        {/* -------------------------------Insurance and Certificates------------------------------------- */}
        <div id='InsuranceTab' className="Insurance-wrapper">

          <h2 className="page-title">Insurance and Certificates</h2>

          {
            this.state.IsAdmin ?
            <>
              <div className='InsuranceButton'>
                <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Insurance%20and%20Certificates/Forms/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
                  <PrimaryButton className='AddInsurance' text="Add Insurance and Certificates" />
                </a>
              </div>
            </>
            :
            <>
            </>
          }

          <div className="category-tabs">
            <Pivot selectedKey={this.state.SelectInsuranceCertificatePivot} onLinkClick={(item) => this._onPivotChange("insurance", item)}>
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="Axis Certificates" itemKey="Axis Certificates" />
              <PivotItem headerText="CLC Certificates" itemKey="CLC Certificates" />
              <PivotItem headerText="Concept Certificates" itemKey="Concept Certificates" />
              <PivotItem headerText="Insurances" itemKey="Insurances" />
            </Pivot>
          </div>

          <div className="content-grid">

            <div className="doc-card">
              <div className="doc-header">Latest</div>

              <ul className="doclist">
                {
                  Array.isArray(this.state.InsuranceandCertificateData)
                    ? [...this.state.InsuranceandCertificateData]
                      .sort((a, b) => {
                        const dateA = a.Modified ? new Date(a.Modified).getTime() : 0;
                        const dateB = b.Modified ? new Date(b.Modified).getTime() : 0;
                        return dateB - dateA;
                      })
                      .slice(0, 5)
                      .map((doc) => (
                        <li key={doc.Id} className="docs-items">
                          <div className="docfile-info">
                            <span className="doc-icon">
                              <img src={require('../assets/Vector.png')} />
                            </span>
                            <div className="doc-name">{doc.FileLeafRef}</div>
                          </div>
                          <a href={doc.FileRef} target="_blank" data-interception="off" download className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
                        </li>
                      ))
                    : null
                }
              </ul>
            </div>

            <div className="documents-card">
              <ul className="doc-list scroll">
                {
                  this.state.InsuranceandCertificateFilterdData.length > 0 &&
                  this.state.InsuranceandCertificateFilterdData.map((doc) => {
                    return (
                      <li key={doc.Id} className="docsPolice-items">
                        <div className="docsPolice-info">
                          <span className="docsPolice-icon">
                            <img src={require('../assets/Vector.png')} />
                          </span>
                          <div className="docsPolice-name">{doc.FileLeafRef}</div>
                        </div>
                        <a href={doc.FileRef} target="_blank" data-interception="off" download className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
                      </li>
                    );
                  })
                }

              </ul>
            </div>

          </div>
        </div>

        {/* -------------------------------Business Application Catalogue (H&S) ------------------------------------- */}
        <div id='HSTab' className="Business-wrapper">

          <h2 className="page-title">H&S Documents</h2>

          {
            this.state.IsAdmin ?
            <>
              <div className='HsButton'>
                <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Business%20Application%20Catalogue/Forms/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
                  <PrimaryButton className='AddBusinessDoc' text="Add H&S Documents" />
                </a>
              </div>
            </>
            :
            <>
            </>
          }

          <div className="category-tabs">
            <Pivot selectedKey={this.state.SelectBusinessApplicationPivot} onLinkClick={(item) => this._onPivotChange("business", item)}>
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="Access" itemKey="HS Risk Information – Access" />
              <PivotItem headerText="Asbestos" itemKey="HS Risk Information – Asbestos" />
              <PivotItem headerText="Construction" itemKey="HS Risk Information – Construction" />
              <PivotItem headerText="Demolition" itemKey="HS Risk Information – Demolition" />
              <PivotItem headerText="Electrical" itemKey="HS Risk Information – Electrical" />
              <PivotItem headerText="Kitchen & Bathroom" itemKey="HS Risk Information – Kitchen & Bathroom" />
              <PivotItem headerText="Work Equipment" itemKey="HS Risk Information – Work Equipment" />
              <PivotItem headerText="HSF Axis CLC" itemKey="HSF Axis CLC" />
              <PivotItem headerText="HSP Axis CLC" itemKey="hsp axis clc" />
              <PivotItem headerText="SHEQ Axis" itemKey="SHEQ Axis" />
              <PivotItem headerText=" SHEQ CLC" itemKey=" SHEQ CLC" />
              <PivotItem headerText="SHEQ Concept" itemKey="SHEQ Concept" />
            </Pivot>
          </div>

          <div className="content-grid">

            <div className="doc-card">
              <div className="doc-header">Latest</div>

              <ul className="doclist">
                {
                  Array.isArray(this.state.BusinessApplicationData)
                    ? [...this.state.BusinessApplicationData]
                      .sort((a, b) => {
                        const dateA = a.Modified ? new Date(a.Modified).getTime() : 0;
                        const dateB = b.Modified ? new Date(b.Modified).getTime() : 0;
                        return dateB - dateA;
                      })
                      .slice(0, 5)
                      .map((doc) => (
                        <li key={doc.Id} className="docs-items">
                          <div className="docfile-info">
                            <span className="doc-icon">
                              <img src={require('../assets/Vector.png')} />
                            </span>
                            <div className="doc-name">{doc.FileLeafRef}</div>
                          </div>
                          <a href={doc.FileRef} target="_blank" data-interception="off" download className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
                        </li>
                      ))
                    : null
                }
              </ul>
            </div>

            <div className="documents-card">
              <ul className="doc-list scroll">
                {
                  this.state.BusinessApplicationFilterdData.length > 0 &&
                  this.state.BusinessApplicationFilterdData.map((doc) => {
                    return (
                      <li key={doc.Id} className="docsPolice-items">
                        <div className="docsPolice-info">
                          <span className="docsPolice-icon">
                            <img src={require('../assets/Vector.png')} />
                          </span>
                          <div className="docsPolice-name">{doc.FileLeafRef}</div>
                        </div>
                        <a href={doc.FileRef} target="_blank" data-interception="off" download className="download"><img src={require('../assets/DownloadIcon.png')} /></a>
                      </li>
                    );
                  })
                }

              </ul>
            </div>

          </div>
        </div>

        {/* ---------------------------------------Brochures-------------------------------------------------- */}

        <div id="TempCatalog" className="temps-wrapper">

          <h2 className="page-title">Brochures</h2>

          {
            this.state.IsAdmin ?
            <>
              <div className='BrochureButton'>
                <a href="https://axiseuropeplc.sharepoint.com/sites/GroupIntranet/Lists/Brochures/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: "none", color: 'inherit' }}>
                  <PrimaryButton className='AddDocPolicie' text="Add Brochures" />
                </a>
              </div>
            </>
            :
            <>
            </>
          }

          <div className="category-tabs">

            <Pivot selectedKey={this.state.SelectBrochurePivot} onLinkClick={(item) => this._onPivotChange("brochure", item)}>
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="General Brochure" itemKey="general brochure" />
              <PivotItem headerText="Decarbonisation" itemKey="decarbonisation" />
              <PivotItem headerText="Fire Safety" itemKey="fire safety" />
              <PivotItem headerText="Damp and Mould" itemKey="damp and mould" />
              <PivotItem headerText="M&E" itemKey="m&e" />
              <PivotItem headerText="Healthcare" itemKey="healthcare" />
              <PivotItem headerText="Education" itemKey="education" />
            </Pivot>
          </div>

          <div className="brochures-container">

            {/* LEFT: LATEST BROCHURES */}
            <div className="brochures-latest-card">
              <div className="brochures-header">Latest</div>

              <ul className="brochures-latest-list">
                {Array.isArray(this.state.BrochuresData) &&
                  [...this.state.BrochuresData]
                    .sort((a, b) => {
                      const dateA = a.Modified ? new Date(a.Modified).getTime() : 0;
                      const dateB = b.Modified ? new Date(b.Modified).getTime() : 0;
                      return dateB - dateA;
                    })
                    .slice(0, 5)
                    .map((item) => (
                      <a
                        href={item.Link.Url}
                        target="_blank"
                        rel="noopener noreferrer"
                        data-interception="off"
                        style={{ textDecoration: "none", color: "inherit" }}
                      >
                        <li key={item.Id} className="brochures-latest-item">

                          <div className="brochures-file-info">
                            <img
                              src={require("../assets/Vector.png")}
                              className="brochures-file-icon"
                            />
                            <span className="brochures-file-name">{item.Title}</span>
                          </div>

                          <a
                            href={item.Link}
                            target="_blank"
                            rel="noopener noreferrer"
                            data-interception="off"
                            download
                            className="brochures-download-btn"
                          >
                            {/* <img src={require("../assets/DownloadIcon.png")} /> */}
                          </a>

                        </li>

                      </a>


                    ))}
              </ul>
            </div>

            {/* RIGHT: BROCHURE PREVIEW GRID */}
            <div className="brochures-preview-card">
              <ul className="brochures-grid scroll">

                {this.state.BrochuresFilterdData.length > 0 &&
                  this.state.BrochuresFilterdData.map((item) => {
                    return (
                      <a
                        href={item.Link.Url}
                        target="_blank"
                        rel="noopener noreferrer"
                        data-interception="off"
                        style={{ textDecoration: "none", color: "inherit" }}
                      >
                        <li key={item.Id} className="brochures-grid-item">

                          <div className="brochures-image-card">
                            <img src={item.Image} alt={item.Title} />
                            <p title={item.Title}>{item.Title}</p>
                          </div>

                          <a download className="brochures-download-icon">
                            {/* <img src={require("../assets/DownloadIcon.png")} /> */}
                          </a>

                        </li>
                      </a>
                    );
                  }
                  )}
              </ul>
            </div>

          </div>

        </div>

      </section>
    );
  }

  public async componentDidMount() {
    this.getBusinessAnnouncementData();
    this.getLatestResourcesData();
    this.getDocumentPoliciesData();
    this.getTemplatesData();
    this.getInsuranceData();
    this.getGuidesData();
    this.getBusinessAppData();
    this.getBrochuresData();
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
      console.error("Error checking admin status:", error);
    }
  }

  public async getBusinessAnnouncementData(): Promise<void> {
    try {
      const items: any[] = await sp.web.lists
        .getByTitle("Business Announcement")
        .items
        .select(
          "ID",
          "Title",
          "Description",
          "Source",
          "Images",
          "Videos",
          "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .get();

      let AllData: any[] = [];

      if (items && items.length > 0) {

        items.forEach((item: any) => {

          let imageUrl: string = "";
          let videoUrl: string = "";

          if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {

            const file = item.AttachmentFiles[0];
            const fileName = file.FileName.toLowerCase();

            if (fileName.match(/\.(jpg|jpeg|png|gif)$/)) {
              imageUrl = file.ServerRelativeUrl;
            }
            else if (fileName.match(/\.(mp4|webm|ogg|mov|avi|m4v)$/)) {
              videoUrl = file.ServerRelativeUrl;
            }
          }

          let videoColumnUrl: string = "";

          if (item.Videos) {

            // Case 1: Hyperlink field object
            if (typeof item.Videos === "object" && item.Videos.Url) {
              videoColumnUrl = item.Videos.Url;
            }

            // Case 2: Direct string
            else if (typeof item.Videos === "string") {
              videoColumnUrl = item.Videos;
            }
          }

          AllData.push({
            ID: item.ID || "",
            Title: item.Title || "",
            Description: item.Description || "",
            Source: item.Source || "",
            Images: imageUrl,
            Videos: videoUrl || videoColumnUrl,
            // link: item.link ? (item.link.Url ? item.link.Url : item.link) : ""
          });

        });

        this.setState({
          BusinessAnnouncementsData: AllData
        });
      }

    } catch (error) {
      console.log("Error Fetching details :", error);
    }
  }

  public async getLatestResourcesData() {
    try {
      const doc = await sp.web.lists.getByTitle("Company handbook").items.orderBy("Created", false).top(5).select(
        "Id", "FileLeafRef", "FileRef", "Modified", "Editor/Title", "File_x0020_Type"
      ).expand("Editor").get();
      this.setState({ CompanyDocumentsData: doc });
    } catch (error) {
      console.log("Error fetching Company Documents data: ", error);
    }

  }

  public async getDocumentPoliciesData() {
    try {
      const docPolicie = await sp.web.lists.getByTitle("Documents and Policies").items.select(
        "Id", "FileLeafRef", "FileRef", "Modified", "Editor/Title", "File_x0020_Type", "Category", "File",
        "EncodedAbsUrl", "File/ServerRelativeUrl"
      ).expand("Editor,File").getAll();

      this.setState({ DocumentsAndPoliciesData: docPolicie, DocumentsAndPoliciesFilterdData: docPolicie });
      console.log("Documents and Policies data: ", docPolicie);
    } catch (error) {
      console.log("Error fetching Documents and Policies data: ", error);
    }
  }

  // public async getTemplatesData() {
  //   try {
  //     const docPolicie = await sp.web.lists.getByTitle("Templates and Branding").items.select(
  //       "Id","FileLeafRef","FileRef","Modified","Editor/Title","File_x0020_Type","Category"
  //     ).expand("Editor").get();

  //     this.setState({ TemplatesandBrandingData : docPolicie , TemplatesandBrandingFilterdData : docPolicie });
  //   } catch (error) {
  //     console.log("Error fetching Documents and Policies data: ", error);
  //   }
  // }

  // public async getTemplatesData() {
  //   try {
  //     const docPolicie = await sp.web.lists
  //       .getByTitle("Templates and Branding")

  //       .items.select(
  //         "Id",
  //         "FileLeafRef",
  //         "FileRef",
  //         "UniqueId",
  //         "Modified",
  //         "Editor/Title",
  //         "File_x0020_Type",
  //         "Category",
  //         "File",
  //         "EncodedAbsUrl",
  //         "File/ServerRelativeUrl"
  //       ).expand("Editor,File")
  //       .get();

  //     this.setState({
  //       TemplatesandBrandingData: docPolicie,
  //       TemplatesandBrandingFilterdData: docPolicie
  //     });
  //     console.log("Templates and Branding data: ", docPolicie);
  //   } catch (error) {
  //     console.log("Error fetching Documents and Policies data: ", error);
  //   }
  // }

  public async getTemplatesData() {
    try {
      const docPolicie = await sp.web.lists
        .getByTitle("Templates and Branding")
        .items.select(
          "Id",
          "FileLeafRef",
          "FileRef",
          "UniqueId",
          "Modified",
          "Editor/Title",
          "File_x0020_Type",
          "Category",
          "EncodedAbsUrl"
        )
        .expand("Editor")
        .getAll();

      this.setState({
        TemplatesandBrandingData: docPolicie,
        TemplatesandBrandingFilterdData: docPolicie
      });

      console.log("Templates and Branding data: ", docPolicie);

    } catch (error) {
      console.log("Error fetching Documents and Policies data: ", error);
    }
  }

  // public async getInsuranceData() {
  //   try {
  //     const insurancecerti = await sp.web.lists.getByTitle("Insurance and Certificates").items.select(
  //       "Id", "FileLeafRef", "FileRef", "Modified", "Editor/Title", "File_x0020_Type", "Category", "File",
  //       "EncodedAbsUrl", "File/ServerRelativeUrl"
  //     ).expand("Editor,File").orderBy("FileLeafRef", true).getAll();

  //     this.setState({ InsuranceandCertificateData: insurancecerti, InsuranceandCertificateFilterdData: insurancecerti });
  //     console.log("Insurance and Certificates data: ", insurancecerti);
  //   } catch (error) {
  //     console.log("Error fetching Documents and Policies data: ", error);
  //   }
  // }

  public async getInsuranceData(): Promise<void> {
    try {
      const insurancecerti = await sp.web.lists
        .getByTitle("Insurance and Certificates")
        .items
        .select(
          "Id",
          "FileLeafRef",
          "FileRef",
          "Modified",
          "Editor/Title",
          "File_x0020_Type",
          "Category",
          "EncodedAbsUrl",
          "File/ServerRelativeUrl",
          "File/Name"
        )
        .expand("Editor", "File")
        // Sort by the actual file name from the File object
        .orderBy("File/Name", true)
        .getAll();
  
      // Optional: client-side sorting to ensure proper alphabetical order
      insurancecerti.sort((a, b) => {
        const nameA = (a.File?.Name || a.FileLeafRef || "").toLowerCase();
        const nameB = (b.File?.Name || b.FileLeafRef || "").toLowerCase();
        return nameA.localeCompare(nameB);
      });
  
      this.setState({
        InsuranceandCertificateData: insurancecerti,
        InsuranceandCertificateFilterdData: insurancecerti
      });
  
      console.log("Insurance and Certificates data:", insurancecerti);
    } catch (error) {
      console.error("Error fetching Insurance and Certificates data:", error);
    }
  }

  public async getGuidesData() {
    try {
      const guide = await sp.web.lists.getByTitle("How to Guides").items.select(
        "Id", "FileLeafRef", "FileRef", "Modified", "Editor/Title", "File_x0020_Type", "Category", "File",
        "EncodedAbsUrl", "File/ServerRelativeUrl"
      ).expand("Editor,File").getAll();

      this.setState({ HowtoGuidesData: guide, HowtoGuidesFilterdData: guide });
      console.log("How to Guides data: ", guide);
    } catch (error) {
      console.log("Error fetching Documents and Policies data: ", error);
    }
  }

  public async getBusinessAppData() {
    try {
      const appbusiness = await sp.web.lists.getByTitle("H&S Documents").items.select(
        "Id", "FileLeafRef", "FileRef", "Modified", "Editor/Title", "File_x0020_Type", "Category", "File",
        "EncodedAbsUrl", "File/ServerRelativeUrl"
      ).expand("Editor,File").getAll();

      this.setState({ BusinessApplicationData: appbusiness, BusinessApplicationFilterdData: appbusiness });
      console.log("Business Application Catalogue data: ", appbusiness);
    } catch (error) {
      console.log("Error fetching Documents and Policies data: ", error);
    }
  }

  public async getBrochuresData() {

    const brouchers = await sp.web.lists.getByTitle("Brochures").items.select(
      "ID",
      "Title",
      "Image",
      "Link",
      "Category"
    ).expand("AttachmentFiles").get().then((data) => {
      let AllData = [];
      // console.log(data);
      // console.log(brouchers);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Image: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : item.Image ? JSON.parse(item.Image).serverRelativeUrl : require(`../assets/Frame14.png`),
            Link: item.Link ? item.Link : "#",
            Category: item.Category ? item.Category : ""
          });
        });
        this.setState({ BrochuresData: AllData, BrochuresFilterdData: AllData });
      }
    }).catch((error) => {
      console.log("Error fetching Brochures data: ", error);
    });

    // try {
    //   const brouhers = await sp.web.lists.getByTitle("Brochures Documents").items.select(
    //     "Id", "FileLeafRef", "FileRef", "UniqueId", "Modified", "Editor/Title", "File_x0020_Type", "Category",
    //     "EncodedAbsUrl"
    //   ).expand("Editor").getAll();

    //   this.setState({ BrochuresData: brouhers, BrochuresFilterdData: brouhers });
    //   console.log("Business Application Catalogue data: ", brouhers);
    // } catch (error) {
    //   console.log("Error fetching Documents and Policies data: ", error);
    // }

  }

  private _onPivotChange = (
    section:
      | "docs"
      | "templates"
      | "insurance"
      | "howto"
      | "business"
      | "brochure",
    item?: PivotItem
  ): void => {

    if (!item) return;

    const key = item.props.itemKey || "all";

    const filterByCategory = (data: any[]) =>
      key === "all"
        ? data
        : data.filter(d => d.Category?.toLowerCase() === key.toLowerCase());

    switch (section) {

      case "docs":
        this.setState({
          SelectedDocsPivot: key,
          DocumentsAndPoliciesFilterdData:
            filterByCategory(this.state.DocumentsAndPoliciesData)
        });
        break;

      case "templates":
        this.setState({
          SelectedTemplatesPivot: key,
          TemplatesandBrandingFilterdData:
            filterByCategory(this.state.TemplatesandBrandingData)
        });
        break;

      case "insurance":
        this.setState({
          SelectInsuranceCertificatePivot: key,
          InsuranceandCertificateFilterdData:
            filterByCategory(this.state.InsuranceandCertificateData)
        });
        break;

      case "howto":
        this.setState({
          SelectHowtoGuidePivot: key,
          HowtoGuidesFilterdData:
            filterByCategory(this.state.HowtoGuidesData)
        });
        break;

      case "business":
        this.setState({
          SelectBusinessApplicationPivot: key,
          BusinessApplicationFilterdData:
            filterByCategory(this.state.BusinessApplicationData)
        });
        break;

      case "brochure":
        this.setState({
          SelectBrochurePivot: key,
          BrochuresFilterdData:
            filterByCategory(this.state.BrochuresData)
        });
    }
  }

  private searchDocuments = async () => {

    const libraries = [
      "Company handbook",
      "Documents and Policies",
      "Templates and Branding",
      "Insurance and Certificates",
      // "How to Guides",
      "H&S Documents",
      "Brochures"
    ];

    let allResults: any[] = [];

    for (let lib of libraries) {

      const items = await sp.web.lists
        .getByTitle(lib)
        .items
        .select("Title", "FileLeafRef", "FileRef", "Created")
        .filter(`substringof('${this.state.searchText}',FileLeafRef)`)
        .top(50)
        .get();

      const formatted = items.map(item => ({
        ...item,
        LibraryName: lib
      }));

      allResults = [...allResults, ...formatted];
    }

    this.setState({
      results: allResults
    });

  }

}
