import * as React from 'react';
import styles from './HomePagePollRequest.module.scss';
import { IHomePagePollRequestProps } from './IHomePagePollRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { Chart } from 'chart.js';

export interface IHomePagePollRequestState {
  options: any;
  question: any;
  hasVoted: boolean;
  selectedOption: any;
  TotalResponses: any;
  counts: any;
  SurveyData: any;
  SurveyResponseData: any;
}

require('../assets/style.css');

export default class HomePagePollRequest extends React.Component<IHomePagePollRequestProps, IHomePagePollRequestState> {

  constructor(props: IHomePagePollRequestProps, state: IHomePagePollRequestState) {
    super(props);

    this.state = {
      options: [],
      question: "",
      hasVoted: false,
      selectedOption: "",
      TotalResponses: "",
      counts: "",
      SurveyData: "",
      SurveyResponseData: ""
    };


  }

  public render(): React.ReactElement<IHomePagePollRequestProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="homePagePollRequest">

        <div className="Poll-card">
          <div className="Poll-header">
            <h3>Poll</h3>
            <span className="arrow">⌃</span>
          </div>


          {
            this.state.hasVoted ?
              <>
                <h3>{this.state.question}</h3>

                <canvas id="pollChart" height="250"></canvas>

              </> : <>

                <h3>{this.state.question}</h3>

                {
                  this.state.options.map((opt, index) => {

                    return (
                      <div key={index} style={{ marginBottom: 8 }}>
                        <input
                          type="radio"
                          name="poll"
                          value={opt}
                          onChange={(e) => this.setState({ selectedOption: e.target.value })}
                        />
                        <span style={{ marginLeft: 8 }}>{opt}</span>
                      </div>
                    );

                  })
                }

                <button
                  style={{
                    marginTop: 12,
                    padding: "6px 16px",
                    cursor: "pointer"
                  }}
                  className='PollButton'
                  onClick={() => this.submitVote()}
                >
                  Submit
                </button>

              </>
          }

        </div>


        {/* ---------------------------------------------------------------- */}


        <div className="user-widget">

          <h3 className="hello">Hello {userDisplayName}</h3>
          <div className="hello-underline"></div>


          <a href="https://axiseuropeplc.sharepoint.com/sites/AxisLMS/SitePages/My-Training-Dashboard.aspx?isSPOFile=1" style={{ textDecoration: "none", color: "black" }}>
            <div className="stat-card">
              <div className='context'>
                <img src={require("../assets/checkcirclebroken.png")} /><span> My Training</span>
              </div>
              {/* <strong>3</strong> */}
            </div>
          </a>


          {/* <div className="stat-card">
                <div className='context'>
                  <img src={require("../assets/icon2.png")} /><span> My Approvels</span>
                </div>
                <strong>4</strong>
              </div> */}

          <a href='https://servicedesk.axisclc.com/portal/tickets?btn=60&viewid=1' style={{ textDecoration: "none", color: "black" }}>
            <div className="stat-card">
              <div className='context'>
                <img src={require("../assets/ticket01.png")} /><span> My IT Tickets</span>
              </div>
              {/* <strong>5</strong> */}
            </div>
          </a>


          {/* <h4>My Favorite Articles</h4>

              <ul className="fav-list">
                <li>Better Understanding your patients needs</li>
                <li>401k Updates fpr 2020</li>
                <li className="active">Covid Frequently Asked Questions</li>
                <li>HR Polices and Procedures Guidelines</li>
              </ul>

              <a className="views-all" href="#">View all</a> */}



          <h4>Useful Apps</h4>

          <div className="apps-grid">

            {/* <a href='https://axiseurope.crm4.dynamics.com/main.aspx?appid=9fa6e94b-63a5-4a31-89d6-6298402f0d3e&pagetype=dashboard&type=system&_canOverride=true' style={{ textDecoration: "none" }}><div className="app-card">Dynamics CE <img className='next-i' src={require("../assets/icon.png")} /></div></a> */}
            <a href='https://uk.sheassure.net/clc' style={{ textDecoration: "none" }}><div className="app-card">Evotix <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://bit.ly/4l6gNQc' style={{ textDecoration: "none" }}><div className="app-card" >Outlook <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://go.accessacloud.com/o/repbp/workspaces/98d34671c16d4d2e9e1429a2fd965ec2/Access.PeopleXDEmpMain/2f1f1cba97924b2b891fc2f51a13677a?location=https%3A%2F%2Fmy.xd.accessacloud.com%2Fpls%2Fcoreportal_repbp%2Fi%23EmpMain%2Fmytime' style={{ textDecoration: "none" }}><div className="app-card">PeopleXD <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://teams.cloud.microsoft/' style={{ textDecoration: "none" }}><div className="app-card">Teams <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://servicedesk.axisclc.com/' style={{ textDecoration: "none" }}><div className="app-card">Halo <img className='next-i' src={require("../assets/icon.png")} /></div></a>
            <a href='https://go.accessacloud.com/o/repbp/workspaces/28609fedc58441bdbf7a8a4cbe52b1c7/Access.Product.Learning/f899dafaf580404cb513eeab0849d751?location=https%3A%2F%2Faxisclcgroup.lms.accessacloud.com%2Fw%2Fhome' style={{ textDecoration: "none" }}><div className="app-card">Training (LMS) <img className='next-i' src={require("../assets/icon.png")} /></div></a>
          </div>

        </div>


      </section>
    );
  }

  public async componentDidMount() {

    await this.loadSurvey();
    await this.checkUserVote();
    await this.loadResults();
     this.getSurveyInfo();


    Chart.pluginService.register({
      beforeDraw: function (chart) {
        if (chart.config.options.centerText) {
          const ctx = chart.chart.ctx;
          const txt = chart.config.options.centerText;

          ctx.save();
          ctx.font = "bold 18px Arial";
          ctx.fillStyle = "#ffffff";
          ctx.textAlign = "center";
          ctx.textBaseline = "middle";

          const centerX = chart.chart.width / 2;
          const centerY = chart.chart.height / 1.8;

          ctx.fillText(txt, centerX, centerY);
          ctx.restore();
        }
      }
    });

  }

  private async loadSurvey() {

    const items = await sp.web.lists
      .getByTitle("Poll")
      .items
      .select("Id", "Question", "Options")
      .top(1)
      .get();

    if (items.length === 0) return;

    const item = items[0];

    /* If Options column is Choice (single) with semicolon values */
    let optionsArray: string[] = [];

    if (item.Options) {
      optionsArray = item.Options;
    }

    this.setState({
      question: item.Question,
      options: optionsArray
    });

  }

  /* Check if user already voted */
  private async checkUserVote() {

    const userId = this.props.context.pageContext.legacyPageContext.userId;

    const items = await sp.web.lists
      .getByTitle("Poll Response")
      .items
      .select("Id", "Author/Id", "Title")
      .expand("Author")
      .filter(`Author/Id eq ${userId} and Title eq '${this.state.question}'`)
      .get();

    if (items.length > 0) {
      this.setState({ hasVoted: true });
    }

  }

  private async submitVote() {

    if (!this.state.selectedOption) {
      alert("Please select option");
      return;
    }

    const email = this.props.context.pageContext.user.email;

    await sp.web.lists.getByTitle("Poll Response").items.add({
      Title: this.state.question,
      // UserEmail: email,
      Option: this.state.selectedOption
    });

    this.setState({ hasVoted: true });

    await this.loadResults();
  }

  /* Load aggregated results */
  // private async loadResults() {

  //   const items = await sp.web.lists
  //     .getByTitle("Survey Response")
  //     .items
  //     .select("Option")
  //     .get();

  //   let counts = [0, 0, 0, 0];

  //   items.forEach((item: any) => {

  //     if (item.Option === "test 1") counts[0]++;
  //     if (item.Option === "test 2") counts[1]++;
  //     if (item.Option === "Option 3") counts[2]++;
  //     if (item.Option === "Option 4") counts[3]++;

  //   });

  //   this.setState({ counts: counts }, () => {

  //     if (this.state.hasVoted) {
  //       this.renderChart();
  //     }

  //   });
  // }

  private async loadResults() {

    const items = await sp.web.lists
      .getByTitle("Poll Response")
      .items
      .select("Option")
      .get();

    let counts: number[] = [];

    for (let i = 0; i < this.state.options.length; i++) {
      counts.push(0);
    }
    items.forEach((item: any) => {

      this.state.options.forEach((opt, index) => {
        if (item.Option === opt) {
          counts[index]++;
        }
      });

    });

    this.setState({ counts: counts, TotalResponses: items.length }, () => {

      if (this.state.hasVoted && this.state.options.length > 0) {
        this.renderChart();
      }

    });

  }

  /* Render Pie Chart */
  private renderChart() {

    const backgroundColors = [
      'rgba(95, 255, 214, 0.6)',
      'rgba(217, 198, 255, 0.6)',
      'rgba(255, 210, 168, 0.6)',
      'rgba(41, 199, 217, 0.6)',
    ];

    setTimeout(() => {

      const canvas = document.getElementById("pollChart") as HTMLCanvasElement;
      if (!canvas) return;

      new Chart(canvas, {
        type: 'doughnut',
        data: {
          labels: this.state.options,
          datasets: [{
            data: this.state.counts,
            backgroundColor: backgroundColors.slice(0, this.state.options.length),
          }]
        },
        options: {
          cutoutPercentage: 80,
          legend: {
            labels: {
              fontColor: "#ffffff"
            }
          },
          centerText: this.state.TotalResponses + " Responses"
        }
      });

    }, 300);
  }

  public async getSurveyInfo() {
    const poll = await sp.web.lists.getByTitle("Poll").items.select(
      "ID",
      "Question",
      "Answer",
      "Options"
    ).get().then((data) => {
      let AllData = [];
      console.log(poll);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Question: item.Question ? item.Question : "",
            Answer: item.Answer ? item.Answer : "",
            Options: item.Options ? item.Options : ""
          });
        });
        this.setState({ SurveyData: AllData });
      }
    }).catch((error) => {
      console.log("Error Fetching details ", error);
    });
  }

  public async getSurveyresponse() {
    const response = await sp.web.lists.getByTitle("Poll Response").items.select(
      "ID",
      "Title",
      "Option",
      "PersonName"
    ).get().then((data) => {
      let AllData = [];
      console.log(response);
      console.log(data);
      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.ID ? item.ID : "",
            Title: item.Title ? item.Title : "",
            Option: item.Option ? item.Option : "",
            PersonName: item.PersonName ? item.PersonName : ""
          });
        });
        this.setState({ SurveyResponseData: AllData });
      }
    });
  }

}
