import * as React from "react";
//import styles from './Kudo.module.scss';
import { IKudoProps } from "./IKudoProps";
import styles from "./Kudo.module.scss";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
//import { escape } from '@microsoft/sp-lodash-subset';
interface IKudoDetails {
  Title: string;
  Person: {
    Title: string;
    EMail: string;
  };
  Details: string;
  Author: {
    Title: string;
    EMail: string;
  };
  Status: string;
}
interface IAllKudo {
  AllKudos: IKudoDetails[];
}

export default class Kudo extends React.Component<IKudoProps, IAllKudo> {
  properties: any;
  AuthorNews: any;
  constructor(props: IKudoProps, state: IAllKudo) {
    super(props);
    this.state = {
      AllKudos: [],
    };
  }
  componentDidMount() {
    //alert ("Componenet Did Mount Called...");
    //console.log("First Call.....");
    this.getKudoData();
  }

  public getKudoData = () => {
    console.log("This is Kudos Detail function");
    let listName = this.props.listName;
    let selectecolumns = "*,Person/Title,Person/EMail";
    //console.log(selectecolumns);
    let expandcolumn = `Person`;
    let filterQuery = `Status eq 'Approve'`;
    let top = this.props.noofKudos;
    let orderQuery = `Modified desc`;
    //api call
    let listURL = `${this.props.siteURL}/_api/web/lists/getbytitle('${listName}')/items?$select=${selectecolumns}&$expand=${expandcolumn}&$filter=${filterQuery}&$top=${top}&$orderby=${orderQuery}`;
    console.log(listURL);
    this.props.context.spHttpClient
      .get(listURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          //console.log(responseJSON);
          this.setState({ AllKudos: responseJSON.value });
        });
        console.log(this.state.AllKudos);
      });
  };
  public render(): React.ReactElement<IKudoProps> {
    return (
      <div>
        <div className={styles.component}>
          <p>
            {this.props.componentTitle}{" "}
            <a href={this.props.seeAllPageURL} style={{ fontSize: "15px" }}>
              SeeAll
            </a>
          </p>
        </div>

        {/*   Empty Message */}
        <div
          style={{ display: this.state.AllKudos.length === 0 ? "" : "none" }}
        >
          <p>{this.props.emptyMessage}</p>
        </div>

        <div
          style={{
            height: this.props.webHeight,
          }}
        >
          {this.state.AllKudos.map((kDet) => {
            return (
              <div className={styles.maincontainer}>
                <div className={styles.KudoTo}>
                  <div className={styles.row}>
                    <img
                      src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?accountname=${kDet.Person.EMail}&size=L`}
                      alt=""
                    />
                    <p>{kDet.Person.Title}</p>
                  </div>
                </div>
                <div className={styles.Paragraph}>
                  <h3>{kDet.Title}</h3>
                  <p>{kDet.Details}</p>
                  {/* <button onClick={} >See All</button> */}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}

/*   <>
        <div>{this.props.componentTitle}

        <div
          style={{ display: this.state.AllKudos.length === 0 ? "" : "none" }}
        >
          <p>{this.props.emptyMessage}</p>
        </div>
        <div className={styles.maincontainer}>
        {this.state.AllKudos.map((kDet) => {
            return (

              <><div className={styles.KudoTo}>

                <img
                  src={kDet.KudoTo == null
                    ? require("./Image/images1.png")
                    : window.location.origin +
                    JSON.parse(kDet.KudoTo).serverRelativeUrl}
                  alt="" />
              </div><div>Details</div></>
            )},
        </div></div>
      </>
    );
  }
}
 */
