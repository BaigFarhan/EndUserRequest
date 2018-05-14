import * as React from 'react';
import styles from './WebAtcHr.module.scss';
import { IWebAtcHrProps } from './IWebAtcHrProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Animate from 'react-simple-animate';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { CommandButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { GridForm, Fieldset, Row, Field } from 'react-gridforms'

export default class WebAtcHr extends React.Component<IWebAtcHrProps, {}> {

  public state: IWebAtcHrProps;
  constructor(props, context) {
    super(props);
    this.state = {
      description: "",
      SiteUrl: "",
      PassportRequest: 0,
      LeaveRequest: 0,
      FormIsEnabled: 0,
    }
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');
  }


  public render(): React.ReactElement<IWebAtcHrProps> {
    return (
      <div className={styles.webAtcHr}>

        <div className={styles.containerLeave}>
          <div className={styles.HeaderPassport} ></div>
          <div>
            <p className={styles.pragraph}> Leave Request </p>
          </div>
        </div>

        <div className={styles.Passport}>
          <div className={styles.HeaderDivLeave} ></div>
          <div>
            <p className={styles.pragraph}> Passport Request </p>
          </div>
        </div>

        <div className={styles.Server}>
          <div className={styles.HeaderPassport} ></div>
          <div>
            <p className={styles.pragraph}> Server Request </p>
          </div>
        </div>

        <div className={styles.Sick}>
          <div className={styles.HeaderSick} ></div>
          <div>
            <p className={styles.pragraph}> Sick Request </p>
          </div>
        </div>


        <div className={styles.Allowance}>
          <div className={styles.HeaderAllownce} ></div>
          <div>
            <p className={styles.pragraph}> Allowance Request </p>
          </div>
        </div>

        <div className={styles.Users}>
          <div className={styles.HeaderUsers} ></div>
          <div>
            <p className={styles.pragraph}> User Creation Request </p>
          </div>
        </div>
      


      </div>
    );
  }
}
