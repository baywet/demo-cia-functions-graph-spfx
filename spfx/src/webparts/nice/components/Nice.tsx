require("./Nice.norename.css");
import * as React from "react";
import styles from "./Nice.module.scss";
import { INiceProps } from "./INiceProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {ProgressIndicator } from "office-ui-fabric-react";

export default class Nice extends React.Component<INiceProps, {}> {
  public render(): React.ReactElement<INiceProps> {
    const score: number = Math.round(this.props.score * 100);
    return (
      <div className={ styles.nice }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to CIA!</span>
              <h1>vincent@baywet.onmicrosoft.com</h1>
              <p className={ styles.subTitle }>Evaluate the impact you have on people around you.</p>
              <ProgressIndicator
                label="Canadian Completeness."
                description={`${String(score)} % Not bad, eh?`}
                percentComplete={(score / 100)}
              />
              <button onClick={this.props.registerWebHook} className={ styles.button }>Register WebHook</button>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
