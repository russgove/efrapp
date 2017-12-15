import * as React from 'react';
import styles from './EfrApp.module.scss';
import { IEfrAppProps } from './IEfrAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {PageContext } from "@microsoft/sp-page-context";

export default class EfrApp extends React.Component<IEfrAppProps, {}> {


  public render(): React.ReactElement<IEfrAppProps> {
    debugger;
    return (
      <div className={styles.efrApp}>
        <div>{this.props.task.Title}</div>
      </div>
    );
  }
}
