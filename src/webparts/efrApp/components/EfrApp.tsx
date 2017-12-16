import * as React from 'react';
import styles from './EfrApp.module.scss';
import { IEfrAppProps } from './IEfrAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PageContext } from "@microsoft/sp-page-context";
import { PrimaryButton, ButtonType } from "office-ui-fabric-react/lib/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";

import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { Label } from "office-ui-fabric-react/lib/Label";

import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
// switch to fabric  ComboBox on next upgrade
// let Select = require("react-select") as any;
//import "react-select/dist/react-select.css";
import pnp from "sp-pnp-js";
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from "office-ui-fabric-react/lib/DetailsList";
import { DatePicker, } from "office-ui-fabric-react/lib/DatePicker";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";

import Dropzone from 'react-dropzone';

export default class EfrApp extends React.Component<IEfrAppProps, {}> {
  public createSummaryMarkup(html) {
    return { __html: html };
  }
  private onDrop(acceptedFiles, rejectedFiles) {
    debugger;
    acceptedFiles.forEach(file => {
           this.props.uploadFile(file, this.props.task.EFRLibrary);
    });
  }
  public getDateString( dateString:string):string{
    var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    let year =parseInt(dateString.substr(0,4));
    let month =parseInt(dateString.substr(5,2));
    let day =parseInt(dateString.substr(8,2));
    return new Date(year,month,day)
    .toLocaleDateString(this.props.cultureInfo.currentCultureName,options)
    
  }
  public render(): React.ReactElement<IEfrAppProps> {
    debugger;
  
;

    return (
      <div className={styles.efrApp}>
        <div className={styles.headerArea}>

          <TextField label="Referenece"
            className={styles.inline}
            style={{ width: 100 }}
            disabled={true}
            value={this.props.task.Title} />


          <TextField 
          className={styles.inline}
           label="DueDate" 
           disabled={true} 
           style={{ width: 180 }} 
           value={this.getDateString(this.props.task.EFRDueDate)}
            />


          <TextField label="Library" className={styles.inline} disabled={true} style={{ width: 120 }} value={this.props.task.EFRLibrary} />


          <TextField label="Period" className={styles.inline} disabled={true} style={{ width: 80 }} value={this.props.task.EFRPeriod} />


          <TextField className={styles.inline} disabled={true} style={{ width: 80 }} value={this.props.task.AssignedTo} />

          <Label >Information Requested: </Label>
          <div dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRInformationRequested)} />
          <Dropzone  onDrop={this.onDrop.bind(this)} >
            <div>
              Drag and drop files here to upload
          </div>
          </Dropzone>
        </div >
      </div>
    );
  }
}
