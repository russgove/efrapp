import * as React from 'react';
import styles from './EfrApp.module.scss';
import { DocumentIframe } from './DocumentIframe';
import { IEfrAppProps } from './IEfrAppProps';
import { IEfrAppState } from './IEfrAppState';
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
import { PBCTask, Document } from "../model";
import Dropzone from 'react-dropzone';

export default class EfrApp extends React.Component<IEfrAppProps, IEfrAppState> {
  public constructor(props: IEfrAppProps) {
    super();
    console.log("in Construrctor");
    this.state = {
      documentCalloutIframeUrl: "",
      documents: props.documents
    }
  }
  public createSummaryMarkup(html) {
    console.log("in createSummaryMarkup");
    return { __html: html };
  }
  private onDrop(acceptedFiles, rejectedFiles) {
    console.log("in onDrop");
    debugger;
    let promises :Array<Promise<any>>=[];
    acceptedFiles.forEach(file => {
      promises.push(this.props.uploadFile(file, this.props.task.EFRLibrary, this.props.task.Title));
    });
    Promise.all(promises).then((x)=>{
      this.props.getDocuments(this.props.task.EFRLibrary).then((dox) => {
        this.setState((current) => ({ ...current, documents: dox }));
      });
  
    });

  }
  public uploadFile(e: any) {

    let file: any = e.target["files"][0];
    this.props.uploadFile(file, this.props.task.EFRLibrary, this.props.task.Title).then((response) => {
      this.props.getDocuments(this.props.task.EFRLibrary).then((dox) => {
        this.setState((current) => ({ ...current, documents: dox }));
      });
    }).catch((error) => {
      console.error("an error occurred uploading the file");
      console.error(error);
    });
  }
  public getDateString(dateString: string): string {
    console.log("in getDateString");

    var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    let year = parseInt(dateString.substr(0, 4));
    let month = parseInt(dateString.substr(5, 2));
    let day = parseInt(dateString.substr(8, 2));
    return new Date(year, month, day)
      .toLocaleDateString(this.props.cultureInfo.currentCultureName, options)

  }
  public documentRowMouseEnter(document: Document, e: any) {
    console.log("in documentRowMouseEnter");

    // mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 3, this.props.task.EFRLibrary).then(url => {
      if (!url || url === "") {
        url = document.serverRalativeUrl;
      }
      // this.state.documentCalloutIframeUrl = url;
      // this.state.documentCalloutTarget = e.target;
      // this.state.documentCalloutVisible = true;
      this.setState((current) => ({
        ...current,
        documentCalloutIframeUrl: url,
        documentCalloutTarget: e.target,
        documentCalloutVisible: true
      }));

    });
  }
  public documentRowMouseOut(item: Document, e: any) {
    console.log("in documentRowMouseOut");

    // this.state.documentCalloutTarget = null;
    // this.state.documentCalloutVisible = false;
    this.setState((current) => ({ ...current, documentCalloutTarget: null, documentCalloutVisible: false }));
    console.log("mouse exit for " + item.title);
  }
  public editDocument(document: Document): void {
    console.log("in editDocument");

    debugger;

    // mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 1, this.props.task.EFRLibrary).then(url => {
      console.log("wopi frame url is " + url)
      if (!url || url === "") {
        window.open(document.serverRalativeUrl, "_blank");
      } else {
        window.open(url, "_blank");
      }
      //    this.state.wopiFrameUrl=url;
      //  this.setState(this.state);
      // window.location.href=url;

    });

  }
  public render(): React.ReactElement<IEfrAppProps> {
    console.log("in render");
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
        </div >
        
        <Label >Information Requested: </Label>
        <div className={styles.informationRequested}
         dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRInformationRequested)} />
          <Label >Please upload the files containing the information requested.
             You can drag and drop file(s) into the area shaded in blue below, or click the 
             'Choose File' button to select the file(s). Files you upload will be prefixed
             with the reference {[this.props.task.Title}}   </Label>
        <Dropzone className={styles.dropzone} onDrop={this.onDrop.bind(this)} disableClick={true} >
          <div>
            Drag and drop files here to upload
          </div>
          <div style={{ float: "left" }}>
            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              items={this.state.documents}
              onRenderRow={(props, defaultRender) => <div
                onMouseEnter={(event) => this.documentRowMouseEnter(props.item, event)}
                onMouseOut={(evemt) => this.documentRowMouseOut(props.item, event)}>
                {defaultRender(props)}
              </div>}
              setKey="id"
              selectionMode={SelectionMode.none}
              columns={[
                {
                  key: "Edit", name: "", fieldName: "Title", minWidth: 20,
                  onRender: (item) => <div>
                    <i onClickCapture={(e) => { debugger; e.preventDefault(); this.editDocument(item); return false; }}
                      className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>

                  </div>
                },
                { key: "title", name: "File Name", fieldName: "title", minWidth: 1, maxWidth: 500 , },

              ]}
            />
            <input type="file" id="uploadfile" onChange={e => { this.uploadFile(e); }} />
          </div>
          <div style={{ float: "right" }}>
            <DocumentIframe src={this.state.documentCalloutIframeUrl}
              height={this.props.documentIframeHeight}
              width={this.props.documentIframeWidth} />
          </div>
          <div style={{ clear: "both" }}></div>

        </Dropzone>

      </div>
    );
  }
}
