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
  private validBrandIcons = " accdb csv docx dotx mpp mpt odp ods odt one onepkg onetoc potx ppsx pptx pub vsdx vssx vstx xls xlsx xltx xsn ";

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
    let promises: Array<Promise<any>> = [];
    acceptedFiles.forEach(file => {
      promises.push(this.props.uploadFile(file, this.props.task.EFRLibrary, this.props.task.Title));
    });
    Promise.all(promises).then((x) => {
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
      .toLocaleDateString(this.props.cultureName, options)

  }
  public documentRowMouseEnter(document: Document, e: any) {
    console.log("in documentRowMouseEnter");

    // mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 0, this.props.task.EFRLibrary).then(url => {
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
    this.props.fetchDocumentWopiFrameURL(document.id, 0, this.props.task.EFRLibrary).then(url => {
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
  public getAssignees(assignees:Array<{}>):string{
    let result ="";
    for (let assignee of assignees){
      result+=assignee["Title"]+";  ";
    }
    return result;
  }
  public renderItemTitle(item?: any, index?: number, column?: IColumn): any {
    debugger;
    let extension = item.title.split('.').pop();
    let classname = "";
    if (this.validBrandIcons.indexOf(" " + extension + " ") !== -1) {
      classname += " ms-Icon ms-BrandIcon--" + extension + " ms-BrandIcon--icon16 ";
    }
    else {
      //classname += " ms-Icon ms-Icon--TextDocument " + styles.themecolor;
      classname += " ms-Icon ms-Icon--TextDocument ";
    }


    return (
      <div>
        <div className={classname} /> &nbsp;
        <a href="#"
          onClickCapture={(e) => { debugger; e.preventDefault(); this.editDocument(item); return false; }}>{item.title}</a>
      </div>);
  }
  public render(): React.ReactElement<IEfrAppProps> {
    console.log("in render");
    return (
      <div className={styles.efrApp}>
        <div className={styles.headerArea}>
          <table>
            <tr>
              <td>
                <Label>Reference:</Label>
              </td>
              <td>
                <TextField label=""

                  disabled={true}
                  value={this.props.task.Title} />
              </td>
            </tr>
            <tr>
              <td>
                <Label>Due Date:</Label>
              </td>
              <td>
                <TextField label=""


                  disabled={true}
                  value={this.getDateString(this.props.task.EFRDueDate)} />
              </td>
            </tr>
            <tr>
              <td>
                <Label>Period:</Label>
              </td>
              <td>
                <TextField label=""


                  disabled={true}
                  value={this.props.task.EFRPeriod} />
              </td>
            </tr>
            <tr>
              <td>
                <Label>Library:</Label>
              </td>
              <td>
                <TextField label=""


                  disabled={true}
                  value={this.props.task.EFRLibrary} />
              </td>
            </tr>
            <tr>
              <td>
                <Label>Assigned To:</Label>
              </td>
              <td>
                <TextField label=""


                  disabled={true}
                  value={this.getAssignees(this.props.task.EFRAssignedTo)} />
              </td>
            </tr>
            <tr>
              <td>
                <Label>Information Requested:</Label>
              </td>
              <td>
                <div className={styles.informationRequested}
                  dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRInformationRequested)} />
              </td>
            </tr>
          </table>
        </div >

        <Label className={styles.uploadInstructions} >Please upload the files containing the information requested on or before the Due Date above.
             You can drag and drop file(s) into the area shaded in blue below, or click the
             'Choose File' button to select the file(s). Uploaded files  will be automatically prefixed
             with the reference {this.props.task.Title}.  </Label>
        <Dropzone className={styles.dropzone} onDrop={this.onDrop.bind(this)} disableClick={true} >
          <div>
            Drag and drop files here to upload, or click Choose File below. Click on a file to view it.
          </div>
          {/* <div style={{ float: "left" }}> */}
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
                  key: "title", name: "File Name",
                  fieldName: "title", minWidth: 1, maxWidth: 500,
                  onRender: this.renderItemTitle.bind(this)
                },

              ]}
            />
            <input type="file" id="uploadfile" onChange={e => { this.uploadFile(e); }} />
          {/* </div>
          <div style={{ float: "right" }}>
            <DocumentIframe src={this.state.documentCalloutIframeUrl}
              height={this.props.documentIframeHeight}
              width={this.props.documentIframeWidth} />
          </div>
          <div style={{ clear: "both" }}></div> */}

        </Dropzone>

      </div>
    );
  }
}
