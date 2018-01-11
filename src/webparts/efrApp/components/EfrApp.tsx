import * as React from 'react';
import styles from './EfrApp.module.scss';
import { DocumentIframe } from './DocumentIframe';
import { TextFieldWithEdit } from './TextFieldWithEdit';
import { RichTextEditor } from './RichTextEditor';
import { IEfrAppProps } from './IEfrAppProps';
import { IEfrAppState } from './IEfrAppState';
import { escape,cloneDeep } from '@microsoft/sp-lodash-subset';
import { PageContext } from "@microsoft/sp-page-context";
import { PrimaryButton, ButtonType, CompoundButton, } from "office-ui-fabric-react/lib/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import {
  NormalPeoplePicker, TagPicker, ITag
} from "office-ui-fabric-react/lib/Pickers";

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
    this.CloseButton = this.CloseButton.bind(this);
    this.CompleteButton = this.CompleteButton.bind(this);
    this.state = {
      documentCalloutIframeUrl: "",
      documents: props.documents,
      taskComments:props.task.EFRComments  // User can only update the comments on this
    };
  }
  /**
   * comverts html to an onject for use in dangerouslySetInnerHTML
   * 
   * @param {any} html 
   * @returns 
   * @memberof EfrApp
   */
  public createSummaryMarkup(html) {
    console.log("in createSummaryMarkup");
    return { __html: html };
  }
  /**
   * Called when a user drops files into the DropZone. It calls 
   * the uploadFile method on the props to upload the files to sharepoint and then adds them to state.
   * 
   * @private
   * @param {any} acceptedFiles 
   * @param {any} rejectedFiles 
   * @memberof EfrApp
   */
  private onDrop(acceptedFiles, rejectedFiles) {
    console.log("in onDrop");
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
  /**
   * This method is called when the user uploads sa file using the Add file button. It calls 
   * the uploadFile method on the props to upload the files to sharepoint and then adds them to state.
   * 
   * @param {*} e 
   * @memberof EfrApp
   */
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
  /**
   * Gets a date, formatted as a string for display. It uses the ussers Culture so that 
   * dates are formatted properly for Australian users/
   * 
   * @param {string} dateString 
   * @returns {string} 
   * @memberof EfrApp
   */
  public getDateString(dateString: string): string {
    console.log("in getDateString");

    var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    let year = parseInt(dateString.substr(0, 4));
    let month = parseInt(dateString.substr(5, 2));
    let day = parseInt(dateString.substr(8, 2));
    return new Date(year, month, day)
      .toLocaleDateString(this.props.cultureName, options);

  }
  /**
   * This is called when the user hovers over a document in the list. It callse the fetchDocumentWopiFrameURL
   * in the props to het th url, and then sets the url in state toi have the iframe display the document.
   * 
   * @param {Document} document 
   * @param {*} e 
   * @memberof EfrApp
   */
  public documentRowMouseEnter(document: Document, e: any) {
    console.log("in documentRowMouseEnter");
  
    // mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 0, this.props.task.EFRLibrary).then(url => {
      // if (!url || url === "") {  // is this causing the download when i hove over a non office doc?
      //   url = document.serverRalativeUrl;
      // }
      this.setState((current) => ({
        ...current,
        documentCalloutIframeUrl: url,
        documentCalloutTarget: e.target,
        documentCalloutVisible: true
      }));

    });
  }
  /**
   * Determines if a task is assigned to the current user
   * 
   * @param {PBCTask} task 
   * @returns {boolean} 
   * @memberof EfrApp
   */
  public userIsAssignedTask(task: PBCTask): boolean {
  
    for (var assignee of task.EFRAssignedTo) {
      if (assignee.UserName.toUpperCase() === this.props.currentUserLoginName.toUpperCase()) {
        return true;
      }
    }
    return false;
  }
  /**
   * called when a user mouses out of a document row. Sets the url to null in state so th eiframe no longer
   * shows the documentt
   * 
   * @param {Document} item 
   * @param {*} e 
   * @memberof EfrApp
   */
  public documentRowMouseOut(item: Document, e: any) {
    console.log("in documentRowMouseOut");
    this.setState((current) => ({
      ...current,
      documentCalloutTarget: null,
      documentCalloutVisible: false
    }));
    console.log("mouse exit for " + item.title);
  }
  public openDocument(document: Document): void {
    console.log("in editDocument");
    // mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 0, this.props.task.EFRLibrary).then(url => {
      console.log("wopi frame url is " + url);
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
  public getAssignees(assignees: Array<{}>): Array<ITag> {
    let result: ITag[] = [];
    for (let assignee of assignees) {
      result.push({ key: assignee["Title"], name: assignee["Title"] });
    }
    return result;
  }
  public renderItemTitle(item?: any, index?: number, column?: IColumn): any {
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
          onClickCapture={(e) => {

            e.preventDefault();
            this.openDocument(item); return false;
          }}><span className={styles.documentTitle} > {item.title}</span></a>
      </div>);
  }
  private completeTask(e): void {
    this.props.completeTask(this.props.task);
  }
  private closeWindow(): void {
    this.props.closeWindow();
  }
  /**
   * This button  marks the task as complete and closes the window, retruning to whatever
   * the url specified in the source  query string parameters. 
   * An email  is sent to the users of a group specified in the proertypane to notify them that the task is
   * complete.
   * 
   * @private
   * @returns {JSX.Element} 
   * @memberof EfrApp
   */
  private CompleteButton(): JSX.Element {
  
    if (!this.userIsAssignedTask(this.props.task)) { // dont show the button if its not my task
      return (<div />);
    }
    if (this.props.task.EFRCompletedByUser === "Yes") { // dont show the button if task is complete
      return (<div />);
    }
    return (
      <div>
        <CompoundButton onClick={this.completeTask.bind(this)} description='Click here to close this window and  mark this task as complete.'  >
          I have uploaded all the information requested
  </CompoundButton>

      </div>
    );
  }
  /**
   *  closes the window, retuning to whatever
   * the url specified in the source  query string parameters. 
   * 
   * @private
   * @returns {JSX.Element} 
   * @memberof EfrApp
   */
  private CloseButton(): JSX.Element {
    

    return (
      <div>
        <CompoundButton onClick={this.closeWindow.bind(this)} description='Click here to close this window.'  >
          Close Window
  </CompoundButton>

      </div>
    );
  }
  public commentsChanged(oldValue, newValue):Promise<any>{

    return this.props.updateTaskComments(this.props.task.Id,oldValue,newValue).then(()=>{
      debugger;
      this.setState((current)=>({...current, taskComments:newValue}));
    });
  }
  
  /**
   * renders the page
   * 
   * @returns {React.ReactElement<IEfrAppProps>} 
   * @memberof EfrApp
   */
  public render(): React.ReactElement<IEfrAppProps> {
    console.log("ckEditorConfig follows");

    console.log(this.props.ckEditorConfig)

   debugger;
   try{
    return (
      <div className={styles.efrApp}>
        
        <div className={styles.headerArea}>
          <div style={{ float: "left", width: "50%" }}>
            <table>
              <tr>
                <td>
                  <Label>Reference #:</Label>
                </td>
                <td>
                  <TextField label=""

                    disabled={true}
                    value={this.props.task.Title} />
                </td>
              </tr>

              <tr>
                <td>
                  <Label>Assigned To:</Label>
                </td>
                <td>
                  <TagPicker
                    disabled={true}
                    onResolveSuggestions={null}
                    defaultSelectedItems={this.getAssignees(this.props.task.EFRAssignedTo)}
                  />
              </td>
              </tr>
              <tr>
                <td>
                  <Label>Comments:</Label>
                </td>
                <td>
                  <TextFieldWithEdit
                    value={this.state.taskComments}
                    onValueChanged={this.commentsChanged.bind(this)}
                    ckEditorUrl={this.props.ckEditorUrl}
                    ckEditorConfig={JSON.parse(this.props.ckEditorConfig)}
                  />
             
                </td>
              </tr>


            </table>
          </div >
          <div style={{ float: "left", width: "50%" }}>
            <this.CompleteButton />
            <this.CloseButton />
          </div>
          <div style={{ clear: "both" }}></div>
        </div>

        <div>
          Please upload the files containing :
          <div className={styles.informationRequested} dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRInformationRequested)} />

          for the period:
          <div className={styles.informationRequested} dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRPeriod)} />


          to the {this.props.task.EFRLibrary} library on or before {this.getDateString(this.props.task.EFRDueDate)} .<br />
          You can drag and drop file(s) into the area shaded in blue below, or click the
             'Choose File' button to select the file(s). Files uploaded this way will be automatically prefixed
             with the Reference {this.props.task.Title}.
       </div>
        <Label className={styles.uploadInstructions} >
          Alternatively, you can navigate to the {this.props.task.EFRLibrary} using the navigation bar to the left
       and upload the files to the library using the SharePoint upload function. Note that if you upload this way
       the files will not be automatically prefixed with {this.props.task.Title}  YOU must prefix the files with  {this.props.task.Title} before
       uploading OR ELSE!!
       </Label>
        <Dropzone className={styles.dropzone} onDrop={this.onDrop.bind(this)} disableClick={true} >
          <div>
            Drag and drop files here to upload, or click Choose File below. Click on a file to view it.
          </div>
          <div style={{ float: "left", width: "310px" }}>
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
                  fieldName: "title", minWidth: 1, maxWidth: 200,
                  onRender: this.renderItemTitle.bind(this)
                },
              ]}
            />
          </div>
          <div style={{ float: "right" }}>
            <DocumentIframe src={this.state.documentCalloutIframeUrl}
              height={this.props.documentIframeHeight}
              width={this.props.documentIframeWidth} />
          </div>
          <div style={{ clear: "both" }}></div>

          <input type="file" id="uploadfile" onChange={e => { this.uploadFile(e); }} />
        </Dropzone>

      </div>
    );
  }catch(error){
    console.error("An error occurred renering EFrapp.");
    console.error(error);
    return (<div>An error occurred rendering the EFR application</div>);
  }
  }
}
