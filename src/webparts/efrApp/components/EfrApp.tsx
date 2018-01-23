import * as React from 'react';
import styles from './EfrApp.module.scss';
import { TextFieldWithEdit } from './TextFieldWithEdit';
import { IEfrAppProps } from './IEfrAppProps';
import { IEfrAppState } from './IEfrAppState';
import { CompoundButton, } from "office-ui-fabric-react/lib/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { TagPicker, ITag } from "office-ui-fabric-react/lib/Pickers";
import { Label } from "office-ui-fabric-react/lib/Label";
import { PBCTask } from "../model";
import FileList from "./FileList";
export default class EfrApp extends React.Component<IEfrAppProps, IEfrAppState> {
  public constructor(props: IEfrAppProps) {
    super();
    console.log("in Construrctor");
    this.CloseButton = this.CloseButton.bind(this);
    this.CompleteButton = this.CompleteButton.bind(this);
    this.state = {
      documentCalloutIframeUrl: "",
      documents: props.documents,
      taskComments: props.task.EFRComments  // User can only update the comments on this
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

  public getAssignees(assignees: Array<{}>): Array<ITag> {
    let result: ITag[] = [];
    for (let assignee of assignees) {
      result.push({ key: assignee["Title"], name: assignee["Title"] });
    }
    return result;
  }
  private completeTask(e): void {
    this.props.completeTask(this.props.task);
  }
  private reopenTask(e): void {
    this.props.reopenTask(this.props.task).then(() => {

    }).catch((err) => {
      debugger;
      alert("An error occurred reopining this task.")
      console.error(err);
    })
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
      return (<div>
        <CompoundButton onClick={this.reopenTask.bind(this)} description='Click here to Reopen  this task.'  >
          I have additional files I need to upload.
</CompoundButton>

      </div>);
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
  public commentsChanged(oldValue, newValue): Promise<any> {
    return this.props.updateTaskComments(this.props.task.Id, oldValue, newValue).then(() => {
      this.setState((current) => ({ ...current, taskComments: newValue }));
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
    console.log(this.props.ckEditorConfig);

    try {
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
          <FileList
            uploadFile={this.props.uploadFile}
            getDocuments={this.props.getDocuments}
            fetchDocumentWopiFrameURL={this.props.fetchDocumentWopiFrameURL}
            EFRLibrary={this.props.task.EFRLibrary}
            TaskTitle={this.props.task.Title}
            documents={this.props.documents}
            documentIframeHeight={this.props.documentIframeHeight}
            documentIframeWidth={this.props.documentIframeWidth}
            enableUpload={this.userIsAssignedTask(this.props.task) && this.props.task.EFRCompletedByUser === "No"}
          />
        </div>
      );
    } catch (error) {
      console.error("An error occurred renering EFrapp.");
      console.error(error);
      return (<div>An error occurred rendering the EFR application</div>);
    }
  }
}
