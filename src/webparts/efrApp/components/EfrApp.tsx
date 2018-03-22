import * as React from 'react';
import styles from './EfrApp.module.scss';
import { TextFieldWithEdit } from './TextFieldWithEdit';
import { IEfrAppProps } from './IEfrAppProps';
import { IEfrAppState } from './IEfrAppState';
import { CompoundButton, } from "office-ui-fabric-react/lib/Button";
//import { TextField } from "office-ui-fabric-react/lib/TextField";
import { TagPicker, ITag } from "office-ui-fabric-react/lib/Pickers";
import { Label } from "office-ui-fabric-react/lib/Label";
import { PBCTask,HelpLink } from "../model";
import FileList from "./FileList";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import {map} from "lodash";
export default class EfrApp extends React.Component<IEfrAppProps, IEfrAppState> {
  public constructor(props: IEfrAppProps) {
    super();

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

    var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    let year = parseInt(dateString.substr(0, 4));
    let month = parseInt(dateString.substr(5, 2)) - 1;
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
      alert("An error occurred reopining this task.");
      console.error(err);
    });
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
  public showHelp(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    debugger;
    window.open(item.data.url,item.data.name,item.data.specs);
  }
  /**
   * renders the page
   * 
   * @returns {React.ReactElement<IEfrAppProps>} 
   * @memberof EfrApp
   */
  public render(): React.ReactElement<IEfrAppProps> {

    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "UPLOAed",
        name: "Upload Files",
        icon: "Upload",
        title: this.props.uploadFilesHoverText,
        disabled: (this.props.task.EFRCompletedByUser === "Yes" || !this.userIsAssignedTask(this.props.task)),
        onClick: (e) => {
          debugger;
          var input: any = document.getElementById("TronoxEFRUploadfile");
          input.click();

        }

        // the below only works in chrome
        // onClick: (e) => {
        //   debugger;
        //   var input: HTMLInputElement = document.createElement("input");
        //   input.type = "file";
        //   // add onchange handler if you wish to get the file :)
        //   console.log(input);

        //   try {
        //     input.click(); // opening dialog
        //     input.onchange = (element) => {
        //       let file: any = element.target["files"][0];
        //       console.log("uplopading file");
        //       this.props.uploadFile(file, this.props.task.EFRLibrary, this.props.task.Title).then((response) => {
        //         console.log("getting documents");
        //         this.props.getDocuments(this.props.task.EFRLibrary).then((dox) => {
        //           console.log("got documents " + dox.length);

        //           this.setState((current) => ({ ...current, documents: dox }));
        //         });
        //       }).catch((error) => {
        //         console.error("an error occurred uploading the file");
        //         console.error(error);
        //       });
        //     };
        //     return true; // avoiding navigation
        //   }
        //   catch (err) {
        //     debugger
        //   };

        // },
        // END OF the below only works in chrome

      }
    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Save", name: "Save", icon: "Save", onClick: this.closeWindow.bind(this),
        title: this.props.saveHoverText,
        disabled: (this.props.task.EFRCompletedByUser === "Yes" || !this.userIsAssignedTask(this.props.task))

      },
      {
        key: "Task Complete", name: "Task Complete", icon: "Completed", onClick: this.completeTask.bind(this),
        title: this.props.taskCompleteHoverText,
        disabled: (this.props.task.EFRCompletedByUser === "Yes" || !this.userIsAssignedTask(this.props.task))
      },
      {
        key: "Reopen Task", name: "Reopen Task", icon: "Refresh", onClick: this.reopenTask.bind(this),

        disabled: (this.props.task.EFRCompletedByUser === "No" || !this.userIsAssignedTask(this.props.task)),
        title: this.props.reopenTaskHoverText
      },
      {
        key: "helpLinks", name: "Help", icon: "help",
        title: this.props.helpHoverText,
        items: map(this.props.helpLinks,(hl):IContextualMenuItem=>{
          debugger;
          return{
            key:hl.Id.toString(), // this is the id of the listitem
            href:hl.Url.Url, 
            title:hl.Url.Description,
            icon:hl.IconName,
            name:hl.Title,
            target:hl.Target

          }
        })
      }
    ];
    try {
      return (
        <div className={styles.efrApp}>
          <input type="file" id="TronoxEFRUploadfile" style={{ "display": "none" }} onChange={element => {
            let file: any = element.target["files"][0];
            console.log("uplopading file");
            this.props.uploadFile(file, this.props.task.EFRLibrary, this.props.task.Title).then((response) => {
              console.log("getting documents");
              this.props.getDocuments(this.props.task.EFRLibrary).then((dox) => {
                console.log("got documents " + dox.length);

                this.setState((current) => ({ ...current, documents: dox }));
              });
            });
          }}
          />
          <CommandBar
            isSearchBoxVisible={false}
            items={itemsNonFocusable}
            farItems={farItemsNonFocusable}

          />
          <div className={styles.headerArea}>
            <div style={{ float: "left", width: "70%" }}>
              <table>
                <tr>
                  <td>
                    <Label>Reference #:</Label>
                  </td>
                  <td>
                    {this.props.task.Title}
                  </td>
                </tr>
                <tr>
                  <td>
                    <Label>Information Requested:</Label>
                  </td>
                  <td>
                    <span className={styles.informationRequested} dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRInformationRequested)} />
                  </td>
                </tr>
                <tr>
                  <td>
                    <Label>Reporting Period:</Label>
                  </td>
                  <td>
                    <span className={styles.informationRequested} dangerouslySetInnerHTML={this.createSummaryMarkup(this.props.task.EFRPeriod)} />
                  </td>
                </tr>
                <tr>
                  <td>
                    <Label>Due Date:</Label>
                  </td>
                  <td>
                    {this.getDateString(this.props.task.EFRDueDate)}
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
            <div style={{ float: "left", width: "30%" }}>

              <table>

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
                    <Label>Library:</Label>
                  </td>
                  <td>
                    {this.props.task.EFRLibrary}
                  </td>
                </tr>
              </table>
            </div>
            <div style={{ clear: "both" }}></div>
          </div>
          <div dangerouslySetInnerHTML={(this.props.task.EFRCompletedByUser === "Yes") ? this.createSummaryMarkup(this.props.efrFormInstructionsClosed) : this.createSummaryMarkup(this.props.efrFormInstructionsOpen)}>
          </div>

          <FileList
            uploadFile={this.props.uploadFile}
            getDocuments={this.props.getDocuments}
            fetchDocumentWopiFrameURL={this.props.fetchDocumentWopiFrameURL}
            EFRLibrary={this.props.task.EFRLibrary}
            TaskTitle={this.props.task.Title}
            documents={this.state.documents}
            documentIframeHeight={this.props.documentIframeHeight}
            documentIframeWidth={this.props.documentIframeWidth}
            enableUpload={this.userIsAssignedTask(this.props.task) && this.props.task.EFRCompletedByUser === "No"}
            dropZoneText={this.props.dropZoneText}
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
