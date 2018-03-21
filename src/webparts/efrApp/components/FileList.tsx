import * as React from 'react';
import styles from './EfrApp.module.scss';
import { DocumentIframe } from './DocumentIframe';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Document } from "../model";
import Dropzone from 'react-dropzone';
export interface IFileListProps {
  uploadFile: (file: any, Library: string, filePrefix: string) => Promise<any>;
  getDocuments: (library: string) => Promise<Array<Document>>;
  fetchDocumentWopiFrameURL: (id: number, mode: number, library: string) => Promise<string>;
  EFRLibrary: string;
  TaskTitle: string;
  documents: Array<Document>;
  documentIframeHeight: number;
  documentIframeWidth: number;
  enableUpload: boolean;
  dropZoneText:string;
}
export interface IFileListState {
  documents: Array<Document>;
  documentCalloutIframeUrl?: string;
}
export default class FileList extends React.Component<IFileListProps, IFileListState> {
  private validBrandIcons = " accdb csv docx dotx mpp mpt odp ods odt one onepkg onetoc potx ppsx pptx pub vsdx vssx vstx xls xlsx xltx xsn ";

  public constructor(props: IFileListProps) {
    super();
    console.log("in Construrctor");
    this.state = {
      documents: props.documents
    };
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
      promises.push(this.props.uploadFile(file, this.props.EFRLibrary, this.props.TaskTitle));
    });
    Promise.all(promises).then((x) => {
      this.props.getDocuments(this.props.EFRLibrary).then((dox) => {
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
    console.log("uplopading file");
    this.props.uploadFile(file, this.props.EFRLibrary, this.props.TaskTitle).then((response) => {
      console.log("getting documents");
      this.props.getDocuments(this.props.EFRLibrary).then((dox) => {
        console.log("got documents"+dox.length);
        this.setState((current) => ({ ...current, documents: dox }));
      });
    }).catch((error) => {
      console.error("an error occurred uploading the file");
      console.error(error);
    });
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
   

    // mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 0, this.props.EFRLibrary).then(url => {
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
   * called when a user mouses out of a document row. Sets the url to null in state so th eiframe no longer
   * shows the documentt
   * 
   * @param {Document} item 
   * @param {*} e 
   * @memberof EfrApp
   */
  public documentRowMouseOut(item: Document, e: any) {

    this.setState((current) => ({
      ...current,
      documentCalloutTarget: null,
      documentCalloutVisible: false
    }));

  }
  public openDocument(document: Document): void {

    // mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(document.id, 0, this.props.EFRLibrary).then(url => {
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
public componentWillReceiveProps(nextProps: IFileListProps){
  this.setState((current)=>({...current, documents:nextProps.documents}));
}
  public render(): React.ReactElement<IFileListProps> {
   

    if (this.props.enableUpload) {
      try {
        return (
          <Dropzone className={styles.dropzone} onDrop={this.onDrop.bind(this)} disableClick={true} >
            <div>
             {this.props.dropZoneText}
          </div>
            <div style={{ float: "left", width: "310px" }}>
              <DetailsList key="files"
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

            {/* <input type="file" id="uploadfile" onChange={e => { this.uploadFile(e); }} /> */}
          </Dropzone>


        );
      } catch (error) {
        console.error("An error occurred rendering FileList.");
        console.error(error);
        return (<div>An error occurred rendering the EFR application</div>);
      }
    }
    else {
      try {
        return (
          <div>
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
                    key: "title", name: "Role Name",
                    fieldName: "Role_x0020_Name", minWidth: 1, maxWidth: 200,
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

          </div>


        );
      } catch (error) {
        console.error("An error occurred rendering FileList.");
        console.error(error);
        return (<div>An error occurred rendering the EFR application</div>);
      }
    }
  }
}
