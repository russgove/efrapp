import * as React from 'react';
import { SPComponentLoader } from "@microsoft/sp-loader";
export interface IRichTextEditorProps {
    value: string;
    ckEditorUrl:string;
    ckEditorConfig:any;
}
export interface IRichTextEditorState {
    text: string;
}
export class RichTextEditor extends React.Component<IRichTextEditorProps, IRichTextEditorState>{
    private ckeditor;
    public componentDidMount() {
        debugger;
        // see https://github.com/SharePoint/sp-de//cdn.ckeditor.com/4.6.2/full/ckeditor.jsv-docs/issues/374
        // var ckEditorCdn: string = "//cdn.ckeditor.com/4.6.2/full/ckeditor.js";
        var ckEditorCdn: string = this.props.ckEditorUrl;
        SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: "CKEDITOR" }).then((CKEDITOR: any): void => {
          this.ckeditor = CKEDITOR;
          // replaces the title with a ckeditor. the other textareas are not visible yet. They will be replaced when the tab becomes active
          this.ckeditor.replace("someuniqueFieldName", this.props.ckEditorConfig);
    
        });
    
      }
    constructor(props: IRichTextEditorProps) {
        super(props);
        // set initial state
        this.state = { text: props.value };
    }
    public getValue(){
        debugger;
        let instance=this.ckeditor.instances["someuniqueFieldName"];
        let data = instance.getData();
        return data;

    }
    public render() {
        debugger;
        return (
            <div>
                <textarea name="someuniqueFieldName" id="someuniqueFieldName" style={{ display: "none" }}>
                    {this.state.text}
                </textarea>
            </div>

        );
    }
}