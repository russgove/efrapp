import * as React from 'react';
export interface IDocumentIframeProps {
    src: string;
    height:number;
    width: number;
}
export class DocumentIframe extends React.Component<IDocumentIframeProps, {}>{

    public render() {
        const heightAttr:string=this.props.height+"px";
        const widthAttr:string=this.props.width+"px";
            console.log("iframe source set to " + this.props.src);
        return (<iframe src={this.props.src} height={heightAttr} width={widthAttr}/>);
    }
}