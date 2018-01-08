import * as React from 'react';
import styles from './EfrLockAndHideSites.module.scss';
import { IEfrLockAndHideSitesProps } from './IEfrLockAndHideSitesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from "moment";
import {
  NormalPeoplePicker, TagPicker, ITag
} from "office-ui-fabric-react/lib/Pickers";

import { PrimaryButton, Button, ButtonType, ActionButton } from "office-ui-fabric-react/lib/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";

import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { Label } from "office-ui-fabric-react/lib/Label";

import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
// switch to fabric  ComboBox on next upgrade
// let Select = require("react-select") as any;
//import "react-select/dist/react-select.css";
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from "office-ui-fabric-react/lib/DetailsList";
import { DatePicker, } from "office-ui-fabric-react/lib/DatePicker";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";
import { SPComponentLoader } from "@microsoft/sp-loader";
export default class EfrLockAndHideSites extends React.Component<IEfrLockAndHideSitesProps, {}> {
  /**
 * Renders a formatted date in the UI
 * 
 * @param {*} [item]  The item the field resides in
 * @param {number} [index] The index of the item in the list of items
 * @param {IColumn} [column] The column that contains the date to dispplay
 * @returns {*} 
 * 
 * @memberof TrForm
 */
  public renderDate(item?: any, index?: number, column?: IColumn): any {

    return moment(item[column.fieldName]).format("MMM Do YYYY");
  }
  public renderBoolean(item?: any, index?: number, column?: IColumn): JSX.Element {
    return (<Checkbox checked={item[column.fieldName]} />)
  }
  public renderOnTopNav(item?: any, index?: number, column?: IColumn): JSX.Element {

    // let x =  this.props.topNav as Array<any>;
    // this.props.topNav.toArray;
    for (let navnode of this.props.topNav) {
      if (navnode.Title === item.title) {
        return (<Checkbox checked={true} />)
      }
    }
    return (<Checkbox checked={false} />)
  }


  public renderRemoveFromTopNav(item?: any, index?: number, column?: IColumn): JSX.Element {

    // let x =  this.props.topNav as Array<any>;
    // this.props.topNav.toArray;
    for (let navnode of this.props.topNav) {
      if (navnode.Title === item.title) {
        return (<ActionButton onClick={(e) => {
          // find the navitem for the selected web
          for (let ni of this.props.topNav) {
            if (ni.Title === item.title) {
              debugger;
              this.props.removeSiteFromTopNav(ni)
            }
          }
        }
        } >
          Remove from Navigation
        </ActionButton>)
      }
    }
    return (<div />)
  }
  public renderLockSite(item?: any, index?: number, column?: IColumn): JSX.Element {

    if (item.isLocked) {
      return (<div />);
    }
    else {
      return (<ActionButton onClick={(e) => {
        // find the navitem for the selected web
        this.props.lockSite(item);
      }}
      >

        Lock Site
        </ActionButton>)
    }
  }

  public render(): React.ReactElement<IEfrLockAndHideSitesProps> {

    let columns: IColumn[] = [
      { key: "Title", name: "Site", fieldName: "title", minWidth: 60, maxWidth: 60 },
      { key: "locked", name: "Locked", fieldName: "isLocked", minWidth: 60, maxWidth: 60, onRender: this.renderBoolean.bind(this) },
      { key: "onTopNav", name: "onTopNav", fieldName: "onTopNav", minWidth: 90, maxWidth: 60, onRender: this.renderOnTopNav.bind(this) },
      { key: "date", onRender: this.renderDate, name: "Last Modified", fieldName: "lastModifiedDate", minWidth: 120, maxWidth: 120 },
      { key: "Remove from Menu", name: "Remove from Menu", fieldName: "Title", minWidth: 60, maxWidth: 60, onRender: this.renderRemoveFromTopNav.bind(this) },
      { key: "lock site", name: "Lock Site", fieldName: "Title", minWidth: 60, maxWidth: 60, onRender: this.renderLockSite.bind(this) },
    ];
    return (
      <div className={styles.efrLockAndHideSites} >
        <DetailsList selectionMode={SelectionMode.none} items={this.props.efrWebs} columns={columns}>

        </DetailsList>
      </div >
    );
  }
}
