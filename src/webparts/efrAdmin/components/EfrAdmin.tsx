import * as React from 'react';
import styles from './EfrAdmin.module.scss';
import { IEfrAdminProps } from './IEfrAdminProps';
import { IEfrAdminState } from './IEfrAdminState';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, ButtonType } from "office-ui-fabric-react/lib/Button";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";
import { List } from "office-ui-fabric-react/lib/List";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import pnp, { WebAddResult } from "sp-pnp-js";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { Label } from "office-ui-fabric-react/lib/Label";
import { RoleDefinitions } from 'sp-pnp-js/lib/sharepoint/roles';
import {find}  from "lodash";
export default class EfrAdmin extends React.Component<IEfrAdminProps, IEfrAdminState> {
  public constructor(props) {
    super();
    console.log("in Construrctor");

    this.state = {
      messages: ["Enter the site name and click the create site button"],
      siteName: ""
    };
  }
  public addMessage(msg): void {
    console.log("in addMessage there are " + this.state.messages.length + " messages");

    this.setState((current: IEfrAdminState) => {
      let newState = current;
      console.log("in addMessage callback  there are " + current.messages.length + " messages");
      newState.messages.push(msg);
      return newState;

    })

  }
  public createSite(): Promise<any> {
    debugger;
    return pnp.sp.web.webs.add(this.state.siteName, this.state.siteName, this.state.siteName, "STS#0").then((war: WebAddResult) => {
      this.addMessage("CreatingSite");

      // show the response from the server when adding the web
      console.log(war.data);

      this.addMessage("Site created, adding lists");

      let newweb = war.web;
      console.log("got web");
      // now add the lists
      return pnp.sp.web.lists.getByTitle("EFRLibraries").items.get().then((libraries) => {
        debugger;
        this.addMessage("got list of libraries");
        return pnp.sp.web.roleDefinitions.get().then((RoleDefinitions) => {
          this.addMessage("got roledefinitions");
          debugger;
          return pnp.sp.web.siteGroups.get().then((siteGroups)=>{
            this.addMessage("got Site Groups");
            let listPromises: Array<Promise<any>> = [];
            for (const library of libraries) {
              listPromises.push(newweb.lists.add(library["Title"], library["Title"], 101, false).then(listResponse => {
                this.addMessage("Created Library " + library["Title"])
                debugger;
                let list = listResponse.list;
                return list.breakRoleInheritance(true).then((e) => {
                  this.addMessage("broke role inheritance on "+library["Title"]);
                  debugger;
                  let group=find(siteGroups,(sg=>{return sg["Title"]===library["EFRsecurityGroup"]}));
                  let principlaID=group["Id"];
                  let roledef=find(RoleDefinitions,(rd=>{return rd["Name"]==="Content Authors without delete or modify"}));
                  let roleDefId=roledef["Id"];
                  return list.roleAssignments.add(principlaID,roleDefId).then(xxx=>{
                    this.addMessage("granted "+library["EFRsecurityGroup"] + "access to " +library["Title"])
                  });
                });
  
                // break role inheraitanc and add the new group
  
  
              }));
  
            }
            return Promise.all(listPromises);
          });
        
        });

      })
    }).catch(error => {
      debugger;
    });
  }
  private _onRenderCell(item: any, index: number, isScrolling: boolean): JSX.Element {
    return (
      <div className='ms-ListGhostingExample-itemName'>{item}</div>

    );
  }

  public render(): React.ReactElement<IEfrAdminProps> {

    console.log("in render there are " + this.state.messages.length + " messages");
    return (
      <div className={styles.efrAdmin} >
        <TextField label="Site Name" onChanged={(e) => {
          this.setState((current) => ({ ...current, siteName: e }));
        }} />

        <PrimaryButton onClick={this.createSite.bind(this)} title="Create Site">Create Site</PrimaryButton>

        <div>{this.state.messages.join("; ")}</div>
      </div>
    );
  }
}
