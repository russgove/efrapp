import * as React from 'react';
import styles from './EfrAdmin.module.scss';
import { IEfrAdminProps } from './IEfrAdminProps';
import { IEfrAdminState } from './IEfrAdminState';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, ButtonType } from "office-ui-fabric-react/lib/Button";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";

import { TextField } from "office-ui-fabric-react/lib/TextField";
import pnp, { WebAddResult, Web, RoleDefinitionBindings, List, ListAddResult } from "sp-pnp-js";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { Label } from "office-ui-fabric-react/lib/Label";
import { RoleDefinitions } from 'sp-pnp-js/lib/sharepoint/roles';
import { find, map } from "lodash";
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
    console.log(msg);
    this.setState((current: IEfrAdminState) => {
      let newState = current;
      newState.messages.push(msg);
      return newState;
    });

  }
  public async createSite(): Promise<any> {
    debugger;
    let newWeb: Web;
    let libraryList: Array<any>;
    let roleDefinitions: Array<any>;
    let siteGroups: Array<any>;
    this.addMessage("CreatingSite");
    debugger;
    await pnp.sp.web.webs.add(this.state.siteName, this.state.siteName, this.state.siteName, "STS#0").then((war: WebAddResult) => {
      this.addMessage("CreatedSite");
      // show the response from the server when adding the web
      console.log(war.data);
      newWeb = war.web;
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error creating site</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);

      console.error(error);
      return;
    });

    // now get  the list of libraries
    debugger;
    await pnp.sp.web.lists.getByTitle("EFRLibraries").items.get().then((libraries) => {
      this.addMessage("got list of libraries");
      libraryList = libraries;
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error fetching library list</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    debugger;
    await pnp.sp.web.roleDefinitions.get().then((roleDefs) => {
      this.addMessage("got roledefinitions");
      roleDefinitions = roleDefs;
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error fetching roledefs</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    debugger;
    await pnp.sp.web.siteGroups.get().then((sg) => {
      this.addMessage("got Site Groups");
      siteGroups = sg;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error getting site groups</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    })
    debugger;
    for (const library of libraryList) {
      this.addMessage("Creating library " + library["Title"]);
      debugger;
      await newWeb.lists.add(library["Title"], library["Title"], 101, false).then(async (listResponse) => {
        this.addMessage("Created Library " + library["Title"]);
        debugger;
        let list = listResponse.list;
        await list.breakRoleInheritance(true).then((e) => {
          this.addMessage("broke role inheritance on " + library["Title"]);
        }).catch(error => {
          debugger;
          this.addMessage("<h1>error breaking role inheritance on  library " + library["Title"] + "</h1>");
          this.addMessage(error.data.responseBody["odata.error"].message.value);
          console.error(error);
          return;
        });
        let group = find(siteGroups, (sg => { return sg["Title"] === library["EFRsecurityGroup"]; }));
        let principlaID = group["Id"];
        let roledef = find(roleDefinitions, (rd => { return rd["Name"] === "Content Authors without delete or modify"; }));
        let roleDefId = roledef["Id"];
        debugger;
        await list.roleAssignments.add(principlaID, roleDefId).then(xxx => {
          this.addMessage("granted " + library["EFRsecurityGroup"] + "access to " + library["Title"]);
        }).catch(error => {
          debugger;
          this.addMessage("<h1>error adding role asisigment to  library " + library["Title"] + "</h1>");
          this.addMessage(error.data.responseBody["odata.error"].message.value);
          console.error(error);
          return;
        });
      }).catch(error => {
        debugger;
        this.addMessage("<h1>error creating library" + library["Title"] + "</h1>");
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        console.error(error);
        return;
      });

      // break role inheraitanc and add the new group

    }
  }
  private displayMessages(): any {
    const messages = map(this.state.messages, (m) => {
      debugger;
      return "<div>" + m + "</div>";

    });
    return { __html: messages };
  }
  public render(): React.ReactElement<IEfrAdminProps> {

    return (
      <div className={styles.efrAdmin} >
        <TextField label="Site Name" onChanged={(e) => {
          this.setState((current) => ({ ...current, siteName: e }));
        }} />

        <PrimaryButton onClick={this.createSite.bind(this)} title="Create Site">Create Site</PrimaryButton>

        <div dangerouslySetInnerHTML={this.displayMessages()} />
      </div>
    );
  }
}
