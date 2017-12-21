import * as React from 'react';
import styles from './EfrAdmin.module.scss';
import { IEfrAdminProps } from './IEfrAdminProps';
import { IEfrAdminState } from './IEfrAdminState';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, ButtonType } from "office-ui-fabric-react/lib/Button";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";

import { TextField } from "office-ui-fabric-react/lib/TextField";
import pnp, { WebAddResult, Web, RoleDefinitionBindings, List, ListAddResult, TypedHash, ViewAddResult } from "sp-pnp-js";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { Label } from "office-ui-fabric-react/lib/Label";
import { RoleDefinitions } from 'sp-pnp-js/lib/sharepoint/roles';
import { find, map } from "lodash";
// use jsom to add webpart to editform
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
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
  public async AddWebPartToEditForm(webRelativeUrl: string, editformUrl,webPartXml:string) {
    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var oFile = clientContext.get_web().getFileByServerRelativeUrl(editformUrl);
    
    var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
    let webparts = limitedWebPartManager.get_webParts();
    clientContext.load(webparts, 'Include(WebPart)');
    clientContext.load(limitedWebPartManager);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
    debugger;
    let cnt = webparts.get_count();
    let originalWebPartDef = webparts.get_item(0);
    let originalWebPart = originalWebPartDef.get_webPart();
    originalWebPart.set_hidden(true);
    originalWebPartDef.saveWebPartChanges();
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the webpart was hidden");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
    debugger;
      let oWebPartDefinition = limitedWebPartManager.importWebPart(webPartXml);
      let oWebPart = oWebPartDefinition.get_webPart();
   
      limitedWebPartManager.addWebPart(oWebPart, 'Main', 1);
   
      clientContext.load(oWebPart);
   
      await new Promise((resolve, reject) => {
        clientContext.executeQueryAsync((x) => {
          console.log("the new webpart was added");
          resolve();
        }, (error) => {
          console.log(error);
          reject();
        });
      });
  }
  

  public async createSite(): Promise<any> {

    let newWeb: Web;  // the web that gets created
    let libraryList: Array<any>; // the list of libraries we need to create in the new site. has the library name and the name of the group that should get access
    let roleDefinitions: Array<any>;// the roledefs for the site, we need to grant 'contribute no delete'
    let siteGroups: Array<any>;// all the sitegroups in the site
    let tasks: Array<any>; // the list of tasks in the TaskMaster list. We need to create on e task for each of these in tye EFRTasks list in the new site
    let taskList: List; // the task list we created  in the new site
    let contentTypes: Array<any>; /// the content types in the site. We need to add the pnctask content type to the taskList
    let webServerRelativeUrl: string;
    let editformurl: string;
    let editform: any;
    this.addMessage("CreatingSite");

    // create the site
    await pnp.sp.web.webs.add(this.state.siteName, this.state.siteName, this.state.siteName, "STS#0").then((war: WebAddResult) => {
      this.addMessage("CreatedSite");

      // show the response from the server when adding the web
      webServerRelativeUrl = war.data.ServerRelativeUrl;
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
    // now get  the list of libraries we need to create on the new site
    await pnp.sp.web.lists.getByTitle("EFRLibraries").items.top(2).get().then((libraries) => {
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
    // get the role definitions
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
    // get the site Groups
    await pnp.sp.web.siteGroups.get().then((sg) => {
      this.addMessage("got Site Groups");
      siteGroups = sg;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error getting site groups</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    // create the libraries and assign permissions
    for (const library of libraryList) {
      if (!library["EFRsecurityGroup"]) {
        this.addMessage("bypassing Library " + library["Title"] + "because it has no security group");
      } else {
        this.addMessage("Creating library " + library["Title"]);
        await newWeb.lists.add(library["Title"], library["Title"], 101, false).then(async (listResponse) => {
          this.addMessage("Created Library " + library["Title"]);
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
      }

    }
    // get the master list of tasks
    await pnp.sp.web.lists.getByTitle("PBCMaster").items.expand("EFRLibrary").select("*,EFRLibrary/Title")
      .top(2).get().then((efrtasks) => {
        this.addMessage("got tsakMaster list");
        tasks = efrtasks;
        return;
      }).catch(error => {
        debugger;
        this.addMessage("<h1>error fetching taskmaster</h1>");
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        console.error(error);
        return;
      });
    //  create the task list in the site
    this.addMessage("Creating taskList ");

    await newWeb.lists.add("EFRTasks", "EFRTasks", 100, true).then(async (listResponse) => {
      this.addMessage("Created List EFRTasks ");

      taskList = listResponse.list;

      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error creating tasklist</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    debugger;
    // set the custom form
    // await taskList.forms.get().then(async (forms)=>{
    //   debugger;
    //   find(forms,(f:any)=>{return f.FormType === 6})["ServerRelativeUrl"]=webServerRelativeUrl+"/SiteAssets/testForm.aspx";
    //   debugger;
    //   await taskList.update({Forms:forms}).then(async (f)=>{
    //     debugger;

    //     this.addMessage("updatedf forms ");
    //     return;
    //   }).catch(error => {
    //     debugger;
    //     this.addMessage("<h1>error updaing forms</h1>");
    //     this.addMessage(error.data.responseBody["odata.error"].message.value);
    //     console.error(error);
    //     return;
    //   });;
    // }).catch(error => {
    //   debugger;
    //   this.addMessage("<h1>error fetching forms</h1>");
    //   this.addMessage(error.data.responseBody["odata.error"].message.value);
    //   console.error(error);
    //   return;
    // });
    await taskList.forms.get().then(async (forms) => {
      debugger;
      editformurl = find(forms, (f: any) => { return f.FormType === 6 })["ServerRelativeUrl"];
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error fetching forms</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    await this.AddWebPartToEditForm(webServerRelativeUrl, editformurl,this.props.webPartXml);
    //add the PBC Task content type
    await taskList.contentTypes.addAvailableContentType("0x0100F2A5ABE2D8166E4E9A3C888E1DB4DC8B").then(ct => {
      this.addMessage("Added EFR Task content type");
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error adding content type to task list</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    //add the default view to show only open items assigned to me sorted bt date descening

    await taskList.views.add("My Open Tasks", false, {
      RowLimit: 10,
      ViewQuery: '<OrderBy><FieldRef Name="EFRDueDate" Ascending="FALSE" /></OrderBy><Where><And><Eq><FieldRef Name="EFRAssignedTo" /><Value Type="Integer"><UserID Type="Integer" /></Value></Eq><Eq><FieldRef Name="EFRVerifiedByAdmin" /><Value Type="Text">No</Value></Eq></And></Where>'
    }).then(async (v: ViewAddResult) => {

      // set this as the homePage
      let homepage = v.data.ServerRelativeUrl.substr(webServerRelativeUrl.length + 1);
      this.context
      await newWeb.rootFolder.update({ "WelcomePage": homepage }).then((x) => {
        this.addMessage("Set Site homepage to this view");
      }).catch(error => {
        debugger;
        this.addMessage("<h1>error setting site home page</h1>");
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        console.error(error);
        return;
      });
      // manipulate the view's fields
      await v.view.fields.removeAll().then(async _ => {

        await Promise.all([
          v.view.fields.add("Title"),
          v.view.fields.add("EFRInformationRequested"),
          v.view.fields.add("EFRDueDate"),
          v.view.fields.add("EFRAssignedTo"),
          v.view.fields.add("EFRCompletedByUser"),
          v.view.fields.add("EFRVerifiedByAdmin"),
        ]).then(_ => {

          this.addMessage("Added My Tasks View");
          return;
        });
        return;
      });
      return;

    });
    //add the a view to show alln items assigned to me sorted bt date descening
    //add the default view to show only open items assigned to me sorted bt date descening
    await taskList.views.add("My Tasks", false, {
      RowLimit: 10,
      DefaultView: true,
      ViewQuery: '<OrderBy><FieldRef Name="EFRDueDate" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="EFRAssignedTo" /><Value Type="Integer"><UserID Type="Integer" /></Value></Eq></Where>'
    }).then(async (v: ViewAddResult) => {

      // manipulate the view's fields
      await v.view.fields.removeAll().then(async _ => {

        await Promise.all([

          v.view.fields.add("Title"),
          v.view.fields.add("EFRInformationRequested"),
          v.view.fields.add("EFRDueDate"),
          v.view.fields.add("EFRAssignedTo"),
          v.view.fields.add("EFRCompletedByUser"),
          v.view.fields.add("EFRVerifiedByAdmin"),
        ]).then(_ => {

          this.addMessage("Added My Tasks View");
          return;
        });
        return;
      });
      return
    });


    // remove the item  and folder  Content types
    // await taskList.rootFolder.contentTypeOrder.get().then(async (ctypes) => {
    //   debugger;

    //   ctypes.push(ctypes.shift(0));
    //   const stuff:TypedHash= {
    //     "ContentTypeOrder":
    //       {
    //         __metadata:
    //           {
    //             'type': 'Collection(SP.ContentTypeId)'
    //           },
    //         results: ctypes
    //       }
    //   };
    //   await taskList.rootFolder.update({ "ContentTypeOrder": stuff })
    //     .then((x) => {
    //       this.addMessage("Holy crap!");
    //       return;
    //     }).catch((error) => {
    //       this.addMessage("NFG");
    //       this.addMessage(error.data.responseBody["odata.error"].message.value);
    //       return;
    //     });
    //   ;

    //   return;
    // });
    // // remove the item  and folder  Content types
    // await taskList.contentTypes.get().then(async (ctypes) => {
    //   debugger;
    //   for (const ctype of ctypes) {
    //     debugger;
    //     if (ctype.Id !== "0x0100F2A5ABE2D8166E4E9A3C888E1DB4DC8B") {
    //       await taskList.contentTypes.getById(ctype.Id).delete().then(x => {
    //         debugger;
    //         this.addMessage("deleted contenttype " + ctype.Name)
    //         return;
    //       }).catch(error => {
    //         debugger;
    //         this.addMessage("<h1>error deleting  content type" + ctype.Name + " from task list</h1>");
    //         this.addMessage(error.data.responseBody["odata.error"].message.value);
    //         console.error(error);
    //         return;
    //       });;
    //     }
    //   }
    //   return;
    // });

    // create the tasks in the new task list

    for (const task of tasks) {

      let itemToAdd = {
        "ContentTypeId": "0x0100F2A5ABE2D8166E4E9A3C888E1DB4DC8B",
        "Title": task.Title,
        "EFRDueDate": task.DueDate,
        "EFRAssignedToId": {
          "results": task.EFRAssignedToId
        },
        "EFRInformationRequested": task.InformationRequested,
        "EFRLibraryId": task.EFRLibraryId,
        "EFRPeriod": task.Period,
        "EFRCompletedByUser": "No",
        "EFRVerifiedByAdmin": "No"
      }
      await taskList.items.add(itemToAdd).then((results) => {
        this.addMessage("added task" + task.Title);
        return;
      }).catch(error => {
        debugger;
        this.addMessage("<h1>error adding task library" + task["Title"] + "</h1>");
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        console.error(error);
        return;
      });
    }
    this.addMessage("DONE!!");
  }
  private displayMessages(): any {
    const messages = map(this.state.messages, (m) => {
      return "<div>" + m + "</div>";
    });
    return { __html: messages.join('') };
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
