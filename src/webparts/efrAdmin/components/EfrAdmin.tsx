import * as React from 'react';
import styles from './EfrAdmin.module.scss';
import { IEfrAdminProps } from './IEfrAdminProps';
import { IEfrAdminState } from './IEfrAdminState';


import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";


//import { load, exec, toArray } from "../../JsomHelpers"
import { TextField } from "office-ui-fabric-react/lib/TextField";
import {
  sp,
  WebAddResult, Web,
  ContextInfo, List, ViewAddResult
} from "@pnp/sp";



import { find, map } from "lodash";

// use jsom to add webpart to editform
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('sp-workflow');

export default class EfrAdmin extends React.Component<IEfrAdminProps, IEfrAdminState> {

  public constructor(props: IEfrAdminProps) {
    super();
    console.log("in Construrctor");

    this.state = {
      messages: ["Enter the site name and click the create site button"],
      siteName: "",
      pbcMasterList: props.PBCMasterLists[0].text

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
  /**
  *  Adds a custom webpart to the edit form located at editformUrl
  * 
  * @param {string} webRelativeUrl -- The web containing the list
  * @param {any} editformUrl -- the url of the editform page
  * @param {string} webPartXml  -- the xml for the webpart to add
  * @memberof EfrAdmin
  */
  public async SetWebToUseSharedNavigation(webRelativeUrl: string) {

    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var currentWeb = clientContext.get_web();
    var navigation = currentWeb.get_navigation();
    navigation.set_useShared(true);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the web was set to use shared navigation");
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });
  }
  public async AddQuickLaunchItem(webUrl: string, title: string, url: string, isExternal: boolean) {
    let nnci: SP.NavigationNodeCreationInformation = new SP.NavigationNodeCreationInformation();
    nnci.set_title(title);
    nnci.set_url(url);
    nnci.set_isExternal(isExternal);
    const clientContext: SP.ClientContext = new SP.ClientContext(webUrl);
    const web = clientContext.get_web();
    web.get_navigation().get_quickLaunch().add(nnci);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });

  }
  public async RemoveQuickLaunchItem(webUrl: string, titlesToRemove: string[]) {
    const clientContext: SP.ClientContext = new SP.ClientContext(webUrl);
    const ql: SP.NavigationNodeCollection = clientContext.get_web().get_navigation().get_quickLaunch();
    clientContext.load(ql);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });

    let itemsToDelete = [];
    let itemCount = ql.get_count();
    for (let x = 0; x < itemCount; x++) {
      let item = ql.getItemAtIndex(x);
      let itemtitle = item.get_title();
      if (titlesToRemove.indexOf(itemtitle) !== -1) {
        itemsToDelete.push(item);
      }
    }
    for (let item of itemsToDelete) {
      item.deleteObject();
    }
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (error) => {
        console.log(error);
        reject();
      });
    });


  }

  public async fixUpLeftNav(webUrl: string, homeUrl: string) {

    await this.AddQuickLaunchItem(webUrl, "EFR Home", homeUrl, true);
    await this.RemoveQuickLaunchItem(webUrl, ["Pages", "Documents"]);

  }

  /**
   *  Adds a custom webpart to the edit form located at editformUrl
   * 
   * @param {string} webRelativeUrl -- The web containing the list
   * @param {any} editformUrl -- the url of the editform page
   * @param {string} webPartXml  -- the xml for the webpart to add
   * @memberof EfrAdmin
   */
  public async AddWebPartToEditForm(webRelativeUrl: string, editformUrl, webPartXml: string) {
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
  /**
   *  Adds a custom webpart to the edit form located at editformUrl
   * 
   * @param {string} webRelativeUrl -- The web containing the list
   * @param {any} editformUrl -- the url of the editform page
   * @param {string} webPartXml  -- the xml for the webpart to add
   * @memberof EfrAdmin
   */
  public async AddAdminEditForm(newWeb: Web, taskList: List,taskListId:string, taskListName: string, webServerRelativeUrl: string, webRelativeUrl: string, adminWebPartXml: string) {

   debugger;
    let listUrl = await taskList.rootFolder.serverRelativeUrl.get();
    let listFullUrl = document.location.protocol+"//"+
    document.location.hostname+"/"+listUrl;
    
    debugger;
    let editformUrl: string = "";
    let form = await taskList.rootFolder
      .files
      .addTemplateFile(listUrl + "/AdminEdit.aspx", 2).then((result) => { // 2 is a form page, 0 is a webpart page, 1 is a wiki page
        editformUrl = result.data.ServerRelativeUrl;
        //  Neeed to figure how  to use this let lwpm = result.file.getLimitedWebPartManager();
      }).catch((err) => {
        debugger;
        this.addMessage("<h1>error adding admin edit form</h1>");
        console.log(err);
        this.addMessage(err.data.responseBody["odata.error"].message.value);
      
      });
    debugger;
    const clientContext: SP.ClientContext = new SP.ClientContext(webRelativeUrl);
    var oFile = clientContext.get_web().getFileByServerRelativeUrl(editformUrl);
    var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
    let webparts = limitedWebPartManager.get_webParts();
    clientContext.load(webparts, 'Include(WebPart)');
    clientContext.load(limitedWebPartManager);
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        resolve();
      }, (err) => {
        this.addMessage("<h1>error getting webpartmanager on adminedit page</h1>");
        console.log(err);
        this.addMessage(err.data.responseBody["odata.error"].message.value);
        reject();
      });
    });
    // let originalWebPartDef = webparts.get_item(0);
    // let originalWebPart = originalWebPartDef.get_webPart();
    // originalWebPart.set_hidden(true);
    // originalWebPartDef.saveWebPartChanges();
    // await new Promise((resolve, reject) => {
    //   clientContext.executeQueryAsync((x) => {
    //     console.log("the webpart was hidden");
    //     resolve();
    //   }, (error) => {
    //     console.log(error);
    //     reject();
    //   });
    // });
    debugger;
    let webpartXml=adminWebPartXml.replace(/__LISTURL__/g, listFullUrl);
     webpartXml=webpartXml.replace(/__LISTID__/g, taskListId); //{5D85429B-96FD-4E2F-ACDF-A586B25A22F1}
    let oWebPartDefinition = limitedWebPartManager.importWebPart(webpartXml);
    let oWebPart = oWebPartDefinition.get_webPart();
    debugger;
    limitedWebPartManager.addWebPart(oWebPart, 'Main', 1);

    clientContext.load(oWebPart);
    debugger;
    await new Promise((resolve, reject) => {
      clientContext.executeQueryAsync((x) => {
        console.log("the new webpart was added");
        debugger;
        resolve();
      }, (err) => {
        debugger;
        this.addMessage("<h1>error adding adminedit webpart</h1>");
        console.log(err);
        this.addMessage(err.data.responseBody["odata.error"].message.value);
        reject();
      });
    });

    // form created now add a custom action for it
    taskList.userCustomActions.add({
      "Location": "EditControlBlock",
      "Title": "Admin Edit",
      "Url": listUrl + "/AdminEdit.aspx?ID={ItemId}",
      "Sequence": "1004",
      "Rights": ({ "High": "0", "Low": "2048" }) as any
    });
  }


  /**
   * Creates an EFR Quarterly subsite including secured libraries and an efr tsak list
   * 
   * @returns {Promise<any>} 
   * @memberof EfrAdmin
   */
  public async createSite(): Promise<any> {


    let newWeb: Web;  // the web that gets created
    let libraryList: Array<any>; // the list of libraries we need to create in the new site. has the library name and the name of the group that should get access
    // let foldersList: Array<string>; // the list of folders to create in each of the libraries.
    let roleDefinitions: Array<any>;// the roledefs for the site, we need to grant 'contribute no delete'
    let siteGroups: Array<any>;// all the sitegroups in the site
    let tasks: Array<any>; // the list of tasks in the TaskMaster list. We need to create on e task for each of these in tye EFRTasks list in the new site
    let taskList: List; // the task list we created  in the new site
    let taskListId: string; // the ID of task list we created  in the new site
    let webServerRelativeUrl: string; // the url of the subweb
    let contextInfo: ContextInfo;
    let editformurl: string;





    this.addMessage("CreatingSite");
    await sp.site.getContextInfo().then((context: ContextInfo) => {
      contextInfo = context;
    });
    // create the site
    await sp.web.webs.add(this.state.siteName, this.state.siteName, this.state.siteName, this.props.templateName).then((war: WebAddResult) => {
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

    await this.SetWebToUseSharedNavigation(webServerRelativeUrl);

    await this.fixUpLeftNav(webServerRelativeUrl, this.props.siteUrl);
    // now get  the list of libraries we need to create on the new site,

    await sp.web.lists.getByTitle(this.props.EFRLibrariesListName).items
      //   .top(2)
      .get().then((libraries) => {
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
    // now get  the list of folders  we need to create in each library
    // foldersList = await pnp.sp.web.lists.getByTitle(this.props.EFRFoldersListName).items.get().then((folders) => {
    //   this.addMessage("got list of folders");
    //   return map(folders, (f) => { return f["Title"]; });
    // }).catch(error => {
    //   debugger;
    //   this.addMessage("<h1>error fetching folder list</h1>");
    //   this.addMessage(error.data.responseBody["odata.error"].message.value);
    //   console.error(error);
    //   return null;
    // });
    // get the role definitions
    await sp.web.roleDefinitions.get().then((roleDefs) => {
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
    await sp.web.siteGroups.get().then((sg) => {
      this.addMessage("got Site Groups");
      siteGroups = sg;
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error getting site groups</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    // create the libraries and assign permissions
    // Also give the groups access to the site
    for (const library of libraryList) {
      if (!library["EFRsecurityGroup"]) {
        this.addMessage("bypassing Library " + library["Title"] + "because it has no security group");
      } else {
        this.addMessage("Creating library " + library["Title"]);
        await newWeb.lists.add(library["Title"], library["Title"], 101, false).then(async (listResponse) => {
          this.addMessage("Created Library " + library["Title"]);
          let list = listResponse.list;

          // for (const folder of foldersList) {
          //   await list.rootFolder.folders.add(folder)
          //     .then((results) => {
          //       console.log("created folder");
          //     })
          //     .catch((error) => {
          //       debugger;
          //       console.log("error creating folder");
          //     });
          // }

          // await folderBatch.execute().then((results) => {
          //   console.log("executed batch");
          // }).catch((error) => {
          //   console.log("error executing batch");
          //   return;

          // });

          let viewUrl: string;
          await list.views.getByTitle("All Documents").get().then((view) => {

            viewUrl = view.ServerRelativeUrl;
            return;
          }).catch(error => {
            debugger;
            this.addMessage("<h1>error getting AllDocuments view</h1>");
            this.addMessage(error.data.responseBody["odata.error"].message.value);
            console.error(error);
            return;
          });
          // Remove Libraries from Left Nav
          // await newWeb.navigation.quicklaunch.add(library["Title"], viewUrl, true).then((response) => {
          //   return;
          // }).catch(error => {
          //   debugger;
          //   this.addMessage("<h1>error adding list to quicklaunch " + library["Title"] + "</h1>");
          //   this.addMessage(error.data.responseBody["odata.error"].message.value);
          //   console.error(error);
          //   return;
          // });

          // Setup security on the library. First, break role inheritance
          await list.breakRoleInheritance(false).then((e) => {
            this.addMessage("broke role inheritance on " + library["Title"]);
          }).catch(error => {
            debugger;
            this.addMessage("<h1>error breaking role inheritance on  library " + library["Title"] + "</h1>");
            this.addMessage(error.data.responseBody["odata.error"].message.value);
            console.error(error);
            return;
          });
          // second , add the library-specific group
          let group = find(siteGroups, (sg => { return sg["Title"] === library["EFRsecurityGroup"]; }));
          let principlaID = group["Id"];
          let roledef = find(roleDefinitions, (rd => { return rd["Name"] === this.props.permissionToGrantToLibraries; }));
          if (!roledef) {
            this.addMessage("<h1>Role Definition  '" + this.props.permissionToGrantToLibraries + "' was not found</h1>");

          }
          let roleDefId = roledef["Id"];
          await list.roleAssignments.add(principlaID, roleDefId).then(() => {
            this.addMessage("granted " + library["EFRsecurityGroup"] + " read access to " + library["Title"]);
          }).catch(error => {
            debugger;
            this.addMessage("<h1>error adding role asisigment to  library " + library["Title"] + "</h1>");
            this.addMessage(error.data.responseBody["odata.error"].message.value);
            console.error(error);
            return;
          });
          // third  , add the global read access grouops
          for (let readgroupname of this.props.ReadAccessGroups.split(',')) {

            let readgroup = find(siteGroups, (sg => { return sg["Title"] === readgroupname; }));
            let readprinciplaID = readgroup["Id"];
            let readroledef = find(roleDefinitions, (rd => { return rd["Name"] === "Read"; }));
            let readroleDefId = readroledef["Id"];
            await list.roleAssignments.add(readprinciplaID, readroleDefId).then(() => {
              this.addMessage("granted " + readgroupname + "access to " + library["Title"]);
            }).catch(error => {
              debugger;
              this.addMessage("<h1>error adding role asisigment to  library " + library["Title"] + "</h1>");
              this.addMessage(error.data.responseBody["odata.error"].message.value);
              console.error(error);
              return;
            });
          }

          // fourth   , add the global write  access grouops
          for (let writegroupname of this.props.WriteAccessGroups.split(',')) {

            let writegroup = find(siteGroups, (sg => { return sg["Title"] === writegroupname; }));
            let writeprinciplaID = writegroup["Id"];
            let writeroledef = find(roleDefinitions, (rd => { return rd["Name"] === "Contribute"; }));
            let writeroleDefId = writeroledef["Id"];
            await list.roleAssignments.add(writeprinciplaID, writeroleDefId).then(() => {
              this.addMessage("granted " + writegroupname + " Contribute  access to " + library["Title"]);
            }).catch(error => {
              debugger;
              this.addMessage("<h1>error adding role asisigment to  library " + library["Title"] + "</h1>");
              this.addMessage(error.data.responseBody["odata.error"].message.value);
              console.error(error);
              return;
            });
          }
        });
      }

    }
    // get the master list of tasks, Either PBCMaster or PBCMaster YearEnd from the rootweb
    await sp.web.lists.getByTitle(this.state.pbcMasterList).items.expand("EFRLibrary").select("*,EFRLibrary/Title")
      .top(this.props.PBCMaximumTasks)
      .get().then((efrtasks) => {
        this.addMessage("got PBC MASTER list");
        tasks = efrtasks;
        return;
      }).catch(error => {
        debugger;
        this.addMessage("<h1>error fetching PBC MASTER list</h1>");
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        console.error(error);
        return;
      });
    //  create the task list in the site
    this.addMessage("Creating taskList ");

    await newWeb.lists.add("EFRTasks", "EFRTasks", 100, true,{"EnableAttachments":false}).then(async (listResponse) => {
      this.addMessage("Created List EFRTasks ");
      taskList = listResponse.list;
      taskListId = listResponse.data.Id;

      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error creating tasklist</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });

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

      editformurl = find(forms, (f: any) => { return f.FormType === 6; })["ServerRelativeUrl"];
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error fetching forms</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    await this.AddWebPartToEditForm(webServerRelativeUrl, editformurl, this.props.webPartXml);
    //add the PBC Task content type
    await taskList.contentTypes.addAvailableContentType(this.props.PBCTaskContentTypeId).then(ct => {
      this.addMessage("Added EFR Task content type");
      return;
    }).catch(error => {
      debugger;
      this.addMessage("<h1>error adding content type to task list</h1>");
      this.addMessage(error.data.responseBody["odata.error"].message.value);
      console.error(error);
      return;
    });
    // add the admin editforrm
    debugger;
    await this.AddAdminEditForm(newWeb, taskList, taskListId, "EFRTasks", webServerRelativeUrl, webServerRelativeUrl, this.props.adminWebPartXml);

    //add the default view to show only open items assigned to me sorted bt date descening
    await taskList.views.add("My Open Tasks", false, {
      RowLimit: 10,
      ViewQuery: '<OrderBy><FieldRef Name="EFRDueDate" Ascending="TRUE" /></OrderBy><Where><And><Eq><FieldRef Name="EFRAssignedTo" /><Value Type="Integer"><UserID Type="Integer" /></Value></Eq><Eq><FieldRef Name="EFRCompletedByUser" /><Value Type="Text">No</Value></Eq></And></Where>'
    }).then(async (v: ViewAddResult) => {
      // add this view to the quicklauch
      await this.AddQuickLaunchItem(window.location.origin + webServerRelativeUrl, "My Open Tasks", window.location.origin + v.data.ServerRelativeUrl, false);
      // set this as the homePage
      let homepage = v.data.ServerRelativeUrl.substr(webServerRelativeUrl.length + 1);
      await newWeb.rootFolder.update({ "WelcomePage": homepage }).then(() => {
        this.addMessage("Set Site homepage to this view");
      }).catch(error => {
        debugger;
        this.addMessage("<h1>error setting site home page</h1>");
        this.addMessage(error.data.responseBody["odata.error"].message.value);
        console.error(error);
        return;
      });
      // manipulate the view's fields

      await v.view.fields.removeAll().catch((err) => { debugger; });
      await v.view.fields.add("LinkTitle").catch((err) => { debugger; });
      await v.view.fields.add("EFRInformationRequested").catch((err) => { debugger; });
      await v.view.fields.add("EFRPeriod").catch((err) => { debugger; });
      await v.view.fields.add("EFRDueDate").catch((err) => { debugger; });
      await v.view.fields.add("EFRLibrary").catch((err) => { debugger; });
      await v.view.fields.add("EFRAssignedTo").catch((err) => { debugger; });
      await v.view.fields.add("EFRCompletedByUser").catch((err) => { debugger; });
      this.addMessage("Added My Open Tasks View");
      return;


    });
    // add the ALL OPEN TASKS VIEW
    await taskList.views.add("All Open Tasks", false, {
      RowLimit: 10,
      ViewQuery: '<OrderBy><FieldRef Name="EFRDueDate" Ascending="TRUE" /></OrderBy><Where><Eq><FieldRef Name="EFRVerifiedByAdmin" /><Value Type="Text">No</Value></Eq></Where>'
    }).then(async (v: ViewAddResult) => {
      // manipulate the view's fields

      await v.view.fields.removeAll().catch((err) => { debugger; });
      await v.view.fields.add("LinkTitle").catch((err) => { debugger; });
      await v.view.fields.add("EFRInformationRequested").catch((err) => { debugger; });
      await v.view.fields.add("EFRPeriod").catch((err) => { debugger; });

      await v.view.fields.add("EFRDueDate").catch((err) => { debugger; });
      await v.view.fields.add("EFRLibrary").catch((err) => { debugger; });

      await v.view.fields.add("EFRAssignedTo").catch((err) => { debugger; });
      await v.view.fields.add("EFRCompletedByUser").catch((err) => { debugger; });
      this.addMessage("Added All  Open Tasks View");
      return;


    });

    //add the a view to show alln items assigned to me sorted bt date descening
    //add the default view to show only open items assigned to me sorted bt date descening
    await taskList.views.add("All My Tasks", false, {
      RowLimit: 10,
      DefaultView: true,
      ViewQuery: '<OrderBy><FieldRef Name="EFRDueDate" Ascending="TRUE" /></OrderBy><Where><Eq><FieldRef Name="EFRAssignedTo" /><Value Type="Integer"><UserID Type="Integer" /></Value></Eq></Where>'
    }).then(async (v: ViewAddResult) => {

      await this.AddQuickLaunchItem(window.location.origin + webServerRelativeUrl, "My Tasks", window.location.origin + v.data.ServerRelativeUrl, false);
      // manipulate the view's fields
      await v.view.fields.removeAll().catch((err) => { debugger; });
      await v.view.fields.add("LinkTitle").catch((err) => { debugger; });
      await v.view.fields.add("EFRInformationRequested").catch((err) => { debugger; });

      await v.view.fields.add("EFRPeriod").catch((err) => { debugger; });

      await v.view.fields.add("EFRDueDate").catch((err) => { debugger; });
      await v.view.fields.add("EFRLibrary").catch((err) => { debugger; });

      await v.view.fields.add("EFRAssignedTo").catch((err) => { debugger; });
      await v.view.fields.add("EFRCompletedByUser").catch((err) => { debugger; });

    });


    // manipulate the All Items view's fields
    await taskList.views.getByTitle("All Items").fields.add("EFRInformationRequested").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRPeriod").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRDueDate").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRLibrary").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRAssignedTo").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRCompletedByUser").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRVerifiedByAdmin").catch((err) => { debugger; });
    await taskList.views.getByTitle("All Items").fields.add("EFRDateCompleted").catch((err) => { debugger; });

    // create the tasks in the new task list

    for (const task of tasks) {

      if (task.IsActive !== "No") {

        let itemToAdd = {
          "ContentTypeId": this.props.PBCTaskContentTypeId,
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
        };
        await taskList.items.add(itemToAdd).then((results) => {
          this.addMessage("added task " + task.Title);
          return;
        }).catch(error => {
          debugger;
          this.addMessage("<h1>error adding task " + task["Title"] + "</h1>");
          this.addMessage(error.data.responseBody["odata.error"].message.value);
          console.error(error);
          return;
        });
      }
    }

    // add the workflow to the task list
    // 2010 workflow is associate with the content type
    //await this.addNotificationWorkflow(contextInfo.SiteFullUrl, taskListId)


    //now we need to gramt permissions to the task list
    // each library has an associated group that gives them access to the library. 
    // those groups need access to the task list as well
    await taskList.breakRoleInheritance(true)
      .then(async (results) => {
        debugger;
        this.addMessage("Broke role inheritance on task list");
        /////////////////////////////////////////////////////////////////////////////////////////////
        for (const library of libraryList) {
          if (!library["EFRsecurityGroup"]) {
            this.addMessage(`Bypassing Library ${library["Title"]} because it has no security group`);
          } else {
            this.addMessage(`Granting group ${library["EFRsecurityGroup"]} access to the task list.`);

            let group = find(siteGroups, (sg => { return sg["Title"] === library["EFRsecurityGroup"]; }));
            let principlaID = group["Id"];
            let roledef = find(roleDefinitions, (rd => { return rd["Name"] === this.props.permissionToGrantToTaskList; }));
            if (!roledef) {
              this.addMessage("<h1>Role Definition  '" + this.props.permissionToGrantToTaskList + "' was not found</h1>");

            }
            let roleDefId = roledef["Id"];
            await taskList.roleAssignments.add(principlaID, roleDefId)
              .then(() => {
                this.addMessage(`Granted group ${library["EFRsecurityGroup"]} access to the task list.`);
                return;
              })
              .catch(error => {
                debugger;
                this.addMessage(`<h1>Error granting group ${library["EFRsecurityGroup"]} access to the task list.</h1>`);
                this.addMessage(error.data.responseBody["odata.error"].message.value);
                console.error(error);
                return;
              });
          }
        }
      })
      .catch(error => {
        console.log(error);

        this.addMessage("<h1>Error breaking role inheritance on task list</h1>");
        debugger;
        this.addMessage(error.data.responseBody["odata.error"].message.value);
      });

    /////////////////////////////////////////////////////////////////////////////////////////////






    this.addMessage("DONE!!");
  }
  // public async addNotificationWorkflow(webServerRelativeUrl, efrTaskListId: string): Promise<any> {
  //   debugger;
  //   let wf: SP.WorkflowServices.WorkflowDefinition;
  //   let historyListId: string;
  //   let taskListId: string;
  //   let workflowID: SP.Guid;
  //   const context: SP.ClientContext = new SP.ClientContext(webServerRelativeUrl);
  //   var workflowServicesManager = SP.WorkflowServices.WorkflowServicesManager.newObject(context, context.get_web());
  //   // connect to the deployment service
  //   var workflowDeploymentService = workflowServicesManager.getWorkflowDeploymentService();
  //   // get all installed workflows
  //   var publishedWorkflowDefinitions = workflowDeploymentService.enumerateDefinitions(true);

  //   context.load(publishedWorkflowDefinitions);

  //   await new Promise((resolve, reject) => {
  //     context.executeQueryAsync((x) => {
  //       resolve();
  //     }, (error) => {
  //       console.log(error);
  //       reject();
  //     });
  //   });

  //   debugger;
  //   var pwe = publishedWorkflowDefinitions.getEnumerator();
  //   console.log("wourkflowcount " + publishedWorkflowDefinitions.get_count());
  //   while (pwe.moveNext()) {
  //     debugger;
  //     let publishedWorkflowDefinition = pwe.get_current();
  //     debugger;
  //     console.log(publishedWorkflowDefinition.get_displayName());
  //     if (publishedWorkflowDefinition.get_displayName() === this.props.workflowName) {
  //       wf = publishedWorkflowDefinition;
  //       let wfid: string = wf.get_id().toString();
  //       workflowID = new SP.Guid(wfid);
  //     }
  //   }
  //   debugger;
  //   await pnp.sp.web.lists.getByTitle("Workflow History").get()
  //     .then((list) => {
  //       debugger;
  //       historyListId = list.Id;
  //     }).catch(error => {
  //       debugger;
  //       this.addMessage("<h1>error getting Workflow History listy</h1>");
  //       this.addMessage(error.data.responseBody["odata.error"].message.value);
  //       console.error(error);
  //       return;
  //     });;
  //   await pnp.sp.web.lists.getByTitle("Tasks").get()
  //     .then((list) => {
  //       debugger;
  //       taskListId = list.Id;
  //     }).catch(error => {
  //       debugger;
  //       this.addMessage("<h1>error creating workflow task list</h1>");
  //       this.addMessage(error.data.responseBody["odata.error"].message.value);
  //       console.error(error);
  //       return;
  //     });;

  //   debugger;



  //   // connect to the deployment service

  //   // connect to the subscription service
  //   var workflowSubscriptionService = workflowServicesManager.getWorkflowSubscriptionService();
  //   // create a new association / subscription
  //   let newSubscription = new SP.WorkflowServices.WorkflowSubscription(context, null);
  //   newSubscription.set_definitionId(workflowID);
  //   newSubscription.set_enabled(true);
  //   newSubscription.set_name("EFR Notifications");


  //   var startupOptions = new Array<string>();
  //   // automatic start
  //   // manual start
  //   startupOptions.push("WorkflowStart");

  //   // set the workflow start settings
  //   newSubscription.set_eventTypes(startupOptions);


  //   // set the associated task and history lists
  //   newSubscription.setProperty("HistoryListId", historyListId);
  //   newSubscription.setProperty("TaskListId", taskListId);

  //   // OPTIONAL: add any association form values
  //   //    newSubscription.SetProperty("Prop1", "Value1");
  //   //    newSubscription.SetProperty("Prop2", "Value2");

  //   // create the association
  //   workflowSubscriptionService.publishSubscriptionForList(newSubscription, taskListId);
  //   await new Promise((resolve, reject) => {
  //     context.executeQueryAsync((x) => {
  //       resolve();
  //       debugger;
  //     }, (request, error) => {
  //       console.log(error);
  //       reject();
  //     });
  //   });
  //   debugger;
  // }
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

        <Dropdown
          options={this.props.PBCMasterLists}
          selectedKey={this.state.pbcMasterList}

          label="PBC Master List"
          onChanged={(e) => {
            this.setState((current) => ({ ...current, pbcMasterList: e.text }));
          }} />
        <PrimaryButton onClick={this.createSite.bind(this)} title="Create Site">Create Site</PrimaryButton>

        <div dangerouslySetInnerHTML={this.displayMessages()} />

      </div >
    );
  }
}
