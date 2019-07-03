import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import  {  sp,   Web } from "@pnp/sp";
import * as strings from 'EfrLockAndHideSitesWebPartStrings';
import EfrLockAndHideSites from './components/EfrLockAndHideSites';
import { IEfrLockAndHideSitesProps } from './components/IEfrLockAndHideSitesProps';
import { efrWeb, topNavItem } from "./model";
export interface IEfrLockAndHideSitesWebPartProps {
  libraryToTestForLockedSite: string;
  EFRLibariesList: string;
  permissionTotestForLockedSite: string;
  permissionToReplaceWith:string;
}

export default class EfrLockAndHideSitesWebPart extends BaseClientSideWebPart<IEfrLockAndHideSitesWebPartProps> {
  private efrWebs: Array<efrWeb> = [];
  private topNav: Array<topNavItem>;
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {

      sp.setup({
        spfxContext: this.context,
      });

      return this.loadData();
    });
  }
  public async loadData(): Promise<any> {
    let webInfos: Array<{ Title: string, ServerRelativeUrl: string, LastItemModifiedDate: string, Id: string }>;
    await sp.site.rootWeb.webinfos.get().then((webInfosResults) => {
      webInfos = webInfosResults;
    }).catch(err => {
      debugger;
      console.log(err);
      alert("error loading web infos");
      alert(err.data.responseBody["odata.error"].message.value);
    });

    await sp.site.rootWeb.navigation.topNavigationBar.get().then((topNavigationBarResults) => {
      debugger;
      this.topNav = topNavigationBarResults;
    }).catch(err => {
      debugger;
      console.log(err);
      alert("error loading topnav");
      alert(err.data.responseBody["odata.error"].message.value);
    });
    let securityGroupTotest: string; // this is just the name of the group
    // get the item in the libraryList to see the securtitygroup it uses. 
    await sp.web.lists.getByTitle(this.properties.EFRLibariesList)
      .items.filter("Title eq '" + this.properties.libraryToTestForLockedSite + "'").get().then(libraryItem => {
        console.log("security group to test is " + libraryItem[0].EFRsecurityGroup);
        securityGroupTotest = libraryItem[0].EFRsecurityGroup;
      }).catch(err => {
        debugger;
        console.log(err);
        alert("error getting security group for testlibrary");
        alert(err.data.responseBody["odata.error"].message.value);
      });
    let roleDefinitionId: number;
    await sp.web.roleDefinitions.getByName(this.properties.permissionTotestForLockedSite).get().then(roleDef => {
      console.log("role defintion to  " + roleDef.Id);
      roleDefinitionId = roleDef.Id;
    }).catch(err => {
      debugger;
      console.log(err);
      alert("error getting role definition id");
      alert(err.data.responseBody["odata.error"].message.value);
    });
    let securityGroupTotestPrinciipalId: number;
    // get the principalId for that security group/
    await sp.web.siteGroups.getByName(securityGroupTotest)
      .get().then(spGroup => {
        securityGroupTotestPrinciipalId = spGroup.Id;
      }).catch(err => {
        debugger;
        console.log(err);
        alert("error getting security group for testlibrary");
        alert(err.data.responseBody["odata.error"].message.value);
      });

    for (let webInfo of webInfos) {

      let subweb: Web;
      await sp.site.openWebById(webInfo.Id).then(sw => {
        subweb = sw.web;
      }).catch(err => {
        debugger;
        console.log(err);
        alert("error openning web  " + webInfo.Title);
        alert(err.data.responseBody["odata.error"].message.value);
      });


      await subweb.lists
        .getByTitle(this.properties.libraryToTestForLockedSite)
        //.roleAssignments.filter("PrincipalId eq " + securityGroupTotestPrinciipalId+ " and RoleDefinitionBindings/Description eq '"+this.properties.permissionTotestForLockedSite +"'")
        .roleAssignments.filter("PrincipalId eq " + securityGroupTotestPrinciipalId)
        .expand('RoleDefinitionBindings')
        .get()
        .then(roleAsssignments => {

          // see if the group has the required permission
          let islocked = true;
          for (let roleAsssignment of roleAsssignments) {
            for (let roleDefinitionBinding of roleAsssignment.RoleDefinitionBindings) {
              if (roleDefinitionBinding.Name === this.properties.permissionTotestForLockedSite) {
                islocked = false;
              }
            }
          }
          this.efrWebs.push({
            title: webInfo.Title,
            isLocked: islocked,
            url: webInfo.ServerRelativeUrl,
            lastModifiedDate: webInfo.LastItemModifiedDate,
            id: webInfo.Id
          });
        }).catch(err => {
          debugger;
          console.log(err);
          alert("error fetching list " + this.properties.libraryToTestForLockedSite + " from site " + webInfo.ServerRelativeUrl);
          alert(err.data.responseBody["odata.error"].message.value);
        });
    }


  }
  // public removeSiteFromTopNav(navItem: topNavItem): Promise<any> {
  //   debugger;
  //   return pnp.sp.site.rootWeb.navigation.topNavigationBar.getById(navItem.Id).Update({ IsVisible: false }).then((results) => {
  //     debugger;
  //   }).catch(err => {
  //     debugger;
  //   });
  // }
  public async lockSite(web: efrWeb): Promise<any> {
    debugger;
    // the roledefinition top replace witj 
    let replacementRoleDefinitionId: number;
    await sp.web.roleDefinitions.getByName(this.properties.permissionToReplaceWith).get().then(roleDef => {
      console.log("role defintion to  " + roleDef.Id);
      replacementRoleDefinitionId = roleDef.Id;
    }).catch(err => {
      debugger;
      console.log(err);
      alert("error getting role definition id");
      alert(err.data.responseBody["odata.error"].message.value);
    });
    let libs: Array<any>;
    await sp.web.lists.getByTitle(this.properties.EFRLibariesList).items.get().then(results => {
      libs = results;
    }).catch((err) => {
      console.error(err);
      alert("error loading list of libraries");
    });
    let subweb: Web;
    await sp.site.openWebById(web.id).then(sw => {
      subweb = sw.web;
    }).catch(err => {
      debugger;
      console.log(err);
      alert("error openning web  " + web.title);
      alert(err.data.responseBody["odata.error"].message.value);
    });
    for (let lib of libs) {
      let secGroup = lib.EFRsecurityGroup;
      let principalId: number;
      // get the grouip assiciated with this library
      await sp.web.siteGroups.getByName(secGroup).get().then((sg) => {
        principalId = sg.Id;
      }).catch((err) => {
        debugger;
        console.log(err);
        alert("error openning web  " + web.title);
        alert(err.data.responseBody["odata.error"].message.value);
      });
      await subweb.lists
        .getByTitle(lib.Title)
        .roleAssignments.getById(principalId).bindings.get()
        .then(async rdbs => {
          debugger;
          for (let rdb of rdbs) {
            if (rdb.Name === this.properties.permissionTotestForLockedSite) {
              await subweb.lists.getByTitle(lib.Title).roleAssignments.getById(principalId).delete()
              .then(() => {})
              .catch((err)=>{
                debugger;
                console.error(err);
                alert("an error occurred removing the permission from library "+lib.Title);
                return;
              });

              await subweb.lists.getByTitle(lib.Title).roleAssignments.add(principalId, replacementRoleDefinitionId).then(() => {
                
              }).catch(error => {
                debugger;
                alert("an error occurred granting " + this.properties.permissionToReplaceWith+" permission to library "+lib.Title);
                console.error(error);
                return;
              });
            }
          }
        }).catch((err) => {
          debugger;
          console.log(err);
          alert("error fetching list " +lib.Title +"from site " +web.title);
          alert(err.data.responseBody["odata.error"].message.value);
        });

    }


    //  // return pnp.sp.site.rootWeb.navigation.topNavigationBar.getById(navItem.Id).update({ IsVisible: false }).then((results) => {
    //     debugger;
    //   }).catch(err => {
    //     debugger;
    //   });
  }
  public render(): void {
    const element: React.ReactElement<IEfrLockAndHideSitesProps> = React.createElement(
      EfrLockAndHideSites,
      {
        efrWebs: this.efrWebs,
        topNav: this.topNav,
        removeSiteFromTopNav:null,// this.removeSiteFromTopNav.bind(this),
        lockSite: this.lockSite.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('libraryToTestForLockedSite', {
                  label: "Library to test for Locked Site"
                }),
                PropertyPaneTextField("EFRLibariesList", {
                  label: "The list of libraries  in each subsite (EFRLibraries)"
                }),
                PropertyPaneTextField("permissionTotestForLockedSite", {
                  label: "The Permission used to test for if a site is locked (Content Authors without delete or modify)"
                }),
                PropertyPaneTextField("permissionToReplaceWith", {
                  label: "The Permission to replace with to lock a site(Read)"
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
