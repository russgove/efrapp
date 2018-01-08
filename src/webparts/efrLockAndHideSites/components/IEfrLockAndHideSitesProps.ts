import {efrWeb,topNavItem}from "../model";
import pnp, { RoleDefinitionBindings, NavigationNodes, SearchQuery, SearchResults, SortDirection, EmailProperties, Items, Web } from "sp-pnp-js";
export interface IEfrLockAndHideSitesProps {
  efrWebs:Array<efrWeb>;
  topNav:Array<topNavItem>;
  removeSiteFromTopNav: (topNavItem)=>Promise<any>;
  lockSite: (efrWeb)=>Promise<any>;

}
