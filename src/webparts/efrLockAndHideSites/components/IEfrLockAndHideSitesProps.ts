import { efrWeb, topNavItem } from "../model";
export interface IEfrLockAndHideSitesProps {
  efrWebs: Array<efrWeb>;
  topNav: Array<topNavItem>;
  removeSiteFromTopNav: (topNavItem) => Promise<any>;
  lockSite: (efrWeb) => Promise<any>;
}
