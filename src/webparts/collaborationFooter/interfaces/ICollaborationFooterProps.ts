import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICollaborationFooterProps {
  description: string;
  siteUrl: string;
  context: WebPartContext;
  adminList:string;
  settingsList:string;
}
export interface ICollaborationFooterState {
  currentUserId:any;
  isAdmin:boolean;
  logoUrl:string;
  logolink:string;
  linkedinLogoUrl:string;
  linkedinLogolink:string;
  footerdata:any[];
  openEditModal:boolean;
  footerdataitem:any[];
  openEditItemModal:boolean;
  title:string;
  itemLink:string;
  itemId:any;
  openAddModal:boolean;
  itemLogo:any;
}