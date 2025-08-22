// export interface ICollaborationOrgAnnouncementProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }



import { IPropertyPaneAccessor, WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';

export interface ICollaborationOrgAnnouncementProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  layout: string;
  context: WebPartContext
  spHttpClient: SPHttpClient;
  loggerName: string;
  webUrl: string;
  listName: string;
  colCount: number;
  isAutorotate: boolean;
  duration: number;
  imagesCount: number;
  width: number;
  height: number;
  propertyPane: IPropertyPaneAccessor;
  WebpartTitle: string;
  ColumnSection: string;
}

export interface ICollaborationOrgAnnouncementState {
  imageCount?: number;
  imageInfo?: IImageDetails;
}

export class Constants {

  public static get CaroselMax(): number {
    return 10;
  }
  public static get ListMax(): number {
    return 25;
  }
  public static get LightboxMax(): number {
    return 50;
  }
}
export interface IImageDetails {
  info: IImageDetail[];
}
export interface IImageDetail {
  name: string;
  redirectLink: string;
  caption: string;
  path: string;
  description?: string;
}
