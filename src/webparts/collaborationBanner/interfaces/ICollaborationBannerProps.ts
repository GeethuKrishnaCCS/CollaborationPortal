import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICollaborationBannerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  BannerImageUrl: string;
  welcomeText?: string;
  missionText?: string;
  missionDescription?: string;
  context: WebPartContext;
  selectDesign?: string;
  imageOverlapBackgroundColor?: string;
  welcomeheadingfontcolor?: string;
  usernameheadingfontcolor?: string;
  descriptionfontcolor?: string;
}

export interface ICollaborationBannerWebPartProps {
  description: string;
  BannerImageUrl: string;
  welcomeText?: string;
  missionText?: string;
  missionDescription?: string;
  selectDesign?: string;
  imageOverlapBackgroundColor?: string;
  welcomeheadingfontcolor?: string;
  usernameheadingfontcolor?: string;
  descriptionfontcolor?: string;

}