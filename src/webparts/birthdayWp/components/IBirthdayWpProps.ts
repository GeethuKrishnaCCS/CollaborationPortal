import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBirthdayWpProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  layout:string;
  context:WebPartContext;
}