import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEtihadBannerProps {
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
}

export interface IEtihadBannerWebPartProps {
  description: string;
  BannerImageUrl: string;
  welcomeText?: string;
  missionText?: string;
  missionDescription?: string;

}