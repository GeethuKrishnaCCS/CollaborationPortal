import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICollaborationEventsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;

  userData: { mail: string, title: string };
  fetchGroupData: () => Promise<any>;
  context: WebPartContext;
}
export interface ICollaborationEventsModelProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
}
export interface ICollaborationEventsModelState {
  userData: { mail: string, title: string };
}

export interface ICollaborationEventsWebPartProps {
  description: string;
}

export interface ICollaborationEventsState {
  // showPanel: boolean;
  // showModal: boolean;
  // groupData: any[];

}