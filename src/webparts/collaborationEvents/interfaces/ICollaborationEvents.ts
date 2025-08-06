import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICollaborationEventsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;

  userData: { mail: string, title: string };
  fetchGroupData: () => Promise<any>;
  context: WebPartContext;

  dataSource: string;
  location: string;
  Cardlayout: string;
  StylesForCards: string;
  StylesForImages: string;
  Events: string;
  Field: string;
  SharePointLists: { key: string; text: string }[];
  headingfontcolor: string;





}

export interface ICollaborationEventsWebPartProps {
  description: string;
  dataSource: string;
  location: string;
  headingfontcolor: string;



  Cardlayout: string;
  StylesForCards: string;
  StylesForImages: string;
  Events: string;
  Field: string
  SharePointLists: { key: string; text: string }[];




}
export interface ICollaborationEventsModelProps {
  description: string;
  Cardlayout: string;
  dataSource: string;
  location: string;
  headingfontcolor: string;


  StylesForCards: string;
  StylesForImages: string;
  Events: string;
  Field: string
  SharePointLists: { key: string; text: string }[];





  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  context: WebPartContext;

}
export interface ICollaborationEventsModelState {
  userData: { mail: string, title: string };
}





export interface ICollaborationEventsState {
  EventCategory: any[];
  EventTitle: string;
  Eventlocation: string;
  EventDescription: string;
  EventLink: string;
  selectedCategory: IDropdownOption;

  StartDate: Date;
  StartTime: Date;

  EndDate: Date;
  EndTime: Date;

  AlldayNeeded: boolean;
  EnableEventTemplate: boolean;

  ListDetails: any[];

  isPopupVisible: boolean;
  selectedImage: any;
  EventImage: string;
}