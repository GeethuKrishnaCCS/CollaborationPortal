import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IProductAnniversaryProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  dataSource: string;
  Cardlayout: string;
  Anniversary: string;
  SharePointLists: { key: string; text: string }[];

  headingfontcolor: string;
  headingFontFamily: string;
  contentfontcolor: string;
  contentFontfamily: string;
  StylesForCards: string;
  AnniversayNoOfItemDisplay: string;

  BdayToggleValue: boolean;
  WorkToggleValue: boolean;
  WeddingToggleValue: boolean;

  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;


  displayListFields: any;

  ShapeForImages: string;
  // displayItems: string;
}

export interface IProductAnniversaryWebPartProps {
  description: string;
  dataSource: string;
  Cardlayout: string;
  Anniversary: string;
  SharePointLists: { key: string; text: string }[];

  headingfontcolor: string;
  contentfontcolor: string;
  headingFontFamily: string;
  contentFontfamily: string;
  StylesForCards: string;
  AnniversayNoOfItemDisplay: string;

  BdayToggleValue: boolean;
  WorkToggleValue: boolean;
  WeddingToggleValue: boolean;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;

  displayListFields: any;
  ShapeForImages: string;
  // displayItems: string;
}

export interface IProductAnniversaryModelProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
  dataSource: string;
  Cardlayout: string;
  Anniversary: string;
  SharePointLists: { key: string; text: string }[];
  userDisplayName: string;

  headingfontcolor: string;
  contentfontcolor: string;
  headingFontFamily: string;
  contentFontfamily: string;
  
  StylesForCards: string;
  AnniversayNoOfItemDisplay: string;

  BdayToggleValue: boolean;
  WorkToggleValue: boolean;
  WeddingToggleValue: boolean;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;
  ShapeForImages: string;
  // displayItems: string;

  displayListFields: any;
}



export interface IProductAnniversaryState { 
  listItems: any;
  today: string;
  greetings: any;

}