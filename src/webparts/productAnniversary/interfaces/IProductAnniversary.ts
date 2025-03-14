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
  contentfontcolor: string;
  StylesForCards: string;
  AnniversayNoOfItemDisplay: string;

  BdayToggleValue: boolean;
  WorkToggleValue: boolean;
  WeddingToggleValue: boolean;

  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;

  // displayName: boolean;
  // displayJobTitle: boolean;
  // displayLocation: boolean;
  // displayEmail: boolean;
  // displayType: boolean;
  // displayDOB: boolean;
  // displayJoiningDate: boolean;
  // displayWeddingDate: boolean;

  // ListFields: { key: string; text: string }[];

  ListFields: Array<{ key: string; text: string }>;
  [key: string]: any;

  ShapeForImages: string;
  displayItems: string;
}

export interface IProductAnniversaryWebPartProps {
  description: string;
  dataSource: string;
  Cardlayout: string;
  Anniversary: string;
  SharePointLists: { key: string; text: string }[];

  headingfontcolor: string;
  contentfontcolor: string;
  StylesForCards: string;
  AnniversayNoOfItemDisplay: string;

  BdayToggleValue: boolean;
  WorkToggleValue: boolean;
  WeddingToggleValue: boolean;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;

  // displayName: boolean;
  // displayJobTitle: boolean; 
  // displayLocation: boolean; 
  // displayEmail: boolean;
  // displayType: boolean;
  // displayDOB: boolean;
  // displayJoiningDate: boolean;
  // displayWeddingDate: boolean;

  // ListFields: { key: string; text: string }[];

  ListFields: Array<{ key: string; text: string }>;
  [key: string]: any;


  ShapeForImages: string;
  displayItems: string;
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
  
  StylesForCards: string;
  AnniversayNoOfItemDisplay: string;

  BdayToggleValue: boolean;
  WorkToggleValue: boolean;
  WeddingToggleValue: boolean;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;

  // displayName: boolean;
  // displayJobTitle: boolean;
  // displayLocation: boolean;
  // displayEmail: boolean;
  // displayType: boolean;
  // displayDOB: boolean;
  // displayJoiningDate: boolean;
  // displayWeddingDate: boolean;

  // ListFields: { key: string; text: string }[];
  
  ListFields: Array<{ key: string; text: string }>;
  [key: string]: any;

  ShapeForImages: string;
  displayItems: string;
}



export interface IProductAnniversaryState { 
  listItems: any;
  today: string;
  greetings: any;

  // RenderedGreetings?: any;  
  // Next?: any; 
  // Count?: any;
}