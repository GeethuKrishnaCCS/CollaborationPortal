import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface ICollaborationCalendarProps {
  title: string;
  siteUrl: string;
  listName: string; 
 
  context: WebPartContext;
  showOutlookEvents: boolean;
  showSpEventsOnly: boolean;
  outlookEventColor?: string;
  spEventColor?: string;
  calendarBackgroundColor?: string;

}