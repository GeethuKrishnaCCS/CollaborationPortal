export interface IEventItem {
  Id: number; 
    Title: string;
    EventDate: Date;
    EndDate: Date;
    Location?: string;
    Description?: string;
    Source: 'SharePoint' | 'Outlook';
    TeamMembers: string[];
    resource?: any;
  }
    
  