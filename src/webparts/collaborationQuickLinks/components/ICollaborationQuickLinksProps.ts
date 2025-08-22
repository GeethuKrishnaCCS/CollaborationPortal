// export interface ICollaborationQuickLinksProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }

import { CollaborationQuickLinksModel} from "../models/CollaborationQuickLinksModel"; 
import { CollaborationQuickLinksService } from "../services/CollaborationQuickLinksService";
import { LayoutType } from "../models/CollaborationQuickLinksLayout";
import { ILayoutSettings } from "../models/CollaborationQuickLinksLayout";

export interface ICollaborationQuickLinksProps {
  title: string;
  links: CollaborationQuickLinksModel[];
  layoutType: LayoutType;
  service: CollaborationQuickLinksService; 
  layoutSettings: Record<LayoutType, ILayoutSettings>;
  onManageLink: (id: string) => void;
  context: any;
  defaultBorderRadius: number
  imageOpacity: number;
  color: string;
  iconColor: string;
  textColor: string;
  cardColor: string;
  filmstripSize?: 'small' | 'medium' | 'large'; // Added filmstripSize prop
  TargetAudiences?: any[];
  enableAudienceTargeting: boolean;
  quickAccessHeader?: string;
  quickAccessBackgroundColor?: string;
  webPartInstanceId: string;
  TitleColor?: string;
  BottonColor?: string;      
  cardAlignment?: 'left' | 'center' | 'right';       
  cardHoverColor?: string;   
  TitleFontSize?: number; 
  showImageBackground?: boolean;
  quickAccessHeaderAlignment?: 'left' | 'center' | 'right';
}
