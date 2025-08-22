export interface CollaborationQuickLinksModel {
  Id: string;
  Title: string;
  URL: string;
  Icon: string;
  ThumbnailOption?: 'default' | 'file' | 'icon' | 'text-only';
  SortWeight: number;
  Description: string;
  ImageOpacity?: number;
  FillType?: 'color' | 'image';
  FillColor?: string;
  color?: string;
  textColor?: string;
  iconColor?: string;
  TitleColor?: string;
  BottonColor?: string;
  quickAccessHeader?: string;
  quickAccessBackgroundColor?: string;
  filmstripSize?: 'small' | 'medium' | 'large';
  TextWrap?: boolean;
  BorderRadius?: number;
  Target?: '_self' | '_blank'
  TargetAudiences:any[];
  WebPartInstanceId: string;
  cardAlignment?: 'left' | 'center' | 'right';
  cardHoverColor?: string;
  TitleFontSize?: number; 
  showImageBackground?: boolean;
}