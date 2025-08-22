export interface CollaborationNewsModel {
  id: number;
  title: string;
  url: string;
  imageUrl?: string; 
  author?: string;  
  publishedDate: Date;
  views: number;  
  newsMeta: any;
  viewsLifeTime?: number;
  likes: number;
  likeCount?: number;
  Description: string; // Added for description
  categories?: string[];
}