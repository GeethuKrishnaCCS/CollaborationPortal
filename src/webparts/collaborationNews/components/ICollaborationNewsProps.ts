import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { CollaborationNewsModel } from "../models/CollaborationNewsModel";

export interface ICollaborationNewsProps {
  StyleToggle: "list" | "topstory" | "sidebyside" | "hubnews" | "carousel" | "tiles";
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
  LikeToggle: boolean;
  CommentToggle: boolean;
  selectedSites: string[];
  selectedCategories: string[];
  selectedGroups?: IPropertyFieldGroupOrPerson[];
  numberOfNews: number;
  context: WebPartContext;
  placeholderImageUrl: string;
  headingColor: string 
  backgroundColor: string 
  customTitle: string;
  isAdmin: boolean; 
  // canAddNews: boolean;
  newsContainerColor: string; 
  initialNews?: CollaborationNewsModel[];
  handleAddNewsPost?: (siteUrl: string) => void;
  handleAddNewsLink?: (siteUrl: string) => void;
  buttonBackgroundColor?: string;
  textColor?: string;
}
