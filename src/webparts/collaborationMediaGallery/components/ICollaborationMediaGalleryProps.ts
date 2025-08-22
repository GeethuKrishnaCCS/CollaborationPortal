// export interface ICollaborationMediaGalleryProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }



export interface ICollaborationMediaGalleryProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  listId: string;
  numberOfItems: number;
  description: string;
  headingfontcolor: string;
}

export interface ICollaborationMediaGalleryState {
  items: IVideo[];
  selectedMedia?: IVideo;
  loading: boolean;
  error: string
}

export interface IVideo {
  ID: number;
  title: string;
  description: string;
  serverRelativeUrl: string;
  isSelected: boolean;
  pURL: string;
  thumbnailURL: string;
}
