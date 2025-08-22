export enum LayoutType {
    Compact = "Compact",
    FilmStrip = "Filmstrip",
    Grid = "Grid",
    Button = "Button",
    List = "List",
    Tiles = "Tiles",
    ImageWithTitle= "ImageWithTitle",
}

export enum IconSize {
    Small = 22,
    Medium = 30,
    Large = 36,
  
}


export interface ILayoutSettings {
    hideText?: boolean;
    iconSize?: IconSize;
    tileSize?: {
       tileWidth: number; 
  tileHeight: number;
  numColumns: number;
    }
    imageSize?: {
        width: number;
        height: number;
    };
    containerWidth?: number; 
    containerHeight?: number;
  
}
