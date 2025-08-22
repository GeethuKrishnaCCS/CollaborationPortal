import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle,
  PropertyPaneButton,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import CollaborationQuickLinks from './components/CollaborationQuickLinks';
import { ICollaborationQuickLinksProps } from './components/ICollaborationQuickLinksProps';
import { CollaborationQuickLinksModel } from './models/CollaborationQuickLinksModel';
import { LayoutType, IconSize, ILayoutSettings } from './models/CollaborationQuickLinksLayout';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import {
  PropertyPaneSlider,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { CollaborationQuickLinksService } from './services/CollaborationQuickLinksService';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { PropertyFieldFilePicker, IFilePickerResult } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

const CONTAINER_SIZES: { [key: string]: { width: number; height: number } } = {
  [LayoutType.ImageWithTitle]: { width: 350, height: 250 },
  [LayoutType.Compact]: { width: 50, height: 50 },
  [LayoutType.FilmStrip]: { width: 50, height: 50 },
  [LayoutType.Grid]: { width: 50, height: 50 },
  [LayoutType.Button]: { width: 50, height: 50 },
  [LayoutType.List]: { width: 50, height: 50 },
  [LayoutType.Tiles]: { width: 50, height: 50 },
};

const TILE_SIZES: { [key: string]: { tileWidth: number; tileHeight: number } } = {
  [LayoutType.ImageWithTitle]: { tileWidth: 350, tileHeight: 250 },
  [LayoutType.Compact]: { tileWidth: 500, tileHeight: 80 },
  [LayoutType.FilmStrip]: { tileWidth: 300, tileHeight: 150 },
  [LayoutType.Grid]: { tileWidth: 200, tileHeight: 150 },
  [LayoutType.Button]: { tileWidth: 500, tileHeight: 80 },
  [LayoutType.List]: { tileWidth: 180, tileHeight: 80 },
  [LayoutType.Tiles]: { tileWidth: 150, tileHeight: 150 },
};

const TILE_MIN_WIDTHS: { [key: string]: number } = {
  [LayoutType.ImageWithTitle]: 150,
  [LayoutType.Compact]: 150,
  [LayoutType.FilmStrip]: 120,
  [LayoutType.Grid]: 150,
  [LayoutType.Button]: 150,
  [LayoutType.List]: 150,
  [LayoutType.Tiles]: 100,
};

const TILE_MIN_HEIGHTS: { [key: string]: number } = {
  [LayoutType.ImageWithTitle]: 80,
  [LayoutType.Compact]: 50,
  [LayoutType.FilmStrip]: 60,
  [LayoutType.Grid]: 100,
  [LayoutType.Button]: 50,
  [LayoutType.List]: 60,
  [LayoutType.Tiles]: 100,
};

const IMAGE_MIN_WIDTHS: { [key: string]: number } = {
  [LayoutType.ImageWithTitle]: 10,
  [LayoutType.Compact]: 10,
  [LayoutType.FilmStrip]: 10,
  [LayoutType.Grid]: 10,
  [LayoutType.Button]: 10,
  [LayoutType.List]: 10,
  [LayoutType.Tiles]: 10,
};

const IMAGE_MIN_HEIGHTS: { [key: string]: number } = {
  [LayoutType.ImageWithTitle]: 10,
  [LayoutType.Compact]: 10,
  [LayoutType.FilmStrip]: 10,
  [LayoutType.Grid]: 10,
  [LayoutType.Button]: 10,
  [LayoutType.List]: 10,
  [LayoutType.Tiles]: 10,
};

export interface ICollaborationQuickLinksWebPartProps {
  title: string;
  links: CollaborationQuickLinksModel[];
  selectedLinkId?: string;
  tenantUrl: string;
  layoutType: LayoutType;
  defaultBorderRadius: number;
  enableDescription: boolean
  imageOpacity: number;
  listName: string;
  color: string;
  layoutSettings: Record<LayoutType, ILayoutSettings>;
  iconSize?: IconSize;
  enableAudienceTargeting: boolean;
  TargetAudiences: any[];
  quickAccessHeader: string; // Added for Quick Access header text
  quickAccessBackgroundColor: string;
  TitleColor: string;
  BottonColor: string;
  iconColor: string;
  textColor: string;
  cardColor: string;
  numColumns: number;
  webPartInstanceId: string;
  filmstripSize?: 'small' | 'medium' | 'large';
  cardAlignment?: 'left' | 'center' | 'right';
  cardHoverColor?: string;
  TitleFontSize?: number;
  quickAccessHeaderAlignment?: 'left' | 'center' | 'right';
}

export default class CollaborationQuickLinksWebPart extends BaseClientSideWebPart<ICollaborationQuickLinksWebPartProps> {
  private quickLinkService: CollaborationQuickLinksService;
  private selectedItemId: string | null = null;

  public async onInit(): Promise<void> {
    this.quickLinkService = new CollaborationQuickLinksService(this.context, this.properties.listName);
    if (typeof this.properties.defaultBorderRadius === 'undefined') {
      this.properties.defaultBorderRadius = 4; // Default border radius
    }
    if (typeof this.properties.imageOpacity === 'undefined') {
      this.properties.imageOpacity = 100; // Default opacity (100%)
    }
    if (!this.properties.color) {
      this.properties.color = '#C8EFFE'; // Default color
    }


    if (!this.properties.links || this.properties.links.length === 0) {
      this.properties.links = await this.quickLinkService.getLinks(this.instanceId);
    }
    if (!this.properties.layoutType) {
      this.properties.layoutType = LayoutType.ImageWithTitle; // set your desired default layout here
    }
    return super.onInit();



  }


  public render(): void {

    const currentLayoutType = this.properties.layoutType || LayoutType.ImageWithTitle;

    // Initialize layoutSettings if undefined
    if (!this.properties.layoutSettings) {
      this.properties.layoutSettings = {} as Record<LayoutType, ILayoutSettings>;
    }

    // Initialize settings for the current layout type if not set
    if (!this.properties.layoutSettings[currentLayoutType]) {
      this.properties.layoutSettings[currentLayoutType] = {
        hideText: false,
        iconSize: IconSize.Medium,
        imageSize: { width: 50, height: 50 },
      };
    }


    const element: React.ReactElement<ICollaborationQuickLinksProps> = React.createElement(
      CollaborationQuickLinks, {
      title: this.properties.title,
      links: this.properties.links,
      layoutType: currentLayoutType,
      layoutSettings: this.properties.layoutSettings,
      service: this.quickLinkService,
      defaultBorderRadius: this.properties.defaultBorderRadius,
      imageOpacity: this.properties.imageOpacity,
      color: this.properties.color,
      context: this.context,
      iconColor: this.properties.iconColor,
      textColor: this.properties.textColor,
      cardColor: this.properties.cardColor,
      TitleColor: this.properties.TitleColor,
      BottonColor: this.properties.BottonColor,
      enableAudienceTargeting: this.properties.enableAudienceTargeting,
      quickAccessHeader: this.properties.quickAccessHeader,
      quickAccessBackgroundColor: this.properties.quickAccessBackgroundColor,
      webPartInstanceId: this.instanceId,
      filmstripSize: this.properties.filmstripSize || 'medium',
      cardAlignment: this.properties.cardAlignment,
      cardHoverColor: this.properties.cardHoverColor,
      TitleFontSize: this.properties.TitleFontSize,
      onManageLink: async (id: string) => {
        this.selectedItemId = id;
        // Refresh links before opening property pane
        this.properties.links = await this.quickLinkService.getLinks(this.instanceId);
        this.context.propertyPane.open();
      },

    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // In your web part class, modify the onPropertyPaneFieldChanged method
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  if (propertyPath === 'cardAlignment') {
    this.properties.cardAlignment = newValue;
    this.context.propertyPane.refresh();
    this.render();
  } else if (propertyPath.startsWith('links')) {
    const match = /links\[(\d+)\]\.(.+)/.exec(propertyPath);
    if (match) {
      const linkIndex = parseInt(match[1], 10);
      const field = match[2];

      if (field === 'Target') {
        this.properties.links[linkIndex].Target = newValue;
        this.context.propertyPane.refresh();
        this.render();
      } else if (field === 'ThumbnailOption') {
        this.properties.links[linkIndex].ThumbnailOption = newValue;
        if (newValue === 'default') {
          this.properties.links[linkIndex].Icon = 'Globe';
        } else if (newValue === 'text-only') {
          this.properties.links[linkIndex].Icon = '';
        }
        this.context.propertyPane.refresh();
        this.render();
      }
    }
  } else if (propertyPath === 'enableAudienceTargeting') {
    this.context.propertyPane.refresh();
    this.render();
  }
  super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
}


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const index = this.selectedItemId
      ? this.properties.links.findIndex((x) => x.Id === this.selectedItemId)
      : -1;

    if (index !== -1) {
      return this.getLinkManagementPane(index);
    }

    return this.getDefaultPropertyPane();

  }

  private getLinkManagementPane(index: number): IPropertyPaneConfiguration {

    type ThumbnailOption = "default" | "file" | "icon" | "text-only";
    const link = this.properties.links[index];
    const currentIcon = link.Icon;

    // Determine thumbnail option
    const inferredOption: ThumbnailOption =
      link.ThumbnailOption === 'text-only' ? 'text-only' :
        currentIcon && currentIcon.startsWith("http") ? "file" :
          currentIcon ? "icon" :
            "default";

    const selectedThumbnailOption = link.ThumbnailOption as ThumbnailOption || inferredOption;

    const thumbnailOptions = [
      { key: 'default', text: 'Default' },
      { key: 'file', text: 'Custom image' },
      { key: 'icon', text: 'Icon' },
      { key: 'text-only', text: 'Text only' },
    ];


    // Prepare file picker result for existing custom images
    const filePickerResult: IFilePickerResult = currentIcon && currentIcon.startsWith('http')
      ? {
        fileAbsoluteUrl: currentIcon,
        fileName: currentIcon.split('/').pop() || 'custom-image',
        fileNameWithoutExtension: (currentIcon.split('/').pop() || 'custom-image').split('.').slice(0, -1).join('.') || 'custom-image',
        downloadFileContent: async () => {
          try {
            const response = await fetch(currentIcon);
            if (!response.ok) throw new Error('Failed to fetch image');
            const blob = await response.blob();
            return new File([blob], currentIcon.split('/').pop() || 'custom-image');
          } catch (error) {
            console.error('Error loading image:', error);
            return new File([], 'error');
          }
        }
      }
      : {
        fileAbsoluteUrl: '',
        fileName: '',
        fileNameWithoutExtension: '',
        downloadFileContent: async () => new File([], ''),
      };



    return {
      pages: [
        {
          header: { description: 'Manage Link' },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField(`links[${index}].Title`, {
                  label: 'Title',
                  value: this.properties.links[index].Title
                }),
                PropertyPaneTextField(`links[${index}].URL`, {
                  label: 'URL',
                  value: this.properties.links[index].URL
                }),

                PropertyPaneChoiceGroup(`links[${index}].ThumbnailOption`, {
                  label: 'Thumbnail',
                  options: thumbnailOptions,
                }),

                ...(selectedThumbnailOption === 'text-only' ? [
                  PropertyPaneTextField(`links[${index}].Title`, {
                    label: 'Text',
                    placeholder: 'Enter link text'
                  })
                ] : []),
                ...(selectedThumbnailOption === 'default' ? [
                  PropertyFieldIconPicker('defaultIconPicker', {
                    key: `defaultIconPicker-${index}`,
                    onSave: (name: string) => {
                      this.properties.links[index].Icon = name;
                      this.context.propertyPane.refresh();
                    },
                    currentIcon: this.properties.links[index].Icon || 'Globe',
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    buttonLabel: 'Select default icon',
                    label: 'Default Icon'
                  })
                ] : []),


                ...(selectedThumbnailOption === 'file' ? [PropertyFieldFilePicker('filePicker', {
                  context: this.context as any,
                  filePickerResult: filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (result: IFilePickerResult) => {
                    if (result?.fileAbsoluteUrl) {
                      this.properties.links[index].Icon = result.fileAbsoluteUrl;
                      this.context.propertyPane.refresh();
                      this.render();
                    }
                  },
                  onChanged: (result: IFilePickerResult) => {
                    if (result?.fileAbsoluteUrl) {
                      this.properties.links[index].Icon = result.fileAbsoluteUrl;
                      this.context.propertyPane.refresh();
                      this.render();
                    }
                  },
                  accepts: [
                    ".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"
                  ],
                  key: 'filePickerId',
                  buttonLabel: 'Select image',
                  label: 'Custom image'
                })] : []),

                ...(selectedThumbnailOption === 'icon' ? [PropertyFieldIconPicker('iconPicker', {
                  key: `iconPicker-${index}`,
                  onSave: (name: string) => {
                    this.properties.links[index].Icon = name;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  currentIcon: this.properties.links[index].Icon,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  buttonLabel: 'Select icon',
                  label: 'Icon'
                })] : []),
                PropertyPaneChoiceGroup(`links[${index}].Target`, {
                  label: "Open link in",
                  options: [
                    { key: "_self", text: "This tab" },
                    { key: "_blank", text: "New tab" }
                  ]
                }),



                PropertyFieldToggleWithCallout('enableDescription', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'toggleDescription',
                  label: 'Enable Description',
                  calloutContent: 'Show additional descriptive text for this link',
                  onText: 'Enabled',
                  offText: 'Disabled',
                  checked: this.properties.enableDescription
                }),
                ...(this.properties.enableDescription ? [
                  PropertyPaneTextField(`links[${index}].Description`, {
                    label: 'Description',
                    multiline: true,
                    rows: 3,
                    value: link.Description,
                    placeholder: 'Enter link description'
                  })
                ] : []),
                ...(this.properties.enableAudienceTargeting ? [
                  PropertyFieldPeoplePicker('TargetAudiences', {
                    label: 'Users or Groups Allowed to Add/Edit Events',
                    initialData: this.properties.links[index].TargetAudiences || [],
                    allowDuplicate: false,
                    principalType: [
                      PrincipalType.Users,
                      PrincipalType.SharePoint,
                      PrincipalType.Security
                    ],
                    context: this.context as any,
                    properties: this.properties,
                    onPropertyChange: (path, oldVal, newVal) => {
                      this.properties.links[index].TargetAudiences = (newVal || [])
                        .filter((p: any) => p && p.id)
                        .map((p: any) => ({
                          id: p.id,
                          email: p.email,
                          login: p.login,
                          fullName: p.fullName,
                        }));
                    },
                    key: `allowedGroupsFieldId-${index}`
                  }),
                ] : []),

                PropertyPaneButton('', {
                  text: 'Apply',
                  onClick: async () => {
                    const link = this.properties.links[index];
                    if (this.properties.enableAudienceTargeting) {
                      link.TargetAudiences = Array.isArray(link.TargetAudiences) ? link.TargetAudiences : [];
                    } else {
                      link.TargetAudiences = [];
                    }
                    try {
                      await this.quickLinkService.saveLink(link);
                      this.properties.links[index] = link;
                      this.properties.links = [...this.properties.links];
                      this.context.propertyPane.refresh();
                      this.render();
                    } catch (error) {
                      console.error('Error saving link:', error);
                    }
                  },
                }),
                PropertyPaneButton('', {
                  text: 'Delete',
                  onClick: async () => {
                    try {
                      await this.quickLinkService.deleteLink(this.properties.links[index].Id);
                      const updatedLinks = await this.quickLinkService.getLinks(this.instanceId);
                      this.properties.links = updatedLinks;
                      this.selectedItemId = null;
                      this.context.propertyPane.refresh();
                      this.render();
                    } catch (error) {
                      console.error('Error deleting link:', error);
                    }
                  },
                }),
                PropertyPaneButton('', {
                  text: 'Close',
                  onClick: () => {
                    this.selectedItemId = null;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }


  private getDefaultPropertyPane(): IPropertyPaneConfiguration {

    const currentLayoutType = this.properties.layoutType || LayoutType.ImageWithTitle;
    const containerSize = CONTAINER_SIZES[currentLayoutType] || { width: 50, height: 50 };
    const TileSize = TILE_SIZES[currentLayoutType] || { width: 50, height: 50 };


    if (!this.properties.layoutSettings) {
      this.properties.layoutSettings = {} as Record<LayoutType, ILayoutSettings>;
    }
    if (!this.properties.layoutSettings[currentLayoutType]) {
      this.properties.layoutSettings[currentLayoutType] = {
        hideText: false,
        iconSize: IconSize.Medium,
        imageSize: { width: 100, height: 100 },
      };
    }
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: 'Layout options',
              groupFields: [
                PropertyPaneChoiceGroup('layoutType', {
                  label: 'Select Layout',
                  options: [
                    { key: LayoutType.ImageWithTitle, text: 'ImageWithTitle', iconProps: { officeFabricIconFontName: 'FileImage' } },
                    { key: LayoutType.Compact, text: 'Compact', iconProps: { officeFabricIconFontName: 'BacklogList' } },
                    { key: LayoutType.FilmStrip, text: 'FilmStrip', iconProps: { officeFabricIconFontName: 'AspectRatio' } },
                    { key: LayoutType.Grid, text: 'Grid', iconProps: { officeFabricIconFontName: 'GridViewSmall' } },
                    { key: LayoutType.Button, text: 'Button', iconProps: { officeFabricIconFontName: 'Calculator' } },
                    { key: LayoutType.List, text: 'List', iconProps: { officeFabricIconFontName: 'List' } },
                    { key: LayoutType.Tiles, text: 'Tiles', iconProps: { officeFabricIconFontName: 'PictureTile' } }
                  ],
                })
              ]
            },

            {
              groupName: "SharePoint Configuration",
              groupFields: [
                PropertyPaneTextField('tenantUrl', {
                  label: 'Tenant URL',
                  value: this.context.pageContext.web.absoluteUrl
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name'
                }),
                PropertyPaneSlider('defaultBorderRadius', {
                  label: 'Default Corner Radius (px)',
                  min: 0,
                  max: 50,
                  value: this.properties.defaultBorderRadius,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneSlider('imageOpacity', {
                  label: 'Image Opacity (%)',
                  min: 0,
                  max: 100,
                  value: this.properties.imageOpacity,
                  showValue: true,
                  step: 1
                }),
                PropertyFieldColorPicker('color', {
                  label: 'Color',
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyPaneTextField('quickAccessHeader', {
                  label: 'Quick Access Header Text',
                  placeholder: 'Header',
                  value: this.properties.quickAccessHeader,
                }),
              //   PropertyPaneChoiceGroup('quickAccessHeaderAlignment', {
              //   label: 'Quick Access Header Alignment',
              //   options: [
              //     { key: 'left', text: 'Left', checked: this.properties.quickAccessHeaderAlignment === 'left' },
              //     { key: 'center', text: 'Center', checked: this.properties.quickAccessHeaderAlignment === 'center' },
              //     { key: 'right', text: 'Right', checked: this.properties.quickAccessHeaderAlignment === 'right' }
              //   ],
              // }),
                PropertyFieldColorPicker('quickAccessBackgroundColor', {
                  label: 'Quick Access Background Color',
                  selectedColor: this.properties.quickAccessBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'quickAccessBackgroundColorFieldId',
                }),
                PropertyFieldColorPicker('cardColor', {
                  label: 'Card Color',
                  selectedColor: this.properties.cardColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'cardColorId',
                }),

                PropertyPaneDropdown(`layoutSettings.${currentLayoutType}.iconSize`, {
                  label: 'Icon Size',
                  options: [
                    { key: IconSize.Small, text: `Small (${IconSize.Small}px)` },
                    { key: IconSize.Medium, text: `Medium (${IconSize.Medium}px)` },
                    { key: IconSize.Large, text: `Large (${IconSize.Large}px)` },
                  ],
                  selectedKey: this.properties.layoutSettings[currentLayoutType]?.iconSize || IconSize.Medium,
                }),

                PropertyFieldColorPicker('TitleColor', {
                  label: 'Title Color',
                  selectedColor: this.properties.TitleColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 200,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'titleColorField'
                }),
                PropertyPaneSlider('TitleFontSize', {
                  label: 'Title Font Size (px)',
                  min: 12,
                  max: 40,
                  value: this.properties.TitleFontSize || 32,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneToggle('showImageBackground', {
                  label: 'Show white background behind images'
                }),
                PropertyFieldColorPicker('BottonColor', {
                  label: 'Button Color',
                  selectedColor: this.properties.BottonColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 200,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'buttonColorField'
                }),
                PropertyFieldColorPicker('textColor', {
                  label: 'Text Color',
                  selectedColor: this.properties.textColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 200,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'textColorField'
                }),
                PropertyFieldColorPicker('iconColor', {
                  label: 'Icon Color',
                  selectedColor: this.properties.iconColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 200,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'iconColorField'

                }),

                // Image Width for the current layout
                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.imageSize.width`, {
                  label: 'Image Width (px)',
                  min: IMAGE_MIN_WIDTHS[currentLayoutType],
                  max: containerSize.width,
                  value: this.properties.layoutSettings[currentLayoutType]?.imageSize?.width || containerSize.width,
                  showValue: true,
                  step: 4,
                }),
                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.imageSize.height`, {
                  label: 'Image Height (px)',
                  min: IMAGE_MIN_HEIGHTS[currentLayoutType],
                  max: containerSize.height,
                  value: this.properties.layoutSettings[currentLayoutType]?.imageSize?.height || containerSize.height,
                  showValue: true,
                  step: 4,
                }),
                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.tileSize.tileWidth`, {
                  label: 'Tile Width (px)',
                  min: TILE_MIN_WIDTHS[currentLayoutType] || 150,
                  max: TileSize.tileWidth,
                  value: this.properties.layoutSettings[currentLayoutType]?.tileSize?.tileWidth || TileSize.tileWidth,
                  showValue: true,
                  step: 4,
                }),
                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.tileSize.tileHeight`, {
                  label: 'Tile Height (px)',
                  min: TILE_MIN_HEIGHTS[currentLayoutType] || 50,
                  max: TileSize.tileHeight,
                  value: this.properties.layoutSettings[currentLayoutType]?.tileSize?.tileHeight || TileSize.tileHeight,
                  showValue: true,
                  step: 4,
                }),
                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.tileSize.numColumns`, {
                  label: 'Number of Columns',
                  min: 1,
                  max: 5,
                  value: this.properties.layoutSettings[currentLayoutType]?.tileSize?.numColumns || 5,
                  showValue: true,
                  step: 1,
                }),

                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.containerWidth`, {
                  label: "Container Width (px)",
                  min: 200,
                  max: 1200,
                  value: this.properties.layoutSettings[currentLayoutType]?.containerWidth || 800,
                  showValue: true,
                  step: 10,
                }),
                PropertyPaneSlider(`layoutSettings.${currentLayoutType}.containerHeight`, {
                  label: "Container Height (px)",
                  min: 200,
                  max: 1000,
                  value: this.properties.layoutSettings[currentLayoutType]?.containerHeight || 400,
                  showValue: true,
                  step: 10,

                }),
                PropertyPaneChoiceGroup('filmstripSize', {
                  label: 'Filmstrip Size',
                  options: [
                    { key: 'small', text: 'Small', checked: this.properties.filmstripSize === 'small' },
                    { key: 'medium', text: 'Medium', checked: this.properties.filmstripSize === 'medium' },
                    { key: 'large', text: 'Large', checked: this.properties.filmstripSize === 'large' }
                  ],
                }),
                PropertyPaneChoiceGroup('cardAlignment', {
                  label: 'Grid Alignment',
                  options: [
                    { key: 'left', text: 'Left', checked: this.properties.cardAlignment === 'left' },
                    { key: 'center', text: 'Center', checked: this.properties.cardAlignment === 'center' },
                    { key: 'right', text: 'Right', checked: this.properties.cardAlignment === 'right' }
                  ],
                }),
                PropertyFieldColorPicker('cardHoverColor', {
                  label: 'Card Hover Color',
                  selectedColor: this.properties.cardHoverColor || '#f0f0f0',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 200,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'cardHoverColorField'
                }),


                PropertyFieldToggleWithCallout('enableAudienceTargeting', {
                  key: 'enableAudienceTargetingToggle',
                  label: 'Enable Audience Target',
                  calloutContent: React.createElement("div", {}, "Toggle this to enable the feature."),
                  calloutTrigger: CalloutTriggers.Click,
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.enableAudienceTargeting,
                  disabled: false,
                }),
              ]
            }
          ]
        }
      ]
    };
  }

}