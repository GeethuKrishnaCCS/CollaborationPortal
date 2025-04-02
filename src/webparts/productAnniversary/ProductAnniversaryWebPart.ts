import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  // IPropertyPaneDropdownOption,

} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProductAnniversaryWebPartStrings';
import ProductAnniversary from './components/ProductAnniversary';
import { IProductAnniversaryProps, IProductAnniversaryWebPartProps } from './interfaces/IProductAnniversary';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
// import { ProductAnniversaryService } from './services/ProductAnniversaryService';

export interface IPropertyControlsTestWebPartProps {
  headingfontcolor: string;
  contentfontcolor: string;
  collectionData: any[];
}

export default class ProductAnniversaryWebPart extends BaseClientSideWebPart<IProductAnniversaryWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
    // private _service: ProductAnniversaryService;
 
  public render(): void {
    const element: React.ReactElement<IProductAnniversaryProps> = React.createElement(
      ProductAnniversary,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        Cardlayout: this.properties.Cardlayout,
        dataSource: this.properties.dataSource,
        Anniversary: this.properties.Anniversary,
        SharePointLists: this.properties.SharePointLists,
        AnniversayNoOfItemDisplay: this.properties.AnniversayNoOfItemDisplay,

        BdayToggleValue: this.properties.BdayToggleValue,
        WorkToggleValue: this.properties.WorkToggleValue,
        WeddingToggleValue: this.properties.WeddingToggleValue,
        bdayGreetingWish: this.properties.bdayGreetingWish,
        WorkGreetingWish: this.properties.WorkGreetingWish,
        weddingGreetingWish: this.properties.weddingGreetingWish,

        headingfontcolor: this.properties.headingfontcolor,
        contentfontcolor: this.properties.contentfontcolor,

        StylesForCards: this.properties.StylesForCards,
        
        ShapeForImages: this.properties.ShapeForImages,       
        collectionData: this.properties.collectionData,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  // private async getTimeZones(): Promise<Array<IPropertyPaneDropdownOption>> {
  //   const result: Array<IPropertyPaneDropdownOption> = [];
  //   try {
  //     const listFields = await this._service(this.properties.Anniversary); // Fetch fields from the specified list
  //     for (const field of listFields) {
  //       result.push({ key: field.InternalName, text: field.Title }); // Use InternalName as key and Title as text
  //     }
  //   } catch (error) {
  //     console.error("Error fetching list fields:", error);
  //   }
  //   return result;
  // }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),

                PropertyPaneDropdown('dataSource', {
                  label: "Select Data Source",
                  options: [
                    { key: 'SharepointList', text: 'Sharepoint List' },
                    { key: 'AzureDirectory', text: 'Azure Directory' },
                    { key: 'OtherExternalSources', text: 'Other External Sources' },
                  ],
                  selectedKey: 'SharepointList'
                }),

                ...(this.properties.dataSource === 'SharepointList'
                  ? [
                    PropertyPaneTextField('Anniversary', {
                      label: "Anniversary List Url"
                    }),
                  ]
                  : []),
              ]
            },

            {
              groupName: "Select Fields to Display",
              groupFields: [
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: "Location",
                      title: "Location",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                  
                    {
                      id: "SelectTimeZone",
                      title: "Select Time Zone",
                      type: CustomCollectionFieldType.dropdown,
                      // options: this.getTimeZones(),
                      required: true
                    },                    
                  ],
                  disabled: false
                })
              ]
            },

            {
              groupName: "Wishing",
              groupFields: [
                PropertyPaneToggle('BdayToggleValue', {
                  label: "Birthday",
                  checked: this.properties.BdayToggleValue
                }),
                PropertyPaneToggle('WorkToggleValue', {
                  label: "Work Anniversary",
                  checked: this.properties.WorkToggleValue
                }),

                PropertyPaneToggle('WeddingToggleValue', {
                  label: "Wedding Anniversary",
                  checked: this.properties.WeddingToggleValue
                }),

                PropertyPaneTextField('bdayGreetingWish', {
                  label: "Birthday Greeting"
                }),
                PropertyPaneTextField('WorkGreetingWish', {
                  label: "Work Anniversary Greeting"
                }),
                PropertyPaneTextField('weddingGreetingWish', {
                  label: "Wedding Anniversary Greeting"
                }),

                PropertyPaneChoiceGroup('Cardlayout', {
                  label: "Select Card Layout",
                  options: [
                    { key: 'Filmstrip', text: 'Filmstrip' },
                    { key: 'Cards', text: 'Cards' },
                    { key: 'Compact', text: 'Compact' },
                  ],
                }),

                ...(this.properties.Cardlayout === 'Cards' ? [
                  PropertyPaneChoiceGroup('StylesForCards', {
                    label: "Styles For Cards",
                    options: [
                      { key: 'Rectangle', text: 'Rectangle' },
                      { key: 'Square', text: 'Square' },
                      { key: 'Circle', text: 'Circle' },
                    ],
                  }),
                ] : []),


                PropertyPaneChoiceGroup('ShapeForImages', {
                  label: "Shape for Images",
                  options: [
                    { key: 'Rectangle', text: 'Rectangle' },
                    { key: 'Square', text: 'Square' },
                    { key: 'Circle', text: 'Circle' },
                  ],
                })
              ]
            },


            // Heading
            {
              groupName: "Styles for Headiing",
              groupFields: [
                PropertyFieldColorPicker('headingfontcolor', {
                  label: 'Font Color',
                  selectedColor: this.properties.headingfontcolor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'headingfontcolorFieldId'
                }),

                PropertyPaneDropdown('headingFontfamily', {
                  label: "Select Font Family",
                  options: [
                    { key: 'TimesNewRoman', text: 'Times New Roman' },
                    { key: 'AzureDirectory', text: 'Azure Directory' },
                    { key: 'OtherExternalSources', text: 'Other External Sources' },
                  ],
                  selectedKey: 'TimesNewRoman'
                }),

              ]
            },

            // Contents
            {
              groupName: "Styles for Contents",
              groupFields: [
                PropertyFieldColorPicker('contentfontcolor', {
                  label: 'Font Color',
                  selectedColor: this.properties.contentfontcolor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'contentfontcolorFieldId'
                }),
                PropertyPaneDropdown('contentFontfamily', {
                  label: "Select Content  Font Family",
                  options: [
                    { key: 'TimesNewRoman', text: 'Times New Roman' },
                    { key: 'AzureDirectory', text: 'Azure Directory' },
                    { key: 'OtherExternalSources', text: 'Other External Sources' },
                  ],
                  selectedKey: 'TimesNewRoman'
                }),
              ]
            },

          ],


        }
      ]
    };
  }
}


