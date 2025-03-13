import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneTextField,

} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CollaborationEventsWebPartStrings';
import CollaborationEventsModel from './models/CollaborationEventsModel';
import { ICollaborationEventsModelProps, ICollaborationEventsWebPartProps } from './interfaces/ICollaborationEvents';
import { CollaborationEventsService } from './services/CollaborationEventsService';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';



export default class CollaborationEventsWebPart extends BaseClientSideWebPart<ICollaborationEventsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _service: CollaborationEventsService;



  public async onInit(): Promise<void> {
    this._service = new CollaborationEventsService(this.context, this.context.pageContext.web.absoluteUrl);
    const lists = await this._service.getSharePointLists();
    this.properties.SharePointLists = lists.map(list => ({ key: list.Title, text: list.Title }));

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


  public render(): void {
    const element: React.ReactElement<ICollaborationEventsModelProps> = React.createElement(
      CollaborationEventsModel,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        context: this.context,

        Cardlayout: this.properties.Cardlayout,
        dataSource: this.properties.dataSource,
        StylesForCards: this.properties.StylesForCards,
        StylesForImages: this.properties.StylesForImages,
        Events: this.properties.Events,
        Field: this.properties.Field,
        SharePointLists: this.properties.SharePointLists,

        location: this.properties.location,
        headingfontcolor: this.properties.headingfontcolor,


      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
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
                    PropertyPaneDropdown('Events', {
                      label: "Select List",
                      options: this.properties.SharePointLists || [],
                      selectedKey: this.properties.Events,
                    })
                  ]
                  : []),

                PropertyPaneDropdown('Field', {
                  label: "Select Field",
                  options: [
                    { key: 'Title', text: 'Title' },
                    { key: 'Category', text: 'Category' },
                    { key: 'StartDate', text: 'StartDate' }
                  ],
                  selectedKey: this.properties.Field,
                }),









                PropertyPaneDropdown('location', {
                  label: "Select Location",
                  options: [
                    { key: 'Kochi', text: 'Kochi' },
                    { key: 'Kottayam', text: 'Kottayam' }
                  ],
                  selectedKey: this.properties.location,
                }),

                PropertyPaneChoiceGroup('Cardlayout', {
                  label: "Select Card Layout",
                  options: [
                    { key: 'Filmstrip', text: 'Filmstrip' },
                    { key: 'Cards', text: 'Cards' },
                    { key: 'List', text: 'List' },
                    { key: 'Compact', text: 'Compact' },
                  ],
                }),

                PropertyPaneChoiceGroup('StylesForImages', {
                  label: "Styles for Images",
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

          ]
        }
      ]
    };
  }
}
