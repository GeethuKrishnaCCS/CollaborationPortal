import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CollaborationBannerWebPartStrings';
import CollaborationBanner from './components/CollaborationBanner';
import { ICollaborationBannerProps, ICollaborationBannerWebPartProps } from './interfaces/ICollaborationBannerProps';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls';


export default class CollaborationBannerWebPart extends BaseClientSideWebPart<ICollaborationBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICollaborationBannerProps> = React.createElement(
      CollaborationBanner,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        BannerImageUrl: this.properties.BannerImageUrl,
        context: this.context,
        welcomeText: this.properties.welcomeText,
        missionDescription: this.properties.missionDescription,
        selectDesign: this.properties.selectDesign,

        imageOverlapBackgroundColor: this.properties.imageOverlapBackgroundColor,
        welcomeheadingfontcolor: this.properties.welcomeheadingfontcolor,
        usernameheadingfontcolor: this.properties.usernameheadingfontcolor,
        descriptionfontcolor: this.properties.descriptionfontcolor,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }


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
                PropertyPaneTextField('BannerImageUrl', {
                  label: "Banner Image URL",
                }),
                PropertyPaneTextField('welcomeText', {
                  label: "welcome",
                }),
                PropertyPaneTextField('missionDescription', {
                  label: "Mission Description",
                }),
                PropertyPaneChoiceGroup('selectDesign', {
                  label: "Select Design",
                  options: [
                    { key: 'ImageOverlap', text: 'ImageOverlap' },
                    { key: 'New', text: 'New' },
                  ],
                }),

                ...(this.properties.selectDesign === 'ImageOverlap' ? [
                  PropertyFieldColorPicker('imageOverlapBackgroundColor', {
                    label: 'Background Color',
                    selectedColor: this.properties.imageOverlapBackgroundColor,
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
                ] : []),

              ]
            },

            // Heading
            {
              groupName: "Styles for Welcome Text",
              groupFields: [
                PropertyFieldColorPicker('welcomeheadingfontcolor', {
                  label: 'Font Color',
                  selectedColor: this.properties.welcomeheadingfontcolor,
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

              ]
            },
            {
              groupName: "Styles for UserName",
              groupFields: [
                PropertyFieldColorPicker('usernameheadingfontcolor', {
                  label: 'Font Color',
                  selectedColor: this.properties.usernameheadingfontcolor,
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

              ]
            },

            {
              groupName: "Styles for Description",
              groupFields: [
                PropertyFieldColorPicker('descriptionfontcolor', {
                  label: 'Font Color',
                  selectedColor: this.properties.descriptionfontcolor,
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

              ]
            },
          ]
        }
      ]
    };
  }
}
