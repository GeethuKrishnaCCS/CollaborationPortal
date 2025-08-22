import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CollaborationMediaGalleryWebPartStrings';
import CollaborationMediaGallery from './components/CollaborationMediaGallery';
import { ICollaborationMediaGalleryProps } from './components/ICollaborationMediaGalleryProps';
import { spService } from './services/sp';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle, PropertyFieldNumber } from '@pnp/spfx-property-controls';

export interface ICollaborationMediaGalleryWebPartProps {
  list: string;
  numberOfItems: number;
  description: string;
  headingfontcolor: string;
}

export default class CollaborationMediaGalleryWebPart extends BaseClientSideWebPart<ICollaborationMediaGalleryWebPartProps> {
  private lists: IPropertyPaneDropdownOption[] = [];
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICollaborationMediaGalleryProps> = React.createElement(
      CollaborationMediaGallery,
      {
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        listId: this.properties.list,
        numberOfItems: this.properties.numberOfItems,
        description: this.properties.description,
         headingfontcolor: this.properties.headingfontcolor,
      }
    );

    ReactDom.render(element, this.domElement);
  }

   protected async onInit(): Promise<void> {
    const packageSolution = await require('../../../config/package-solution.json');
    console.log(`CCS-MediaGalleryWebPart: v.${packageSolution.solution.version}`);
    return this._getEnvironmentMessage().then(async message => {
      spService.initSP(this.context)
      const _lists = await this.loadLists()
      if (_lists.length > 0) {
        this.lists = _lists;
        // this.properties.list = this.lists[0].key.toString();
        this.context.propertyPane.refresh();
      }
      this._environmentMessage = message;
      return Promise.resolve()
    });
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
  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    const _lists: IPropertyPaneDropdownOption[] = [];
    try {
      const results = await spService.getSiteLists();
      for (const list of results) {
        _lists.push({ key: list.Id, text: list.Title });
      }
    } catch (error) {
      console.error("Error occurred: ", error);
      this.context.propertyPane.refresh();
    }
    return _lists;
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
                PropertyPaneDropdown('list', {
                  ariaDescription: `Selected Library must have "MediaDescription" Field`,
                  label: "Library",
                  options: this.lists,
                }),
                PropertyFieldNumber("numberOfItems", {
                  key: "numberOfItems",
                  label: "Number of items to load",
                  description: "Number between 1 and 50",
                  value: this.properties.numberOfItems,
                  maxValue: 50,
                  minValue: 1,
                  disabled: false
                }),
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
              ]
            },
          ]
        }
      ]
    };
  }
}
