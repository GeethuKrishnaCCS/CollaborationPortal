// import * as React from 'react';
// import * as ReactDom from 'react-dom';
// import { Version } from '@microsoft/sp-core-library';
// import {
//   type IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-property-pane';
// import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';

// import * as strings from 'CollaborationCalendarWebPartStrings';
// import CollaborationCalendar from './components/CollaborationCalendar';
// import { ICollaborationCalendarProps } from './components/ICollaborationCalendarProps';

// export interface ICollaborationCalendarWebPartProps {
//   description: string;
// }

// export default class CollaborationCalendarWebPart extends BaseClientSideWebPart<ICollaborationCalendarWebPartProps> {

//   private _isDarkTheme: boolean = false;
//   private _environmentMessage: string = '';

//   public render(): void {
//     const element: React.ReactElement<ICollaborationCalendarProps> = React.createElement(
//       CollaborationCalendar,
//       {
//         description: this.properties.description,
//         isDarkTheme: this._isDarkTheme,
//         environmentMessage: this._environmentMessage,
//         hasTeamsContext: !!this.context.sdks.microsoftTeams,
//         userDisplayName: this.context.pageContext.user.displayName
//       }
//     );

//     ReactDom.render(element, this.domElement);
//   }

//   protected onInit(): Promise<void> {
//     return this._getEnvironmentMessage().then(message => {
//       this._environmentMessage = message;
//     });
//   }



//   private _getEnvironmentMessage(): Promise<string> {
//     if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
//       return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
//         .then(context => {
//           let environmentMessage: string = '';
//           switch (context.app.host.name) {
//             case 'Office': // running in Office
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
//               break;
//             case 'Outlook': // running in Outlook
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
//               break;
//             case 'Teams': // running in Teams
//             case 'TeamsModern':
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
//               break;
//             default:
//               environmentMessage = strings.UnknownEnvironment;
//           }

//           return environmentMessage;
//         });
//     }

//     return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
//   }

//   protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
//     if (!currentTheme) {
//       return;
//     }

//     this._isDarkTheme = !!currentTheme.isInverted;
//     const {
//       semanticColors
//     } = currentTheme;

//     if (semanticColors) {
//       this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
//       this.domElement.style.setProperty('--link', semanticColors.link || null);
//       this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
//     }

//   }

//   protected onDispose(): void {
//     ReactDom.unmountComponentAtNode(this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: strings.PropertyPaneDescription
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('description', {
//                   label: strings.DescriptionFieldLabel
//                 })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }



import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

import { setupPnP, getSP } from './pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { ICollaborationCalendarProps } from './components/ICollaborationCalendarProps';
import CollaborationCalendar from './components/CollaborationCalendar';

export interface ICollaborationCalendarWebPartProps {
  title: string;
  siteUrl: string;
  listName: string;
  showOutlookEvents: boolean;
  showSpEventsOnly: boolean;
  outlookEventColor: string;
  spEventColor: string;
  heading: string;
  backgroundColor: string;
  calendarBackgroundColor: string;
  allowedGroups: any[];
  applyGroupsToggle: boolean;
  selectedDate: IDateTimeFieldValue;
  context: WebPartContext;
}

export default class CollaborationCalendarWebPart extends BaseClientSideWebPart<ICollaborationCalendarWebPartProps> {
  private listOptions: IPropertyPaneDropdownOption[] = [];
  private tempAllowedGroups: any[] = [];

  public async onInit(): Promise<void> {
    setupPnP(this.context);
    this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
    this.properties.allowedGroups = this.properties.allowedGroups || [];
    this.properties.applyGroupsToggle = false;
    this.tempAllowedGroups = [...this.properties.allowedGroups];
    await this.loadLists();
    return super.onInit();
  }

  private async loadLists(): Promise<void> {
    try {
      const sp: SPFI = getSP();
      const lists = await sp.web.lists
        .select('Title', 'BaseTemplate', 'Hidden')
        .filter(`(BaseTemplate eq 106 or BaseTemplate eq 100) and Hidden eq false`)();

      this.listOptions = lists.map((list: { Title: string }) => ({
        key: list.Title,
        text: list.Title
      }));

      if (this.listOptions.length === 0) {
        this.listOptions = [{ key: '', text: 'No Lists Found' }];
      }

      if (this.context.propertyPane.isPropertyPaneOpen()) {
        this.context.propertyPane.refresh();
      } else {
        this.context.propertyPane.open();
      }
    } catch (error) {
      console.error('Error fetching SharePoint lists:', error);
      this.listOptions = [{ key: '', text: 'Error fetching lists' }];
    }
  }

  public render(): void {
    const element: React.ReactElement<ICollaborationCalendarProps> = React.createElement(
      CollaborationCalendar,
      {
        title: this.properties.title,
        siteUrl: this.properties.siteUrl,
        listName: this.properties.listName,
        context: this.context,
        showOutlookEvents: this.properties.showOutlookEvents,
        showSpEventsOnly: this.properties.showSpEventsOnly,
        outlookEventColor: this.properties.outlookEventColor,
        spEventColor: this.properties.spEventColor,
        calendarBackgroundColor: this.properties.calendarBackgroundColor,
        allowedGroups: this.properties.allowedGroups,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'allowedGroups') {
      this.tempAllowedGroups = newValue || [];
    } else if (propertyPath === 'applyGroupsToggle' && newValue === true) {
      this.properties.allowedGroups = [...this.tempAllowedGroups];
      this.context.propertyPane.refresh();

      setTimeout(() => {
        this.properties.applyGroupsToggle = false;
        this.context.propertyPane.refresh();
        this.render();
      }, 1500);
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  public onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Web Part Settings' },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyFieldColorPicker('calendarBackgroundColor', {
                  label: 'Calendar Background Color',
                  selectedColor: this.properties.calendarBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  key: 'calendarBackgroundColorField',
                  style: PropertyFieldColorPickerStyle.Full
                }),
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL (Auto-fetched)',
                  disabled: true
                }),
                PropertyPaneDropdown('listName', {
                  label: 'Select a Calendar List',
                  options: this.listOptions.length ? this.listOptions : [{ key: '', text: 'Loading...' }],
                  selectedKey: this.properties.listName
                }),
                PropertyPaneToggle('showOutlookEvents', {
                  label: 'Show Outlook Calendar Events?',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.showOutlookEvents
                }),
                PropertyPaneToggle('showSpEventsOnly', {
                  label: 'Show SharePoint Events Only?',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.showSpEventsOnly
                }),
                PropertyFieldColorPicker('outlookEventColor', {
                  label: 'Outlook Event Color',
                  selectedColor: this.properties.outlookEventColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  key: 'outlookColorField',
                  style: PropertyFieldColorPickerStyle.Full
                }),
                PropertyFieldColorPicker('spEventColor', {
                  label: 'SharePoint Event Color',
                  selectedColor: this.properties.spEventColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  key: 'spColorField',
                  style: PropertyFieldColorPickerStyle.Full
                }),
                PropertyFieldPeoplePicker('allowedGroups', {
                  label: 'Users or Groups Allowed to Add/Edit Events',
                  initialData: this.properties.allowedGroups || [],
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  context: this.context as any,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  key: 'allowedGroupsFieldId'
                }),
                PropertyPaneToggle('applyGroupsToggle', {
                  label: 'Apply Selected Users/Groups',
                  onText: 'Apply',
                  offText: 'Select Users/Groups',
                  checked: this.properties.applyGroupsToggle,
                  key: 'applyGroupsToggle'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}