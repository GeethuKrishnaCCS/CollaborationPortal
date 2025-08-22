
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneToggle,
    PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CollaborationDirectoryWebPartStrings';
import CollaborationDirectory from './components/CollaborationDirectory';
import { ICollaborationDirectoryProps } from './components/ICollaborationDirectoryProps';

export interface ICollaborationDirectoryWebPartProps {
   title: string;
    searchFirstName: boolean;
    searchProps: string;
    clearTextSearchProps: string;
    pageSize: number;
    justifycontent: boolean;
    
    showUserPhoto: boolean;
    showUserDept: boolean;
    showUserJobTitle: boolean;
    showUserPhone: boolean;
    showUserEmail: boolean;
    showUserLocation: boolean;

    hideUsersWithoutDept: boolean;
    hideUsersWithoutJobTitle: boolean;
    hideUsersWithoutPhone: boolean;
    hideUsersWithoutEmail: boolean;
    hideUsersWithoutLocation: boolean;

    refiners: string;
    filterDepartment: string;
}

export default class CollaborationDirectoryWebPart extends BaseClientSideWebPart<ICollaborationDirectoryWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ICollaborationDirectoryProps> = React.createElement(
      CollaborationDirectory,
      {
        title: this.properties.title,
                context: this.context,
                searchFirstName: this.properties.searchFirstName,
                displayMode: this.displayMode,
                updateProperty: (value: string) => {
                    this.properties.title = value;
                },
                searchProps: this.properties.searchProps,
                clearTextSearchProps: this.properties.clearTextSearchProps,
                pageSize: this.properties.pageSize,
                useSpaceBetween: this.properties.justifycontent,
                cardSettings: {
                    showUserPhoto: this.properties.showUserPhoto,
                    showUserDept: this.properties.showUserDept,
                    showUserJobTitle: this.properties.showUserJobTitle,
                    showUserPhone: this.properties.showUserPhone,
                    showUserEmail: this.properties.showUserEmail,
                    showUserLocation: this.properties.showUserLocation
                },
                filterSettings: {
                    hideUsersWithoutDept: this.properties.hideUsersWithoutDept,
                    hideUsersWithoutJobTitle: this.properties.hideUsersWithoutJobTitle,
                    hideUsersWithoutPhone: this.properties.hideUsersWithoutPhone,
                    hideUsersWithoutEmail: this.properties.hideUsersWithoutEmail,
                    hideUsersWithoutLocation: this.properties.hideUsersWithoutLocation,
                    refiners: this.properties.refiners,
                    filterDepartment: this.properties.filterDepartment,
                }
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
    protected get disableReactivePropertyChanges(): boolean {
        return true;
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
                                PropertyPaneTextField("title", {
                                    label: strings.TitleFieldLabel
                                }),
                                PropertyPaneToggle("searchFirstName", {
                                    checked: false,
                                    label: "Search on First Name ?"
                                }),
                                PropertyPaneToggle("justifycontent", {
                                    checked: false,
                                    label: "Result Layout",
                                    onText: "SpaceBetween",
                                    offText: "Center"
                                }),
                                PropertyPaneTextField('searchProps', {
                                    label: strings.SearchPropsLabel,
                                    description: strings.SearchPropsDesc,
                                    value: this.properties.searchProps,
                                    multiline: false,
                                    resizable: false
                                }),
                                PropertyPaneTextField('clearTextSearchProps', {
                                    label: strings.ClearTextSearchPropsLabel,
                                    description: strings.ClearTextSearchPropsDesc,
                                    value: this.properties.clearTextSearchProps,
                                    multiline: false,
                                    resizable: false
                                }),
                                PropertyPaneSlider('pageSize', {
                                    label: 'Results per page',
                                    showValue: true,
                                    max: 20,
                                    min: 1,
                                    step: 2,
                                    value: this.properties.pageSize
                                })
                            ]
                        },
                        {
                            groupName: "Show/Hide Details",
                            isCollapsed: false,
                            groupFields: [
                                PropertyPaneToggle("showUserPhoto", {
                                    checked: false,
                                    label: "Show User Photo?"
                                }),
                                PropertyPaneToggle("showUserDept", {
                                    checked: false,
                                    label: "Show Department ?"
                                }),
                                PropertyPaneToggle("showUserJobTitle", {
                                    checked: false,
                                    label: "Show Job Title ?"
                                }),
                                PropertyPaneToggle("showUserPhone", {
                                    checked: false,
                                    label: "Show Phone ?"
                                }),
                                PropertyPaneToggle("showUserEmail", {
                                    checked: false,
                                    label: "Show Email ?"
                                }),
                                PropertyPaneToggle("showUserLocation", {
                                    checked: false,
                                    label: "Show Location ?"
                                }),
                            ]
                        },
                        {
                            groupName: "Filter",
                            groupFields: [
                                PropertyPaneToggle("hideUsersWithoutDept", {
                                    checked: false,
                                    label: "Hide users without department set?"
                                }),
                                PropertyPaneToggle("hideUsersWithoutJobTitle", {
                                    checked: false,
                                    label: "Hide users without job title set?"
                                }),
                                PropertyPaneToggle("hideUsersWithoutPhone", {
                                    checked: false,
                                    label: "Hide users without phone set?"
                                }),
                                PropertyPaneToggle("hideUsersWithoutEmail", {
                                    checked: false,
                                    label: "Hide Users without email set?"
                                }),
                                PropertyPaneToggle("hideUsersWithoutLocation", {
                                    checked: false,
                                    label: "Hide users without office location set?"
                                }),
                                PropertyPaneTextField('refiners', {
                                    label: strings.SearchPropsLabel,
                                    // description: strings.SearchPropsDesc,
                                    value: this.properties.refiners,
                                    multiline: false,
                                    resizable: false
                                }),
                                PropertyPaneTextField('filterDepartment', {
                                label: "Filter by Department",
                                description: "Enter department name to filter results",
                                value: this.properties.filterDepartment
                            }),
                                
                            ]
                        },
                    ]
                }
            ]
        };
  }
}
