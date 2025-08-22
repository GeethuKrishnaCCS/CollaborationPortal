import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  IPropertyFieldSite,
  PropertyFieldSitePicker,
} from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";
// import * as strings from 'CollaborationNewsWebPartStrings';
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle,
} from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { ICollaborationNewsProps } from './components/ICollaborationNewsProps';
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users";
import CollaborationNews from "./components/CollaborationNews";
import { CollaborationNewsService } from "./services/CollaborationNewsService";
import { CollaborationNewsModel } from "./models/CollaborationNewsModel";

export interface ICollaborationNewsWebPartProps {
  StyleToggle: "list" | "topstory" | "sidebyside" | "hubnews" | "carousel" | "tiles";
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
  LikeToggle: boolean;
  CommentToggle: boolean;
  sites: IPropertyFieldSite[];
  numberOfNews: number;
  selectedCategories: string[];
  headingColor: string;
  backgroundColor: string;
  customTitle: string;
  newsContainerColor: string;
  buttonBackgroundColor: string;
  textColor: string;
}

export default class CollaborationNewsWebPart extends BaseClientSideWebPart<ICollaborationNewsWebPartProps> {

  private _categories: string[] = [];
  private _currentUserEmail: string | undefined;
  private _newsService: CollaborationNewsService;
  private _canAddNews: boolean = false;
  private _initialNews: CollaborationNewsModel[] = [];
  private _initPromise: Promise<void> | null = null;
  private _isLoadingCategories: boolean = false;
  private _renderTimeout: number | null = null;
  handleAddNewsPost: ((siteUrl: string) => void) | undefined;
  handleAddNewsLink: ((siteUrl: string) => void) | undefined;

  protected async onInit(): Promise<void> {
    this._newsService = new CollaborationNewsService(this.context);
    this._currentUserEmail = this.context.pageContext.user.email?.toLowerCase();

    // Initialize default values if undefined
    if (this.properties.backgroundColor === undefined) {
      this.properties.backgroundColor = "#ffffff";
      console.log("onInit: Initialized backgroundColor to #ffffff");
    }
    if (this.properties.newsContainerColor === undefined) {
      this.properties.newsContainerColor = "#ffffff";
      console.log("onInit: Initialized newsContainerColor to #ffffff");
    }
    if (this.properties.buttonBackgroundColor === undefined) {
      this.properties.buttonBackgroundColor = "#0078d4";
      console.log("onInit: Initialized buttonBackgroundColor to #0078d4");
    }

    await this.loadCategories();
    try {
      // this._canAddNews = await this._newsService.isUserInSelectedGroups(this.properties.addNewsGroups || []);
      console.log("onInit: canAddNews:", this._canAddNews, "user:", this._currentUserEmail);

      const selectedSites = this.getSelectedSites();
      console.log("onInit: selectedSites:", selectedSites);

      if (selectedSites.length > 0) {
        const startTime = performance.now();
        this._initialNews = await this._newsService.getNewsFromSites(
          selectedSites,
          this.properties.selectedCategories || [],
          // this.properties.addNewsGroups || []
        );
        console.log(
          "onInit: Pre-fetched news:",
          this._initialNews.length,
          "items in",
          (performance.now() - startTime).toFixed(2),
          "ms",
          this._initialNews
        );
      } else {
        console.warn("onInit: No selected sites, skipping news fetch");
        this._initialNews = [];
      }
    } catch (error) {
      console.error("onInit: Error:", error);
      this._canAddNews = false;
      this._initialNews = [];
    }
  }
  private getSelectedSites(): string[] {
    return this.properties.sites
      ? this.properties.sites
        .map((site) => {
          let url = site.url;
          if (url && !url.startsWith("http")) url = `https://${url}`;
          return url ? url.replace(/\/$/, "") : "";
        })
        .filter((url): url is string => !!url)
      : [];
  }

  private async loadCategories(): Promise<void> {
    if (this._isLoadingCategories) {
      console.log("loadCategories: Already loading, skipping");
      return;
    }
    this._isLoadingCategories = true;
    this._categories = [];

    const siteUrl = "https://ccsdev01.sharepoint.com/sites/ProductDevelopment";
    console.log("Loading categories from site:", siteUrl);

    try {
      const cacheKey = `categories_${siteUrl}`;
      const cachedCategories = sessionStorage.getItem(cacheKey);
      if (cachedCategories) {
        this._categories = JSON.parse(cachedCategories);
        console.log("loadCategories: Loaded from cache:", this._categories);
      } else {
        const startTime = performance.now();
        const sp = spfi(siteUrl).using(SPFx(this.context));
        const list = sp.web.lists.getByTitle("Site Pages");
        const listInfo = await list.select("Title")().catch((err) => {
          throw new Error(`List 'Site Pages' not found: ${err.message}`);
        });
        console.log("Found list:", listInfo.Title);

        const field = await list.fields.getByInternalNameOrTitle("Pagecategory")().catch((err) => {
          throw new Error(`Field 'Pagecategory' not found: ${err.message}`);
        });
        console.log("Field details:", {
          Title: field.Title,
          InternalName: field.InternalName,
          TypeAsString: field.TypeAsString,
          Choices: field.Choices,
        });

        if (field.TypeAsString === "Choice" || field.TypeAsString === "MultiChoice") {
          this._categories = Array.from(new Set(field.Choices || []));
          if (this._categories.length === 0) {
            console.warn("No choices defined in the Pagecategory field.");
          }
          sessionStorage.setItem(cacheKey, JSON.stringify(this._categories));
        } else {
          console.warn(`Pagecategory field is not Choice or MultiChoice. Type: ${field.TypeAsString}`);
        }
        console.log(
          "loadCategories: Fetched in",
          (performance.now() - startTime).toFixed(2),
          "ms"
        );
      }
    } catch (error) {
      console.error(`Failed to load categories from ${siteUrl}:`, error);
      this._categories = [];
    }

    console.log("Loaded Categories:", this._categories);
    this._isLoadingCategories = false;
    this.context.propertyPane.refresh();
  }

  private clearCaches(): void {
    localStorage.clear();
    sessionStorage.clear();
    console.log("clearCaches: Cleared all storage");
  }

  public async render(): Promise<void> {
    if (this._renderTimeout) {
      clearTimeout(this._renderTimeout);
    }

    return new Promise((resolve) => {
      this._renderTimeout = setTimeout(async () => {
        if (!this._initPromise) {
          this._initPromise = this.onInit();
        }
        await this._initPromise;

        const currentUserEmail = this.context.pageContext.user.email?.toLowerCase();
        if (currentUserEmail !== this._currentUserEmail) {
          console.log("render: User changed, clearing caches");
          this._currentUserEmail = currentUserEmail;
          this.clearCaches();
          try {
            // this._canAddNews = await this._newsService.isUserInSelectedGroups(this.properties.addNewsGroups || []);
            console.log("render: canAddNews after user change:", this._canAddNews);
          } catch (error) {
            console.error("render: Error re-checking permissions:", error);
            this._canAddNews = false;
          }
        }

        const selectedSites = this.getSelectedSites();

        // Validate color properties before rendering
        console.log("render: Before validation:", {
          backgroundColor: this.properties.backgroundColor,
          newsContainerColor: this.properties.newsContainerColor,
        });

        if (!this.properties.backgroundColor) {
          this.properties.backgroundColor = "#ffffff";
          console.log("render: Set default backgroundColor to #ffffff");
        }
        if (!this.properties.newsContainerColor) {
          this.properties.newsContainerColor = "#ffffff";
          console.log("render: Set default newsContainerColor to #ffffff");
        }

        console.log("render: Props:", {
          canAddNews: this._canAddNews,
          selectedSites,
          selectedCategories: this.properties.selectedCategories,
          initialNewsCount: this._initialNews.length,
          headingColor: this.properties.headingColor,
          backgroundColor: this.properties.backgroundColor,
          newsContainerColor: this.properties.newsContainerColor,
          buttonBackgroundColor: this.properties.buttonBackgroundColor,
        });

        const newsProps: ICollaborationNewsProps = {
          StyleToggle: this.properties.StyleToggle || "list",
          AuthorToggle: this.properties.AuthorToggle ?? false,
          PublishedDateToggle: this.properties.PublishedDateToggle ?? false,
          LikeToggle: this.properties.LikeToggle ?? false,
          CommentToggle: this.properties.CommentToggle ?? false,
          selectedSites,
          selectedCategories: this.properties.selectedCategories || [],
          // selectedGroups: this.properties.addNewsGroups || [],
          numberOfNews: this.properties.numberOfNews || 5,
          context: this.context,
          placeholderImageUrl: "https://via.placeholder.com/150x150.png?text=No+Image+Available",
          headingColor: this.properties.headingColor,
          backgroundColor: this.properties.backgroundColor,
          newsContainerColor: this.properties.newsContainerColor,
          customTitle: this.properties.customTitle || "News",
           textColor: this.properties.textColor ,
          // canAddNews: this._canAddNews,

          initialNews: this._initialNews,
          isAdmin: false,
          handleAddNewsPost: this.handleAddNewsPost,
          handleAddNewsLink: this.handleAddNewsLink,
          buttonBackgroundColor: this.properties.buttonBackgroundColor,
        };

        const element: React.ReactElement<ICollaborationNewsProps> = React.createElement(CollaborationNews, newsProps);
        ReactDom.render(element, this.domElement);
        resolve();
      }, 100); // Debounce render by 100ms
    });
  }

  protected onDispose(): void {
    if (this._renderTimeout) {
      clearTimeout(this._renderTimeout);
    }
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const showToggles = !["carousel", "tiles"].includes(this.properties.StyleToggle || "list");
    const isCategoriesDisabled = this._categories.length === 0;

    return {
      pages: [
        {
          header: { description: "Property Pane Web Part Settings" },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField("customTitle", {
                  label: "Custom Title",
                  value: this.properties.customTitle || "News",
                }),
                PropertyPaneDropdown("StyleToggle", {
                  label: "Select News Layout",
                  options: [
                    { key: "list", text: "List" },
                    { key: "topstory", text: "Top Story" },
                    { key: "sidebyside", text: "Side by Side" },
                    { key: "hubnews", text: "Hub News" },
                    { key: "carousel", text: "Carousel" },
                    { key: "tiles", text: "Tiles" },
                  ],
                  selectedKey: this.properties.StyleToggle || "list",
                }),
                ...(showToggles
                  ? [
                    PropertyPaneToggle("AuthorToggle", {
                      label: "Show Author",
                      checked: this.properties.AuthorToggle ?? false,
                    }),
                    PropertyPaneToggle("PublishedDateToggle", {
                      label: "Show Published Date",
                      checked: this.properties.PublishedDateToggle ?? false,
                    }),
                    PropertyPaneToggle("LikeToggle", {
                      label: "Show Like Button",
                      checked: this.properties.LikeToggle ?? false,
                    }),
                    PropertyPaneToggle("CommentToggle", {
                      label: "Show Comment Section",
                      checked: this.properties.CommentToggle ?? false,
                    }),
                  ]
                  : []),
                PropertyFieldMultiSelect("selectedCategories", {
                  label: isCategoriesDisabled
                    ? "No categories available (ensure Pagecategory has choices in ProductDevelopment)"
                    : "Select News Categories",
                  key: "selectedCategoriesField",
                  options: this._categories.map((cat) => ({ key: cat, text: cat })),
                  selectedKeys: this.properties.selectedCategories || [],
                  disabled: isCategoriesDisabled,
                }),
                PropertyFieldSitePicker("sites", {
                  label: "Select Sites",
                  context: this.context as any,
                  initialSites: this.properties.sites,
                  onPropertyChange: async (
                    propertyPath: string,
                    oldValue: IPropertyFieldSite[],
                    newValue: IPropertyFieldSite[]
                  ): Promise<void> => {
                    console.log("Sites changed:", newValue);
                    this.properties.sites = newValue;
                    this.clearCaches();
                    this.properties.selectedCategories = [];
                    await this.loadCategories();
                    const selectedSites = this.getSelectedSites();
                    if (selectedSites.length > 0) {
                      const startTime = performance.now();
                      this._initialNews = await this._newsService.getNewsFromSites(
                        selectedSites,
                        this.properties.selectedCategories || [],
                        // this.properties.addNewsGroups || []
                      );
                      console.log(
                        "Sites changed: Fetched news:",
                        this._initialNews.length,
                        "items in",
                        (performance.now() - startTime).toFixed(2),
                        "ms"
                      );
                    } else {
                      this._initialNews = [];
                    }
                    this.context.propertyPane.refresh();
                    await this.render();
                  },
                  properties: this.properties,
                  multiSelect: true,
                  key: "sitePicker",
                }),
                // PropertyFieldPeoplePicker("addNewsGroups", {
                //   label: "Users/Groups Allowed to Add News",
                //   initialData: this.properties.addNewsGroups || [],
                //   allowDuplicate: false,
                //   principalType: [PrincipalType.Users, PrincipalType.Security, PrincipalType.SharePoint], // Added PrincipalType.User
                //   context: this.context,
                //   properties: this.properties,
                //   onPropertyChange: async (
                //     propertyPath: string,
                //     oldValue: IPropertyFieldGroupOrPerson[],
                //     newValue: IPropertyFieldGroupOrPerson[]
                //   ): Promise<void> => {
                //     console.log("addNewsGroups changed:", newValue);
                //     this.properties.addNewsGroups = newValue;
                //     this.clearCaches();
                //     try {
                //       this._canAddNews = await this._newsService.isUserInSelectedGroups(newValue);
                //       console.log("addNewsGroups: canAddNews:", this._canAddNews);
                //     } catch (error) {
                //       console.error("addNewsGroups: Error:", error);
                //       this._canAddNews = false;
                //     }
                //     const selectedSites = this.getSelectedSites();
                //     if (selectedSites.length > 0) {
                //       const startTime = performance.now();
                //       this._initialNews = await this._newsService.getNewsFromSites(
                //         selectedSites,
                //         this.properties.selectedCategories || [],
                //         this.properties.addNewsGroups || []
                //       );
                //       console.log(
                //         "addNewsGroups: Fetched news:",
                //         this._initialNews.length,
                //         "items in",
                //         (performance.now() - startTime).toFixed(2),
                //         "ms"
                //       );
                //     }
                //     this.context.propertyPane.refresh();
                //     await this.render();
                //   },
                //   key: "addNewsGroupsFieldId",
                // }),
                PropertyPaneSlider("numberOfNews", {
                  label: "Number of News Items to Display",
                  min: 1,
                  max: 25,
                  step: 1,
                  value: this.properties.numberOfNews || 5,
                }),
                PropertyFieldColorPicker("headingColor", {
                  label: "Heading Color",
                  selectedColor: this.properties.headingColor || "#000000",
                  onPropertyChange: (propertyPath, oldValue, newValue) => {
                    this.properties.headingColor = newValue as string;
                    console.log("headingColor changed:", newValue);
                    this.render();
                  },
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: "headingColorField",
                }),
                PropertyFieldColorPicker("backgroundColor", {
                  label: "Web Part Background Color",
                  selectedColor: this.properties.backgroundColor || "#ffffff",
                  onPropertyChange: (propertyPath, oldValue, newValue) => {
                    const newColor = newValue as string;
                    console.log("backgroundColor changed in Property Pane:", {
                      oldValue: this.properties.backgroundColor,
                      newValue: newColor,
                      newsContainerColorBefore: this.properties.newsContainerColor,
                    });
                    this.properties.backgroundColor = newColor;
                    // Validate newsContainerColor isn't affected
                    if (this.properties.newsContainerColor === undefined) {
                      this.properties.newsContainerColor = "#ffffff";
                      console.log("backgroundColor change: Reset newsContainerColor to #ffffff");
                    }
                    console.log("backgroundColor changed: After update:", {
                      backgroundColor: this.properties.backgroundColor,
                      newsContainerColor: this.properties.newsContainerColor,
                    });
                    this.render();
                  },
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: "backgroundColorField",
                }),
                PropertyFieldColorPicker("textColor", {
                  label: "Text Color",
                  selectedColor: this.properties.textColor || "#000000",
                  onPropertyChange: (propertyPath, oldValue, newValue) => {
                    const newColor = newValue as string;
                    this.properties.textColor = newColor;
                    this.render();
                  },
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: "textColorField"
                }),
                PropertyFieldColorPicker("newsContainerColor", {
                  label: "News Container Color",
                  selectedColor: this.properties.newsContainerColor || "#ffffff",
                  onPropertyChange: (propertyPath, oldValue, newValue) => {
                    const newColor = newValue as string;
                    console.log("newsContainerColor changed in Property Pane:", {
                      oldValue: this.properties.newsContainerColor,
                      newValue: newColor,
                      backgroundColorBefore: this.properties.backgroundColor,
                    });
                    this.properties.newsContainerColor = newColor;
                    console.log("newsContainerColor changed: After update:", {
                      backgroundColor: this.properties.backgroundColor,
                      newsContainerColor: this.properties.newsContainerColor,
                    });
                    this.render();
                  },
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: "newsContainerColorField",
                }),

                PropertyFieldColorPicker("buttonBackgroundColor", {
                  label: "Button Background Color",
                  selectedColor: this.properties.buttonBackgroundColor || "#0078d4",
                  onPropertyChange: (propertyPath, oldValue, newValue) => {
                    const newColor = newValue as string;
                    console.log("buttonBackgroundColor changed in Property Pane:", {
                      oldValue: this.properties.buttonBackgroundColor,
                      newValue: newColor,
                    });
                    this.properties.buttonBackgroundColor = newColor;
                    this.render();
                  },
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Full,
                  key: "buttonBackgroundColorField",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): Promise<void> {
    console.log(`onPropertyPaneFieldChanged: ${propertyPath}`, { oldValue, newValue });
    let shouldFetchNews = false;

    if (propertyPath === "StyleToggle") {
      this.properties.StyleToggle = newValue as ICollaborationNewsProps["StyleToggle"];
    } else if (propertyPath === "selectedCategories") {
      this.properties.selectedCategories = Array.isArray(newValue) ? (newValue as string[]) : [];
      console.log("selectedCategories changed:", this.properties.selectedCategories);
      this.clearCaches();
      shouldFetchNews = true;
    } else if (propertyPath === "LikeToggle") {
      this.properties.LikeToggle = newValue as boolean;
      this.clearCaches();
    } else if (propertyPath === "customTitle") {
      this.properties.customTitle = newValue as string;
    } else if (propertyPath === "numberOfNews") {
      this.properties.numberOfNews = newValue as number;
      shouldFetchNews = true;
    }

    if (shouldFetchNews) {
      const selectedSites = this.getSelectedSites();
      if (selectedSites.length > 0) {
        const startTime = performance.now();
        this._initialNews = await this._newsService.getNewsFromSites(
          selectedSites,
          this.properties.selectedCategories || [],
          // this.properties.addNewsGroups || []
        );
        console.log(
          "onPropertyPaneFieldChanged: Fetched news:",
          this._initialNews.length,
          "items in",
          (performance.now() - startTime).toFixed(2),
          "ms"
        );
      } else {
        this._initialNews = [];
      }
    }

    this.context.propertyPane.refresh();
    await this.render();
  }
}
