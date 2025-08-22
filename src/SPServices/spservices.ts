import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import { SearchResults, ISearchQuery, SortDirection } from "@pnp/sp/search";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPServices } from "./ISPServices";

export class spservices implements ISPServices {
  private _sp: ReturnType<typeof spfi>;

  private readonly searchProperties: string[] = [
    "FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL",
    "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects",
    "BaseOfficeLocation", "SPS-UserType", "GroupId"
  ];

  constructor(private context: WebPartContext) {
    this._sp = spfi().using(SPFx(this.context));
  }

  public async searchUsers(searchString: string, searchFirstName: boolean): Promise<SearchResults> {
    try {
      const _search = searchFirstName
        ? `FirstName:${searchString}*`
        : `LastName:${searchString}*`;

      const users = await this._sp.search(<ISearchQuery>{
        Querytext: _search,
        RowLimit: 500,
        EnableInterleaving: true,
        SelectProperties: this.searchProperties,
        SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",
        SortList: [{ Property: "LastName", Direction: SortDirection.Ascending }]
      });

      return users;
    } catch (error: any) {
      throw new Error(error.message || "Error in searchUsers");
    }
  }

  public async searchUsers2(): Promise<SearchResults> {
    try {
      const users = await this._sp.search(<ISearchQuery>{
        Querytext: "*",
        RowLimit: 500,
        EnableInterleaving: true,
        SelectProperties: this.searchProperties,
        SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",
        SortList: [{ Property: "LastName", Direction: SortDirection.Ascending }]
      });

      return users;
    } catch (error: any) {
      throw new Error(error.message || "Error in searchUsers2");
    }
  }

  public async _getImageBase64(pictureUrl: string): Promise<string> {
    return new Promise((resolve) => {
      const image = new Image();
      image.crossOrigin = "anonymous";
      image.onload = () => {
        const tempCanvas = document.createElement("canvas");
        tempCanvas.width = image.width;
        tempCanvas.height = image.height;
        tempCanvas.getContext("2d")?.drawImage(image, 0, 0);
        try {
          const base64Str = tempCanvas.toDataURL("image/png");
          resolve(base64Str);
        } catch {
          resolve("");
        }
      };
      image.onerror = () => resolve("");
      image.src = pictureUrl;
    });
  }

  public async searchUsersNew(
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean
  ): Promise<SearchResults> {
    let qrytext = "";

    if (isInitialSearch) {
      qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
    } else {
      qrytext = srchQry || searchString || "*";
    }

    try {
      const users = await this._sp.search(<ISearchQuery>{
        Querytext: qrytext,
        RowLimit: 500,
        EnableInterleaving: true,
        SelectProperties: this.searchProperties,
        SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",
        SortList: [{ Property: "LastName", Direction: SortDirection.Ascending }]
      });

      if (users?.PrimarySearchResults?.length > 0) {
        const processedResults = users.PrimarySearchResults.map((user: any) => {
  if (user.WorkEmail) {
    user.PictureURL = `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}`;
  }
  return user;
});

 (users as any).ProcessedResults = processedResults;

      }

      return users;
    } catch (error: any) {
      throw new Error(error.message || "Error in searchUsersNew");
    }
  }
}