import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/search";
import "@pnp/sp/clientside-pages";
import "@pnp/sp/comments/clientside-page";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export class PnPConfig {
  private context: WebPartContext;
  private sp: SPFI; // Store default spfi instance

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(this.context)); // Initialize default spfi instance
  }

  public getSP(siteUrl?: string): SPFI {
    if (siteUrl) {
      return spfi(siteUrl).using(SPFx(this.context)); // Create new instance for specific site
    }
    return this.sp; // Return cached default instance
  }

  public getContext(): WebPartContext {
    return this.context;
  }
}