import { WebPartContext } from "@microsoft/sp-webpart-base";
// import pnp and pnp logging system
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items/get-all";
import "@pnp/sp/presets/all";

class SPService {
    private _sp: SPFI;
    private _siteUrl: string;

    public initSP = (context: WebPartContext) => {
        if (context !== null) {
            this._sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
            this._siteUrl = context.pageContext.site.serverRelativeUrl;
        } else {
            throw new Error("Context not defined")
        }
    };
    get sp() {
        return this._sp;
    }

    public async getSiteLists(): Promise<any[]> {
        let results: any[] = [];

        if (!this._siteUrl) {
            return [];
        }

        try {
            results = await this._sp.web.lists
                .select("*", "Title", "Id")
                .filter('BaseTemplate eq 101 and Hidden ne true')
                ();
        } catch (error) {
            console.error("Error fetching lists:", error);
            return Promise.reject(error);
        }

        return results;
    }
}

export const spService = new SPService()