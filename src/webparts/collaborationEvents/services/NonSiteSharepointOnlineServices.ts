/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, SPFx, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import "@pnp/sp/lists";
import "@pnp/sp/items";

export class NonSiteSharepointOnlineServices {

    private _sp: SPFI;

    constructor(context: any, siteNme: string) {

        //this._sp = spfi().using(SPFx(context));

        const _tenanturl = context.pageContext.web.absoluteUrl.split('/sites')[0];

        this._sp = spfi(_tenanturl + "/sites/" + siteNme).using(SPFx({
            pageContext: {
                web: {
                    absoluteUrl: _tenanturl + "/sites/" + siteNme
                },
                legacyPageContext: {
                    formDigestTimeoutSeconds: 30,
                    formDigestValue: ''
                }
            }
        }));
    }

    // Get list title from list internal name
    public async getListTitle(listInternalName: string): Promise<any> {

        const result = await this._sp.web.lists.filter(`EntityTypeName eq '${listInternalName}List'`).select("Title")();

        if (result.length > 0)
            return result[0].Title;
        else
            return "";
    }
    public async getitemsfromList(listName: string, listSelect: string[], filter: string): Promise<any[]> {

        let spItems = this._sp.web.lists.getByTitle(listName).items;

        if (listSelect.length > 0) {
            spItems = spItems.select(...listSelect);
        }

        if (filter.length > 0) {
            spItems = spItems.filter(filter);
        }

        const result = await spItems.top(5000)();
        return result;
    }

    // Get all list items from admin list
    public async getListItems(listName: string, listSelect: string[], filter: string, expand: string): Promise<any[]> {

        let spItems = this._sp.web.lists.getByTitle(listName).items;

        if (listSelect.length > 0) {
            spItems = spItems.select(...listSelect);
        }
        if (expand.length > 0) {
            spItems = spItems.expand(expand.toString());
        }

        if (filter.length > 0) {
            spItems = spItems.filter(filter);
        }

        const result = await spItems.orderBy('Id', false).top(5000)();
        return result;
    }

}