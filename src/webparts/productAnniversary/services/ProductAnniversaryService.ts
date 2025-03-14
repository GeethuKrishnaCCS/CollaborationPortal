import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";

export class ProductAnniversaryService extends BaseService {
    private spfi: SPFI;
    constructor(context: WebPartContext, siteUrl: string) {
        super(context, siteUrl);
        this.spfi = getSP(context);
    }

    public getCurrentUser() {
        return this.spfi.web.currentUser();
    }


    public async getSharePointLists(): Promise<any[]> {
        try {
            const lists = await this.spfi.web.lists.filter("BaseTemplate ne 101")();
            return lists;
        } catch (error) {
            console.error("Error fetching SharePoint lists:", error);
            return [];
        }
    }
    public async getListFields(lisname:string): Promise<any[]> {
        try {
            const lists = await this.spfi.web.lists.getByTitle(lisname).fields();
            const editableFields = lists.filter(field => !field.ReadOnlyField)
            return editableFields;
        } catch (error) {
            console.error("Error fetching SharePoint lists:", error);
            return [];
        }
    }
}