import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";
//import { ISPFxArchitectureService } from "../interfaces";
import { ICollaborationEventsService } from "../interfaces/ICollaborationEventsService";
import { IItem } from "@pnp/sp/items";

export class CollaborationEventsService extends BaseService implements ICollaborationEventsService {
    private spfi: SPFI;
    constructor(context: WebPartContext, siteUrl: string) {
        super(context, siteUrl);
        this.spfi = getSP(context);
    }

    public getDocumentIndexItems(url: string, listname: string, id: number, fields: string, expand: string): Promise<any> {
        return this.spfi.web.getList(url + "/Lists/" + listname).items.select(fields).expand(expand).filter("ID eq " + id)();
    }

    public getChoiceListItems(url: string, field: string): Promise<any> {
        return this.spfi.web.getList(url).fields.getByInternalNameOrTitle(field)();
    }

    public addItemRequestForm(data: any, listname: string, url: string): Promise<any> {
        return this.spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    }

    public getListItems(url: string): Promise<any> {
        return this.spfi.web.getList(url).items();
    }

    public async addAttachments(listname: string, id: number, fileName: string, attachment: any): Promise<any> {
        const item: IItem = await this.spfi.web.lists.getByTitle(listname).items.getById(id)
        return await item.attachmentFiles.add(fileName, attachment)
    }

    public async updateItemRequestForm(data: any, listname: string, id: number, url: string): Promise<any> {
        return this.spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
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
}