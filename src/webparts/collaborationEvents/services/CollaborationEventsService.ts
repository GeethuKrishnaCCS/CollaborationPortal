import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";
//import { ISPFxArchitectureService } from "../interfaces";
import { ICollaborationEventsService } from "../interfaces/ICollaborationEventsService";

export class CollaborationEventsService extends BaseService implements ICollaborationEventsService {
    private spfi: SPFI;
    constructor(context: WebPartContext, siteUrl?: string) {
        super(context, siteUrl);
        this.spfi = getSP(context);
    }

    public getDocumentIndexItems(url: string, listname: string, id: number, fields: string, expand: string): Promise<any> {
        return this.spfi.web.getList(url + "/Lists/" + listname).items.select(fields).expand(expand).filter("ID eq " + id)();
    }

}