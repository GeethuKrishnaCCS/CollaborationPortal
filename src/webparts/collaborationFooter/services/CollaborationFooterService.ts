import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from '../../../common/PnP/PnPConfig';

export class CollaborationFooterService {
    private _sp: SPFI;
    constructor(context: WebPartContext) {
        this._sp = getSP(context);
    }
    public getCurrentUser() {
        return this._sp.web.currentUser();
    }
    public getListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items();
    }
    public getAdminListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select('Admin/Title,Admin/ID,Admin/EMail').expand('Admin')();
    }
    public createNewItem(url: string, listname: string, data: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.add(data);
    }    
    public updateItem(url: string, listname: string, data: any, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
}
