
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { BaseService } from "../../../common/services/BaseService";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { CollaborationQuickLinksModel } from "../models/CollaborationQuickLinksModel";

export class CollaborationQuickLinksService extends BaseService {
    private listName: string;
    private context: WebPartContext;


    constructor(context: WebPartContext, listName: string, siteUrl: string = context.pageContext.web.absoluteUrl) {
        super(context, siteUrl);
        this.listName = listName;
        this.context = context;

    }

    public async getLinks(webPartInstanceId: string): Promise<CollaborationQuickLinksModel[]> {
        try {
            const queryUrl = `${this.context.pageContext.web.serverRelativeUrl}/lists/${this.listName}`;
            const filter = `WebPartInstanceId eq '${webPartInstanceId}'`;
            const items: any[] = await this.getItemsSelectFilter(
                queryUrl,
                "ID,Title,URL,Icon,SortWeight,Description,TargetAudiences,Target,WebPartInstanceId",
                filter
            ).then(items => items.sort((a: any, b: any) => a.SortWeight - b.SortWeight));

            const allLinks: CollaborationQuickLinksModel[] = items.map(item => ({
                Id: item.ID.toString(),
                Title: item.Title,
                URL: item.URL,
                Icon: item.Icon,
                SortWeight: item.SortWeight,
                Description: item.Description,
                TargetAudiences: item.TargetAudiences ? JSON.parse(item.TargetAudiences) : [],
                Target: item.Target,
                WebPartInstanceId: item.WebPartInstanceId
            }));

            return allLinks;
        } catch (error) {
            console.error(`Error fetching Quick Links from ${this.listName}:`, error);
            return [];
        }
    }

    public async updateLinkOrder(updatedLinks: CollaborationQuickLinksModel[]): Promise<void> {
        try {
            const queryUrl = `${this.context.pageContext.web.serverRelativeUrl}/lists/${this.listName}`;
            for (const link of updatedLinks) {
                await this.updateItem(queryUrl, { SortWeight: link.SortWeight }, Number(link.Id));
            }
            console.log("Link order updated successfully");
        } catch (error) {
            console.error("Error updating link order:", error);
            throw error;
        }

    }

    public async addLink(link: CollaborationQuickLinksModel): Promise<CollaborationQuickLinksModel> {
        const queryUrl = `${this.context.pageContext.web.serverRelativeUrl}/lists/${this.listName}`;
        try {
            const listExists = await this.sp.web.lists.getByTitle(this.listName).select("Title")().catch(() => null);
            if (!listExists) {
                throw new Error(`List ${this.listName} does not exist or is inaccessible.`);
            }

            const lastItem = await this.getItemsSelectFilter(queryUrl, "SortWeight", "").then(items =>
                items.sort((a: any, b: any) => b.SortWeight - a.SortWeight)[0]
            );
            const newSortWeight = lastItem ? lastItem.SortWeight + 1 : 1;



            const newItem = await this.createNewItem(queryUrl, {
                Title: link.Title,
                URL: link.URL,
                Icon: link.Icon,
                SortWeight: newSortWeight,
                Description: link.Description,
                TargetAudiences: link.TargetAudiences ? JSON.stringify(link.TargetAudiences) : null,
                Target: link.Target,
                WebPartInstanceId: link.WebPartInstanceId,
            });

            console.log("createNewItem response:", newItem);

            const newId = newItem?.data?.ID || newItem?.data?.Id || newItem?.Id;
            if (!newId) {
                throw new Error("Item creation failed. No ID found in response: " + JSON.stringify(newItem));
            }

            console.log(`New link added with Id: ${newId} and SortWeight: ${newSortWeight}`);
            return { ...link, Id: newId.toString(), SortWeight: newSortWeight };
        } catch (error: any) {
            console.error("Error in addLink:", error?.message || error);
            throw error;
        }
    }

    public async updateLink(link: CollaborationQuickLinksModel): Promise<CollaborationQuickLinksModel> {
        if (!link.Id || isNaN(Number(link.Id))) {
            throw new Error("Invalid link ID for update.");
        }

        const queryUrl = `${this.context.pageContext.web.serverRelativeUrl}/lists/${this.listName}`;
        const itemId = Number(link.Id);

        try {

            await this.updateItem(queryUrl, {
                Title: link.Title,
                URL: link.URL,
                Icon: link.Icon,
                SortWeight: Number(link.SortWeight),
                Description: link.Description,
                TargetAudiences: link.TargetAudiences ? JSON.stringify(link.TargetAudiences) : null,
                Target: link.Target,
                WebPartInstanceId: link.WebPartInstanceId,
            }, itemId);

            console.log(`Updated link with Id: ${itemId}`);
            return { ...link, Id: itemId.toString() };
        } catch (error: any) {
            console.error("Error in updateLink:", error?.message || error);
            throw error;
        }
    }

    public async saveLink(link: CollaborationQuickLinksModel): Promise<CollaborationQuickLinksModel> {
        return link.Id && Number.isInteger(Number(link.Id))
            ? this.updateLink(link)
            : this.addLink(link);
    }

    public async deleteLink(linkId: string): Promise<void> {
        try {
            const itemId = Number(linkId);
            if (!itemId) throw new Error("Invalid item Id");

            const queryUrl = `${this.context.pageContext.web.serverRelativeUrl}/lists/${this.listName}`;
            await this.DeleteItem(queryUrl, itemId);
            console.log(`Item with Id: ${itemId} moved to Recycle Bin.`);
        } catch (error) {
            console.error(`Error deleting item with Id ${linkId}:`, error);
            throw error;
        }
    }
}