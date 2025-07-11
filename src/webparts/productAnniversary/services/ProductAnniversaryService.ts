import { BaseService } from "../../../common/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../common/PnP/PnPConfig";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

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
    public async getListFields(lisname: string): Promise<any[]> {
        try {
            const lists = await this.spfi.web.lists.getByTitle(lisname).fields();
            const editableFields = lists.filter(field => !field.ReadOnlyField)
            return editableFields;
        } catch (error) {
            console.error("Error fetching SharePoint lists:", error);
            return [];
        }
    }
    public updateItem(url: string, data: any, id: number): Promise<any> {
        return this.spfi.web.getList(url).items.getById(id).update(data);
    }


    public async toggleLike(url: string, itemId: number, userEmail: string, type: 'Birthday' | 'Wedding Anniversary' | 'Work Anniversary', isLiked: boolean): Promise<void> {
        try {
            // Fetch the current item to get existing likedBy and likesCount
            const item = await this.spfi.web.getList(url).items.getById(itemId).select(
                'BirthdayLikedBy', 'BirthdayLikesCount',
                'WeddingAnniversaryLikedBy', 'WeddingAnniversaryLikesCount',
                'WorkAnniversaryLikedBy', 'WorkAnniversaryLikesCount'
            )();

            let likedByField: string;
            let likesCountField: string;
            let currentLikedBy: string[] = [];
            let currentLikesCount: number = 0;

            // Determine which fields to update based on the type
            switch (type) {
                case 'Birthday':
                    likedByField = 'BirthdayLikedBy';
                    likesCountField = 'BirthdayLikesCount';
                    currentLikedBy = item.BirthdayLikedBy ? item.BirthdayLikedBy.split(';').filter((email: string) => email) : [];
                    currentLikesCount = item.BirthdayLikesCount || 0;
                    break;
                case 'Wedding Anniversary':
                    likedByField = 'WeddingAnniversaryLikedBy';
                    likesCountField = 'WeddingAnniversaryLikesCount';
                    currentLikedBy = item.WeddingAnniversaryLikedBy ? item.WeddingAnniversaryLikedBy.split(';').filter((email: string) => email) : [];
                    currentLikesCount = item.WeddingAnniversaryLikesCount || 0;
                    break;
                case 'Work Anniversary':
                    likedByField = 'WorkAnniversaryLikedBy';
                    likesCountField = 'WorkAnniversaryLikesCount';
                    currentLikedBy = item.WorkAnniversaryLikedBy ? item.WorkAnniversaryLikedBy.split(';').filter((email: string) => email) : [];
                    currentLikesCount = item.WorkAnniversaryLikesCount || 0;
                    break;
                default:
                    throw new Error('Invalid anniversary type');
            }

            let updatedLikedBy: string[];
            let updatedLikesCount: number;

            if (isLiked) {
                // Unlike: Remove user email and decrease count
                updatedLikedBy = currentLikedBy.filter(email => email !== userEmail);
                updatedLikesCount = Math.max(0, currentLikesCount - 1);
            } else {
                // Like: Add user email and increase count
                updatedLikedBy = [...currentLikedBy, userEmail];
                updatedLikesCount = currentLikesCount + 1;
            }

            // Prepare update payload
            const updateData = {
                [likedByField]: updatedLikedBy.join(';'),
                [likesCountField]: updatedLikesCount
            };

            // Update the item
            await this.updateItem(url, updateData, itemId);
        } catch (error) {
            console.error(`Error toggling like for ${type}:`, error);
            throw error;
        }
    }



    public async getComments(url: string, itemId: number, type: 'Birthday' | 'Wedding Anniversary' | 'Work Anniversary'): Promise<string[]> {
        try {
            let commentField: string;
            switch (type) {
                case 'Birthday':
                    commentField = 'BirthdayComments';
                    break;
                case 'Wedding Anniversary':
                    commentField = 'WeddingAnniversaryComments';
                    break;
                case 'Work Anniversary':
                    commentField = 'WorkAnniversaryComments';
                    break;
                default:
                    throw new Error('Invalid anniversary type');
            }

            // Fetch item versions to get appended comments
            const versions = await this.spfi.web.getList(url).items.getById(itemId).versions();
            const comments: string[] = [];

            // Iterate through versions to collect comments
            for (const version of versions) {
                if (version[commentField]) {
                    comments.push(version[commentField]);
                }
            }

            // Return comments in chronological order (oldest first)
            return comments;
        } catch (error) {
            console.error(`Error fetching comments for ${type}:`, error);
            return [];
        }
    }

    public async addComment(url: string, itemId: number, type: 'Birthday' | 'Wedding Anniversary' | 'Work Anniversary', comment: string): Promise<void> {
        try {
            let commentField: string;
            switch (type) {
                case 'Birthday':
                    commentField = 'BirthdayComments';
                    break;
                case 'Wedding Anniversary':
                    commentField = 'WeddingAnniversaryComments';
                    break;
                case 'Work Anniversary':
                    commentField = 'WorkAnniversaryComments';
                    break;
                default:
                    throw new Error('Invalid anniversary type');
            }

            // Prepare update payload
            const updateData = {
                [commentField]: comment
            };

            // Update the item to append the new comment
            await this.updateItem(url, updateData, itemId);
        } catch (error) {
            console.error(`Error adding comment for ${type}:`, error);
            throw error;
        }
    }

    // Add this method to fetch user by email
    public async getUserByEmail(email: string): Promise<any> {
        return await this.spfi.web.siteUsers.getByEmail(email)();
    }
}