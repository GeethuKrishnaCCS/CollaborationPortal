import { spfi, SPFI } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { Birthday } from "../types";

export class BirthdayService {
    private sp: SPFI;

    constructor() {
        this.sp = spfi("https://ccsdev01.sharepoint.com/sites/VrizeTemplate2/");
    }

    public async fetchBirthdays(): Promise<Birthday[]> {
        try {
            const items = await this.sp.web.lists
                .getByTitle("Birthday")
                .items();

            return items.map((item: { Title: string; BirthdayDate: string }) => ({
                name: item.Title,
                date: new Date(item.BirthdayDate), // Assuming BirthdayDate is ISO formatted
            }));
        } catch (error) {
            console.error("Error fetching birthdays:", error);
            return [];
        }
    }
}