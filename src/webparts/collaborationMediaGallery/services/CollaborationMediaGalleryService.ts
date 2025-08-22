import { spService } from './sp';

export default class CollaborationMediaGalleryService {

    static async getAllItems(listId: string, numberOfItems: number): Promise<any[]> {
        let results: any[] = [];
        try {

            const web = spService.sp.web;
            results = await web.lists
                .getById(listId).items
                .select('*')
                .select('ID', 'Title', 'SortOrder', 'MediaDescription', 'File_x0020_Type', 'FileSystemObjectType', 'File/Name', 'File/ServerRelativeUrl', 'File/Title', 'File/Id', 'File/TimeLastModified')
                .top(numberOfItems)
                .expand('File')
                .filter((`File_x0020_Type eq  'mp4'`))
                .orderBy('SortOrder', true)
                ();
        } catch (error) {
            return Promise.reject(error);
        }
        return results;
    }
}