export interface ICollaborationEventsService {
    getCurrentUser(): Promise<any>;
    getDocumentIndexItems(url: string, listname: string, id: number, fields: string, expand: string): Promise<any>;
}