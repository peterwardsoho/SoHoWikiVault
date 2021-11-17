export interface ICommonDataProvider {
    getListItems(listName: string, selectQuery: string, top?: string, filter?: string, expand?: string, orderby?: string): Promise<any>;
    getItemEntityType(listName: string): Promise<string>;
    removeListItem(Id: string, listName: string): Promise<any>;
    getOwnerGroupDetails(): Promise<any>;
}