import { IMainCategory } from "../../model";

export interface IMainCategoryDataProvider {
    saveCategoryListItem(mainCategoryItem: IMainCategory): Promise<any>;
    updateMainCategory(updatedListItem: any, id: any): Promise<any>;
    isUserWikiAdmin(groupName: string): Promise<any>;
    removeCategoryListItem(mainCategoryItem: IMainCategory): Promise<any>;
}