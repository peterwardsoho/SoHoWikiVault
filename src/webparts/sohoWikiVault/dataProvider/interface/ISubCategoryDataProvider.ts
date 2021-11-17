import { ISubCategory } from "../../model";

export interface ISubCategoryDataProvider {
    saveSubCategoryListItem(subCategoryItem: ISubCategory): Promise<any>;
    updateSubCategory(updatedListItem: any, id: any): Promise<any>;
    removeSubCategoryListItem(subCategoryItem: ISubCategory): Promise<any>;
}