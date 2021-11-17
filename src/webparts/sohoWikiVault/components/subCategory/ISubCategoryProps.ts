
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IMainCategory, ISubCategory } from '../../model';

export interface ISubCategoryProps {
    context: WebPartContext;
    mainCategory: IMainCategory;
    searchText: string;
    subCategory: ISubCategory[];
    // for Editing Purpose
    mainCategoryItems: IMainCategory[];
    subCategoryItems: ISubCategory[];
    isAdmin: boolean; //if the user belongs to the admin group
    wikiId: string;
}