import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IMainCategory, ISubCategory, IWikiVault } from '../../model';

export interface IListViewProps {
    mainCategory: IMainCategory;
    subCategory: ISubCategory;
    context: WebPartContext;
    searchText: string;
    // For Wiki Edit Form
    mainCategoryItems: IMainCategory[];
    subCategoryItems: ISubCategory[];

    isAdmin: boolean; //if the user belongs to the admin group
    wikiId: string;
}