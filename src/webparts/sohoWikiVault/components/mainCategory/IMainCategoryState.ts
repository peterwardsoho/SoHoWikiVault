import { IMainCategory, ISubCategory } from '../../model';

export interface IMainCategoryState {
    isAdmin: boolean; //if the user belongs to the admin group
    searchText: string; // search text for searching in the wiki List
    mainCategoryItems: IMainCategory[]; // List of Category.
    subCategoryItems: ISubCategory[]; // List of subcategory. 
    mainCategory: IMainCategory[]; // List used for tab display. It include 'All' Categoy
    subCategory: ISubCategory[]; // List used for tab display. It include 'All' Categoy
    adminControlDisabled: boolean;  // when an action is taking place it disables all the other button on the form
    adminGroupId: string;
}