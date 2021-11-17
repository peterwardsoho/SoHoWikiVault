import { IMainCategory, ISubCategory } from "../../model/index";

export interface IAdminCenterState {
    catError: string;
    catorderError: string;
    subCatError: string;
    subCatorderError: string;
    showDeleteDialog: boolean;
    dialogTitle: string;
    dialogSubTitle: string;
    currentMCBeingDeleted: IMainCategory;
    currentSCBeingDeleted: ISubCategory;
    disableDialogButtons: boolean;
}