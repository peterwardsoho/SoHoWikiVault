import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IMainCategory, ISubCategory } from '../../model';

export interface IAdminCenterProps {
    mainCategoryItems: IMainCategory[];
    subCategoryItems: ISubCategory[];
    adminGroupID: string;
    adminControlDisabled: boolean;
    context: WebPartContext;
    // main Category Functions
    onEditMCButtonClick: any;
    onSaveMCButtonClick: any;
    onCancelMCButtonClick: any;
    onAddMCButtonClick: any;
    onMCTextBoxChange: any;
    onShowHideMCButtonClick: any;
    onDeleteMCButtonClick: any;
    // SubCategory Functions
    onEditSCButtonClick: any;
    onSaveSCButtonClick: any;
    onCancelSCButtonClick: any;
    onAddSCButtonClick: any;
    onSCTextBoxChange: any;
    onShowHideSCButtonClick: any;
    onDeleteSCButtonClick: any;
}