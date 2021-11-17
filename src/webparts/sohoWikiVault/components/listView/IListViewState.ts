// import { IListViewData } from '../../model';
// Office UI Fabric controls
import { IColumn } from 'office-ui-fabric-react/';
import { IWikiVault } from '../../model';
export interface IListViewState {
    items: any;
    viewColumns: IColumn[];
    allItems: any;
    currentItemBeingEditedorViewed: IWikiVault;
    currentItemBeingDeleted: string;
    showWikiEditForm: boolean;
    showDeleteDialog: boolean;
    showWikiDetailsForm: boolean;
}