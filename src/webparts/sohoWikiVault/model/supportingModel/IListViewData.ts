import { IIconsDetails, IWikiVault } from '../';

export interface IListViewData {
    // subCategory?: string;
    Title?: string;
    BriefLabelDescription?: string;
    URL?: IWikiVault;
    PageOwner?: string;
    Icon?: IIconsDetails;
    Edit?: IWikiVault;
    Delete?: string;
}
