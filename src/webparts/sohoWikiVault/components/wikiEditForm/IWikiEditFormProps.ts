import { WebPartContext } from '@microsoft/sp-webpart-base';

//Model
import { IMainCategory, ISubCategory, IWikiVault } from '../../model';

export interface IWikiEditFormProps {
    mainCategoryItems: IMainCategory[];
    subCategoryItems: ISubCategory[];
    context: WebPartContext;
    wikiItem: IWikiVault;
    onCloseButtonClick?: any;
}