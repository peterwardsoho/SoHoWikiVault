import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWikiVault } from '../../model';

export interface IViewWikiDetailsProps {
    context: WebPartContext;
    wikiItem: IWikiVault;
    onCloseButtonClick?: any;
}
