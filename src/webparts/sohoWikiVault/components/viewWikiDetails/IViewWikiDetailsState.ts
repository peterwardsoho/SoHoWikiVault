import { IWikiVaultPassword } from "../../model";

export interface IViewWikiDetailsState {
    wikiPassword: IWikiVaultPassword[];
    needPerission: string;
    subCategory: string;
    pageOwner: string;
    userID: string;
    sitePassword: string;
    restrictedLabel: string;
    userNameCopied: string;
    passwordCopied: string;
    urlCopied: string;
    showBriefDesc: boolean;
    showRestDesc: boolean;
}