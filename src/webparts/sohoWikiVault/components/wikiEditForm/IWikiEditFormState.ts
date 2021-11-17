import { IDropdownOption } from 'office-ui-fabric-react';
import { IPeoplePicker } from '../../model';
export interface IWikiEditFormState {
    controlDisabled: boolean;
    showNew: boolean;
    wikiMainCategoryChoices: IDropdownOption[];
    wikiSubCategoryChoices: IDropdownOption[];
    wikiSubscriptionTypeChoices: IDropdownOption[];
    wikiLabelName: string;
    selectedMainCategory: string;
    selectedSubscriptionType: string;
    selectedSubCategory: string[];
    briefLabelDesc: string;
    subscriptionExpirationDate: Date;
    isSubscriptionPaid: boolean;
    userId: string;
    password: string;
    canThePasswordBeCutAndPasted: boolean;
    restrictedLabelDesc: string;
    url: string;
    istheServiceBeingUsed: boolean;
    pageOwners: IPeoplePicker[];
    defaultSelectedPageOwners: string[];
    wikiId: string;
    error: string;
    labelError: string;
    userIDError: string;
    sitePasswordError: string;
}