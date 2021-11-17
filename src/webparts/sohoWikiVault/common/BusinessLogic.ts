import { IIconsDetails, IListDetails } from '../model';
// Constant
import { Constants } from '../common';

export class BusinessLogic {
    public setIconsforWiki = (subCategory: IListDetails[], expirationDate: string, constants: Constants, isSubscriptionPaid: boolean, title: string, isServiceBeingUsed: boolean) => {
        console.log("I am here");
        let iconDetails: IIconsDetails = { isExpired: false, isSubscription: false, isMembership: false, isServiceBeingUsed: false };
        if (isSubscriptionPaid) {
            iconDetails.isSubscription = true;
        }
        // if (subCategory.indexOf(constants.subCategory.membership) > -1) {
        //     iconDetails.isMembership = true;
        // }
        if (expirationDate) {
            let dateToBeCheckOut = new Date(expirationDate);
            let today = new Date();
            if (isSubscriptionPaid && dateToBeCheckOut < today) {
                iconDetails.isExpired = true;
            }
        }
        if (isServiceBeingUsed != null && !isServiceBeingUsed) {
            iconDetails.isServiceBeingUsed = true;
        }
        return iconDetails;
    }
    public getDescription = (description: string) => {
        let continuedot: string = '';
        let briefDesc = description;
        if (description && description.length > 0) {
            if (briefDesc.length > 200) {
                continuedot = ' ...';
            }
            briefDesc = briefDesc.replace(/<style([\s\S]*?)<\/style>/gi, '');
            briefDesc = briefDesc.replace(/<script([\s\S]*?)<\/script>/gi, '');
            briefDesc = briefDesc.replace(/<\/div>/ig, '\n');
            briefDesc = briefDesc.replace(/<\/li>/ig, '\n');
            briefDesc = briefDesc.replace(/<li>/ig, '  *  ');
            briefDesc = briefDesc.replace(/<\/ul>/ig, '\n');
            briefDesc = briefDesc.replace(/<\/p>/ig, '\n');
            briefDesc = briefDesc.replace(/<br\s*[\/]?>/gi, "\n");
            briefDesc = briefDesc.replace(/<[^>]+>/ig, '');
            briefDesc = briefDesc.replace(/<[^>]+>/ig, '');
            briefDesc = briefDesc.replace(/<[^>]+>/ig, '');
            briefDesc = briefDesc.replace(/&#160;/gi, '');
            briefDesc = briefDesc.replace(/&#58;/gi, '');
            briefDesc = briefDesc.substring(0, 200);
            briefDesc = briefDesc + continuedot;
        }
        return briefDesc;
    }
    public checkValidURL = (str: string): boolean => {
        let res = str.match(/(http(s)?:\/\/.)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/g);
        return (res !== null);
    }
}