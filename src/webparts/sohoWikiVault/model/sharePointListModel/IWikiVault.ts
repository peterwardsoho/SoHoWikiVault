export interface IWikiVault {
    Id?: string;
    Title?: string;
    BriefLabelDescription?: string;
    MainCategory?: {
        Id: string;
        Title: string;
    };
    URL?: {
        Description: string,
        Url: string
    };
    PageOwner: {
        EMail: string;
        FirstName: string;
        LastName: string;
        Name: string;
        Id: string;
    }[];
    SubscriptionExpirationDate: string;
    SubCategory: {
        Id: string;
        Title: string;
    }[];
    IsSubscriptionPaid: boolean;
    IsServiceBeingUsed: boolean;
    SubscriptionType: string;
    CanThePasswordBeCutAndPasted: boolean;
    Modified?: string;
    Editor?: {
        FirstName: string;
        LastName: string;
    };
}
