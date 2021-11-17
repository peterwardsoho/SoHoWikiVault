// SharePoint Http Call
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// DataProvider
import { IWikiEditFormDataProvider, CommonDataProvider } from '../';
// Constants
import { Constants } from '../../common';
//Modal
import { IDropdownValues, IWikiVault, IMainCategory, IWikiPasswordCreateItem, IWikiVaultPassword } from '../../model';
//Compnents
import { IWikiEditFormState } from '../../components';
export class WikiEditFormDataProvider implements IWikiEditFormDataProvider {
    private CommonDataProvider: CommonDataProvider;
    private context: WebPartContext;
    private constant: Constants;
    constructor(context: WebPartContext) { this.context = context; this.constant = new Constants(); this.CommonDataProvider = new CommonDataProvider(this.context); }

    public getFieldDDValue(): Promise<IDropdownValues[]> {
        let promise: Promise<IDropdownValues[]> = new Promise<IDropdownValues[]>((resolve, reject) => {
            let query = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiVault.listName}')/fields?$filter=EntityPropertyName eq '${this.constant.wikiVault.dropdownColumn}'`;

            this.context.spHttpClient.get(
                query,
                SPHttpClient.configurations.v1,
                this.constant.spHttpOptions.getNoMetadata,
            )
                .then((response: SPHttpClientResponse): Promise<{ value: IDropdownValues[] }> => {
                    return response.json();
                })
                .then((response: { value: IDropdownValues[] }) => {
                    resolve(response.value);
                })
                .catch((error: any) => {
                    reject(error);
                });
        });
        return promise;
    }
    public onSaveClick = async (wikiState: IWikiEditFormState, mainCategory: IMainCategory): Promise<string> => {
        const spEntityType = await this.CommonDataProvider.getItemEntityType(this.constant.wikiVault.listName);
        // create item to create
        let newListItem: any = this.setWikiItem(wikiState);
        // add SP-required metadata
        newListItem['@odata.type'] = spEntityType;
        // build request
        let requestDetails: any = this.constant.spHttpOptions.postNoMetadata;
        requestDetails.body = JSON.stringify(newListItem);
        // create the item
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiVault.listName}')/items`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
        const newWikiItem: IWikiVault = await response.json();
        const wikiPasswordId = await this.createWikiPasswordItem(wikiState, newWikiItem.Id, mainCategory);
        const id: string = newWikiItem.Id;
        return id;
    }
    public onUpdateClick = async (wikiState: IWikiEditFormState, mainCategory: IMainCategory): Promise<string> => {
        let filter = `&$filter=Id eq ${wikiState.wikiId}`;
        const wikiVault: IWikiVault[] = await this.CommonDataProvider.getListItems(this.constant.wikiVault.listName, this.constant.wikiVault.selectQuery, '', filter, this.constant.wikiVault.expand, '');
        // Update Wiki Vault List Item
        let updateListItem: any = this.setWikiItem(wikiState);
        let requestDetails: any = this.constant.spHttpOptions.updateNoMetadata;
        requestDetails.body = JSON.stringify(updateListItem);
        const updateERequestLink = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiVault.listName}')/items(${wikiState.wikiId})`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
        // Get WikiVault Password List Item
        filter = `&$filter=SohoWikiVaultId eq ${wikiState.wikiId}`;
        const wikiPassword: IWikiVaultPassword[] = await this.CommonDataProvider.getListItems(this.constant.WikiVaultPassword.listName, this.constant.WikiVaultPassword.selectQuery, '', filter, '', '');
        // If wiki main category is changed
        if (wikiVault[0].MainCategory.Title != mainCategory.Title) {
            const wikiPasswordId = await this.createWikiPasswordItem(wikiState, wikiState.wikiId, mainCategory);
            this.CommonDataProvider.removeListItem(wikiPassword[0].Id, this.constant.WikiVaultPassword.listName);
        } else {
            this.updateWikiPassword(wikiState, wikiPassword[0].Id);
        }
        return '';
    }
    private updateWikiPassword = async (wikiState: IWikiEditFormState, Id: string) => {
        let updateListItem: any = { Title: wikiState.wikiLabelName, UserId: wikiState.userId, SitePassword: wikiState.password, RestrictedLabelDescription: wikiState.restrictedLabelDesc };
        let requestDetails: any = this.constant.spHttpOptions.updateNoMetadata;
        requestDetails.body = JSON.stringify(updateListItem);
        const updateERequestLink = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/items(${Id})`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    private setWikiItem = (wikiState: IWikiEditFormState): any => {
        let pageOwnerId: number[] = [];
        if (wikiState.pageOwners && wikiState.pageOwners.length > 0) {
            wikiState.pageOwners.forEach(owner => {
                pageOwnerId.push(owner.id);
            });
        }
        const wikiListItem: any = {
            Title: wikiState.wikiLabelName,
            MainCategoryId: +wikiState.selectedMainCategory,
            BriefLabelDescription: wikiState.briefLabelDesc,
            URL: {
                Description: 'URL',
                Url: wikiState.url
            },
            PageOwnerId: pageOwnerId,
            CanThePasswordBeCutAndPasted: wikiState.canThePasswordBeCutAndPasted ? true : false,
            IsSubscriptionPaid: wikiState.isSubscriptionPaid ? true : false,
            SubscriptionExpirationDate: wikiState.subscriptionExpirationDate ? wikiState.subscriptionExpirationDate : null,
            SubscriptionType: wikiState.selectedSubscriptionType,
            SubCategoryId: wikiState.selectedSubCategory,
            IsServiceBeingUsed: wikiState.istheServiceBeingUsed ? true : false,
        };
        return wikiListItem;
    }
    private createWikiPasswordItem = async (wikiState: IWikiEditFormState, wikiId: string, mainCategory: IMainCategory) => {
        const requestURL = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/AddValidateUpdateItemUsingPath`;
        const folderName = `${mainCategory.CategoryFolderName.split('_')[0]}_${mainCategory.CategoryFolderName.split('_')[1]}_${mainCategory.CategoryFolderName.split('_')[2]}`;
        const spEntityType = await this.CommonDataProvider.getItemEntityType(this.constant.WikiVaultPassword.listName);
        // create item to create
        let newListItem: any = {
            listItemCreateInfo: {
                FolderPath: {
                    DecodedUrl: `${this.context.pageContext.web.absoluteUrl}/lists/${this.constant.WikiVaultPassword.listInternalName}/${folderName}`
                },
                UnderlyingObjectType: 0
            },
            formValues: [
                {
                    "FieldName": "Title",
                    "FieldValue": wikiState.wikiLabelName
                },
                {
                    "FieldName": "UserId",
                    "FieldValue": wikiState.userId
                },
                {
                    "FieldName": "SitePassword",
                    "FieldValue": wikiState.password
                },
                {
                    "FieldName": "RestrictedLabelDescription",
                    "FieldValue": wikiState.restrictedLabelDesc
                },
                {
                    "FieldName": "SohoWikiVaultId",
                    "FieldValue": `${wikiId}`
                }
            ],
            bNewDocumentUpdate: false
        };
        // build request
        let requestDetails: any = this.constant.spHttpOptions.postNoMetadata;
        requestDetails.body = JSON.stringify(newListItem);
        // create the item
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(requestURL,
            SPHttpClient.configurations.v1,
            requestDetails
        );
        const newWikiPasswordItem: IWikiPasswordCreateItem = await response.json();
        let wikiPasswordId: string = '';
        if (newWikiPasswordItem && newWikiPasswordItem.value.length > 0) {
            newWikiPasswordItem.value.forEach(element => {
                if (element.FieldName == "Id") {
                    wikiPasswordId = element.FieldValue;
                }
            });
        }
        return wikiPasswordId;
    }
}



// private setPermission = async (itemId: string, mainCategory: IMainCategory) => {
    //     const groupID = mainCategory.CategoryGroupName.Url.split('=')[1];
    //     let requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/getItemById(${itemId})/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
    //     // build request
    //     let requestDetails: any = {};
    //     // break Permission on the item
    //     const responseBreakPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
    //         SPHttpClient.configurations.v1,
    //         requestDetails
    //     );
    //     // set permission on the item
    //     requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/getItemById(${itemId})/roleassignments/addroleassignment(principalid=${groupID}, roleDefId=1073741826)`;
    //     // break Permission on the item
    //     const responseGivePermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
    //         SPHttpClient.configurations.v1,
    //         requestDetails
    //     );
    // }

    // if (newWikiPasswordItem && newWikiPasswordItem.Id) {
    //     this.setPermission(newWikiPasswordItem.Id, mainCategory);
    // }