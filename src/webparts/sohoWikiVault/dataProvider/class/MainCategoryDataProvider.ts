// SharePoint Http Call
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// DataProvider
import { IMainCategoryDataProvider, CommonDataProvider } from '../';

// Constants
import { Constants } from '../../common';
import { IListDetails, IMainCategory, IWikiVault, IWikiVaultPassword } from '../../model';

export class MainCategoryDataProvider implements IMainCategoryDataProvider {
    private context: WebPartContext;
    private constant: Constants;
    private CommonDataProvider: CommonDataProvider;
    constructor(context: WebPartContext) { this.context = context; this.constant = new Constants(); this.CommonDataProvider = new CommonDataProvider(this.context); }

    /*
    Saves and updates Main category Item. 
    **/
    public saveCategoryListItem = async (mainCategoryItem: IMainCategory): Promise<any> => {
        if (mainCategoryItem.Id == "-1") {
            const spEntityType = await this.getItemEntityType(this.constant.wikiMainCategory.listName);
            // set value for item to create
            let newListItem: any = { Title: mainCategoryItem.Title.trim(), CategoryOrder: +mainCategoryItem.CategoryOrder.trim(), Status: this.constant.comparingStrings.visible };
            // add SP-required metadata
            newListItem['@odata.type'] = spEntityType;
            // build request
            let requestDetails: any = this.constant.spHttpOptions.postNoMetadata;
            requestDetails.body = JSON.stringify(newListItem);
            // create the item
            const response: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiMainCategory.listName}')/items`,
                SPHttpClient.configurations.v1,
                requestDetails
            );
            const newMainCat: IMainCategory = await response.json();
            // create SPGroup for the Category and create a Folder in the password list for the category and set Permission on the password List
            const groupName = `SohoWiki_${newMainCat.Title}_${newMainCat.Id}`;
            const ownerGroupId = await this.CommonDataProvider.getOwnerGroupDetails();
            const categoryGroupId = await this.CommonDataProvider.createSPGroup(groupName, ownerGroupId);
            const categoryFolderName = await this.createFolderAndSetPermissions(groupName, categoryGroupId, ownerGroupId);

            await this.updateMainCategory(
                {
                    CategoryGroupName: {
                        Description: groupName,
                        Url: `${this.context.pageContext.web.absoluteUrl}${this.constant.URLs.spGroupURL}${categoryGroupId}`
                    },
                    CategoryFolderName: categoryFolderName
                },
                newMainCat.Id);

            mainCategoryItem.Id = newMainCat.Id;
            mainCategoryItem.CategoryGroupName = { Description: groupName, Url: `${this.context.pageContext.web.absoluteUrl}${this.constant.URLs.spGroupURL}${categoryGroupId}` };
            mainCategoryItem.CategoryFolderName = categoryFolderName;
            return mainCategoryItem;
        } else {
            const listItem = { Title: mainCategoryItem.Title.trim(), CategoryOrder: +mainCategoryItem.CategoryOrder.toString().trim() };
            await this.updateMainCategory(listItem, mainCategoryItem.Id);
            return null;
        }
    }
    private createFolderAndSetPermissions = async (groupName: string, categoryGroupId: string, ownerGroupId: string) => {
        const spEntityType = await this.getItemEntityType(this.constant.WikiVaultPassword.listName);
        let requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/items`;
        // FileSystemObjectType: 1, FileLeafRef: groupName, DisplayName: groupName 
        let newListFolderDetails: any = { Title: groupName, ContentTypeId: "0x0120" };
        // add SP-required metadata
        newListFolderDetails['@odata.type'] = spEntityType;
        // build request
        let request: any = this.constant.spHttpOptions.postNoMetadata;
        request.body = JSON.stringify(newListFolderDetails);
        const createFolder: SPHttpClientResponse = await this.context.spHttpClient.post(
            requestUrl,
            SPHttpClient.configurations.v1,
            request
        );
        const folderDetails: IListDetails = await createFolder.json();
        // Updating FieldRef Of the Folder
        requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/items('${folderDetails.Id}')`;
        let updateListFolderDetails = { FileLeafRef: groupName };
        let updateRequestDetails: any = this.constant.spHttpOptions.updateNoMetadata;
        // let updatedListItem: any = { Title: mainCategoryItem.Title, CategoryOrder: mainCategoryItem.CategoryOrder };
        updateRequestDetails.body = JSON.stringify(updateListFolderDetails);
        const updateERequestLink = await this.context.spHttpClient.post(
            requestUrl,
            SPHttpClient.configurations.v1,
            updateRequestDetails
        );
        // setting Permission
        await this.setFolderPermission(folderDetails.Id, categoryGroupId, ownerGroupId);
        return `${folderDetails.Title}_${folderDetails.Id}`;
    }
    private setFolderPermission = async (folderId: string, categoryGroupId: string, ownerGroupId: string) => {
        let requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/getItemById(${folderId})/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
        // build request
        let requestDetails: any = {};
        // break Permission on the item
        const responseBreakPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );
        // set view permisson for Category SP Group
        requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/getItemById(${folderId})/roleassignments/addroleassignment(principalid=${categoryGroupId}, roleDefId=1073741826)`;
        // break Permission on the item
        const categoryPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );
        // set view permisson for wiki owner SP Group
        requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.WikiVaultPassword.listName}')/getItemById(${folderId})/roleassignments/addroleassignment(principalid=${ownerGroupId}, roleDefId=1073741829)`;
        // break Permission on the item
        const ownerPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    /*
    Updates main category list item
    **/
    public updateMainCategory = async (updatedListItem: any, id: any) => {
        let requestDetails: any = this.constant.spHttpOptions.updateNoMetadata;
        requestDetails.body = JSON.stringify(updatedListItem);
        const updateERequestLink = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiMainCategory.listName}')/items(${id})`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    /*
    Gets List Entity Type 
    **/
    private getItemEntityType(listName: string): Promise<string> {
        let promise: Promise<string> = new Promise<string>((resolve, reject) => {
            this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')?$select=ListItemEntityTypeFullName`,
                SPHttpClient.configurations.v1,
                this.constant.spHttpOptions.getNoMetadata
            )
                .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
                    return response.json();
                })
                .then((response: { ListItemEntityTypeFullName: string }): void => {
                    resolve(response.ListItemEntityTypeFullName);
                })
                .catch((error: any) => {
                    reject(error);
                });
        });
        return promise;
    }
    /*
    deletes Main category Item from the list. 
    **/
    public removeCategoryListItem = async (mainCategoryItem: IMainCategory): Promise<any> => {
        if (mainCategoryItem.Id) {
            let filter = `&$filter=MainCategoryId eq ${mainCategoryItem.Id}`;
            const wikiVaultItems: IWikiVault[] = await this.CommonDataProvider.getListItems(this.constant.wikiVault.listName, this.constant.wikiVault.selectQuery, '', filter, this.constant.wikiVault.expand, '');

            await Promise.all(wikiVaultItems.map(async (wikiVaultItem) => {
                this.CommonDataProvider.removeListItem(wikiVaultItem.Id, this.constant.wikiVault.listName);
            }));
            const passwordListFolderId: string = mainCategoryItem.CategoryFolderName.split('_')[3];
            const categoryGroupId: string = mainCategoryItem.CategoryGroupName.Url.split('=')[1];
            this.CommonDataProvider.removeListItem(passwordListFolderId, this.constant.WikiVaultPassword.listName);
            this.CommonDataProvider.removeGroup(categoryGroupId);
            this.CommonDataProvider.removeListItem(mainCategoryItem.Id, this.constant.wikiMainCategory.listName);
        }
    }
    public isUserWikiAdmin = async (groupName: string): Promise<any> => {
        let isaMember: boolean = false;
        let query = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyname('${groupName}')/CanCurrentUserViewMembership`;
        const response = await this.context.spHttpClient.get(query, SPHttpClient.configurations.v1, this.constant.spHttpOptions.getNoMetadata);
        const jsonResponse: any = await response.json();
        isaMember = JSON.parse(jsonResponse.value);
        return isaMember;
    }
}



     // // build request
            // let requestDetails: any = this.constant.spHttpOptions.deleteNoMetadata;
            // // create the item
            // const response: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiMainCategory.listName}')/items('${mainCategoryItem.Id}')`,
            //     SPHttpClient.configurations.v1,
            //     requestDetails
            // );
            // Get WikiVault Password List Item
                // filter = `&$filter=SohoWikiVaultId eq ${wikiVaultItem.Id}`;
                // const wikiPassword: IWikiVaultPassword[] = await this.CommonDataProvider.getListItems(this.constant.WikiVaultPassword.listName, this.constant.WikiVaultPassword.selectQuery, '', filter, '', '');
                // if (wikiPassword.length > 0 && wikiPassword[0].Id) {
                //  this.CommonDataProvider.removeListItem(wikiPassword[0].Id, this.constant.WikiVaultPassword.listName);