// SharePoint Http Call
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// DataProvider
import { ISubCategoryDataProvider, CommonDataProvider } from '../';

// Constants
import { Constants } from '../../common';
import { IMainCategory, ISubCategory, IWikiVault, IWikiVaultPassword } from '../../model';

export class SubCategoryDataProvider implements ISubCategoryDataProvider {
    private CommonDataProvider: CommonDataProvider;
    private context: WebPartContext;
    private constant: Constants;
    constructor(context: WebPartContext) { this.context = context; this.constant = new Constants(); this.CommonDataProvider = new CommonDataProvider(this.context); }
    /*
       Saves and updates Main category Item. 
       **/
    public saveSubCategoryListItem = async (subCategoryItem: ISubCategory): Promise<any> => {
        if (subCategoryItem.Id == "-1") {
            const spEntityType = await this.CommonDataProvider.getItemEntityType(this.constant.wikiSubCategory.listName);
            // create item to create
            let newListItem: any = { Title: subCategoryItem.Title.trim(), SubCategoryOrder: +subCategoryItem.SubCategoryOrder.trim(), Status: this.constant.comparingStrings.visible };
            // add SP-required metadata
            newListItem['@odata.type'] = spEntityType;
            // build request
            let requestDetails: any = this.constant.spHttpOptions.postNoMetadata;
            requestDetails.body = JSON.stringify(newListItem);
            // create the item
            const response: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiSubCategory.listName}')/items`,
                SPHttpClient.configurations.v1,
                requestDetails
            );
            const newMainCat: IMainCategory = await response.json();
            subCategoryItem.Id = newMainCat.Id;

            return subCategoryItem;
        } else {
            await this.updateSubCategory({ Title: subCategoryItem.Title.trim(), SubCategoryOrder: subCategoryItem.SubCategoryOrder.toString().trim() }, subCategoryItem.Id);
            return null;
        }
    }
    /*
       Updates main category list item
       **/
    public updateSubCategory = async (updatedListItem: any, id: any) => {
        let requestDetails: any = this.constant.spHttpOptions.updateNoMetadata;
        requestDetails.body = JSON.stringify(updatedListItem);
        const updateERequestLink = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.constant.wikiSubCategory.listName}')/items(${id})`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    /*
    deletes Main category Item from the list. 
    **/
    public removeSubCategoryListItem = async (subCategoryItem: ISubCategory): Promise<any> => {
        if (subCategoryItem.Id) {
            let filter = `&$filter=SubCategoryId eq ${subCategoryItem.Id}`;
            const wikiVaultItems: IWikiVault[] = await this.CommonDataProvider.getListItems(this.constant.wikiVault.listName, this.constant.wikiVault.selectQuery, '', filter, this.constant.wikiVault.expand, '');

            await Promise.all(wikiVaultItems.map(async (wikiVaultItem) => {
                if (wikiVaultItem.SubCategory.length == 1) {
                    //  Get WikiVault Password List Item
                    filter = `&$filter=SohoWikiVaultId eq ${wikiVaultItem.Id}`;
                    const wikiPassword: IWikiVaultPassword[] = await this.CommonDataProvider.getListItems(this.constant.WikiVaultPassword.listName, this.constant.WikiVaultPassword.selectQuery, '', filter, '', '');
                    if (wikiPassword.length > 0 && wikiPassword[0]) {
                        this.CommonDataProvider.removeListItem(wikiPassword[0].Id, this.constant.WikiVaultPassword.listName);
                    }
                    this.CommonDataProvider.removeListItem(wikiVaultItem.Id, this.constant.wikiVault.listName);
                }
            }));
            this.CommonDataProvider.removeListItem(subCategoryItem.Id, this.constant.wikiSubCategory.listName);
        }
    }

}