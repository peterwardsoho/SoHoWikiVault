// SharePoint Http Call
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// DataProvider
import { ICommonDataProvider } from '../';

// Constants
import { Constants } from '../../common';
// Model
import { IListDetails } from '../../model';

export class CommonDataProvider implements ICommonDataProvider {
    private context: WebPartContext;
    private constant: Constants;
    constructor(context: WebPartContext) { this.context = context; this.constant = new Constants(); }

    public getListItems = async (listName: string, selectQuery: string, top?: string, filter?: string, expand?: string, orderby?: string): Promise<any> => {
        let query = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items${selectQuery}${filter}${expand}${orderby}${top}`;
        const response = await this.context.spHttpClient.get(
            query,
            SPHttpClient.configurations.v1,
            this.constant.spHttpOptions.getNoMetadata,
        );
        const listItems: any = await response.json();
        return listItems.value;
    }

    public getItemEntityType(listName: string): Promise<string> {
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
    public removeListItem = async (Id: string, listName: string): Promise<any> => {
        // build request
        let requestDetails: any = this.constant.spHttpOptions.deleteNoMetadata;
        // create the item
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items('${Id}')`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    public removeGroup = async (groupId: string): Promise<any> => {
        // build request
        let requestDetails: any = this.constant.spHttpOptions.deleteNoMetadata;
        // create the item
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/removebyid('${groupId}')`,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    /*
    Create a SharePoint Group for a new Category
    **/
    public createSPGroup = async (groupName: string, ownerGroupId: string): Promise<any> => {
        let requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups`;
        let dataToPost = { Title: groupName };
        dataToPost['@odata.type'] = 'SP.Group';
        // build request
        let requestDetails: any = this.constant.spHttpOptions.postNoMetadata;
        requestDetails.body = JSON.stringify(dataToPost);
        // create the item
        const response: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );
        const groupDetails: any = await response.json();
        if (ownerGroupId) {
            this.SetOwner(groupDetails.Id, ownerGroupId);
        }
        return groupDetails.Id;
    }
    /*
    sets SohoWikiOwner as the owner of the group
    **/
    private SetOwner = async (groupId: string, ownerGroupId: string): Promise<any> => {
        const site = await this.context.spHttpClient.get(
            `${this.context.pageContext.web.absoluteUrl}/_api/site/id`,
            SPHttpClient.configurations.v1,
            this.constant.spHttpOptions.getNoMetadata,
        );
        const siteDetails: any = await site.json();
        if (siteDetails.value && ownerGroupId) {
            const endpoint = this.context.pageContext.web.absoluteUrl + `/_vti_bin/client.svc/ProcessQuery`;
            const body =
                `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="15.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                <Actions>
                    <SetProperty Id="1" ObjectPathId="2" Name="Owner">
                        <Parameter ObjectPathId="3" />
                    </SetProperty>
                    <Method Name="Update" Id="4" ObjectPathId="2" />
                </Actions>
                <ObjectPaths>
                    <Identity Id="2" Name="740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteDetails.value}:g:${groupId}" />
                    <Identity Id="3" Name="740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteDetails.value}:g:${ownerGroupId}" />
                </ObjectPaths>
            </Request>`;
            var headers = {
                "content-type": "text/xml"
            };
            var options = {
                body: body,
                header: headers
            };
            return this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, options)
                .then((response: SPHttpClientResponse) => { return response; });
        }
    }
    public getOwnerGroupDetails = async (): Promise<any> => {
        try {
            const ownerGroup = await this.context.spHttpClient.get(
                `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getByName('${this.constant.ownerGroupName}')`,
                SPHttpClient.configurations.v1,
                this.constant.spHttpOptions.getNoMetadata,
            );
            const ownerGroupDetails: IListDetails = await ownerGroup.json();
            return ownerGroupDetails.Id;
        } catch {
            return null;
        }
    }
}