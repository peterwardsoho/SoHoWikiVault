// SharePoint Http Call
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// DataProvider
import { IPropertyPaneDataProvider, CommonDataProvider } from '../';
// Constants
import { Constants } from '../../common';
//Modal
import { IListDetails, IWikiVault, IMainCategory, IWikiPasswordCreateItem, IWikiVaultPassword } from '../../model';
//Compnents
import { IWikiEditFormState } from '../../components';
// PnP
import { sp } from "@pnp/sp";
import '@pnp/sp/presets/all';

export class PropertyPaneDataProvider implements IPropertyPaneDataProvider {
    private context: WebPartContext;
    private constant: Constants;
    private CommonDataProvider: CommonDataProvider;
    constructor(context: WebPartContext) { this.context = context; this.constant = new Constants(); this.CommonDataProvider = new CommonDataProvider(this.context); }

    public createConfigLists = async (): Promise<boolean> => {
        const permission: any = await sp.web.getCurrentUserEffectivePermissions();
        if (permission.High == this.constant.fullControl) {
            // const webSite = new Web(siteURL);
            const listDetails: IListDetails[] = await this.getListDetails();
            let ownerGroupId = await this.CommonDataProvider.getOwnerGroupDetails();
            if (!ownerGroupId) {
                ownerGroupId = await this.CommonDataProvider.createSPGroup(this.constant.ownerGroupName, null);
                await this.setOwnerGroupPermissionOnWeb(ownerGroupId);
            }
            // create Wiki Main Category List
            const MCListId = await this.create_WikiMainCategory(listDetails, ownerGroupId);
            // create Wiki Sub Category
            const SCListId = await this.create_WikiSubCategory(listDetails, ownerGroupId);
            // create Wiki Vault
            await this.create_WikiVault(listDetails, ownerGroupId, MCListId, SCListId);
            // create Wiki Vault Password
            await this.create_WikiVaultPassword(listDetails, ownerGroupId);
            return true;
        } else {
            return false;
        }
    }
    public getListDetails(): Promise<IListDetails[]> {
        return new Promise<IListDetails[]>(async (resolve, reject) => {
            try {
                await
                    sp.web.lists.filter('Hidden eq true').get().then((data: any) => {
                        const listDetails: IListDetails[] = data;
                        resolve(listDetails);
                    });
            } catch (e) {
                reject(e);
            }
        });
    }
    private create_WikiMainCategory = async (listDetails: IListDetails[], ownerGroupId: string) => {
        let id: string = '';
        const listCheck: IListDetails[] = listDetails.filter(list => list.Title == this.constant.wikiMainCategory.listName);
        console.log(listCheck);
        if (listCheck.length == 0) {
            // checking if list already exist
            const listEnsureResult = await sp.web.lists.add(this.constant.wikiMainCategory.listName, '', 100, false, { Hidden: true });
            id = listEnsureResult.data.Id;
            // const listEnsureResult = await sp.web.lists.ensure(this.constant.wikiMainCategory.listName);
            // id = listEnsureResult.data.Id;
            // if (listEnsureResult.created) {
            // Adding List Columns
            await sp.web.lists.getByTitle(this.constant.wikiMainCategory.listName).fields.createFieldAsXml(`<Field Type="Number" DisplayName="CategoryOrder" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" ID="{00865eb1-45cf-4b52-96ca-4e5f1fc3e7ad}" SourceID="{47c9e70e-5b74-4f04-a55a-7ee0fd340e9c}" StaticName="CategoryOrder" Name="CategoryOrder" ColName="float1" RowOrdinal="0" CustomFormatter="" Percentage="FALSE" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.wikiMainCategory.listName).fields.createFieldAsXml(`<Field Type="URL" DisplayName="CategoryGroupName" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" ID="{0702599b-6a6f-4354-bdb1-5311c854039d}" SourceID="{47c9e70e-5b74-4f04-a55a-7ee0fd340e9c}" StaticName="CategoryGroupName" Name="CategoryGroupName" ColName="nvarchar4" RowOrdinal="0" ColName2="nvarchar5" RowOrdinal2="0" CustomFormatter="" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.wikiMainCategory.listName).fields.createFieldAsXml(`<Field Type="Choice" DisplayName="Status" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" ID="{f38a2412-4ee7-405f-a366-e13c2536b36c}" SourceID="{47c9e70e-5b74-4f04-a55a-7ee0fd340e9c}" StaticName="Status" Name="Status" ColName="nvarchar6" RowOrdinal="0"><Default>Visible</Default><CHOICES><CHOICE>Visible</CHOICE><CHOICE>Hidden</CHOICE></CHOICES></Field>`);
            await sp.web.lists.getByTitle(this.constant.wikiMainCategory.listName).fields.createFieldAsXml(`<Field Type="Text" DisplayName="CategoryFolderName" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" ID="{b999ff27-ecc2-4618-961c-22366c715902}" SourceID="{47c9e70e-5b74-4f04-a55a-7ee0fd340e9c}" StaticName="CategoryFolderName" Name="CategoryFolderName" ColName="nvarchar7" RowOrdinal="0" CustomFormatter="" Version="1" />`);
            this.setFolderPermission(this.constant.wikiMainCategory.listName, ownerGroupId);
            //  }
        }
        return id;
    }
    private create_WikiSubCategory = async (listDetails: IListDetails[], ownerGroupId: string) => {
        let id: string = '';
        const listCheck: IListDetails[] = listDetails.filter(list => list.Title == this.constant.wikiSubCategory.listName);
        if (listCheck.length == 0) {
            // checking if list already exist
            // const listEnsureResult = await sp.web.lists.ensure(this.constant.wikiSubCategory.listName);
            // id = listEnsureResult.data.Id;
            const listEnsureResult = await sp.web.lists.add(this.constant.wikiSubCategory.listName, '', 100, false, { Hidden: true });
            id = listEnsureResult.data.Id;
            // if (listEnsureResult.created) {
            // Adding List Columns
            await sp.web.lists.getByTitle(this.constant.wikiSubCategory.listName).fields.createFieldAsXml(`<Field Type="Number" DisplayName="SubCategoryOrder" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" ID="{417564d2-1761-4a67-86e3-8052e42f3683}" SourceID="{82cf66d8-58b6-42bf-b662-0e95ac350b53}" StaticName="SubCategoryOrder" Name="SubCategoryOrder" ColName="float1" RowOrdinal="0" CustomFormatter="" Percentage="FALSE" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.wikiSubCategory.listName).fields.createFieldAsXml(`<Field Type="Choice" DisplayName="Status" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" ID="{cb66292a-6a6c-485a-8e9e-b3a5926992ba}" SourceID="{82cf66d8-58b6-42bf-b662-0e95ac350b53}" StaticName="Status" Name="Status" ColName="nvarchar4" RowOrdinal="0"><Default>Visible</Default><CHOICES><CHOICE>Visible</CHOICE><CHOICE>Hidden</CHOICE></CHOICES></Field>`);
            this.setFolderPermission(this.constant.wikiSubCategory.listName, ownerGroupId);
            // }
        }
        return id;
    }
    private create_WikiVault = async (listDetails: IListDetails[], ownerGroupId: string, MClistId: string, SCListId: string) => {
        const listCheck: IListDetails[] = listDetails.filter(list => list.Title == this.constant.wikiVault.listName);
        if (listCheck.length == 0) {
            // Getting Main Category List Details
            // const wikiMainListId = await this.get_LookupColumnID(this.constant.wikiMainCategory.listName);
            // console.log(wikiMainListId);
            // // Getting Sub Category List Details
            // const wikiSubListId = await this.get_LookupColumnID(this.constant.wikiSubCategory.listName);
            // console.log(wikiSubListId);
            // checking if list already exist

            const listEnsureResult = await sp.web.lists.add(this.constant.wikiVault.listName, '', 100, false, { Hidden: true });

            //             const listEnsureResult = await sp.web.lists.ensure(this.constant.wikiVault.listName);
            // if (listEnsureResult.created) {
            // Adding List Columns
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Lookup" DisplayName="MainCategory" Required="FALSE" EnforceUniqueValues="FALSE" List="${MClistId}" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{7624248f-dadb-486e-a7ac-fee1476d2821}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="MainCategory" Name="MainCategory" ColName="int1" RowOrdinal="0" Group="" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Note" DisplayName="BriefLabelDescription" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" NumLines="6" RichText="TRUE" RichTextMode="FullHtml" IsolateStyles="TRUE" Sortable="FALSE" ID="{7f404557-eaa4-4895-b19f-6d761605e5f9}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="BriefLabelDescription" Name="BriefLabelDescription" ColName="ntext2" RowOrdinal="0" CustomFormatter="" RestrictedMode="TRUE" AppendOnly="FALSE" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="URL" DisplayName="URL" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" ID="{c1b05bb0-3b17-4773-8716-dcb599099b58}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="URL" Name="URL" ColName="nvarchar4" RowOrdinal="0" ColName2="nvarchar5" RowOrdinal2="0" />`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="UserMulti" DisplayName="PageOwner" List="UserInfo" Required="FALSE" EnforceUniqueValues="FALSE" ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0" Mult="TRUE" Sortable="FALSE" ID="{46602196-dca8-4d97-875c-ef248303ce7e}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="PageOwner" Name="PageOwner" ColName="int2" RowOrdinal="0" Group="" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Boolean" DisplayName="CanThePasswordBeCutAndPasted" EnforceUniqueValues="FALSE" Indexed="FALSE" ID="{8245db15-1e9a-4de4-92bf-7fa602026ee5}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="CanThePasswordBeCutAndPasted" Name="CanThePasswordBeCutAndPasted" ColName="bit1" RowOrdinal="0" CustomFormatter="" Required="FALSE" Version="2"><Default>1</Default></Field>`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Boolean" DisplayName="IsSubscriptionPaid" EnforceUniqueValues="FALSE" Indexed="FALSE" ID="{1533cd26-51d6-48f2-b4bc-2729f0a764bb}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="IsSubscriptionPaid" Name="IsSubscriptionPaid" ColName="bit2" RowOrdinal="0" CustomFormatter="" Required="FALSE" Version="1"><Default>1</Default></Field>`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="DateTime" DisplayName="SubscriptionExpirationDate" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" FriendlyDisplayFormat="Disabled" ID="{7dac30b0-a897-469f-b485-a5b795b8fb4b}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="SubscriptionExpirationDate" Name="SubscriptionExpirationDate" ColName="datetime1" RowOrdinal="0" CustomFormatter="" CalType="0" Version="2" />`);
            // await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Choice" DisplayName="SubscriptionType" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" ID="{d73acdee-cebb-4334-8ee8-851190e99a3c}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="SubscriptionType" Name="SubscriptionType" ColName="nvarchar6" RowOrdinal="0" CustomFormatter="" Version="1"><Default>Monthly- On Going</Default><CHOICES><CHOICE>Monthly- On Going</CHOICE><CHOICE>Monthly</CHOICE><CHOICE>Yearly-  On Going</CHOICE><CHOICE>Yearly</CHOICE><CHOICE>Canceled</CHOICE></CHOICES></Field>`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Choice" DisplayName="SubscriptionType" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" ID="{d73acdee-cebb-4334-8ee8-851190e99a3c}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="SubscriptionType" Name="SubscriptionType" ColName="nvarchar6" RowOrdinal="0" CustomFormatter="" Version="3"><Default>None</Default><CHOICES><CHOICE>Monthly- On Going</CHOICE><CHOICE>Monthly</CHOICE><CHOICE>Yearly-  On Going</CHOICE><CHOICE>Yearly</CHOICE><CHOICE>Canceled</CHOICE><CHOICE>None</CHOICE></CHOICES></Field>`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="LookupMulti" DisplayName="SubCategory" Required="FALSE" EnforceUniqueValues="FALSE" List="${SCListId}" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{b85e599a-78bc-4b38-a8e3-f993c8c8f15d}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="SubCategory" Name="SubCategory" ColName="int3" RowOrdinal="0" Group="" Version="2" Mult="TRUE" Sortable="FALSE" />`);
            await sp.web.lists.getByTitle(this.constant.wikiVault.listName).fields.createFieldAsXml(`<Field Type="Boolean" DisplayName="IsServiceBeingUsed" EnforceUniqueValues="FALSE" Indexed="FALSE" ID="{db0791ff-65ea-49f4-9b51-e9ee5c82f3f2}" SourceID="{6c4c0874-f96d-4efa-91b2-034b00ee39d6}" StaticName="IsServiceBeingUsed" Name="IsServiceBeingUsed" ColName="bit3" RowOrdinal="0" CustomFormatter="" Required="FALSE" Version="1"><Default>1</Default></Field>`);
            this.setFolderPermission(this.constant.wikiVault.listName, ownerGroupId);
            //}
        }
    }
    private create_WikiVaultPassword = async (listDetails: IListDetails[], ownerGroupId: string) => {
        const listCheck: IListDetails[] = listDetails.filter(list => list.Title == this.constant.WikiVaultPassword.listName);
        if (listCheck.length == 0) {
            // checking if list already exist
            // const listEnsureResult = await sp.web.lists.ensure(this.constant.WikiVaultPassword.listName);
            const listEnsureResult = await sp.web.lists.add(this.constant.WikiVaultPassword.listName, '', 100, false, { Hidden: true });
            // if (listEnsureResult.created) {
            // Adding List Columns
            await sp.web.lists.getByTitle(this.constant.WikiVaultPassword.listName).fields.createFieldAsXml(`<Field Type="Text" DisplayName="UserId" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" ID="{ca374d70-5df4-42ab-b3ce-a10bd6710c2a}" SourceID="{f06f009b-1d6f-494d-8a39-ef1793a7678e}" StaticName="UserId" Name="UserId" ColName="nvarchar4" RowOrdinal="0" CustomFormatter="" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.WikiVaultPassword.listName).fields.createFieldAsXml(`<Field Type="Text" DisplayName="SitePassword" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" ID="{7505e7e4-2613-4d3f-81d7-678ecbfbcb38}" SourceID="{f06f009b-1d6f-494d-8a39-ef1793a7678e}" StaticName="SitePassword" Name="SitePassword" ColName="nvarchar5" RowOrdinal="0" CustomFormatter="" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.WikiVaultPassword.listName).fields.createFieldAsXml(`<Field Type="Note" DisplayName="RestrictedLabelDescription" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" NumLines="6" RichText="TRUE" RichTextMode="FullHtml" IsolateStyles="TRUE" Sortable="FALSE" ID="{2d39d207-8cb7-4303-a411-d1516f4e56c3}" SourceID="{f06f009b-1d6f-494d-8a39-ef1793a7678e}" StaticName="RestrictedLabelDescription" Name="RestrictedLabelDescription" ColName="ntext2" RowOrdinal="0" CustomFormatter="" RestrictedMode="TRUE" AppendOnly="FALSE" Version="1" />`);
            await sp.web.lists.getByTitle(this.constant.WikiVaultPassword.listName).fields.createFieldAsXml(`<Field Type="Text" Name="SohoWikiVaultId" FromBaseType="FALSE" DisplayName="SohoWikiVaultId" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" MaxLength="255" ID="{3ec41f73-a810-4049-bfa0-2f83dc555a32}" Version="3" StaticName="SohoWikiVaultId" SourceID="{f06f009b-1d6f-494d-8a39-ef1793a7678e}" ColName="nvarchar6" RowOrdinal="0" />`);
            this.setFolderPermission(this.constant.WikiVaultPassword.listName, ownerGroupId);
            //}
        }
    }
    private setFolderPermission = async (listName: string, ownerGroupId: string) => {
        let requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`;
        // build request
        let requestDetails: any = {};
        // break Permission on the item
        const responseBreakPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );

        // set view permisson for wiki owner SP Group
        requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/roleassignments/addroleassignment(principalid=${ownerGroupId}, roleDefId=1073741829)`;
        // break Permission on the item
        const ownerPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
    private setOwnerGroupPermissionOnWeb = async (groupId: string) => {
        let requestDetails: any = {};
        // set view permisson for wiki owner SP Group
        const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/roleassignments/addroleassignment(principalid=${groupId}, roleDefId=1073741829)`;
        // break Permission on the item
        const ownerPermission: SPHttpClientResponse = await this.context.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            requestDetails
        );
    }
}



// private get_LookupColumnID = (listName: string): Promise<string> => {
//     return new Promise<string>(async (resolve, reject) => {
//         try {
//             await
//                 sp.web.lists.getByTitle(listName).get().then((data: IListDetails) => {
//                     resolve(data.Id);
//                 });
//         } catch (e) {
//             reject(e);
//         }
//     });
// }