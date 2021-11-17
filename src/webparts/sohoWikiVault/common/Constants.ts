// SharePoint Http Call
import { ISPHttpClientOptions } from '@microsoft/sp-http';
import { IColumn, IDatePickerStrings } from 'office-ui-fabric-react/';
import styles from '../components/sohoWikiVault/SohoWikiVault.module.scss';
// Components
import { IWikiEditFormState } from '../components';
export class Constants {
    public ownerGroupName = 'SohoWikiOwners';
    public fullControl = '2147483647';
    // public spGroupURL = '/_layouts/15/people.aspx?MembershipGroupId=';
    public listNames = {
        wikiMainCategoryList: { internalName: 'MainCategory', displayName: 'Main Category' },
    };
    public URLs = {
        spGroupURL: '/_layouts/15/people.aspx?MembershipGroupId=',
    };
    // Wiki Main Category List Details
    public wikiMainCategory = {
        listName: 'WikiMainCategory',
        selectQuery: '?$select=Id,Title,CategoryGroupName,CategoryOrder,Status,CategoryFolderName',
        expand: '',
        orderby: '&$orderby=CategoryOrder',
    };
    public wikiSubCategory = {
        listName: 'WikiSubCategory',
        selectQuery: '?$select=Id,Title,SubCategoryOrder,Status',
        expand: '',
        orderby: '&$orderby=SubCategoryOrder',
    };
    public wikiVault = {
        listName: 'WikiVault',
        selectQuery: '?$select=Id,Title,MainCategory/Title,MainCategory/Id,BriefLabelDescription,URL,SubscriptionType,CanThePasswordBeCutAndPasted,PageOwner/Id,PageOwner/FirstName,PageOwner/LastName,' +
            'SubCategory/Title,SubCategory/Id,SubscriptionExpirationDate,IsSubscriptionPaid,IsServiceBeingUsed,Modified,Editor/FirstName,Editor/LastName',
        expand: '&$expand=PageOwner,MainCategory/Title,MainCategory/Id,SubCategory/Title,SubCategory/Id,Editor/Id',
        orderby: '&$orderby=Created desc',
        top: '&$top=1000',
        dropdownColumn: 'SubscriptionType'
    };
    public WikiVaultPassword = {
        listName: 'WikiVaultPassword',
        listInternalName: 'WikiVaultPassword',
        selectQuery: '?$select=Id,Title,SitePassword,UserId,SohoWikiVaultId/Id,RestrictedLabelDescription',
        expand: '&$expand=SohoWikiVaultId/Id',
        orderby: '&$orderby=Created desc',
        top: '&$top=1000',
    };
    public spHttpOptions: any = {
        getNoMetadata: <ISPHttpClientOptions>{
            headers: { 'Accept': 'application/json; odata.metadata=none' }
        },
        getWithMetadata: <ISPHttpClientOptions>{
            headers: { 'Accept': 'application/json; odata=verbose' }
        },
        postNoMetadata: <ISPHttpClientOptions>{
            headers: {
                'Accept': 'application/json; odata.metadata=none',
                'CONTENT-TYPE': 'application/json'
            }
        },
        updateNoMetadata: <ISPHttpClientOptions>{
            headers: {
                'ACCEPT': 'application/json; odata.metadata=none',
                'CONTENT-TYPE': 'application/json',
                'X-HTTP-Method': 'MERGE',
                'If-Match': '*'
            }
        },
        deleteNoMetadata: <ISPHttpClientOptions>{
            headers: {
                'ACCEPT': 'application/json; odata.metadata=none',
                'CONTENT-TYPE': 'application/json',
                'X-HTTP-Method': 'DELETE',
                'If-Match': '*'
            }
        },
    };
    public viewColumnsAdmin: IColumn[] = [
        {
            key: 'Title', name: 'Label Name', fieldName: 'Title',
            isSorted: false, isSortedDescending: false, minWidth: 150, maxWidth: 200, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'BriefLabelDescription', name: 'Brief Label Description', fieldName: 'BriefLabelDescription',
            isSorted: false, isSortedDescending: false, minWidth: 250, maxWidth: 300, isResizable: true, isMultiline: true, headerClassName: styles.headerCss
        },
        {
            key: 'PageOwner', name: 'Page Owner', fieldName: 'PageOwner',
            isSorted: false, isSortedDescending: false, minWidth: 150, maxWidth: 200, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'Icon', name: 'Icon(s)', fieldName: 'Icon',
            isSorted: false, isSortedDescending: false, minWidth: 150, maxWidth: 200, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'URL', name: 'URL', fieldName: 'URL',
            isSorted: false, isSortedDescending: false, minWidth: 100, maxWidth: 100, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'Edit', name: 'Edit', fieldName: 'Edit',
            isSorted: false, isSortedDescending: false, minWidth: 40, maxWidth: 70, isResizable: false, headerClassName: styles.headerCss
        },
        {
            key: 'Delete', name: 'Delete', fieldName: 'Delete',
            isSorted: false, isSortedDescending: false, minWidth: 40, maxWidth: 70, isResizable: true, headerClassName: styles.headerCss
        }
    ];
    public viewColumns: IColumn[] = [
        {
            key: 'Title', name: 'Label Name', fieldName: 'Title',
            isSorted: false, isSortedDescending: false, minWidth: 150, maxWidth: 200, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'BriefLabelDescription', name: 'Brief Label Description', fieldName: 'BriefLabelDescription',
            isSorted: false, isSortedDescending: false, minWidth: 250, maxWidth: 300, isResizable: true, isMultiline: true, headerClassName: styles.headerCss
        },
        {
            key: 'URL', name: 'URL', fieldName: 'URL',
            isSorted: false, isSortedDescending: false, minWidth: 100, maxWidth: 100, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'PageOwner', name: 'Page Owner', fieldName: 'PageOwner',
            isSorted: false, isSortedDescending: false, minWidth: 150, maxWidth: 200, isResizable: true, headerClassName: styles.headerCss
        },
        {
            key: 'Icon', name: 'Icon(s)', fieldName: 'Icon',
            isSorted: false, isSortedDescending: false, minWidth: 150, maxWidth: 200, isResizable: true, headerClassName: styles.headerCss
        },

    ];
    public comparingStrings = {
        visible: 'Visible',
        hidden: 'Hidden'
    };
    public dateTimePickerString: IDatePickerStrings = {
        months: [
            'January',
            'February',
            'March',
            'April',
            'May',
            'June',
            'July',
            'August',
            'September',
            'October',
            'November',
            'December'
        ],

        shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

        days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

        shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
        goToToday: 'Go to today',
        prevMonthAriaLabel: 'Go to previous month',
        nextMonthAriaLabel: 'Go to next month',
        prevYearAriaLabel: 'Go to previous year',
        nextYearAriaLabel: 'Go to next year'
    };
    public modules = {
        toolbar: [
            [{ 'header': [1, 2, false] }],
            ['bold', 'italic', 'underline', 'strike', 'blockquote'],
            [{ 'list': 'ordered' }, { 'list': 'bullet' }, { 'indent': '-1' }, { 'indent': '+1' }],
            ['link', 'image'],
            [{ 'color': [] }, { 'background': [] }],          // dropdown with defaults from theme
            [{ 'font': [] }],
            ['clean']
        ],
    };
    public formats = [
        'header',
        'bold', 'italic', 'underline', 'strike', 'blockquote',
        'list', 'bullet', 'indent',
        'link', 'image', 'color', 'background', 'font', 'clean'
    ];
    public wikiformInitialState: IWikiEditFormState = {
        controlDisabled: false,
        showNew: true,
        wikiMainCategoryChoices: [],
        wikiSubCategoryChoices: [],
        wikiSubscriptionTypeChoices: [],
        wikiLabelName: '',
        selectedMainCategory: '',
        selectedSubCategory: [],
        selectedSubscriptionType: '',
        briefLabelDesc: '',
        subscriptionExpirationDate: null,
        isSubscriptionPaid: false,
        userId: '',
        password: '',
        canThePasswordBeCutAndPasted: false,
        restrictedLabelDesc: '',
        url: '',
        istheServiceBeingUsed: false,
        pageOwners: [],
        defaultSelectedPageOwners: [],
        wikiId: '',
        error: '',
        sitePasswordError: '',
        userIDError: '',
        labelError: '',
    };
    public permission = {
        yes: 'yes',
        no: 'no'
    };
}