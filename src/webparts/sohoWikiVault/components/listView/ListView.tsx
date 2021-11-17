import * as React from 'react';
import { IListViewProps, IListViewState } from './';
import styles from '../sohoWikiVault/SohoWikiVault.module.scss';

// Office UI Fabric controls
import {
    DetailsList, DetailsListLayoutMode, IColumn, Spinner, CheckboxVisibility,
    SelectionMode, IconButton, Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton
} from 'office-ui-fabric-react/';
// Constant
import { Constants, BusinessLogic } from '../../common';
//Data Provider
import { CommonDataProvider } from '../../dataProvider';
// Model
import { IListViewData, IWikiVault, IIconsDetails, IWikiVaultPassword } from '../../model';

import { ListViewStyle } from '../';
import { WikiEditForm, ViewWikiDetails } from '../../components';

export class ListView extends React.Component<IListViewProps, IListViewState> {
    private officeControlStyle: ListViewStyle;
    private constants: Constants;
    private commonDataProvider: CommonDataProvider;
    public state: IListViewState;
    public businessLogic: BusinessLogic;

    // setting state to null and setting list view webpasrt
    public async componentWillMount() {
        this.officeControlStyle = new ListViewStyle();
        this.constants = new Constants();
        this.businessLogic = new BusinessLogic;
        this.commonDataProvider = new CommonDataProvider(this.props.context);
        await this.setSateonLoad();
        if (this.props.wikiId) {
            const filter1 = `&$filter=Id eq ${this.props.wikiId}`;
            const wikiViewItem: IWikiVault[] = await this.commonDataProvider.getListItems(this.constants.wikiVault.listName, this.constants.wikiVault.selectQuery, this.constants.wikiVault.top, filter1, this.constants.wikiVault.expand, this.constants.wikiVault.orderby);
            if (wikiViewItem && wikiViewItem.length > 0) {
                this.setState({
                    currentItemBeingEditedorViewed: wikiViewItem[0],
                    showWikiDetailsForm: true,
                });
            }
        }
    }
    private setSateonLoad = async () => {
        let viewColumns: IColumn[] = null;
        if (this.props.isAdmin) {
            viewColumns = this.constants.viewColumnsAdmin;
        } else {
            viewColumns = this.constants.viewColumns;
        }
        this.setState({
            items: null,
            currentItemBeingEditedorViewed: null,
            currentItemBeingDeleted: '',
            showWikiEditForm: false,
            showDeleteDialog: false,
            showWikiDetailsForm: false,
            viewColumns: viewColumns
        });
        let listViewItems: IListViewData[] = [];
        let filter = `&$filter=MainCategoryId eq '${this.props.mainCategory.Id}' and SubCategoryId eq '${this.props.subCategory.Id}'`;
        if (this.props.mainCategory.Title == "All" && this.props.subCategory.Title == "All") {
            filter = '';
        } else if (this.props.mainCategory.Title == "All") {
            filter = `&$filter=SubCategoryId eq '${this.props.subCategory.Id}'`;
        } else if (this.props.subCategory.Title == "All") {
            filter = `&$filter=MainCategoryId eq '${this.props.mainCategory.Id}'`;
        }
        let listItems: IWikiVault[] = await this.commonDataProvider.getListItems(this.constants.wikiVault.listName, this.constants.wikiVault.selectQuery, this.constants.wikiVault.top, filter, this.constants.wikiVault.expand, this.constants.wikiVault.orderby);
        if (listItems && listItems.length > 0) {
            for (let i = 0; i < listItems.length; i++) {
                let subCat: string = '';
                let users: string = '';

                if (listItems[i]['PageOwner']) {
                    for (let user = 0; user < listItems[i]['PageOwner'].length; user++) {
                        users = users + listItems[i]['PageOwner'][user]['FirstName'] + ' ' + listItems[i]['PageOwner'][user]['LastName'] + ', ';
                    }
                }
                if (users) {
                    users = users.substring(0, users.length - 2);
                }
                const briefDesc = this.businessLogic.getDescription(listItems[i]['BriefLabelDescription']);
                const icon: IIconsDetails = this.businessLogic.setIconsforWiki(listItems[i].SubCategory, listItems[i].SubscriptionExpirationDate, this.constants,
                    listItems[i].IsSubscriptionPaid, listItems[i].Title, listItems[i].IsServiceBeingUsed);
                if (listItems[i].SubCategory) {
                    listItems[i].SubCategory.forEach(element => {
                        subCat = subCat + element.Title + ',';
                    });
                }
                //   listViewItems.push({ Title: listItems[i].Title, BriefLabelDescription: briefDesc, PageOwner: users, URL: url, Icon: icon, Edit: listItems[i], Delete: listItems[i].Id });
                listViewItems.push({ Title: listItems[i].Title, BriefLabelDescription: briefDesc, PageOwner: users, URL: listItems[i], Icon: icon, Edit: listItems[i], Delete: listItems[i].Id });
                //subCategory: subCat, 
            }
        }
        this.setState({
            items: listViewItems,
            //  viewColumns: this.constants.viewColumns,
            allItems: listViewItems,
            currentItemBeingEditedorViewed: null,
            currentItemBeingDeleted: '',
            showWikiEditForm: false,
            showDeleteDialog: false,
            showWikiDetailsForm: false
        });
    }
    public componentWillReceiveProps(nextProps) {
        this.state.items = nextProps.searchText ?
            this.state.allItems.filter(i => String(i['Title']).toLowerCase().indexOf(nextProps.searchText.toLowerCase()) > -1 ||
                String(i['subCategory']).toLowerCase().indexOf(nextProps.searchText.toLowerCase()) > -1 ||
                String(i['BriefLabelDescription']).toLowerCase().indexOf(nextProps.searchText.toLowerCase()) > -1)
            : this.state.allItems;
        this.setState({
            items: this.state.items
        });
    }
    // setting URL in render method
    private renderItemColumn = (item: any, index: number, column: IColumn) => {
        const fieldContent = item[column.fieldName as keyof any] as string;
        if (fieldContent) {
            switch (column.key) {
                case 'URL': {
                    const wiki = item[column.fieldName as keyof any] as IWikiVault;
                    return <IconButton className={styles.icons} iconProps={{ iconName: 'OpenInNewWindow' }} title="Open" ariaLabel="Open" onClick={() => this.onViewWikiDetailsClick(wiki)} />;
                }
                case 'Edit': {
                    const wiki = item[column.fieldName as keyof any] as IWikiVault;
                    return <IconButton className={styles.icons} iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => this.onEditclick(wiki)} />;
                }
                case 'Delete': {
                    return <IconButton className={styles.deleteIcon} iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" onClick={() => this.onDeleteclick(fieldContent)} />;
                }
                case 'BriefLabelDescription': {
                    return <div>{fieldContent}</div>;
                }
                case 'Icon': {
                    const icons = item[column.fieldName as keyof any] as IIconsDetails;
                    return this.constructIconHtml(icons);
                }
                default:
                    return <span>{fieldContent}</span>;
            }
        }
    }
    private onViewWikiDetailsClick = (wiki: IWikiVault) => {
        this.setState({
            showWikiDetailsForm: true,
            currentItemBeingEditedorViewed: wiki
        });
    }
    private onEditclick = (wiki: IWikiVault) => {
        this.setState({
            showWikiEditForm: true,
            currentItemBeingEditedorViewed: wiki
        });
    }
    private onWikiEditFormCloseButtonClick = async () => {
        this.setState({
            showWikiEditForm: false,
            showDeleteDialog: false,
            currentItemBeingEditedorViewed: null,
            currentItemBeingDeleted: ''
        });
        await this.setSateonLoad();
    }
    private onViewFormCloseButtonClick = async () => {
        const url = window.location.href.split('?');
        window.open(url[0], "_self");
    }
    private onDeleteclick = async (wikiId: string) => {
        this.setState({
            showDeleteDialog: true,
            currentItemBeingDeleted: wikiId
        });

    }
    private onDeleteYesButtonClick = async () => {
        const filter = `&$filter=SohoWikiVaultId eq ${this.state.currentItemBeingDeleted}`;
        const wikiPassword: IWikiVaultPassword[] = await this.commonDataProvider.getListItems(this.constants.WikiVaultPassword.listName, this.constants.WikiVaultPassword.selectQuery, '', filter, '', '');
        if (wikiPassword && wikiPassword.length > 0)
            await this.commonDataProvider.removeListItem(wikiPassword[0].Id, this.constants.WikiVaultPassword.listName);
        await this.commonDataProvider.removeListItem(this.state.currentItemBeingDeleted, this.constants.wikiVault.listName);
        this.setState({
            showDeleteDialog: false,
            currentItemBeingDeleted: ''
        });
        await this.setSateonLoad();
    }
    private constructIconHtml = (icons: IIconsDetails): JSX.Element => {
        return (
            <div>
                {icons.isSubscription ? <IconButton className={styles.icons} iconProps={{ iconName: 'Money' }} title="Subscription" ariaLabel="Subscription" /> : ''}
                {icons.isMembership ? <IconButton className={styles.icons} iconProps={{ iconName: 'PeopleRepeat' }} title="Membership" ariaLabel="Membership" /> : ''}
                {icons.isExpired ? <IconButton className={styles.icons} iconProps={{ iconName: 'SadSolid' }} title="Expired" ariaLabel="Expired" /> : ''}
                {icons.isServiceBeingUsed ? <IconButton className={styles.icons} iconProps={{ iconName: 'Blocked' }} title="Service" ariaLabel="Service" /> : ''}
            </div >

        );
    }
    // on column header click for sorting
    private onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        let isSortedDescending = column.isSortedDescending;
        // If we've sorted this column, flip it.
        if (column.isSorted) {
            isSortedDescending = !isSortedDescending;
        }
        // Sort the items.
        this.state.items = this.copyAndSort(this.state.items, column.fieldName!, isSortedDescending);
        this.state.viewColumns = this.state.viewColumns.map(col => {
            col.isSorted = col.key === column.key;
            if (col.isSorted) {
                col.isSortedDescending = isSortedDescending;
            }
            return col;
        });
        this.setState({
            items: this.state.items
        });
    }
    // return sorted items
    private copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
        const key = columnKey as keyof T;
        return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    }
    public render(): React.ReactElement<IListViewProps> {
        return (
            <div className={styles.listViewDiv}>
                {
                    (this.state.items) ? (
                        <div>
                            <DetailsList
                                items={this.state.items}
                                columns={this.state.viewColumns}
                                onRenderItemColumn={this.renderItemColumn}
                                checkboxVisibility={CheckboxVisibility.hidden}
                                layoutMode={DetailsListLayoutMode.justified}
                                onColumnHeaderClick={this.onColumnClick}
                                compact={false}
                                selectionMode={SelectionMode.none}
                                isHeaderVisible={true}
                            />
                            <Dialog
                                hidden={!this.state.showWikiEditForm}
                                styles={this.officeControlStyle.editDialogStyle}
                                dialogContentProps={{
                                    type: DialogType.largeHeader,
                                    title: "Wiki Edit Form",
                                    //  subText: "this.state.dialogdetails.subTitle",
                                    styles: this.officeControlStyle.editDialogControlStyle
                                }}
                                modalProps={{
                                    titleAriaId: 'modalWEFId',
                                    subtitleAriaId: 'modalWEFSubID',
                                    isBlocking: true,
                                }}>
                                <DialogFooter>
                                    <div>
                                        <WikiEditForm
                                            context={this.props.context}
                                            mainCategoryItems={this.props.mainCategoryItems}
                                            subCategoryItems={this.props.subCategoryItems}
                                            wikiItem={this.state.currentItemBeingEditedorViewed}
                                            onCloseButtonClick={this.onWikiEditFormCloseButtonClick}>
                                        </WikiEditForm>
                                    </div>
                                </DialogFooter>
                            </Dialog >
                            <Dialog
                                hidden={!this.state.showDeleteDialog}
                                dialogContentProps={{
                                    type: DialogType.largeHeader,
                                    title: "Delete Wiki",
                                    subText: 'Are you sure you want to delete the wiki?',
                                    styles: this.officeControlStyle.errorHeaderStyle
                                }}
                                modalProps={{
                                    titleAriaId: 'modalDFId',
                                    subtitleAriaId: 'modalDFSubID',
                                    isBlocking: true,
                                }}>
                                <DialogFooter>
                                    <DefaultButton text="No"
                                        onClick={this.onWikiEditFormCloseButtonClick} />
                                    <PrimaryButton
                                        text="Yes"
                                        styles={this.officeControlStyle.deleteButtonStyle}
                                        onClick={this.onDeleteYesButtonClick}
                                    />
                                </DialogFooter>
                            </Dialog>
                            <Dialog
                                hidden={!this.state.showWikiDetailsForm}
                                styles={this.officeControlStyle.viewDialogStyle}
                                dialogContentProps={{
                                    type: DialogType.largeHeader,
                                    title: `Wiki Details - ${this.state && this.state.currentItemBeingEditedorViewed ? this.state.currentItemBeingEditedorViewed.Title : ''}`,
                                    styles: this.officeControlStyle.viewHeaderStyle
                                }}
                                modalProps={{
                                    titleAriaId: 'modalWEFId',
                                    subtitleAriaId: 'modalWEFSubID',
                                    isBlocking: false,
                                }}>
                                <DialogFooter>
                                    <div>
                                        <ViewWikiDetails
                                            context={this.props.context}
                                            wikiItem={this.state.currentItemBeingEditedorViewed}
                                            onCloseButtonClick={this.onViewFormCloseButtonClick}>
                                        </ViewWikiDetails>
                                    </div>
                                </DialogFooter>
                            </Dialog >
                        </div>
                    ) : (<div>
                        <Spinner label="Loading Data. Please wait!!!" />
                    </div>)
                }
            </div>
        );
    }
}

//onClick={this.onDialogCancelClick} onClick={this.onDeleteClick} 