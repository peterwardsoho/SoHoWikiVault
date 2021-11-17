import * as React from 'react';
import { IMainCategoryProps, IMainCategoryState } from './';
// Style
import styles from '../sohoWikiVault/SohoWikiVault.module.scss';
// Office UI Fabric Imports
import { Pivot, PivotItem, PivotLinkSize, PivotLinkFormat, TextField, IPivotStyles } from 'office-ui-fabric-react';
// Constant
import { Constants } from '../../common';
//Dara Provider
import { CommonDataProvider, MainCategoryDataProvider, SubCategoryDataProvider } from '../../dataProvider';
// Model
import { IMainCategory, ISubCategory } from '../../model';
// Component
import { SubCategory, AdminCenter } from '../';
export class MainCategory extends React.Component<IMainCategoryProps, IMainCategoryState> {
    public state: IMainCategoryState;
    private constants: Constants;
    private commonDataProvider: CommonDataProvider;
    private mainCategoryProvider: MainCategoryDataProvider;
    private subCategoryProvider: SubCategoryDataProvider;

    // Pivot Style
    private pivotStyle: IPivotStyles = {
        link: {
            borderTopWidth: '1px',
            borderRightWidth: '1px',
            borderLeftWidth: '1px',
            borderStyle: "solid",
            marginRight: '4px',
            borderTopRightRadius: '4px',
            borderTopLeftRadius: '4px'
        },
        linkIsSelected: {
            borderTopWidth: '1px',
            borderRightWidth: '1px',
            borderLeftWidth: '1px',
            borderStyle: "solid",
            marginRight: '4px',
            borderTopRightRadius: '4px',
            borderTopLeftRadius: '4px',
            fontWeight: 400
        },
        root: {
            display: 'flex', flexWrap: 'wrap'
        },
        count: {},
        icon: {},
        linkContent: {},
        text: {}
    };
    /* 
    Setting the initial state. 
    Getting Main Category and Sub Category Items from the list and adding 'All Tab' in the Main Category.
    **/
    public async componentWillMount() {
        this.constants = new Constants();
        this.commonDataProvider = new CommonDataProvider(this.props.context);
        this.mainCategoryProvider = new MainCategoryDataProvider(this.props.context);
        this.subCategoryProvider = new SubCategoryDataProvider(this.props.context);

        this.setState({
            isAdmin: false,
            adminGroupId: '',
            searchText: '',
            mainCategoryItems: [],
            subCategoryItems: [],
            mainCategory: [],
            subCategory: [],
            adminControlDisabled: false
        });
        let mainCategoryItems: IMainCategory[] = await this.commonDataProvider.getListItems(this.constants.wikiMainCategory.listName, this.constants.wikiMainCategory.selectQuery, '', '', '', this.constants.wikiMainCategory.orderby);
        let subCategoryItems: ISubCategory[] = await this.commonDataProvider.getListItems(this.constants.wikiSubCategory.listName, this.constants.wikiSubCategory.selectQuery, '', '', '', this.constants.wikiSubCategory.orderby);
        const isAdmin: boolean = await this.mainCategoryProvider.isUserWikiAdmin(this.constants.ownerGroupName);
        if (isAdmin) {
            const adminGroupId = await this.commonDataProvider.getOwnerGroupDetails();
            this.state.adminGroupId = this.props.context.pageContext.web.absoluteUrl + '/' + this.constants.URLs.spGroupURL + adminGroupId;
        }
        if (mainCategoryItems && mainCategoryItems.length > 0) {
            mainCategoryItems.forEach(item => {
                this.state.mainCategory.push({ Id: item.Id, Title: item.Title, CategoryOrder: item.CategoryOrder, CategoryGroupName: item.CategoryGroupName, Status: item.Status, CategoryFolderName: item.CategoryFolderName });
            });
            this.state.mainCategory.push({ Title: 'All', Id: '0', CategoryOrder: '10000000', Status: this.constants.comparingStrings.visible });
        }
        if (subCategoryItems && subCategoryItems.length > 0) {
            this.state.subCategory.push({ Title: 'All', Id: '0', SubCategoryOrder: '-200', Status: this.constants.comparingStrings.visible });
            subCategoryItems.forEach(item => {
                this.state.subCategory.push({ Id: item.Id, Title: item.Title, SubCategoryOrder: item.SubCategoryOrder, Status: item.Status });
            });
        }
        this.setState({
            searchText: '',
            mainCategoryItems: mainCategoryItems,
            subCategoryItems: subCategoryItems,
            mainCategory: this.state.mainCategory,
            subCategory: this.state.subCategory,
            adminControlDisabled: false,
            isAdmin: isAdmin,
            adminGroupId: this.state.adminGroupId
        });
    }
    // Fires when the search text is changed. 
    private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
        this.setState({
            searchText: text
        });
    }
    //#region Main Category CRUD Operations
    /* 
    Fires on edit button click. 
    Sets the 'editEnabled' property to true which makes textboxs visible. 
    **/
    private onEditMCButtonClick = (mainCategoryRow: IMainCategory) => {
        const index = this.state.mainCategoryItems.indexOf(mainCategoryRow, 0);
        this.state.mainCategoryItems[index].editEnabled = true;
        this.setState({ mainCategoryItems: this.state.mainCategoryItems, adminControlDisabled: true });
    }
    /* 
    Fires on cancel button click. 
    if it was a new item then it deletes it from the maincategory Items list
    If it was a existing item, it sets it value to the original value and sets editenabled to false
    **/
    private onCancelMCButtonClick = (mainCategoryRow: IMainCategory) => {
        const index = this.state.mainCategoryItems.indexOf(mainCategoryRow, 0);
        if (this.state.mainCategoryItems[index].Id == "-1") {
            this.state.mainCategoryItems.splice(index, 1);
        } else {
            this.state.mainCategoryItems[index].editEnabled = false;
            this.state.mainCategoryItems[index].Title = this.state.mainCategory[index].Title;
            this.state.mainCategoryItems[index].CategoryOrder = this.state.mainCategory[index].CategoryOrder;
        }
        this.setState({ mainCategoryItems: this.state.mainCategoryItems, adminControlDisabled: false });
    }
    /* 
    Fires on Delete button click. 
    Deletes item from thr list. Can only trigger on existing item. 
    It also deletes the item from both mainCategory and maincategoryitem List Object and updates the state.
    **/
    private onDeleteMCButtonClick = async (mainCategoryRow: IMainCategory) => {
        let index = this.state.mainCategoryItems.indexOf(mainCategoryRow, 0);
        this.state.mainCategoryItems[index].spinner = true;
        this.setState({ mainCategoryItems: this.state.mainCategoryItems });
        await this.mainCategoryProvider.removeCategoryListItem(mainCategoryRow);

        this.state.mainCategory.splice(index, 1);
        this.state.mainCategoryItems.splice(index, 1);
        this.setState({ mainCategoryItems: this.state.mainCategoryItems, mainCategory: this.state.mainCategory });
    }
    private onShowHideMCButtonClick = async (mainCategoryRow: IMainCategory) => {
        let status = '';
        if (mainCategoryRow.Status == this.constants.comparingStrings.visible) {
            status = this.constants.comparingStrings.hidden;
        } else {
            status = this.constants.comparingStrings.visible;
        }
        let index = this.state.mainCategoryItems.indexOf(mainCategoryRow, 0);
        this.state.mainCategoryItems[index].spinner = true;
        this.setState({ mainCategoryItems: this.state.mainCategoryItems });
        await this.mainCategoryProvider.updateMainCategory({ Status: status }, mainCategoryRow.Id);
        this.state.mainCategory[index].Status = status;
        this.state.mainCategoryItems[index].Status = status;
        this.state.mainCategoryItems[index].spinner = false;
        this.setState({ mainCategoryItems: this.state.mainCategoryItems, mainCategory: this.state.mainCategory });
    }
    /* 
    Fires on save button click. 
    Adds a new item or update the item in the Main Category list.
    It also updates mainCategory and maincategoryitem List Object and then sory them to have correct order and updates the state.
    **/
    private onSaveMCButtonClick = async (mainCategoryRow: IMainCategory) => {
        const index = this.state.mainCategoryItems.indexOf(mainCategoryRow, 0);
        this.state.mainCategoryItems[index].spinner = true;
        this.setState({ mainCategoryItems: this.state.mainCategoryItems });
        const newItem: IMainCategory = await this.mainCategoryProvider.saveCategoryListItem(mainCategoryRow);
        if (newItem && newItem.Id) {
            this.state.mainCategory.push({ Id: newItem.Id, Title: this.state.mainCategoryItems[index].Title.trim(), CategoryOrder: this.state.mainCategoryItems[index].CategoryOrder.trim(), Status: this.state.mainCategoryItems[index].Status });
            this.state.mainCategoryItems[index].Id = newItem.Id;
            this.state.mainCategoryItems[index].CategoryGroupName = { ...newItem.CategoryGroupName };
            this.state.mainCategoryItems[index].CategoryFolderName = newItem.CategoryFolderName;
            if (this.state.mainCategory.length == 1) {
                this.state.mainCategory.push({ Title: 'All', Id: '0', CategoryOrder: '10000000', Status: this.constants.comparingStrings.visible });
            }
        } else {
            this.state.mainCategory[index].Title = this.state.mainCategoryItems[index].Title.trim();
            this.state.mainCategory[index].CategoryOrder = this.state.mainCategoryItems[index].CategoryOrder.toString().trim();
        }
        this.state.mainCategory.sort((a, b) => (+a.CategoryOrder > +b.CategoryOrder) ? 1 : -1);
        this.state.mainCategoryItems[index].Title = this.state.mainCategoryItems[index].Title.trim();
        this.state.mainCategoryItems[index].CategoryOrder = this.state.mainCategoryItems[index].CategoryOrder.toString().trim();
        this.state.mainCategoryItems[index].spinner = false;
        this.state.mainCategoryItems[index].editEnabled = false;
        this.state.mainCategoryItems.sort((a, b) => (+a.CategoryOrder > +b.CategoryOrder) ? 1 : -1);
        this.setState({ mainCategoryItems: this.state.mainCategoryItems, mainCategory: this.state.mainCategory, adminControlDisabled: false });
    }
    /* 
    Fires on textbox change event of the textboxes. 
    Updates the state accordingly.
    isTitle parameter will be true if the changes are made for catefory textbox else it will be false.
    **/
    private onMCTextBoxChange = (newValue: string, mainCategoryRow: IMainCategory, isTitle: boolean) => {
        const index = this.state.mainCategoryItems.indexOf(mainCategoryRow, 0);
        if (isTitle) {
            this.state.mainCategoryItems[index].Title = newValue;
        } else {
            this.state.mainCategoryItems[index].CategoryOrder = newValue;
        }
        this.setState({ mainCategoryItems: this.state.mainCategoryItems });
    }
    /* 
    Fires on Add a New Category button click. 
    Adds a new item to maincategorylist item state variable with id set to -1. If an extra item already there it wont add another new item. 
    if its canceled it will delete the extra added item. 
    **/
    private onAddMCButtonClick = () => {
        const newItem = this.state.mainCategoryItems.filter(i => i.Id == "-1");
        if (newItem && newItem.length == 0) {
            this.state.mainCategoryItems.push({ Id: "-1", Title: '', CategoryOrder: '', CategoryGroupName: { Description: '', Url: '' }, editEnabled: true, Status: this.constants.comparingStrings.visible, CategoryFolderName: '' });
            this.setState({ mainCategoryItems: this.state.mainCategoryItems, adminControlDisabled: true });
        }
    }
    //#endregion

    //#region Sub Category CRUD Operations
    /* 
    Fires on edit button click. 
    Sets the 'editEnabled' property to true which makes textboxs visible. 
    **/
    private onEditSCButtonClick = (subCategoryRow: ISubCategory) => {
        const index = this.state.subCategoryItems.indexOf(subCategoryRow, 0);
        this.state.subCategoryItems[index].editEnabled = true;
        this.setState({ subCategoryItems: this.state.subCategoryItems, adminControlDisabled: true });
    }
    /* 
    Fires on cancel button click. 
    if it was a new item then it deletes it from the maincategory Items list
    If it was a existing item, it sets it value to the original value and sets editenabled to false
    **/
    private onCancelSCButtonClick = (subCategoryRow: ISubCategory) => {
        const index = this.state.subCategoryItems.indexOf(subCategoryRow, 0);
        if (this.state.subCategoryItems[index].Id == "-1") {
            this.state.subCategoryItems.splice(index, 1);
        } else {
            this.state.subCategoryItems[index].editEnabled = false;
            this.state.subCategoryItems[index].Title = this.state.subCategory[index].Title;
            this.state.subCategoryItems[index].SubCategoryOrder = this.state.subCategory[index].SubCategoryOrder;
        }
        this.setState({ subCategoryItems: this.state.subCategoryItems, adminControlDisabled: false });
    }
    private onDeleteSCButtonClick = async (subCategoryRow: ISubCategory) => {
        console.log("main Here");
        let index = this.state.subCategoryItems.indexOf(subCategoryRow, 0);
        this.state.subCategoryItems[index].spinner = true;
        this.setState({ subCategoryItems: this.state.subCategoryItems });
        await this.subCategoryProvider.removeSubCategoryListItem(subCategoryRow);

        this.state.subCategory.splice(index, 1);
        this.state.subCategoryItems.splice(index, 1);
        this.setState({ subCategoryItems: this.state.subCategoryItems, subCategory: this.state.subCategory });
    }
    /* 
    Fires on Delete button click. 
    Deletes item from thr list. Can only trigger on existing item. 
    It also deletes the item from both mainCategory and maincategoryitem List Object and updates the state.
    **/
    private onShowHideSCButtonClick = async (subCategoryRow: ISubCategory) => {
        let status = '';
        if (subCategoryRow.Status == this.constants.comparingStrings.visible) {
            status = this.constants.comparingStrings.hidden;
        } else {
            status = this.constants.comparingStrings.visible;
        }
        let index = this.state.subCategoryItems.indexOf(subCategoryRow, 0);
        this.state.subCategoryItems[index].spinner = true;
        this.setState({ subCategoryItems: this.state.subCategoryItems });
        await this.subCategoryProvider.updateSubCategory({ Status: status }, subCategoryRow.Id);
        this.state.subCategory[index].Status = status;
        this.state.subCategoryItems[index].Status = status;
        this.state.subCategoryItems[index].spinner = false;
        this.setState({ subCategoryItems: this.state.subCategoryItems, subCategory: this.state.subCategory });
    }
    /* 
    Fires on save button click. 
    Adds a new item or update the item in the Main Category list.
    It also updates mainCategory and maincategoryitem List Object and then sory them to have correct order and updates the state.
    **/
    private onSaveSCButtonClick = async (subCategoryRow: ISubCategory) => {
        const index = this.state.subCategoryItems.indexOf(subCategoryRow, 0);
        this.state.subCategoryItems[index].spinner = true;
        this.setState({ subCategoryItems: this.state.subCategoryItems });
        const newItem: ISubCategory = await this.subCategoryProvider.saveSubCategoryListItem(subCategoryRow);
        if (newItem && newItem.Id) {
            this.state.subCategory.push({ Id: newItem.Id, Title: this.state.subCategoryItems[index].Title.trim(), SubCategoryOrder: this.state.subCategoryItems[index].SubCategoryOrder.trim(), Status: this.state.subCategoryItems[index].Status });
            this.state.subCategoryItems[index].Id = newItem.Id;
            if (this.state.subCategory.length == 1) {
                this.state.subCategory.push({ Title: 'All', Id: '0', SubCategoryOrder: '-200', Status: this.constants.comparingStrings.visible });
            }
        } else {
            this.state.subCategory[index].Title = this.state.subCategoryItems[index].Title.trim();
            this.state.subCategory[index].SubCategoryOrder = this.state.subCategoryItems[index].SubCategoryOrder.toString().trim();
        }
        this.state.subCategory.sort((a, b) => (+a.SubCategoryOrder > +b.SubCategoryOrder) ? 1 : -1);
        this.state.subCategoryItems[index].Title = this.state.subCategoryItems[index].Title.trim();
        this.state.subCategoryItems[index].SubCategoryOrder = this.state.subCategoryItems[index].SubCategoryOrder.toString().trim();
        this.state.subCategoryItems[index].spinner = false;
        this.state.subCategoryItems[index].editEnabled = false;
        this.state.subCategoryItems.sort((a, b) => (+a.SubCategoryOrder > +b.SubCategoryOrder) ? 1 : -1);
        this.setState({ subCategoryItems: this.state.subCategoryItems, subCategory: this.state.subCategory, adminControlDisabled: false });
    }
    /* 
    Fires on textbox change event of the textboxes. 
    Updates the state accordingly.
    isTitle parameter will be true if the changes are made for catefory textbox else it will be false.
    **/
    private onSCTextBoxChange = (newValue: string, subCategoryRow: ISubCategory, isTitle: boolean) => {
        const index = this.state.subCategoryItems.indexOf(subCategoryRow, 0);
        if (isTitle) {
            this.state.subCategoryItems[index].Title = newValue;
        } else {
            this.state.subCategoryItems[index].SubCategoryOrder = newValue;
        }
        this.setState({ subCategoryItems: this.state.subCategoryItems });
    }
    /* 
    Fires on Add a New Category button click. 
    Adds a new item to maincategorylist item state variable with id set to -1. If an extra item already there it wont add another new item. 
    if its canceled it will delete the extra added item. 
    **/
    private onAddSCButtonClick = () => {
        const newItem = this.state.subCategoryItems.filter(i => i.Id == "-1");
        if (newItem && newItem.length == 0) {
            this.state.subCategoryItems.push({ Id: "-1", Title: '', SubCategoryOrder: '', editEnabled: true, Status: this.constants.comparingStrings.visible });
            this.setState({ subCategoryItems: this.state.subCategoryItems, adminControlDisabled: true });
        }
    }
    //#endregion

    private _renderTabs = (): JSX.Element => {
        let visibleMainCat = this.state.mainCategory.filter((maincat => maincat.Status == this.constants.comparingStrings.visible));
        return (
            <Pivot
                aria-label="Basic Pivot Example"
                linkSize={PivotLinkSize.normal}
                linkFormat={PivotLinkFormat.tabs}
                styles={this.pivotStyle}>
                {visibleMainCat.map((mainCat, index) => {
                    if (index == 0) {
                        return (
                            <PivotItem
                                headerText={mainCat.Title}
                                headerButtonProps={{ 'data-order': 1, 'data-title': 'mainCat.Title' }}
                                className={styles.mainTab}>
                                <SubCategory
                                    context={this.props.context}
                                    mainCategory={mainCat}
                                    searchText={this.state.searchText}
                                    subCategory={this.state.subCategory}
                                    isAdmin={this.state.isAdmin}

                                    mainCategoryItems={this.state.mainCategoryItems}
                                    subCategoryItems={this.state.subCategoryItems}
                                    wikiId={this.props.wikiId}
                                >
                                </SubCategory>
                            </PivotItem>);
                    } else {
                        return (
                            <PivotItem
                                headerText={mainCat.Title}
                                className={styles.mainTab}>
                                <SubCategory
                                    context={this.props.context}
                                    mainCategory={mainCat}
                                    searchText={this.state.searchText}
                                    subCategory={this.state.subCategory}
                                    isAdmin={this.state.isAdmin}

                                    mainCategoryItems={this.state.mainCategoryItems}
                                    subCategoryItems={this.state.subCategoryItems}
                                    wikiId={this.props.wikiId}>
                                </SubCategory>
                            </PivotItem>);
                    }
                }
                )}
                {this.state && this.state.isAdmin ?
                    <PivotItem
                        headerText="Admin"
                        className={styles.mainTab}>
                        <AdminCenter
                            context={this.props.context}
                            mainCategoryItems={this.state.mainCategoryItems}
                            subCategoryItems={this.state.subCategoryItems}
                            adminGroupID={this.state.adminGroupId}
                            adminControlDisabled={this.state.adminControlDisabled} // when an action is taking place it disables all the other button on the form
                            // Main Category Functions
                            onEditMCButtonClick={this.onEditMCButtonClick}
                            onCancelMCButtonClick={this.onCancelMCButtonClick}
                            onSaveMCButtonClick={this.onSaveMCButtonClick}
                            onAddMCButtonClick={this.onAddMCButtonClick}
                            onMCTextBoxChange={this.onMCTextBoxChange}
                            onShowHideMCButtonClick={this.onShowHideMCButtonClick}
                            onDeleteMCButtonClick={this.onDeleteMCButtonClick}
                            // Sub Category Functions
                            onEditSCButtonClick={this.onEditSCButtonClick}
                            onSaveSCButtonClick={this.onSaveSCButtonClick}
                            onCancelSCButtonClick={this.onCancelSCButtonClick}
                            onAddSCButtonClick={this.onAddSCButtonClick}
                            onSCTextBoxChange={this.onSCTextBoxChange}
                            onShowHideSCButtonClick={this.onShowHideSCButtonClick}
                            onDeleteSCButtonClick={this.onDeleteSCButtonClick}>
                        </AdminCenter>
                    </PivotItem> : ''});
            </Pivot >
        );
    }
    public render(): React.ReactElement<IMainCategoryProps> {
        return (
            <div className={styles.sohoWikiVault}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.mainCategory}>
                            <div className={styles.searchRow}>
                                <div className={styles.columnLeft}></div>
                                <div className={styles.column}>
                                    <TextField
                                        className={styles.controlStyle}
                                        placeholder="Search"
                                        onChange={this._onChangeText}></TextField>
                                </div>
                            </div>
                            <div className={styles.pivot}>
                                {this._renderTabs()}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}