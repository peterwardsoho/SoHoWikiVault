import * as React from 'react';
// Component
import { IAdminCenterProps, IAdminCenterState, WikiEditForm, ListViewStyle } from '../';
import styles from '../sohoWikiVault/SohoWikiVault.module.scss';
// Office UI Fabric Imports
import {
    Pivot, PivotItem, PivotLinkSize, IPivotStyles,
    IconButton, TextField, DefaultButton, Spinner, Label, Dialog,
    DialogFooter, PrimaryButton, DialogType
} from 'office-ui-fabric-react';
// Constants
import { Constants } from '../../common';
//Model
import { IMainCategory, ISubCategory } from '../../model';
import { CommonDataProvider } from '../../dataProvider/index';


export class AdminCenter extends React.Component<IAdminCenterProps, IAdminCenterState> {
    public state: IAdminCenterState;
    private constants: Constants;
    private officeControlStyle: ListViewStyle;
    private commonDataProvider: CommonDataProvider;
    // Pivot Style
    private pivotStyle: IPivotStyles = {
        link: {},
        linkIsSelected: {},
        root: {
            display: 'flex', flexWrap: 'wrap'
        },
        count: {},
        icon: {},
        linkContent: {},
        text: {}
    };
    public async componentWillMount() {
        this.constants = new Constants();
        this.officeControlStyle = new ListViewStyle();
        this.setState({
            catError: '',
            catorderError: '',
            subCatError: '',
            subCatorderError: '',
            dialogSubTitle: '',
            dialogTitle: '',
            showDeleteDialog: false,
            disableDialogButtons: false,
        });
    }
    //#region Main category
    private onMCTextBoxChange = (row: IMainCategory, isTitle: boolean, event) => {
        if (event.target.value) {
            if (isTitle) {
                this.setState({ catError: "" });
            } else {
                this.setState({ catorderError: "" });
            }
        }
        this.props.onMCTextBoxChange(event.target.value, row, isTitle);
    }
    private onSaveMCButtonClick = (row: IMainCategory) => {
        if (!this.state.catError && !this.state.catorderError) {
            if (row.Title && row.CategoryOrder) {
                this.setState({ catError: "", catorderError: "" });
                this.props.onSaveMCButtonClick(row);
            } else {
                if (row.Title) {
                    this.setState({ catError: "" });
                } else {
                    this.setState({ catError: "Required" });
                }
                if (row.CategoryOrder) {
                    this.setState({ catorderError: "" });
                } else {
                    this.setState({ catorderError: "Required" });
                }
            }
        }
    }
    private onCancelMCButtonClick = (row: IMainCategory) => {
        this.setState({ catError: "", catorderError: "" });
        this.props.onCancelMCButtonClick(row);
    }
    private getMCTextLengthErrorMessage = (value: any): string => {
        let error: string = '';
        if (value && value.length > 20) {
            //  error = "";
            this.setState({ catError: "Can not take more than 20 characters" });
        } else {
            this.setState({ catError: "" });
        }
        return error;
    }
    // settin Error Message
    private getMCNumberErrorMessage = (value: any): string => {
        let error: string = '';
        if (value) {
            if (isNaN(+value)) {
                // error = "Only numbers greater than 0";
                this.setState({ catorderError: "Only numbers greater than 0" });
            } else {
                if (value <= 0) {
                    //   error = "Only numbers greater than 0";
                    this.setState({ catorderError: "Only numbers greater than 0" });
                }
                else {
                    this.setState({ catorderError: "" });
                }
            }
        }
        return error;
    }
    private delteMCButtonClick = (row: IMainCategory) => {
        this.setState({
            dialogTitle: 'Delete Main Category',
            dialogSubTitle: 'Are you sure you want to delete the Category? Once you delete it, all wiki related to the category will also be deleted.',
            showDeleteDialog: true,
            currentMCBeingDeleted: row,
            currentSCBeingDeleted: null,
            disableDialogButtons: false
        });
    }
    //#endregion
    //#region Sub Caregory
    private onSCTextBoxChange = (row: ISubCategory, isTitle: boolean, event) => {
        if (event.target.value) {
            if (isTitle) {
                this.setState({ subCatError: "" });
            } else {
                this.setState({ subCatorderError: "" });
            }
        }
        this.props.onSCTextBoxChange(event.target.value, row, isTitle);
    }
    private onSaveSCButtonClick = (row: ISubCategory) => {
        if (!this.state.subCatError && !this.state.subCatorderError) {
            if (row.Title && row.SubCategoryOrder) {
                this.setState({ subCatError: "", subCatorderError: "" });
                this.props.onSaveSCButtonClick(row);
            } else {
                if (row.Title) {
                    this.setState({ subCatError: "" });
                } else {
                    this.setState({ subCatError: "Required" });
                }
                if (row.SubCategoryOrder) {
                    this.setState({ subCatorderError: "" });
                } else {
                    this.setState({ subCatorderError: "Required" });
                }
            }
        }
    }
    private onCancelSCButtonClick = (row: ISubCategory) => {
        this.setState({ subCatError: "", subCatorderError: "" });
        this.props.onCancelSCButtonClick(row);
    }
    private delteSCButtonClick = (row: ISubCategory) => {
        this.setState({
            dialogTitle: 'Delete Sub Category',
            dialogSubTitle: 'Are you sure you want to delete the Sub Category? Once you delete it, all wiki related to the sub category will also be deleted.',
            showDeleteDialog: true,
            currentMCBeingDeleted: null,
            currentSCBeingDeleted: row,
            disableDialogButtons: false
        });
        console.log(this.state);
    }
    private getSCTextLengthErrorMessage = (value: any): string => {
        let error: string = '';
        if (value && value.length > 20) {
            //  error = "";
            this.setState({ subCatError: "Can not take more than 20 characters" });
        } else {
            this.setState({ subCatError: "" });
        }
        return error;
    }
    // settin Error Message
    private getSCNumberErrorMessage = (value: any): string => {
        let error: string = '';
        if (value) {
            if (isNaN(+value)) {
                // error = "Only numbers greater than 0";
                this.setState({ subCatorderError: "Only numbers greater than 0" });
            } else {
                if (value <= 0) {
                    //   error = "Only numbers greater than 0";
                    this.setState({ subCatorderError: "Only numbers greater than 0" });
                }
                else {
                    this.setState({ subCatorderError: "" });
                }
            }
        }
        return error;
    }
    //#endregion
    private onDeleteNoButtonClick = () => {
        this.setState({
            showDeleteDialog: false,
            dialogTitle: '',
            dialogSubTitle: '',
            currentMCBeingDeleted: null,
            currentSCBeingDeleted: null,
            disableDialogButtons: false
        });
    }
    private onDeleteYesButtonClick = async () => {
        this.setState({
            disableDialogButtons: true
        });

        if (this.state.currentMCBeingDeleted) {
            await this.props.onDeleteMCButtonClick(this.state.currentMCBeingDeleted);
        }
        if (this.state.currentSCBeingDeleted) {
            await this.props.onDeleteSCButtonClick(this.state.currentSCBeingDeleted);
        }
        this.setState({
            showDeleteDialog: false,
            dialogTitle: '',
            dialogSubTitle: '',
            currentMCBeingDeleted: null,
            currentSCBeingDeleted: null,
            disableDialogButtons: false
        });
    }
    private renderMainCatTableRows = (): JSX.Element => {
        return (
            <tbody>
                {this.props.mainCategoryItems.map((row, index) => {
                    if (row.editEnabled) {
                        return (
                            <tr key={row.Id} className={row.Status == this.constants.comparingStrings.hidden ? styles.hidden : ''}>
                                <td> </td>
                                <td> </td>
                                <td>
                                    {row.spinner ?
                                        <Spinner ariaLive="assertive" labelPosition="left" /> :
                                        <div>
                                            <div>
                                                <IconButton
                                                    className={styles.editIcon}
                                                    iconProps={{ iconName: 'Save' }}
                                                    title="Save Main Cat" ariaLabel="MCSave"
                                                    onClick={() => { this.onSaveMCButtonClick(row); }} />
                                            </div>
                                            <IconButton
                                                className={styles.editIcon}
                                                iconProps={{ iconName: 'Cancel' }}
                                                title="Cancel Main Cat"
                                                ariaLabel="MCCancel"
                                                onClick={() => { this.onCancelMCButtonClick(row); }} />
                                        </div>
                                    }
                                </td>
                                <td>
                                    <TextField
                                        placeholder="Type here"
                                        value={row.Title}
                                        onChange={(e) => this.onMCTextBoxChange(row, true, e)}
                                        onGetErrorMessage={this.getMCTextLengthErrorMessage}
                                        errorMessage={this.state.catError}>
                                    </TextField>
                                </td>
                                <td>
                                    <TextField
                                        placeholder="Type here"
                                        value={String(row.CategoryOrder)}
                                        onChange={(e) => this.onMCTextBoxChange(row, false, e)}
                                        onGetErrorMessage={this.getMCNumberErrorMessage}
                                        errorMessage={this.state.catorderError}>
                                    </TextField>
                                </td>
                                <td>{row.CategoryGroupName ? <a href={row.CategoryGroupName.Url} target="_blank" data-interception="off">{row.CategoryGroupName.Description}</a> : ''} </td>
                            </tr>
                        );
                    } else {
                        return (
                            <tr key={row.Id} className={row.Status == this.constants.comparingStrings.hidden ? styles.hidden : ''}>
                                <td>
                                    {row.spinner ?
                                        <Spinner ariaLive="assertive" labelPosition="left" /> :
                                        <IconButton
                                            disabled={this.props.adminControlDisabled}
                                            className={styles.deleteIcon}
                                            iconProps={{ iconName: 'Delete' }}
                                            title={'Delete'}
                                            ariaLabel={'SCDelete'}
                                            onClick={() => { this.delteMCButtonClick(row); }}
                                        />}
                                </td>
                                <td>
                                    {row.spinner ?
                                        <Spinner ariaLive="assertive" labelPosition="left" /> :
                                        <IconButton
                                            disabled={this.props.adminControlDisabled}
                                            className={row.Status == this.constants.comparingStrings.visible ? styles.editIcon : styles.deleteIcon}
                                            iconProps={{ iconName: row.Status == this.constants.comparingStrings.visible ? 'View' : 'Hide' }}
                                            title={row.Status == this.constants.comparingStrings.visible ? 'Visible' : 'Hidden'}
                                            ariaLabel={row.Status == this.constants.comparingStrings.visible ? 'MCVisible' : 'MCHidden'}
                                            onClick={() => { this.props.onShowHideMCButtonClick(row); }} />}
                                </td>
                                <td>
                                    <IconButton
                                        disabled={this.props.adminControlDisabled}
                                        className={styles.editIcon}
                                        iconProps={{ iconName: 'Edit' }}
                                        title="Edit Main Cat"
                                        ariaLabel="MCEdit"
                                        onClick={() => { this.props.onEditMCButtonClick(row); }} />
                                </td>
                                <td>{row.Title}</td>
                                <td>{row.CategoryOrder} </td>
                                <td>{row.CategoryGroupName ? <a href={row.CategoryGroupName.Url} target="_blank" data-interception="off">{row.CategoryGroupName.Description}</a> : ''} </td>
                            </tr >
                        );
                    }
                })
                }
            </tbody >
        );
    }
    private renderSubCatTableRows = (): JSX.Element => {
        return (
            <tbody>
                {this.props.subCategoryItems.map((row, index) => {
                    if (row.editEnabled) {
                        return (
                            <tr key={row.Id} className={row.Status == this.constants.comparingStrings.hidden ? styles.hidden : ''}>
                                <td> </td>
                                <td> </td>
                                <td>
                                    {row.spinner ?
                                        <Spinner ariaLive="assertive" labelPosition="left" /> :
                                        <div>
                                            <div>
                                                <IconButton
                                                    className={styles.editIcon}
                                                    iconProps={{ iconName: 'Save' }}
                                                    title="Save Sub Cat" ariaLabel="SaveSC"
                                                    onClick={() => { this.onSaveSCButtonClick(row); }} />
                                            </div>
                                            <IconButton
                                                className={styles.editIcon}
                                                iconProps={{ iconName: 'Cancel' }}
                                                title="Cancel Sub Cat"
                                                ariaLabel="CancelSC"
                                                onClick={() => { this.onCancelSCButtonClick(row); }} />
                                        </div>
                                    }
                                </td>
                                <td>
                                    <TextField
                                        placeholder="Type here"
                                        value={row.Title}
                                        onChange={(e) => this.onSCTextBoxChange(row, true, e)}
                                        onGetErrorMessage={this.getSCTextLengthErrorMessage}
                                        errorMessage={this.state.subCatError}>
                                    </TextField>
                                </td>
                                <td>
                                    <TextField
                                        placeholder="Type here"
                                        value={String(row.SubCategoryOrder)}
                                        onChange={(e) => this.onSCTextBoxChange(row, false, e)}
                                        onGetErrorMessage={this.getSCNumberErrorMessage}
                                        errorMessage={this.state.subCatorderError}>
                                    </TextField>
                                </td>
                            </tr>
                        );
                    } else {
                        return (
                            <tr key={row.Id} className={row.Status == this.constants.comparingStrings.hidden ? styles.hidden : ''}>
                                <td>
                                    {row.spinner ?
                                        <Spinner ariaLive="assertive" labelPosition="left" /> :
                                        <IconButton
                                            disabled={this.props.adminControlDisabled}
                                            className={styles.deleteIcon}
                                            iconProps={{ iconName: 'Delete' }}
                                            title={'Delete'}
                                            ariaLabel={'SCDelete'}
                                            onClick={() => { this.delteSCButtonClick(row); }}
                                        />}
                                </td>
                                <td>
                                    {row.spinner ?
                                        <Spinner ariaLive="assertive" labelPosition="left" /> :
                                        <IconButton
                                            disabled={this.props.adminControlDisabled}
                                            className={row.Status == this.constants.comparingStrings.visible ? styles.editIcon : styles.deleteIcon}
                                            iconProps={{ iconName: row.Status == this.constants.comparingStrings.visible ? 'View' : 'Hide' }}
                                            title={row.Status == this.constants.comparingStrings.visible ? 'Visible' : 'Hidden'}
                                            ariaLabel={row.Status == this.constants.comparingStrings.visible ? 'SCVisible' : 'SCHidden'}
                                            onClick={() => { this.props.onShowHideSCButtonClick(row); }} />}
                                </td>
                                <td>
                                    <IconButton
                                        disabled={this.props.adminControlDisabled}
                                        className={styles.editIcon}
                                        iconProps={{ iconName: 'Edit' }}
                                        title="Edit Sub Cat"
                                        ariaLabel="SCEdit"
                                        onClick={() => { this.props.onEditSCButtonClick(row); }} />
                                </td>
                                <td>{row.Title}</td>
                                <td>{row.SubCategoryOrder} </td>
                            </tr >
                        );
                    }
                })
                }
            </tbody >
        );
    }
    public render(): React.ReactElement<IAdminCenterProps> {
        return (
            <div className={styles.subCategory}>
                <div className={styles.subCategorydiv}>
                    <div className={styles.AddAdminLinkstyle}><a target="_blank" data-interception="off" href={this.props.adminGroupID}>Click here to add an Admin</a></div>
                    <Pivot aria-label="Basic Pivot Example" linkSize={PivotLinkSize.normal} styles={this.pivotStyle}>
                        <PivotItem
                            headerText="Category"
                            style={{ padding: 10 }}
                            headerButtonProps={{
                                'data-order': 1,
                                'data-title': 'Category',
                            }}>
                            <DefaultButton
                                disabled={this.props.adminControlDisabled || (this.props.mainCategoryItems && this.props.mainCategoryItems.length >= 12)}
                                onClick={() => { this.props.onAddMCButtonClick(); }}
                                text="Add a Category" />
                            {this.props.mainCategoryItems && this.props.mainCategoryItems.length >= 12 ?
                                <div>
                                    <Label
                                        className={styles.columnErrorlabelstyle}>
                                        Can not create more than 12 Categories.
                                    </Label>
                                </div> : ''}
                            {this.props.mainCategoryItems && this.props.mainCategoryItems.length > 0 ?
                                <div className={styles.divStyle}>
                                    <table className="ms-Table" cellSpacing="0">
                                        <thead>
                                            <tr>
                                                <th>Delete</th>
                                                <th>Status</th>
                                                <th>Edit</th>
                                                <th>Category</th>
                                                <th>Category Order</th>
                                                <th>Category Group Name</th>
                                            </tr>
                                        </thead>
                                        {this.renderMainCatTableRows()}
                                    </table>
                                </div>
                                : ''};
                        </PivotItem>);
                        <PivotItem headerText="Sub Category" style={{ padding: 10 }}>
                            <DefaultButton
                                disabled={this.props.adminControlDisabled || (this.props.subCategoryItems && this.props.subCategoryItems.length >= 12)}
                                onClick={() => { this.props.onAddSCButtonClick(); }}
                                text="Add a Sub Category" />
                            {this.props.subCategoryItems && this.props.subCategoryItems.length >= 12 ?
                                <div>
                                    <Label
                                        className={styles.columnErrorlabelstyle}>
                                        Can not create more than 12 Categories.
                                    </Label>
                                </div> : ''}
                            {this.props.subCategoryItems && this.props.subCategoryItems.length > 0 ?
                                <div className={styles.divStyle}>
                                    <table className="ms-Table" cellSpacing="0">
                                        <thead>
                                            <tr>
                                                <th>Delete</th>
                                                <th>Status</th>
                                                <th>Edit</th>
                                                <th>Sub Category</th>
                                                <th>Sub Category Order</th>
                                            </tr>
                                        </thead>
                                        {this.renderSubCatTableRows()}
                                    </table>
                                </div>
                                : ''};
                        </PivotItem>
                        <PivotItem headerText="Add a Wiki" style={{ padding: 10 }}>
                            <WikiEditForm
                                context={this.props.context}
                                mainCategoryItems={this.props.mainCategoryItems}
                                subCategoryItems={this.props.subCategoryItems}
                                wikiItem={null}>
                            </WikiEditForm>
                        </PivotItem>
                    </Pivot>
                    <Dialog
                        hidden={!this.state.showDeleteDialog}
                        dialogContentProps={{
                            type: DialogType.largeHeader,
                            title: this.state.dialogTitle,
                            subText: this.state.dialogSubTitle,
                            styles: this.officeControlStyle.errorHeaderStyle
                        }}
                        modalProps={{
                            titleAriaId: 'modalDFId',
                            subtitleAriaId: 'modalDFSubID',
                            isBlocking: true,
                        }}>
                        <DialogFooter>
                            <DefaultButton text="No"
                                onClick={this.onDeleteNoButtonClick}
                            />
                            <PrimaryButton
                                text="Yes"
                                styles={this.officeControlStyle.deleteButtonStyle}
                                onClick={this.onDeleteYesButtonClick}
                            />
                        </DialogFooter>
                    </Dialog>
                </div>
            </div>
        );
    }
}