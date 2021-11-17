import * as React from 'react';
// Component
import { IWikiEditFormProps, IWikiEditFormState } from '../';
//import styles from '../sohoWikiVault/SohoWikiVault.module.scss';
import styles from './WikiEditForm.module.scss';
// Office UI Fabric Imports
import {
    Label, TextField, Dropdown, DatePicker, DayOfWeek, IDropdownOption,
    Checkbox, PrimaryButton, DefaultButton, Spinner
} from 'office-ui-fabric-react';
// Constants
import { Constants, BusinessLogic } from '../../common';
//Model
import { IDropdownValues, IMainCategory, IWikiVault, IWikiVaultPassword } from '../../model';
//PnP controls
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// ReactQuill
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css';
//DataProvider
import { WikiEditFormDataProvider, CommonDataProvider } from '../../dataProvider';


export class WikiEditForm extends React.Component<IWikiEditFormProps, IWikiEditFormState> {
    public state: IWikiEditFormState;
    private constants: Constants;
    private commonDataProvider: CommonDataProvider;
    private wikiEditFormDataProvider: WikiEditFormDataProvider;

    public async componentWillMount() {
        this.constants = new Constants();
        this.commonDataProvider = new CommonDataProvider(this.props.context);
        this.wikiEditFormDataProvider = new WikiEditFormDataProvider(this.props.context);
        this.state = { ...this.constants.wikiformInitialState };
        const subscriptionType: IDropdownValues[] = await this.wikiEditFormDataProvider.getFieldDDValue();
        if (subscriptionType && subscriptionType.length > 0) {
            subscriptionType[0].Choices.forEach(choice => {
                this.state.wikiSubscriptionTypeChoices.push({ key: choice, text: choice });
            });
        }
        this.props.mainCategoryItems.forEach(element => {
            this.state.wikiMainCategoryChoices.push({ key: element.Id, text: element.Title });
        });
        this.props.subCategoryItems.forEach(element => {
            this.state.wikiSubCategoryChoices.push({ key: element.Id, text: element.Title });
        });
        this.setState({
            wikiMainCategoryChoices: this.state.wikiMainCategoryChoices,
            wikiSubCategoryChoices: this.state.wikiSubCategoryChoices,
            wikiSubscriptionTypeChoices: this.state.wikiSubscriptionTypeChoices,
            selectedSubscriptionType: "None"
        });
        if (this.props.wikiItem) {
            this.setStateonEditClick(this.props.wikiItem);
        }
    }
    private setStateonEditClick = async (wikiItem: IWikiVault) => {
        const filter = `&$filter=SohoWikiVaultId eq ${wikiItem.Id}`;
        const wikiPassword: IWikiVaultPassword[] = await this.commonDataProvider.getListItems(this.constants.WikiVaultPassword.listName, this.constants.WikiVaultPassword.selectQuery, '', filter, '', '');
        this.state.showNew = false;
        this.state.wikiLabelName = wikiItem.Title;
        this.state.selectedMainCategory = wikiItem.MainCategory.Id;
        wikiItem.SubCategory.forEach(subcat => {
            this.state.selectedSubCategory.push(subcat.Id);
        });
        this.state.selectedSubscriptionType = wikiItem.SubscriptionType;
        this.state.briefLabelDesc = wikiItem.BriefLabelDescription;
        this.state.subscriptionExpirationDate = new Date(wikiItem.SubscriptionExpirationDate);
        this.state.isSubscriptionPaid = wikiItem.IsSubscriptionPaid;
        this.state.canThePasswordBeCutAndPasted = wikiItem.CanThePasswordBeCutAndPasted;
        if (wikiItem.URL && wikiItem.URL.Url)
            this.state.url = wikiItem.URL.Url;

        this.state.istheServiceBeingUsed = wikiItem.IsServiceBeingUsed;
        if (wikiItem.PageOwner && wikiItem.PageOwner.length > 0) {
            wikiItem.PageOwner.forEach(owner => {
                this.state.pageOwners.push({ id: +owner.Id, text: `${owner.FirstName} ${owner.LastName}` });
                this.state.defaultSelectedPageOwners.push(`${owner.FirstName} ${owner.LastName}`);
            });
        }
        this.state.wikiId = wikiItem.Id;
        this.state.userId = wikiPassword[0].UserId;
        this.state.password = wikiPassword[0].SitePassword;
        this.state.restrictedLabelDesc = wikiPassword[0].RestrictedLabelDescription;
        this.setState(this.state);
    }
    private onLabelNameChange = (ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
        this.setState({ wikiLabelName: newValue });
    }
    private onCategoryChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        this.setState({ selectedMainCategory: item.key as string });
    }
    private onSubCategoryChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        if (item) {
            const itemKey = item.key as string;
            if (item.selected) {
                this.state.selectedSubCategory.push(itemKey);
            } else {
                this.state.selectedSubCategory = this.state.selectedSubCategory.filter(key => key !== item.key);
            }
            this.forceUpdate();
        }
    }
    private onSubscriptionChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        this.setState({ selectedSubscriptionType: item.key as string });
    }
    private onSubscriptionDateChange = (date: Date): void => {
        this.setState({ subscriptionExpirationDate: date });
    }
    private onisSubscriptionPaidChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) => {
        this.setState({ isSubscriptionPaid: isChecked });
        if (!isChecked) {
            this.setState({ selectedSubscriptionType: "None" });
        }
    }
    private oncanThePasswordBeCutAndPasted = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) => {
        this.setState({ canThePasswordBeCutAndPasted: isChecked });
    }
    private onIstheServiceBeingUsedChange = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) => {
        this.setState({ istheServiceBeingUsed: isChecked });
    }
    private onURLChange = (ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
        this.setState({ url: newValue });
    }
    private onUserIDChange = (ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
        this.setState({ userId: newValue });
    }
    private onPasswordChange = (ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
        this.setState({ password: newValue });
    }
    private getPeoplePickerItems = (items: any[]) => {
        const defaultOwners = [];
        console.log(items);
        items.forEach(element => {
            defaultOwners.push(element.text);
        });
        this.setState({ pageOwners: items, defaultSelectedPageOwners: defaultOwners });
        // items.forEach(element => {
        //     // const checkUser: string[] = this.state.defaultSelectedPageOwners.filter(name => name == element.text);
        //     // if (checkUser.length == 0) {
        //     this.state.defaultSelectedPageOwners.push(element.text);
        //     // }
        // });
        // this.setState({ pageOwners: items, defaultSelectedPageOwners: [] });
        console.log(this.state.pageOwners);
        console.log(this.state.defaultSelectedPageOwners);
    }
    //ButtonClick
    // on Clear Button Click
    private onClearClick = async () => {
        this.setState({
            controlDisabled: false,
            showNew: true,
            wikiLabelName: '',
            selectedMainCategory: '',
            selectedSubscriptionType: '',
            selectedSubCategory: [],
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
            labelError: ''
        });
    }
    // on Save Click
    private onSaveClick = async () => {
        if (!this.state.labelError && !this.state.userIDError && !this.state.sitePasswordError) {
            this.setState({ controlDisabled: true });
            const error: boolean = this.checkForError();
            if (!error) {
                if (this.state.subscriptionExpirationDate && this.state.subscriptionExpirationDate < new Date()) {
                    this.setState({ error: 'Expiration Date can not be in the past.' });
                } else {
                    this.setState({ error: '' });
                    const mainCategory: IMainCategory[] = this.props.mainCategoryItems.filter(item => item.Id == this.state.selectedMainCategory);
                    if (mainCategory && mainCategory.length > 0) {
                        const ids = await this.wikiEditFormDataProvider.onSaveClick(this.state, mainCategory[0]);
                        if (ids) {
                            this.setState({
                                showNew: false,
                                wikiId: ids,
                                error: ''
                            });
                        }
                    }
                }
            }
            this.setState({ controlDisabled: false });
        }
    }
    private onUpdateClick = async () => {
        if (!this.state.labelError && !this.state.userIDError && !this.state.sitePasswordError) {
            this.setState({ controlDisabled: true });
            const error: boolean = this.checkForError();
            if (!error) {
                const mainCategory: IMainCategory[] = this.props.mainCategoryItems.filter(item => item.Id == this.state.selectedMainCategory);
                if (mainCategory && mainCategory.length > 0) {
                    await this.wikiEditFormDataProvider.onUpdateClick(this.state, mainCategory[0]);
                    this.setState({ error: '' });
                }
            }
            this.setState({ controlDisabled: false });
        }
    }
    private checkForError = (): boolean => {
        let error: boolean = false;
        if (this.state.url) {
            try {
                const url = new URL(this.state.url);
                this.setState({ error: '' });
            } catch (_) {
                this.setState({ error: 'Incorrect URL. Please provide a valid URL.' });
                error = true;
                return error;
            }
        }
        if (!this.state.wikiLabelName || !this.state.selectedMainCategory ||
            this.state.selectedSubCategory.length <= 0 || !this.state.selectedSubscriptionType ||
            this.state.pageOwners.length <= 0) {
            this.setState({ error: 'Please enter all required Fields (*)' });
            error = true;
            return error;
        } else {
            this.setState({ error: '' });
        }
        if (this.state.isSubscriptionPaid) {
            if (this.state.selectedSubscriptionType == "None" || !this.state.subscriptionExpirationDate) {
                this.setState({ error: `With paid subscription, Subscription Type cannot be 'None' and Subscription expiration date is required.` });
                error = true;
                return error;
            } else {
                this.setState({ error: '' });
            }
        }
        return error;
    }
    private getLebelErrorMessage = (value: any): string => {
        if (value && value.length > 255) {
            this.setState({ labelError: "Can not take more than 254 characters" });
        } else {
            this.setState({ labelError: "" });
        }
        return '';
    }
    private getUserErrorMessage = (value: any): string => {
        if (value && value.length > 255) {
            this.setState({ userIDError: "Can not take more than 254 characters" });
        } else {
            this.setState({ userIDError: "" });
        }
        return '';
    }
    private getPasswordErrorMessage = (value: any): string => {
        if (value && value.length > 255) {
            this.setState({ sitePasswordError: "Can not take more than 254 characters" });
        } else {
            this.setState({ sitePasswordError: "" });
        }
        return '';
    }
    public render(): React.ReactElement<IWikiEditFormProps> {
        return (
            <div className={styles.wikiEditForm}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <form>
                            <div className={styles.mainformdiv}>
                                {this.state && this.state.error ?
                                    <div className={styles.formRow}>
                                        <Label className={styles.columnErrorlabelstyle}>{this.state.error}</Label>
                                    </div>
                                    : ''
                                }
                                <div className={styles.formRow}>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle} required={true}>Label Name</Label>
                                        <TextField
                                            placeholder="Type here"
                                            value={this.state.wikiLabelName}
                                            onChange={this.onLabelNameChange}
                                            disabled={this.state.controlDisabled}
                                            onGetErrorMessage={this.getLebelErrorMessage}
                                            errorMessage={this.state.labelError} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle} required={true}>Main Category</Label>
                                        <Dropdown
                                            selectedKey={this.state.selectedMainCategory}
                                            options={this.state.wikiMainCategoryChoices}
                                            placeholder="Select a Category"
                                            onChange={this.onCategoryChange}
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle} required={true}>Sub Category</Label>
                                        <Dropdown
                                            placeholder="Select a Sub-Category"
                                            selectedKeys={[...this.state.selectedSubCategory]}
                                            multiSelect={true}
                                            options={this.state.wikiSubCategoryChoices}
                                            onChange={this.onSubCategoryChange}
                                            disabled={this.state.controlDisabled}
                                        />
                                    </div>
                                </div>
                                <div className={styles.formRow}>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}>URL</Label>
                                        <TextField
                                            placeholder="Type here"
                                            value={this.state.url}
                                            onChange={this.onURLChange}
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}></Label>
                                        <Checkbox
                                            onChange={this.onIstheServiceBeingUsedChange}
                                            checked={this.state.istheServiceBeingUsed}
                                            label="Is this service being used?"
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                    <div className={styles.formColumn}>
                                        <Label className={styles.columnlabelstyle} required={true}>Page Owner</Label>
                                        <PeoplePicker
                                            context={this.props.context as any}
                                            personSelectionLimit={5}
                                            groupName={""} // Leave this blank in case you want to filter from all users    
                                            showtooltip={false}
                                            ensureUser={true}
                                            onChange={this.getPeoplePickerItems}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                            defaultSelectedUsers={[...this.state.defaultSelectedPageOwners]}
                                            disabled={this.state.controlDisabled}
                                        />
                                    </div>
                                </div>
                                <div className={styles.formRow}>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle} required={true}>Subscription Type</Label>
                                        <Dropdown
                                            selectedKey={this.state.selectedSubscriptionType}
                                            options={this.state.wikiSubscriptionTypeChoices}
                                            placeholder="Select a Subscription"
                                            onChange={this.onSubscriptionChange}
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}>Subscription Expiration Date</Label>
                                        <DatePicker
                                            firstDayOfWeek={DayOfWeek.Monday}
                                            //  showWeekNumbers={true}
                                            firstWeekOfYear={1}
                                            showMonthPickerAsOverlay={true}
                                            placeholder="Select a date..."
                                            ariaLabel="Select a date"
                                            // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                            strings={this.constants.dateTimePickerString}
                                            onSelectDate={this.onSubscriptionDateChange}
                                            value={this.state.subscriptionExpirationDate}
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}></Label>
                                        <Checkbox
                                            onChange={this.onisSubscriptionPaidChange}
                                            checked={this.state.isSubscriptionPaid}
                                            label="Is subscription paid?"
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                </div>
                                <div className={styles.formRow}>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}>User ID</Label>
                                        <TextField
                                            placeholder="Type here"
                                            value={this.state.userId}
                                            onChange={this.onUserIDChange}
                                            disabled={this.state.controlDisabled}
                                            onGetErrorMessage={this.getUserErrorMessage}
                                            errorMessage={this.state.userIDError} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}>Site Password</Label>
                                        <TextField
                                            placeholder="Type here"
                                            value={this.state.password}
                                            onChange={this.onPasswordChange}
                                            disabled={this.state.controlDisabled}
                                            onGetErrorMessage={this.getPasswordErrorMessage}
                                            errorMessage={this.state.sitePasswordError} />
                                    </div>
                                    <div className={styles.formColumn} >
                                        <Label className={styles.columnlabelstyle}></Label>
                                        <Checkbox
                                            onChange={this.oncanThePasswordBeCutAndPasted}
                                            checked={this.state.canThePasswordBeCutAndPasted}
                                            label="Can the password be cut and pasted?"
                                            disabled={this.state.controlDisabled} />
                                    </div>
                                </div>
                                <div className={styles.formRow}>
                                    <div className={styles.formColumnWhole} >
                                        <Label className={styles.columnlabelstyle}>Brief Label Description</Label>
                                        <ReactQuill
                                            value={this.state.briefLabelDesc}
                                            modules={this.constants.modules}
                                            formats={this.constants.formats}
                                            onChange={(newvalue) => this.setState({ briefLabelDesc: newvalue })} />
                                    </div>
                                    <div className={styles.formColumnWhole} >
                                        <Label className={styles.columnlabelstyle}>Restricted Label Description</Label>
                                        <ReactQuill
                                            value={this.state.restrictedLabelDesc}
                                            modules={this.constants.modules}
                                            formats={this.constants.formats}
                                            onChange={(newvalue) => this.setState({ restrictedLabelDesc: newvalue })} />
                                    </div>
                                </div>
                                <div className={styles.formRow}>

                                    <div className={styles.buttomMainColumn} >
                                        {
                                            this.props.wikiItem ?
                                                <DefaultButton
                                                    className={styles.button}
                                                    data-automation-id="Close"
                                                    allowDisabledFocus={true}
                                                    text="Close"
                                                    onClick={this.props.onCloseButtonClick}
                                                    disabled={this.state.controlDisabled}> Close
                                                </DefaultButton> :
                                                <DefaultButton
                                                    className={styles.button}
                                                    data-automation-id="Clear"
                                                    allowDisabledFocus={true}
                                                    text="Clear"
                                                    onClick={this.onClearClick}
                                                    disabled={this.state.controlDisabled}> Clear
                                                </DefaultButton>
                                        }
                                        {this.state && this.state.showNew ?
                                            <PrimaryButton
                                                className={styles.button}
                                                data-automation-id="Save"
                                                allowDisabledFocus={true}
                                                onClick={this.onSaveClick}
                                                disabled={this.state.controlDisabled}> {this.state.controlDisabled ? <Spinner ariaLive="assertive" labelPosition="left" /> : "Save"}
                                            </PrimaryButton> :
                                            <PrimaryButton
                                                className={styles.button}
                                                data-automation-id="Update"
                                                allowDisabledFocus={true}
                                                onClick={this.onUpdateClick}
                                                disabled={this.state.controlDisabled}> {this.state.controlDisabled ? <Spinner ariaLive="assertive" labelPosition="left" /> : "Update"}
                                            </PrimaryButton>
                                        }
                                    </div>
                                </div>
                            </div>
                        </form>
                    </div >
                </div >
            </div >
        );
    }
}



{/* <Dropdown
                                            selectedKey={this.state.selectedSubCategory}
                                            options={this.state.wikiSubCategoryChoices}
                                            placeholder="Select a Sub Category"
                                            multiSelect
                                            onChange={this.onSubCategoryChange}
                                            disabled={this.state.controlDisabled} /> */}