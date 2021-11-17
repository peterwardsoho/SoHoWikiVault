import * as React from 'react';
import styles from './ViewWikiDetails.module.scss';
import { IViewWikiDetailsProps, IViewWikiDetailsState } from './';
import { escape } from '@microsoft/sp-lodash-subset';
// data Provider
import { CommonDataProvider } from '../../dataProvider';
// constants
import { Constants } from '../../common';
// Office UI fabric
import { IconButton, MessageBar, Label, DefaultButton, Link } from 'office-ui-fabric-react/lib';
import { IListDetails, IWikiVault, IWikiVaultPassword, IUsers } from '../../model';

export class ViewWikiDetails extends React.Component<IViewWikiDetailsProps, IViewWikiDetailsState> {
    public state: IViewWikiDetailsState;
    private constants: Constants;
    private commonDataProvider: CommonDataProvider;

    public async componentWillMount() {
        console.log(this.props.wikiItem);
        this.constants = new Constants();
        this.commonDataProvider = new CommonDataProvider(this.props.context);
        this.state = {
            wikiPassword: null,
            needPerission: '',
            subCategory: '',
            pageOwner: '',
            restrictedLabel: '',
            sitePassword: '',
            userID: '',
            userNameCopied: 'Copy',
            passwordCopied: 'Copy',
            urlCopied: 'Copy',
            showBriefDesc: false,
            showRestDesc: false
        };
        this.state = await this.setInitialState(this.state, this.props.wikiItem);
        this.setState(this.state);
    }
    private copyUserNametoClipboard = (val: string) => {
        this.copytoClipboard(val);
        this.setState({ userNameCopied: 'CheckMark' });
        setTimeout(() => { this.setState({ userNameCopied: 'Copy' }); }, 2000);
    }
    private copyPasswordtoClipboard = (val: string) => {
        this.copytoClipboard(val);
        this.setState({ passwordCopied: 'CheckMark' });
        setTimeout(() => { this.setState({ passwordCopied: 'Copy' }); }, 2000);
    }
    private copyURLtoClipboard = () => {
        const url = window.location.href.split('?')[0] + '?wikiId=' + this.props.wikiItem.Id;
        this.copytoClipboard(url);
        this.setState({ urlCopied: 'CheckMark' });
        setTimeout(() => { this.setState({ urlCopied: 'Copy' }); }, 2000);
    }
    private copytoClipboard = (val: string): void => {
        const selBox = document.createElement('textarea');
        selBox.style.position = 'fixed';
        selBox.style.left = '0';
        selBox.style.top = '0';
        selBox.style.opacity = '0';
        selBox.value = val;
        document.body.appendChild(selBox);
        selBox.focus();
        selBox.select();
        document.execCommand('copy');
        document.body.removeChild(selBox);

    }
    private setSecondCategory = (secondCategoryArray: IListDetails[]) => {
        let secondCategory: string = '';
        secondCategoryArray.forEach(secondCategoryValue => {
            secondCategory = secondCategory + secondCategoryValue.Title + ',';
        });
        if (secondCategory) {
            secondCategory = secondCategory.slice(0, -1);
        }
        return secondCategory;
    }
    private setPageOwner = (pageOwnerArray: IUsers[]) => {
        let pageOwner: string = '';
        if (pageOwnerArray && pageOwnerArray.length > 0) {
            pageOwnerArray.forEach(pageOwnerValue => {
                pageOwner = pageOwner + pageOwnerValue.FirstName + ' ' + pageOwnerValue.LastName + ',';
            });
            if (pageOwner) {
                pageOwner = pageOwner.slice(0, -1);
            }
        }
        return pageOwner;
    }
    public setInitialState = async (stateValue: IViewWikiDetailsState, wiki: IWikiVault) => {
        if (wiki && wiki.Id) {
            const filter = `&$filter=SohoWikiVaultId eq ${this.props.wikiItem.Id}`;
            const wikiPassword: IWikiVaultPassword[] = await this.commonDataProvider.getListItems(this.constants.WikiVaultPassword.listName, this.constants.WikiVaultPassword.selectQuery, '', filter, '', '');
            stateValue.subCategory = this.setSecondCategory(wiki.SubCategory);
            stateValue.pageOwner = this.setPageOwner(wiki.PageOwner);
            if (wikiPassword && wikiPassword.length > 0) {
                stateValue.needPerission = this.constants.permission.no;
                stateValue.userID = wikiPassword[0].UserId;
                stateValue.sitePassword = wikiPassword[0].SitePassword;
                // stateValue.restrictedLabel = this.getDescription(wikiCredentials[0].RestrictedLabelDescription);
                stateValue.restrictedLabel = wikiPassword[0].RestrictedLabelDescription;
            } else {
                stateValue.needPerission = this.constants.permission.yes;
            }
        }
        return stateValue;
    }

    private renderDescription = (desc: string): JSX.Element => {
        return (
            <div dangerouslySetInnerHTML={{ __html: desc }}></div>
        );
    }
    private openURL = (url: string) => {
        let a = document.createElement('a');
        a.target = '_blank';
        a.href = url;
        a.click();
    }
    private showBriefDesc = (showHide: boolean) => {
        this.setState({ showBriefDesc: showHide });
    }
    private showRestDesc = (showHide: boolean) => {
        this.setState({ showRestDesc: showHide });
    }
    public render(): React.ReactElement<IViewWikiDetailsProps> {
        return (
            <div>
                <div className={styles.viewWikiDetails}>
                    <div className={styles.container}>
                        <div className={styles.rowTop}>
                            <div className={styles.modifiedDetails}>
                                Modified by {this.props.wikiItem.Editor.FirstName} {this.props.wikiItem.Editor.LastName} on {(new Date(this.props.wikiItem.Modified)).toLocaleDateString("en-US")}
                                <IconButton className={styles.fileButton} iconProps={{ iconName: this.state.urlCopied }} title="Copy" ariaLabel="Copy to clipboard" onClick={() => this.copyURLtoClipboard()} />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.maindiv}>
                                <form>
                                    <div className={styles.formrowTop}>
                                        <div className={styles.columnleft}>Main Category</div>
                                        <div className={styles.columnright}>{this.props.wikiItem.MainCategory.Title}</div>
                                    </div>
                                    <div className={styles.formrow}>
                                        <div className={styles.columnleft}>2nd Category</div>
                                        <div className={styles.columnright}>{this.state.subCategory}</div>
                                    </div>
                                    <div className={styles.formrow}>
                                        <div className={styles.columnleft}>Brief Label Description</div>
                                        {this.props.wikiItem.BriefLabelDescription && this.props.wikiItem.BriefLabelDescription.length > 350 ?
                                            <div className={styles.columnright}>
                                                <div className={styles.showAlltext}>
                                                    <Link onClick={() => this.showBriefDesc(!this.state.showBriefDesc)}>{this.state.showBriefDesc ? "Show Less..." : "Show More..."}</Link>
                                                </div>
                                                <div className={this.state.showBriefDesc ? styles.columnrightDescMore : styles.columnrightDesc}>
                                                    {this.renderDescription(this.props.wikiItem.BriefLabelDescription)}
                                                </div>
                                            </div> :
                                            <div className={styles.columnright}>
                                                {this.renderDescription(this.props.wikiItem.BriefLabelDescription)}
                                            </div>
                                        }
                                    </div>
                                    <div className={styles.formrow}>
                                        <div className={styles.columnleft}>Subscription Expiration Date</div>
                                        <div className={styles.columnright}>{this.props.wikiItem.SubscriptionExpirationDate ?
                                            <div>
                                                {(new Date(this.props.wikiItem.SubscriptionExpirationDate)).toLocaleDateString("en-US")}
                                                {this.props.wikiItem.IsSubscriptionPaid && new Date(this.props.wikiItem.SubscriptionExpirationDate) < new Date() ? <span className={styles.redText}>   Expired !!!</span> : ''}
                                            </div>
                                            : ''}</div>
                                    </div>
                                    <div className={styles.formrow}>
                                        <div className={styles.columnleft}>Is Service being Used</div>
                                        <div className={styles.columnright}>{this.props.wikiItem.IsServiceBeingUsed ? "Yes" : <div className={styles.redText}>No</div>}</div>
                                    </div>
                                    <div className={styles.formrow}>
                                        <div className={styles.columnleft}>Page Owner</div>
                                        <div className={styles.columnright}>{this.state.pageOwner}</div>
                                    </div>
                                    {this.state.needPerission == this.constants.permission.no ?
                                        <div className={styles.formrowSolid}>
                                            <div className={styles.columnleft}>User ID</div>
                                            <div className={styles.columnrightSolid}>{this.state.userID} {this.state.userID ?
                                                <IconButton className={styles.fileButton} iconProps={{ iconName: this.state.userNameCopied }} title="Copy" ariaLabel="Copy to clipboard" onClick={() => this.copyUserNametoClipboard(this.state.userID)} />
                                                : ''}</div>
                                        </div> : ''}
                                    {this.state.needPerission == this.constants.permission.no ?
                                        <div className={styles.formrowSolid}>
                                            <div className={styles.columnleft}>Site Password</div>
                                            <div className={styles.columnrightSolid}>{this.state.sitePassword} {this.state.sitePassword ?
                                                <IconButton className={styles.fileButton} iconProps={{ iconName: this.state.passwordCopied }} title="Copy" ariaLabel="Copy to clipboard" onClick={() => this.copyPasswordtoClipboard(this.state.sitePassword)} />
                                                : ''}</div>
                                        </div> : ''}
                                    <div className={styles.formrowSolid}>
                                        <div className={styles.columnleft}>URL</div>
                                        <div className={styles.columnrightSolid}>
                                            {this.props.wikiItem.URL ?
                                                <IconButton className={styles.icons} iconProps={{ iconName: 'OpenInNewWindow' }} title="Open" ariaLabel="Open" onClick={() => this.openURL(this.props.wikiItem.URL.Url)} />
                                                : ''}</div>
                                    </div>
                                    {this.state && this.state.needPerission == this.constants.permission.no ?
                                        <div className={styles.formrow}>
                                            <div className={styles.columnleft}>Can the password be cut and pasted</div>
                                            <div className={styles.columnright}>{this.props.wikiItem.CanThePasswordBeCutAndPasted ? 'Yes' : 'No'}</div>
                                        </div> : ''}
                                    {this.state.needPerission == this.constants.permission.no ?
                                        <div className={styles.formrow}>
                                            <div className={styles.columnleft}>Restricted Label Description</div>
                                            {this.state && this.state.restrictedLabel && this.state.restrictedLabel.length > 350 ?
                                                <div className={styles.columnright}>
                                                    <div className={styles.showAlltext}>
                                                        <Link onClick={() => this.showRestDesc(!this.state.showRestDesc)}>{this.state.showRestDesc ? "Show Less..." : "Show More..."}</Link>
                                                    </div>
                                                    <div className={this.state.showRestDesc ? styles.columnrightDescMore : styles.columnrightDesc}>
                                                        {this.renderDescription(this.state.restrictedLabel)}
                                                    </div>
                                                </div> :
                                                <div className={styles.columnright}>
                                                    {this.renderDescription(this.state.restrictedLabel)}
                                                </div>
                                            }
                                        </div> : ''}
                                </form>
                            </div>
                        </div>
                    </div>
                </div>
                <div className={styles.rowButton}>
                    <div className={styles.permissionText}>
                        <DefaultButton
                            className={styles.button}
                            data-automation-id="Close"
                            allowDisabledFocus={true}
                            text="Close"
                            onClick={this.props.onCloseButtonClick}> Close </DefaultButton>
                    </div>
                </div>
            </div>
        );
    }
}


{/* <a href={this.props.wikiItem.URL.Url} target="_blank" data-interception="off"><img className={styles.iconImage} src={ } alt="Click for more details" title="click for more details" /></a> */ }
{/* <div className={styles.permissionText}>
{this.state && this.state.needPerission == this.constants.permission.yes ? <div>You dont have permission to view restricted credentials. Please contact <a href='mailto:helpdesk@sohodragon.com?subject=Wiki%20Access%20request'>admin!!</a></div> :
''}
</div> */}