import * as React from 'react';
// Component
import { ISubCategoryProps } from '../';
import styles from '../sohoWikiVault/SohoWikiVault.module.scss';
// Office UI Fabric Imports
import { Pivot, PivotItem, PivotLinkSize, IPivotStyles, PivotLinkFormat } from 'office-ui-fabric-react';
// Components
import { ListView } from '../';
//Model
import { ISubCategory } from '../../model';
// Common
import { Constants } from '../../common';

export class SubCategory extends React.Component<ISubCategoryProps, {}> {
    private constants: Constants;
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
    }
    private _renderTabs = (): JSX.Element => {
        let visibleSubCat: ISubCategory[] = this.props.subCategory.filter((maincat => maincat.Status == this.constants.comparingStrings.visible));
        return (
            <Pivot aria-label="Basic Pivot Example" linkSize={PivotLinkSize.normal} styles={this.pivotStyle}>
                {visibleSubCat.map((subCategory, index) => {
                    if (index == 0) {
                        return (
                            <PivotItem headerText={subCategory.Title} headerButtonProps={{
                                'data-order': 1,
                                'data-title': 'All',
                            }}>
                                <ListView
                                    context={this.props.context}
                                    mainCategory={this.props.mainCategory}
                                    subCategory={subCategory}
                                    searchText={this.props.searchText}
                                    isAdmin={this.props.isAdmin}

                                    mainCategoryItems={this.props.mainCategoryItems}
                                    subCategoryItems={this.props.subCategoryItems}
                                    wikiId={this.props.wikiId}>
                                </ListView>
                            </PivotItem>);
                    } else {
                        return (
                            <PivotItem headerText={subCategory.Title}>
                                <ListView
                                    context={this.props.context}
                                    mainCategory={this.props.mainCategory}
                                    subCategory={subCategory}
                                    searchText={this.props.searchText}
                                    isAdmin={this.props.isAdmin}

                                    mainCategoryItems={this.props.mainCategoryItems}
                                    subCategoryItems={this.props.subCategoryItems}
                                    wikiId={this.props.wikiId}>
                                </ListView>
                            </PivotItem>);
                    }
                })
                }
            </Pivot>
        );
    }
    public render(): React.ReactElement<ISubCategoryProps> {
        return (
            <div className={styles.subCategory}>
                <div className={styles.subCategorydiv}>
                    {this._renderTabs()}
                </div>
            </div>
        );
    }
}




 // let subCategory: ISubCategory[] = [];
        // this.setState({ subCategoryItems: [] });

        // if (this.props.subCategoryItems && this.props.subCategoryItems.length > 0) {
        //     subCategory.push({ Title: 'All', Id: '0' });
        //     this.props.subCategoryItems.forEach(item => {
        //         subCategory.push(item);
        //     });
        // }
        // this.setState({
        //     subCategoryItems: subCategory
        // });