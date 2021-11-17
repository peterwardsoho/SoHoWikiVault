import { IWikiEditFormState } from '../../components';
import { IDropdownValues, IMainCategory, IWikiVault } from '../../model';
export interface IWikiEditFormDataProvider {
    getFieldDDValue(): Promise<IDropdownValues[]>;
    onSaveClick(wikiState: IWikiEditFormState, mainCategory: IMainCategory): Promise<string>;
    onUpdateClick(wikiState: IWikiEditFormState, mainCategory: IMainCategory): Promise<string>;
}