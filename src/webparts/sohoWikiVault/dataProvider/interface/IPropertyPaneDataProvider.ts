export interface IPropertyPaneDataProvider {
  createConfigLists(): Promise<boolean>;
}