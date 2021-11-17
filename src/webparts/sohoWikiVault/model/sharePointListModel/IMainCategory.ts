export interface IMainCategory {
  Id?: string;
  Title?: string;
  CategoryGroupName?: {
    Description: string;
    Url: string;
  };
  CategoryOrder?: string;
  editEnabled?: boolean;
  spinner?: boolean;
  Status?: string;
  CategoryFolderName?: string;
}