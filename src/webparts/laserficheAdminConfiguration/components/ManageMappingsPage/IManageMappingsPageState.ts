import { IListItem } from './IListItem';

export interface IManageMappingsPageState {
  mappingRows: any;
  sharePointContentTypes: any;
  laserficheContentTypes: any;
  listItem: IListItem[];
  showDeleteModal: boolean;
  deleteSharePointcontentType: string;
}
