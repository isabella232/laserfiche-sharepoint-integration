import { IListItem } from './IListItem';

export interface IManageConfigurationPageState {
  configurationRows: any;
  listItem: IListItem[];
  showDeleteModal: boolean;
  configurationName: string;
}
