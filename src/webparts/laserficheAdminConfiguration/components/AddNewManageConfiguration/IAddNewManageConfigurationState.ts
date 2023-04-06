import { TemplateFieldInfo } from '@laserfiche/lf-repository-api-client';
import { MappedFields, SPFieldData, FieldMappingError, ProfileConfiguration } from '../EditManageConfiguration/IEditManageConfigurationState';
import { IListItem } from './IListItem';

interface ILfSelectedFolder {
  //selectedNodeUrl: string; // url to open the selected node in Web Client
  selectedFolderPath: string; // path of selected folder
  //selectedFolderName: string; // name of the selected folder
}

export interface IAddNewManageConfigurationState {
  mappingList: (MappedFields)[];
  sharePointFields: SPFieldData[];
  laserficheFields: TemplateFieldInfo[];
  laserficheTemplates: any;
  documentNames: any;
  loadingContent: boolean;
  hideContent: boolean;
  showFolderModal: boolean;
  showtokensModal: boolean;
  deleteModal: JSX.Element | undefined;
  showConfirmModal: boolean;
  columnError: FieldMappingError;
  profileConfig: ProfileConfiguration; 
}
