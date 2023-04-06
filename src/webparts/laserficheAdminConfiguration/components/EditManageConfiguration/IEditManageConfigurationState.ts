import { TemplateFieldInfo } from '@laserfiche/lf-repository-api-client';
import { IListItem } from './IListItem';

interface ILfSelectedFolder {
  //selectedNodeUrl: string; // url to open the selected node in Web Client
  selectedFolderPath: string; // path of selected folder
  //selectedFolderName: string; // name of the selected folder
}


export interface SPFieldData {
  Title: string;
  TypeAsString: string;
  InternalName: string;
}

export interface ProfileConfiguration {
  ConfigurationName: string;
  DocumentName: string;
  DocumentTemplate: any;
  DestinationPath: string;
  EntryId: string;
  Action: string;
  SharePointFields: SPFieldData[];
  LaserficheFields: TemplateFieldInfo[];
}

export enum FieldMappingError {
  CONTENT_TYPE= 'CONTENT_TYPE',
  SELECT_TEMPLATE= 'SELECT_TYPE',
}

export interface MappedFields {
  id: string;
  lfField: TemplateFieldInfo | undefined;
  spField: SPFieldData | undefined;
}


export interface IEditManageConfigurationState {
  mappingList: (MappedFields)[];
  laserficheTemplates: any;
  sharePointFields: SPFieldData[];
  laserficheFields: TemplateFieldInfo[];
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
