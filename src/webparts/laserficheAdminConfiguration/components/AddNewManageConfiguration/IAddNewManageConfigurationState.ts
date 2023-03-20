import { IListItem } from "./IListItem";

interface ILfSelectedFolder {
    //selectedNodeUrl: string; // url to open the selected node in Web Client
    selectedFolderPath: string; // path of selected folder
    //selectedFolderName: string; // name of the selected folder
  }

export interface IAddNewManageConfigurationState{
    action:string;
    listItem:IListItem[];
    mappingList:any;
    sharePointFields:any;
    laserficheTemplates:any;
    laserficheFields:any;
    documentNames:any;
    loadingContent:boolean;
    hideContent:boolean;
    showFolderModal:boolean;
    showtokensModal:boolean;
    showDeleteModal:boolean;
    showConfirmModal:boolean;
    lfSelectedFolder:ILfSelectedFolder;
    shouldShowOpen: boolean; 
    shouldShowSelect: boolean; 
    shouldDisableSelect: boolean;
    region: string;
}
   