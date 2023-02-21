import {IDocument} from "./ILaserficheRepositoryAccessDocument";
import { IColumn,Selection } from "office-ui-fabric-react";

export interface ILaserficheRepositoryAccessWebPartState
{
    columns:IColumn[];
    items:IDocument[];
    announcedMessage?: string;
    selectionDetails:string;
    selection: Selection;
    checkeditemid: number;
    checkeditemfolderornot: boolean;
    parentItemId:number;
    loading: boolean;
    uploadProgressBar:boolean;
    fileUploadPercentage: number;
    webClientUrl:string;
    showUploadModal:boolean;
    showCreateModal:boolean;
    showAlertModal:boolean;
}