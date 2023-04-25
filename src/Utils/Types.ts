import { IPostEntryWithEdocMetadataRequest } from '@laserfiche/lf-repository-api-client';
import { ActionTypes } from '../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';

export interface ISPDocumentData {
    metadata: IPostEntryWithEdocMetadataRequest;
    fileName: string;
    documentName: string;
    templateName: string;
    action: ActionTypes;
    fileUrl: string;
    entryId: string;
    contextPageAbsoluteUrl: string;
    pageOrigin: string;
}