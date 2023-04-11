import { IDocument } from './ILaserficheRepositoryAccessDocument';

export interface ILaserficheRepositoryAccessWebPartState {
  items: IDocument[];
  announcedMessage?: string;
  selectionDetails: string;
  checkeditemid: string;
  checkeditemfolderornot: boolean;
  parentItemId: string;
  loading: boolean;
  uploadProgressBar: boolean;
  fileUploadPercentage: number;
  webClientUrl: string;
  showUploadModal: boolean;
  showCreateModal: boolean;
  showAlertModal: boolean;
  region: string;
}
