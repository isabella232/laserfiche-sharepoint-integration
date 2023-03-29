import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILaserficheRepositoryAccessWebPartProps {
  context: WebPartContext;
  webPartTitle: string;
  laserficheRedirectUrl: string;
  devMode: boolean;
}
