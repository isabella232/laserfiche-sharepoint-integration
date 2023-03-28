import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEditManageConfigurationProps {
  context: WebPartContext;
  match: any;
  laserficheRedirectPage: string;
  devMode: boolean;
}
