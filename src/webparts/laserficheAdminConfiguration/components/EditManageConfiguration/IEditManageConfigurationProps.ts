import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';

export interface IEditManageConfigurationProps {
  context: WebPartContext;
  match: any;
  laserficheRedirectPage: string;
  loggedIn: boolean;
  repoClient: IRepositoryApiClientExInternal;
}
