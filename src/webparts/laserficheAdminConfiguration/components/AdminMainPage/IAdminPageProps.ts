import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';

export interface IAdminPageProps {
  context: WebPartContext;
  loggedIn: boolean;
  repoClient: IRepositoryApiClientExInternal;
}
