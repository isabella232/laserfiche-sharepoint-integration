import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';

export interface IManageMappingsPageProps {
  context: WebPartContext;
  repoClient: IRepositoryApiClientExInternal;
  isLoggedIn: boolean;
}
