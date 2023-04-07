import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';

export interface IAddNewManageConfigurationProps {
  context: WebPartContext;
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
}
