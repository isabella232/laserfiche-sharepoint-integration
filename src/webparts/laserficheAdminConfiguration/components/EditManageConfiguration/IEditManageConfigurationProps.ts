import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';

export interface IEditManageConfigurationProps {
  context: WebPartContext;
  // eslint-disable-next-line
  match: any;
  loggedIn: boolean;
  repoClient: IRepositoryApiClientExInternal;
}
