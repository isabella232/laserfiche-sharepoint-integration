import { IRepositoryApiClientEx } from '@laserfiche/lf-ui-components-services';

export interface IRepositoryApiClientExInternal extends IRepositoryApiClientEx {
  clearCurrentRepo: () => void;
  _repoId?: string;
  _repoName?: string;
}
