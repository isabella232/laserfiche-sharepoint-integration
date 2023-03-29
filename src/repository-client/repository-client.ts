import * as React from 'react';
import {
  IRepositoryApiClient,
  RepositoryApiClient,
} from '@laserfiche/lf-repository-api-client';
import { IRepositoryApiClientExInternal } from './repository-client-types';

export class RepositoryClientExInternal {
  public repoClient: IRepositoryApiClientExInternal;

  constructor(private loginRef: React.RefObject<any>) {}

  public addAuthorizationHeader(
    request: RequestInit,
    accessToken: string | undefined
  ) {
    const headers: Headers | undefined = new Headers(request.headers);
    const AUTH = 'Authorization';
    headers.set(AUTH, 'Bearer ' + accessToken);
    request.headers = headers;
  }

  public beforeFetchRequestAsync = async (
    url: string,
    request: RequestInit
  ) => {
    // TODO trigger authorization flow if no accessToken
    const accessToken =
      this.loginRef.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      this.addAuthorizationHeader(request, accessToken);
      return {
        regionalDomain: this.loginRef.current.account_endpoints.regionalDomain,
      }; // update this if you are using a different region
    } else {
      throw new Error('No access token');
    }
  };

  public afterFetchResponseAsync = async (
    url: string,
    response: ResponseInit,
    request: RequestInit
  ) => {
    if (response.status === 401) {
      const refresh = await this.loginRef.current?.refreshTokenAsync(true);
      if (refresh) {
        const accessToken =
          this.loginRef.current?.authorization_credentials?.accessToken;
        this.addAuthorizationHeader(request, accessToken);
        return true;
      } else {
        this.repoClient?.clearCurrentRepo();
        return false;
      }
    }
    return false;
  };

  public getCurrentRepo = async () => {
    if (this.repoClient) {
      const repos = await this.repoClient.repositoriesClient.getRepositoryList(
        {}
      );
      const repo = repos[0];
      if (repo.repoId && repo.repoName) {
        return { repoId: repo.repoId, repoName: repo.repoName };
      } else {
        throw new Error('Current repoId undefined.');
      }
    } else {
      throw new Error('repoClient undefined.');
    }
  };

  public async createRepositoryClientAsync(): Promise<IRepositoryApiClientExInternal> {
    const partialRepoClient: IRepositoryApiClient =
      RepositoryApiClient.createFromHttpRequestHandler({
        beforeFetchRequestAsync: this.beforeFetchRequestAsync,
        afterFetchResponseAsync: this.afterFetchResponseAsync,
      });
    const clearCurrentRepo = () => {
      if (this.repoClient) {
        this.repoClient._repoId = undefined;
        this.repoClient._repoName = undefined;
      }
    };
    this.repoClient = {
      clearCurrentRepo,
      _repoId: undefined,
      _repoName: undefined,
      getCurrentRepoId: async () => {
        if (this.repoClient?._repoId) {
          console.log('getting id from cache');
          return this.repoClient._repoId;
        } else {
          console.log('getting id from api');
          const repo = (await this.getCurrentRepo()).repoId;
          if (this.repoClient) {
            this.repoClient._repoId = repo;
          }
          return repo;
        }
      },
      getCurrentRepoName: async () => {
        if (this.repoClient?._repoName) {
          return this.repoClient._repoName;
        } else {
          const repo = (await this.getCurrentRepo()).repoName;
          if (this.repoClient) {
            this.repoClient._repoName = repo;
          } else {
            console.debug('Cannot set repoName, repoClient does not exist');
          }
          return repo;
        }
      },
      ...partialRepoClient,
    };
    return this.repoClient;
  }
}
