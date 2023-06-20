import * as React from 'react';
import { useState } from 'react';
import { ILaserficheAdminConfigurationProps } from './ILaserficheAdminConfigurationProps';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { Stack, StackItem } from 'office-ui-fabric-react';
import AdminMainPage from '../components/AdminMainPage/AdminMainPage';
import HomePage from './HomePage/HomePage';
import ManageConfigurationsPage from './ManageConfigurationsPage/ManageConfigurationsPage';
import ManageMappingsPage from './ManageMappingsPage/ManageMappingsPage';
import EditManageConfiguration from './EditManageConfiguration/EditManageConfiguration';
import AddNewManageConfiguration from './AddNewManageConfiguration/AddNewManageConfiguration';
import { clientId, SPDEVMODE_LOCAL_STORAGE_KEY } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { SPComponentLoader } from '@microsoft/sp-loader';

export default function LaserficheAdminConfiguration(
  props: ILaserficheAdminConfigurationProps
) {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.createRef();
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);
  const spDevMode = window?.localStorage.getItem(SPDEVMODE_LOCAL_STORAGE_KEY);
  const region = spDevMode && spDevMode === 'true' ? 'a.clouddev.laserfiche.com' : 'laserfiche.com';
  const redirectPage =
    props.context.pageContext.web.absoluteUrl + props.laserficheRedirectPage;

  React.useEffect(() => {
    SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
    ).then(() => {
      SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      ).then(() => {
        const loginCompleted = async () => {
          await getAndInitializeRepositoryClientAndServicesAsync();
          setLoggedIn(true);
        };
        const logoutCompleted = async () => {
          setLoggedIn(false);
          window.location.href =
            props.context.pageContext.web.absoluteUrl +
            props.laserficheRedirectPage;
        };

        loginComponent.current.addEventListener(
          'loginCompleted',
          loginCompleted
        );
        loginComponent.current.addEventListener(
          'logoutCompleted',
          logoutCompleted
        );
        if (loginComponent.current.authorization_credentials) {
          getAndInitializeRepositoryClientAndServicesAsync().then(() => {
            setLoggedIn(true);
          });
        }
      });
    });
  }, []);

  SPComponentLoader.loadCss(
    'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
  );
  SPComponentLoader.loadCss(
    'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
  );

  async function getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken =
      loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await ensureRepoClientInitializedAsync();
    } else {
      // user is not logged in
    }
  }

  async function ensureRepoClientInitializedAsync(): Promise<void> {
    if (!repoClient) {
      const repoClientCreator = new RepositoryClientExInternal();
      const newRepoClient =
        await repoClientCreator.createRepositoryClientAsync();
      setRepoClient(newRepoClient);
    }
  }

  return (
    <HashRouter>
      <Stack>
        <div className='btnSignOut'>
          <lf-login
            redirect_uri={redirectPage}
            authorize_url_host_name={region}
            redirect_behavior='Replace'
            client_id={clientId}
            ref={loginComponent}
          />
        </div>
        <AdminMainPage
          context={props.context}
          webPartTitle={props.webPartTitle}
          loggedIn={loggedIn}
          repoClient={repoClient}
        />
        <StackItem>
          <Switch>
            <Route
              exact={true}
              component={() => <HomePage />}
              path='/HomePage'
            />
            <Route exact={true} component={() => <HomePage />} path='/' />
            <Route
              exact={true}
              component={() => (
                <ManageConfigurationsPage context={props.context} />
              )}
              path='/ManageConfigurationsPage'
            />
            <Route
              exact={true}
              component={() => <ManageMappingsPage
                context={props.context}
                isLoggedIn={loggedIn}
                repoClient={repoClient}
              />}
              path='/ManageMappingsPage'
            />
            <Route
              exact={true}
              component={() => (
                <AddNewManageConfiguration
                  context={props.context}
                  loggedIn={loggedIn}
                  repoClient={repoClient}
                />
              )}
              path='/AddNewManageConfiguration'
            />
            <Route
              exact={true}
              render={(properties) => (
                <EditManageConfiguration
                  {...properties}
                  context={props.context}
                  laserficheRedirectPage={props.laserficheRedirectPage}
                  loggedIn={loggedIn}
                  repoClient={repoClient}
                />
              )}
              path='/EditManageConfiguration/:name'
            />
          </Switch>
        </StackItem>
      </Stack>
    </HashRouter>
  );
}
