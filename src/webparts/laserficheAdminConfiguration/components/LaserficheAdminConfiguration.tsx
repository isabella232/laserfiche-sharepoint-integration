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
import { clientId } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { getRegion } from '../../../Utils/Funcs';
import { ProblemDetails } from '@laserfiche/lf-repository-api-client';
import styles from './LaserficheAdminConfiguration.module.scss';
import { SPPermission } from '@microsoft/sp-page-context';

const YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN = 'You do not have the necessary rights to view or edit the Laserfiche SharePoint Integration configuration. Please contact your administrator for help.';

export default function LaserficheAdminConfiguration(
  props: ILaserficheAdminConfigurationProps
): JSX.Element {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.createRef();
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);

  const region = getRegion();

  const redirectPage = window.location.origin + window.location.pathname;

  function isAdmin(): boolean {
    const permission = new SPPermission(
      props.context.pageContext.web.permissions.value
    );
    const isFullControl = permission.hasPermission(SPPermission.manageWeb);
    return isFullControl;
  }

  async function getAndInitializeRepositoryClientAndServicesAsync(): Promise<void> {
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

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      SPComponentLoader.loadCss(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
      );
      SPComponentLoader.loadCss(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
      );
      await SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
      );
      await SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      );
      const loginCompleted: () => Promise<void> = async () => {
        await getAndInitializeRepositoryClientAndServicesAsync();
        setLoggedIn(true);
      };
      const logoutCompleted: () => Promise<void> = async () => {
        setLoggedIn(false);
      };

      loginComponent.current.addEventListener('loginCompleted', loginCompleted);
      loginComponent.current.addEventListener(
        'logoutCompleted',
        logoutCompleted
      );
      if (loginComponent.current.authorization_credentials) {
        await getAndInitializeRepositoryClientAndServicesAsync();
        setLoggedIn(true);
      }
    };

    initializeComponentAsync().catch((err: Error | ProblemDetails) => {
      console.warn(
        `Error: ${(err as Error).message ?? (err as ProblemDetails).title}`
      );
    });
  }, []);

  return (
    <React.StrictMode>
      <HashRouter>
        <Stack>
          {isAdmin() && (
            <>
              <div className={styles.loginButton}>
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
                    component={() => (
                      <ManageMappingsPage
                        context={props.context}
                        isLoggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    )}
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
                        loggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    )}
                    path='/EditManageConfiguration/:name'
                  />
                </Switch>
              </StackItem>
            </>
          )}
          {!isAdmin() &&
          <span><b>{YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN}</b></span>}
        </Stack>
      </HashRouter>
    </React.StrictMode>
  );
}
