import * as React from 'react';
import SvgHtmlIcons from '../components/SVGHtmlIcons';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { clientId } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { useEffect, useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import RepositoryViewComponent from './RepositoryViewWebPart';
require('../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');
require('../../../Assets/CSS/bootstrap.min.css');
require('../../../Assets/CSS/custom.css');
import './LaserficheRepositoryAccess.scss';
import { ILaserficheRepositoryAccessWebPartProps } from './ILaserficheRepositoryAccessWebPartProps';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ['lf-field-container']: any;
      ['lf-login']: any;
    }
  }
}

export default function LaserficheRepositoryAccessWebPart(props: ILaserficheRepositoryAccessWebPartProps) {
  const [webClientUrl, setWebClientUrl] = React.useState('');
  let loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);
  const region = props.devMode ? 'a.clouddev.laserfiche.com' : 'laserfiche.com';
  const redirectPage =
    props.context.pageContext.web.absoluteUrl + props.laserficheRedirectPage;

  useEffect(() => {
    SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
    ).then(() => {
      SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      ).then(() => {
        SPComponentLoader.loadCss(
          'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
        );
        SPComponentLoader.loadCss(
          'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
        );
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

  //lf-login will trigger on click on Sign in to Laserfiche

  const getAndInitializeRepositoryClientAndServicesAsync = async () => {
    const accessToken =
      loginComponent?.current?.authorization_credentials?.accessToken;
    setWebClientUrl(loginComponent?.current?.account_endpoints.webClientUrl);
    if (accessToken) {
      await ensureRepoClientInitializedAsync();
    } else {
      // user is not logged in
    }
  };

  const ensureRepoClientInitializedAsync = async () => {
    if (!repoClient) {
      const repoClientCreator = new RepositoryClientExInternal();
      const repoClient = await repoClientCreator.createRepositoryClientAsync();
      setRepoClient(repoClient);
    }
  };

  return (
    <div>
      <div style={{ display: 'none' }}>
        <SvgHtmlIcons />
      </div>
      <div
        className='container-fluid p-3'
        style={{ maxWidth: '100%', marginLeft: '-30px' }}
      >
        <div className='btnSignOut'>
          <lf-login
            redirect_uri={redirectPage}
            redirect_behavior='Replace'
            client_id={clientId}
            authorize_url_host_name={region}
            ref={loginComponent}
          />
        </div>
        <RepositoryViewComponent
          webClientUrl={webClientUrl}
          repoClient={repoClient}
          webPartTitle={props.webPartTitle}
          loggedIn={loggedIn}
        ></RepositoryViewComponent>
      </div>
    </div>
  );
}
