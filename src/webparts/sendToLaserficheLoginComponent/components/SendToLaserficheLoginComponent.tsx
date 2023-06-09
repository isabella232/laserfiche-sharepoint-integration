import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Navigation } from 'spfx-navigation';
import {
  LfLoginComponent,
  LoginState,
} from '@laserfiche/types-lf-ui-components';
import { clientId, SP_LOCAL_STORAGE_KEY } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISPDocumentData } from '../../../Utils/Types';
import SaveToLaserficheCustomDialog from '../../../extensions/savetoLaserfiche/SaveToLaserficheDialog';

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

export default function SendToLaserficheLoginComponent(
  props: ISendToLaserficheLoginComponentProps
) {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();

  const [loggedIn, setLoggedIn] = React.useState<boolean>(false);

  const region = props.devMode ? 'a.clouddev.laserfiche.com' : 'laserfiche.com';

  const spFileMetadata = JSON.parse(
    window.localStorage.getItem(SP_LOCAL_STORAGE_KEY)
  ) as ISPDocumentData;

  React.useEffect(() => {
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
        loginComponent.current.addEventListener(
          'loginCompleted',
          loginCompleted
        );
        loginComponent.current.addEventListener(
          'logoutCompleted',
          logoutCompleted
        );

        const loggedIn: boolean =
          loginComponent.current.state === LoginState.LoggedIn;

        if (loggedIn && spFileMetadata) {
          const dialog = new SaveToLaserficheCustomDialog(spFileMetadata);
          dialog.show().then(() => {
            if (!dialog.successful) {
              console.warn('Could not login successfully');
            }
          });
        }
      });
    });
  }, []);

  const loginCompleted = () => {
    if (spFileMetadata) {
      const dialog = new SaveToLaserficheCustomDialog(spFileMetadata);
      dialog.show().then(() => {
        if (!dialog.successful) {
          console.warn('Could not login successfully');
        }
      });
    }
  };

  const logoutCompleted = () => {
    setLoggedIn(false);
    window.location.href =
      props.context.pageContext.web.absoluteUrl + props.laserficheRedirectUrl;
  };

  function Redirect() {
    const spFileUrl = spFileMetadata.fileUrl;
    const fileNameWithExtension = spFileMetadata.fileName;
    const spFileUrlWithoutFileName = spFileUrl.replace(fileNameWithExtension, '');
    const path = window.location.origin + spFileUrlWithoutFileName;
    Navigation.navigate(path, true);
  }

  return (
    <div>
      <div
        style={{ borderBottom: '3px solid #CE7A14', marginBlockEnd: '32px' }}
      >
        <img
          src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
          height={'46px'}
          width={'45px'}
          style={{ marginTop: '8px', marginLeft: '8px' }}
        />
        <span
          id='remveHeading'
          style={{ marginLeft: '10px', fontSize: '22px', fontWeight: 'bold' }}
        >
          {loggedIn ? 'Sign Out' : 'Sign In'}
        </span>
      </div>
      <p
        id='remve'
        style={{ textAlign: 'center', fontWeight: '600', fontSize: '20px' }}
      >
        {loggedIn ? 'You are signed in to Laserfiche' : 'Welcome to Laserfiche'}
      </p>
      <div style={{ textAlign: 'center' }}>
        <lf-login
          redirect_uri={
            props.context.pageContext.web.absoluteUrl +
            props.laserficheRedirectUrl
          }
          authorize_url_host_name={region}
          redirect_behavior='Replace'
          client_id={clientId}
          sign_in_text='Sign in'
          sign_out_text='Sign out'
          ref={loginComponent}
        />
      </div>
      <div>
        <div
          /* className="lf-component-container lf-right-button" */ style={{
            marginTop: '35px',
            textAlign: 'center',
          }}
        >
          <button style={{ fontWeight: '600' }} onClick={Redirect}>
            Go Back
          </button>
        </div>
      </div>
    </div>
  );
}
