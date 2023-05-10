import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import SendToLaserficheCustomDialog from './SendToLaserficheCustomDialog';
import { Navigation } from 'spfx-navigation';
import {
  CreateEntryResult,
  PostEntryWithEdocMetadataRequest,
  PutFieldValsRequest,
  FileParameter,
  FieldToUpdate,
  ValueToUpdate,
  IPostEntryWithEdocMetadataRequest,
} from '@laserfiche/lf-repository-api-client';
import {
  LfLoginComponent,
  LoginState,
} from '@laserfiche/types-lf-ui-components';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { clientId } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ActionTypes } from '../../laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { TempStorageKeys } from '../../../Utils/Enums';
import { getEntryWebAccessUrl } from '../../../Utils/Funcs';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

const dialog = new SendToLaserficheCustomDialog();
export default function SendToLaserficheLoginComponent(
  props: ISendToLaserficheLoginComponentProps
) {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();
  const [repoClient, setRepoClient] = React.useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);
  const [webClientUrl, setWebClientUrl] = React.useState<string | undefined>(
    undefined
  );
  const [loggedIn, setLoggedIn] = React.useState<boolean>(false);

  const region = props.devMode ? 'a.clouddev.laserfiche.com' : 'laserfiche.com';

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

        const loggedOut: boolean =
          loginComponent.current.state === LoginState.LoggedOut;

        if (!loggedOut) {
          if (window.localStorage.getItem(TempStorageKeys.Filename)) {
            dialog.show();
          }
          getAndInitializeRepositoryClientAndServicesAsync().then(() => {
            GetFileData().then(async (fileData) => {
              saveFileToLaserfiche(fileData);
            });
          });
        }
      });
    });
  }, [repoClient]);

  const loginCompleted = () => {
    if (window.localStorage.getItem(TempStorageKeys.Filename)) {
      dialog.show();
    }
    getAndInitializeRepositoryClientAndServicesAsync().then(() => {
      GetFileData().then(async (fileData) => {
        // TODO repoclient isn't always assigned here bc of state..
        saveFileToLaserfiche(fileData);
      });
    });
  };

  //Laserfiche LF logoutCompleted
  const logoutCompleted = () => {
    //dialog.close();

    setLoggedIn(false);
    window.location.href =
      props.context.pageContext.web.absoluteUrl + props.laserficheRedirectUrl;
  };

  function saveFileToLaserfiche(fileData: Blob) {
    const fileName = window.localStorage.getItem(TempStorageKeys.Filename);
    if (fileName && fileData && repoClient) {
      const LContType = window.localStorage.getItem(TempStorageKeys.LContType);
      if (LContType) {
        SendToLaserficheWithMapping(fileData);
      } else {
        SendtoLaserficheNoMapping(fileData);
      }
    }
  }

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
      setLoggedIn(true);
    }
  };

  async function SendToLaserficheWithMapping(fileData: Blob) {
    const fileDataStuff = getFileDataFromLocalStorage();

    let request: PostEntryWithEdocMetadataRequest;
    if (fileDataStuff.DocTemplate?.length > 0 && fileDataStuff.DocTemplate !== 'undefined') {
      request = getRequestMetadata(fileDataStuff, request);
    } else {
      request = new PostEntryWithEdocMetadataRequest({});
    }

    const fileExtensionPeriod = PathUtils.getCleanedExtension(
      fileDataStuff.Filename
    );
    const filenameWithoutExt = PathUtils.removeFileExtension(
      fileDataStuff.Filename
    );
    const docNameIncludesFileName =
      fileDataStuff.Documentname.includes('FileName');

    const edocBlob: Blob = fileData as unknown as Blob;
    const destinationFolder = (!fileDataStuff.Destinationfolder || fileDataStuff.Destinationfolder==='undefined') ? '1': fileDataStuff.Destinationfolder;
    const parentEntryId = Number(destinationFolder);
    const repoId = await repoClient.getCurrentRepoId();

    let fileName: string | undefined;
    let fileNameInEdoc: string | undefined;
    let extension: string | undefined;
    if (fileDataStuff.Documentname === '') {
      fileName = filenameWithoutExt;
      fileNameInEdoc = fileDataStuff.Filename;
      extension = fileExtensionPeriod;
    } else if (docNameIncludesFileName === false) {
      fileName = fileDataStuff.Documentname;
      fileNameInEdoc = fileDataStuff.Documentname;
      extension = fileExtensionPeriod;
    } else {
      const DocnameReplacedwithfilename = fileDataStuff.Documentname.replace(
        'FileName',
        filenameWithoutExt
      );
      fileName = DocnameReplacedwithfilename;
      fileNameInEdoc = DocnameReplacedwithfilename + `.${fileExtensionPeriod}`;
      extension = fileExtensionPeriod;
    }
    const electronicDocument: FileParameter = {
      fileName: fileNameInEdoc,
      data: edocBlob,
    };
    const entryRequest = {
      repoId,
      parentEntryId,
      fileName,
      autoRename: true,
      electronicDocument,
      request,
      extension,
    };
    try {
      const entryCreateResult: CreateEntryResult =
        await repoClient.entriesClient.importDocument(entryRequest);
      const Entryid = entryCreateResult.operations.entryCreate.entryId ?? 1;
      const fileLink = getEntryWebAccessUrl(
        Entryid.toString(),
        repoId,
        webClientUrl,
        false
      );
      await dialog.close();
      dialog.fileLink = fileLink;
      dialog.isLoading = false;
      dialog.metadataSaved = true;
      dialog.show();

      if (fileDataStuff.Action === ActionTypes.COPY) {
        window.localStorage.removeItem(TempStorageKeys.Filename);
      } else if (fileDataStuff.Action === ActionTypes.MOVE_AND_DELETE) {
        DeleteFile(
          fileDataStuff.PageOrigin,
          fileDataStuff.Fileurl,
          fileDataStuff.Filename
        );
      } else if (fileDataStuff.Action === ActionTypes.REPLACE) {
        deletefileandreplace(
          fileDataStuff.PageOrigin,
          fileDataStuff.Fileurl,
          filenameWithoutExt,
          fileDataStuff.Filename,
          fileLink,
          fileDataStuff.ContextPageAbsoluteUrl
        );
      } else {
        // TODO what should happen?
      }
    } catch (error) {
      const conflict409 =
        error.operations.setFields.exceptions[0].statusCode === 409;
      if (conflict409) {
        const entryidConflict1 = error.operations.entryCreate.entryId;

        const fileLink = getEntryWebAccessUrl(
          entryidConflict1.toString(),
          repoId,
          webClientUrl,
          false
        );
        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = false;
        dialog.show();
        window.localStorage.removeItem(TempStorageKeys.Filename);
      } else {
        window.alert(`Error uploding file: ${JSON.stringify(error)}`);
        window.localStorage.removeItem(TempStorageKeys.Filename);
        dialog.close();
      }
    }
  }

  function getRequestMetadata(
    fileDataStuff: {
      Filename: string;
      Destinationfolder: string;
      Filemetadata: string;
      Action: string;
      Documentname: string;
      Fileurl: string;
      ContextPageAbsoluteUrl: string;
      PageOrigin: string;
      DocTemplate: string;
    },
    request: PostEntryWithEdocMetadataRequest
  ) {
    const Filemetadata: IPostEntryWithEdocMetadataRequest = JSON.parse(
      fileDataStuff.Filemetadata
    );
    const fieldsAlone = Filemetadata?.metadata?.fields;
    const formattedFieldValues:
      | {
          [key: string]: FieldToUpdate;
        }
      | undefined = {};

    for (const key in fieldsAlone) {
      const value = fieldsAlone[key];
      formattedFieldValues[key] = new FieldToUpdate({
        ...value,
        values: value.values.map((val) => new ValueToUpdate(val)),
      });
    }
    request = new PostEntryWithEdocMetadataRequest({
      template: fileDataStuff.DocTemplate,
      metadata: new PutFieldValsRequest({
        fields: formattedFieldValues,
      }),
    });
    return request;
  }

  function getFileDataFromLocalStorage() {
    return {
      Filename: window.localStorage.getItem(TempStorageKeys.Filename),
      Destinationfolder: window.localStorage.getItem(
        TempStorageKeys.Destinationfolder
      ),
      Filemetadata: window.localStorage.getItem(TempStorageKeys.Filemetadata),
      Action: window.localStorage.getItem(TempStorageKeys.Action),
      Documentname: window.localStorage.getItem(TempStorageKeys.Documentname),
      Fileurl: window.localStorage.getItem(TempStorageKeys.Fileurl),
      ContextPageAbsoluteUrl: window.localStorage.getItem(
        TempStorageKeys.ContextPageAbsoluteUrl
      ),
      PageOrigin: window.localStorage.getItem(TempStorageKeys.PageOrigin),
      DocTemplate: window.localStorage.getItem(TempStorageKeys.DocTemplate),
    };
  }

  async function SendtoLaserficheNoMapping(fileData: Blob) {
    const Filenamewithext = window.localStorage.getItem(
      TempStorageKeys.Filename
    );

    const fileNameSplitByDot = (Filenamewithext as string).split('.');
    const fileExtensionPeriod = fileNameSplitByDot.pop();
    const Filenamewithoutext = fileNameSplitByDot.join('.');

    const edocBlob: Blob = fileData as unknown as Blob;
    const parentEntryId = 1;

    try {
      const repoId = await repoClient.getCurrentRepoId();
      const electronicDocument: FileParameter = {
        fileName: Filenamewithext,
        data: edocBlob,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Filenamewithoutext,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod,
      };

      const entryCreateResult: CreateEntryResult =
        await repoClient.entriesClient.importDocument(entryRequest);
      const Entryid14 = entryCreateResult.operations.entryCreate.entryId;
      const fileLink = getEntryWebAccessUrl(
        Entryid14.toString(),
        repoId,
        webClientUrl,
        false
      );

      await dialog.close();
      dialog.fileLink = fileLink;
      dialog.isLoading = false;
      dialog.metadataSaved = true;
      dialog.show();
      window.localStorage.removeItem(TempStorageKeys.Filename);
    } catch (error) {
      window.alert(`Error uploding file: ${JSON.stringify(error)}`);
      window.localStorage.removeItem(TempStorageKeys.Filename);
      dialog.close();
    }
  }

  async function GetFileData() {
    const Fileurl = window.localStorage.getItem(TempStorageKeys.Fileurl);
    const pageOrigin = window.localStorage.getItem(TempStorageKeys.PageOrigin);
    const Filenamewithext2 = window.localStorage.getItem(
      TempStorageKeys.Filename
    );
    const encde = encodeURIComponent(Filenamewithext2);
    const fileur = Fileurl?.replace(Filenamewithext2, encde);
    const Filedataurl = pageOrigin + fileur;
    try {
      const res = await fetch(Filedataurl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = await res.blob();
      return results;
    } catch (error) {
      dialog.close();
      console.log('error occured' + error);
    }
  }

  function Redirect() {
    const pageOrigin = window.localStorage.getItem(TempStorageKeys.PageOrigin);
    const Fileurl = window.localStorage.getItem(TempStorageKeys.Fileurl);
    const Filenamewithext1 = window.localStorage.getItem(
      TempStorageKeys.Filename
    );
    const fileeee = Fileurl.replace(Filenamewithext1, '');
    const path = pageOrigin + fileeee;
    Navigation.navigate(path, true);
  }

  async function DeleteFile(
    pageOrigin: string,
    fileUrl: string,
    filenameWithExt: string
  ) {
    const encde = encodeURIComponent(filenameWithExt);
    const fileur = fileUrl.replace(filenameWithExt, encde);
    const fileUrl1 = pageOrigin + fileur;
    const init: RequestInit = {
      headers: {
        Accept: 'application/json;odata=verbose',
      },
      method: 'DELETE',
    };
    const response = await fetch(fileUrl1, init);
    if (response.ok) {
      window.localStorage.removeItem(TempStorageKeys.Filename);
      //Perform further activity upon success, like displaying a notification
      alert('File deleted successfully');
    } else {
      window.localStorage.removeItem(TempStorageKeys.Filename);
      console.log('An error occurred. Please try again.');
    }
  }

  async function deletefileandreplace(
    pageOrigin: string,
    fileUrl: string,
    filenameWithoutExt: string,
    filenameWithExt: string,
    docFilelink: string,
    contexPageAbsoluteUrl: string
  ) {
    const encde = encodeURIComponent(filenameWithExt);
    const fileur = fileUrl.replace(filenameWithExt, encde);
    const fileUrl1 = pageOrigin + fileur;
    const deleteFile = await fetch(fileUrl1, {
      method: 'DELETE',
      headers: {
        Accept: 'application/json;odata=verbose',
      },
    });
    if (deleteFile.ok) {
      alert('File replaced with link successfully');
      GetFormDigestValue(
        fileUrl,
        filenameWithoutExt,
        filenameWithExt,
        docFilelink,
        contexPageAbsoluteUrl
      );
    } else {
      window.localStorage.removeItem(TempStorageKeys.Filename);
      console.log('An error occurred. Please try again.');
    }
  }
  //
  async function GetFormDigestValue(
    fileUrl: string,
    filenameWithoutExt: string,
    filenameWithExt: string,
    docFileLink: string,
    contextPageAbsoluteUrl: string
  ) {
    const resp = await fetch(contextPageAbsoluteUrl + '/_api/contextinfo', {
      method: 'POST',
      headers: { accept: 'application/json;odata=verbose' },
    });
    if (resp.ok) {
      const data = await resp.json();
      const FormDigestValue = data.d.GetContextWebInformation.FormDigestValue;
      postlink(
        fileUrl,
        filenameWithoutExt,
        filenameWithExt,
        docFileLink,
        contextPageAbsoluteUrl,
        FormDigestValue
      );
    } else {
      window.localStorage.removeItem(TempStorageKeys.Filename);
      console.log('Failed');
    }
  }
  //
  async function postlink(
    fileUrl: string,
    filenameWithoutExt: string,
    filenameWithExt: string,
    docFilelink: string,
    contextPageAbsoluteUrl: string,
    formDigestValue: string
  ) {
    const encde1 = encodeURIComponent(filenameWithoutExt);
    const path = fileUrl.replace(filenameWithExt, '');
    const AddLinkURL =
      contextPageAbsoluteUrl +
      `/_api/web/GetFolderByServerRelativeUrl('${path}')/Files/add(url='${encde1}.url',overwrite=true)`;

    const resp = await fetch(AddLinkURL, {
      method: 'POST',
      body: `[InternetShortcut]\nURL=${docFilelink}`,
      headers: {
        'content-type': 'text/plain',
        accept: 'application/json;odata=verbose',
        'X-RequestDigest': formDigestValue,
      },
    });
    if (resp.ok) {
      window.localStorage.removeItem(TempStorageKeys.Filename);
      console.log('Item Inserted..!!');
      console.log(await resp.json());
    } else {
      window.localStorage.removeItem(TempStorageKeys.Filename);
      console.log('API Error');
      console.log(await resp.json());
    }
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
