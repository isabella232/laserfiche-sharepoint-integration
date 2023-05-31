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
import { getEntryWebAccessUrl } from '../../../Utils/Funcs';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISPDocumentData } from '../../../Utils/Types';

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

  const spFileMetadata = JSON.parse(window.localStorage.getItem('spdocdata')) as ISPDocumentData;

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

        if (loggedIn) {
          if (spFileMetadata) {
            dialog.show();
          }
          getAndInitializeRepositoryClientAndServicesAsync();
        }
      });
    });
  }, []);

  React.useEffect(() => {
    if (repoClient && spFileMetadata) {
      GetFileData().then(async (fileData) => {
        saveFileToLaserfiche(fileData);
      });
    }
  }, [repoClient, spFileMetadata]);

  const loginCompleted = () => {
    if (spFileMetadata) {
      dialog.show();
    }
    getAndInitializeRepositoryClientAndServicesAsync();
  };

  const logoutCompleted = () => {
    setLoggedIn(false);
    window.location.href =
      props.context.pageContext.web.absoluteUrl + props.laserficheRedirectUrl;
  };

  function saveFileToLaserfiche(fileData: Blob) {
    if (spFileMetadata.fileName && fileData && repoClient) {
      const laserficheProfileName = spFileMetadata.lfProfile;
      if (laserficheProfileName) {
        SendToLaserficheWithMapping(fileData, spFileMetadata);
      } else {
        SendtoLaserficheNoMapping(fileData, spFileMetadata);
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
      const newRepoClient = await repoClientCreator.createRepositoryClientAsync();
      setRepoClient(newRepoClient);
      setLoggedIn(true);
    }
  };

  async function SendToLaserficheWithMapping(fileData: Blob, spFileMetadata: ISPDocumentData) {
    let request: PostEntryWithEdocMetadataRequest;
    if (spFileMetadata.templateName) {
      request = getRequestMetadata(spFileMetadata, request);
    } else {
      request = new PostEntryWithEdocMetadataRequest({});
    }

    const fileExtensionPeriod = PathUtils.getCleanedExtension(
      spFileMetadata.fileName
    );
    const filenameWithoutExt = PathUtils.removeFileExtension(
      spFileMetadata.fileName
    );
    const docNameIncludesFileName =
      spFileMetadata.documentName.includes('FileName');

    const edocBlob: Blob = fileData as unknown as Blob;
    const parentEntryId = Number(spFileMetadata.entryId);
    const repoId = await repoClient.getCurrentRepoId();

    let fileName: string | undefined;
    let fileNameInEdoc: string | undefined;
    let extension: string | undefined;
    if (!spFileMetadata.documentName) {
      fileName = filenameWithoutExt;
      fileNameInEdoc = spFileMetadata.fileName;
      extension = fileExtensionPeriod;
    } else if (docNameIncludesFileName === false) {
      fileName = spFileMetadata.documentName;
      fileNameInEdoc = spFileMetadata.documentName;
      extension = fileExtensionPeriod;
    } else {
      const DocnameReplacedwithfilename = spFileMetadata.documentName.replace(
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
      const pageOrigin = spFileMetadata.pageOrigin;
      const fileUrl = spFileMetadata.fileUrl;
      const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
      const path = pageOrigin + fileUrlWithoutDocName;
      await dialog.close();
      dialog.fileLink = fileLink;
      dialog.pathBack = path;
      dialog.isLoading = false;
      dialog.metadataSaved = true;
      dialog.show();

      if (spFileMetadata.action === ActionTypes.COPY) {
        window.localStorage.removeItem('spdocdata');
      } else if (spFileMetadata.action === ActionTypes.MOVE_AND_DELETE) {
        DeleteFile(
          spFileMetadata.pageOrigin,
          spFileMetadata.fileUrl,
          spFileMetadata.fileName
        );
      } else if (spFileMetadata.action === ActionTypes.REPLACE) {
        deletefileandreplace(
          spFileMetadata.pageOrigin,
          spFileMetadata.fileUrl,
          filenameWithoutExt,
          spFileMetadata.fileName,
          fileLink,
          spFileMetadata.contextPageAbsoluteUrl
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
        const pageOrigin = spFileMetadata.pageOrigin;
        const fileUrl = spFileMetadata.fileUrl;
        const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
        const path = pageOrigin + fileUrlWithoutDocName;
        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.pathBack = path;
        dialog.isLoading = false;
        dialog.metadataSaved = false;
        dialog.show();
        window.localStorage.removeItem('spdocdata');
      } else {
        window.alert(`Error uploding file: ${JSON.stringify(error)}`);
        window.localStorage.removeItem('spdocdata');
        dialog.close();
      }
    }
  }

  function getRequestMetadata(
    fileDataStuff: ISPDocumentData,
    request: PostEntryWithEdocMetadataRequest
  ) {
    const Filemetadata: IPostEntryWithEdocMetadataRequest = fileDataStuff.metadata;
    const fieldsAlone = Filemetadata.metadata.fields;
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
      template: Filemetadata.template,
      metadata: new PutFieldValsRequest({
        fields: formattedFieldValues,
      }),
    });
    return request;
  }

  async function SendtoLaserficheNoMapping(fileData: Blob, spFileMetadata: ISPDocumentData) {
    const Filenamewithext = spFileMetadata.fileName;

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
      const pageOrigin = spFileMetadata.pageOrigin;
      const fileUrl = spFileMetadata.fileUrl;
      const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
      const path = pageOrigin + fileUrlWithoutDocName;

      await dialog.close();
      dialog.fileLink = fileLink;
      dialog.pathBack = path;
      dialog.isLoading = false;
      dialog.metadataSaved = true;
      dialog.show();
      window.localStorage.removeItem('spdocdata');
    } catch (error) {
      window.alert(`Error uploding file: ${JSON.stringify(error)}`);
      window.localStorage.removeItem('spdocdata');
      dialog.close();
    }
  }

  async function GetFileData() {
    const Fileurl = spFileMetadata.fileUrl;
    const pageOrigin = spFileMetadata.pageOrigin;
    const Filenamewithext2 = spFileMetadata.fileName;
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
    const Fileurl = spFileMetadata.fileUrl;
    const pageOrigin = spFileMetadata.pageOrigin;
    const Filenamewithext1 = spFileMetadata.fileName;
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
      window.localStorage.removeItem('spdocdata');
      //Perform further activity upon success, like displaying a notification
      alert('File deleted successfully');
    } else {
      window.localStorage.removeItem('spdocdata');
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
      window.localStorage.removeItem('spdocdata');
      console.log('An error occurred. Please try again.');
    }
  }

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
      window.localStorage.removeItem('spdocdata');
      console.log('Failed');
    }
  }
  
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
      window.localStorage.removeItem('spdocdata');
      console.log('Item Inserted..!!');
      console.log(await resp.json());
    } else {
      window.localStorage.removeItem('spdocdata');
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
