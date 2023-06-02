import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import styles from './SendToLaserFiche.module.scss';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import { clientId } from '../../webparts/constants';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { RepositoryClientExInternal } from '../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../repository-client/repository-client-types';
import { ISPDocumentData } from '../../Utils/Types';
import { getEntryWebAccessUrl } from '../../Utils/Funcs';
import {
  PostEntryWithEdocMetadataRequest,
  FileParameter,
  CreateEntryResult,
  IPostEntryWithEdocMetadataRequest,
  FieldToUpdate,
  ValueToUpdate,
  PutFieldValsRequest,
} from '@laserfiche/lf-repository-api-client';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Navigation } from 'spfx-navigation';

export default class SaveToLaserficheCustomDialog extends BaseDialog {
  missingFields?: JSX.Element[];
  spMetadata?: ISPDocumentData;
  successful = false;

  handleCloseClickAsync = async (success: boolean) => {
    this.successful = success;
    await this.close();
  }

  public render(): void {
    const element: React.ReactElement = (
      <SaveToLaserficheDialog
        missingFields={this.missingFields}
        spFileMetadata={this.spMetadata}
        closeClick={this.handleCloseClickAsync}
      />
    );
    ReactDOM.render(element, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false,
    };
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
  }
}

function SaveToLaserficheDialog(props: {
  missingFields?: JSX.Element[];
  closeClick: (success: boolean) => Promise<void>;
  spFileMetadata: ISPDocumentData;
}) {
  const loginComponent = React.createRef<
    NgElement & WithProperties<LfLoginComponent>
  >();

  const [success, setSuccess] = React.useState<
    { fileLink: string; pathBack: string; metadataSaved: boolean } | undefined
  >();

  React.useEffect(() => {
    SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
    ).then(() => {
      SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      ).then(async () => {
        if (loginComponent.current?.authorization_credentials) {
          getAndInitializeRepositoryClientAndServicesAsync();
        } else {
          await props.closeClick(false);
        }
      });
    });
  }, []);

  let textInside = <span>Saving your document to Laserfiche</span>;
  if (props.missingFields) {
    <span>
      The following SharePoint field values are blank and are mapped to required
      Laserfiche fields:
      {props.missingFields}Please fill out these required fields and try again.
    </span>;
  } else if (success) {
    textInside = success.metadataSaved ? (
      <span>Document uploaded</span>
    ) : (
      <span>
        Document uploaded to repository, updating metadata failed due to
        constraint mismatch
        <br />{' '}
        <p style={{ color: 'red' }}>
          The Laserfiche template and fields were not applied to this document
        </p>
      </span>
    );
  }

  async function getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken =
      loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      const repoClient = await ensureRepoClientInitializedAsync();
      const webClientUrl =
        loginComponent?.current?.account_endpoints.webClientUrl;

      if (repoClient && props.spFileMetadata) {
        GetFileData().then(async (fileData) => {
          saveFileToLaserfiche(fileData, repoClient, webClientUrl);
        });
      } else {
        await props.closeClick(false);
      }
    } else {
      // user is not logged in
    }
  }

  async function ensureRepoClientInitializedAsync(): Promise<IRepositoryApiClientExInternal> {
    const repoClientCreator = new RepositoryClientExInternal();
    const newRepoClient = await repoClientCreator.createRepositoryClientAsync();
    // test accessToken validity
    try {
      await newRepoClient.repositoriesClient.getRepositoryList({});
    } catch {
      return undefined;
    }
    return newRepoClient;
  }

  async function GetFileData() {
    const Fileurl = props.spFileMetadata.fileUrl;
    const pageOrigin = props.spFileMetadata.pageOrigin;
    const Filenamewithext2 = props.spFileMetadata.fileName;
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
      console.log('error occured' + error);
    }
  }

  function saveFileToLaserfiche(
    fileData: Blob,
    repoClient: IRepositoryApiClientExInternal,
    webClientUrl: string
  ) {
    if (props.spFileMetadata.fileName && fileData && repoClient) {
      const laserficheProfileName = props.spFileMetadata.lfProfile;
      if (laserficheProfileName) {
        SendToLaserficheWithMapping(fileData, repoClient, webClientUrl);
      } else {
        SendtoLaserficheNoMapping(fileData, repoClient, webClientUrl);
      }
    }
  }

  async function SendToLaserficheWithMapping(
    fileData: Blob,
    repoClient: IRepositoryApiClientExInternal,
    webClientUrl: string
  ) {
    let request: PostEntryWithEdocMetadataRequest;
    if (props.spFileMetadata.templateName) {
      request = getRequestMetadata(props.spFileMetadata, request);
    } else {
      request = new PostEntryWithEdocMetadataRequest({});
    }

    const fileExtensionPeriod = PathUtils.getCleanedExtension(
      props.spFileMetadata.fileName
    );
    const filenameWithoutExt = PathUtils.removeFileExtension(
      props.spFileMetadata.fileName
    );
    const docNameIncludesFileName =
      props.spFileMetadata.documentName.includes('FileName');

    const edocBlob: Blob = fileData as unknown as Blob;
    const parentEntryId = Number(props.spFileMetadata.entryId);
    const repoId = await repoClient.getCurrentRepoId();

    let fileName: string | undefined;
    let fileNameInEdoc: string | undefined;
    let extension: string | undefined;
    if (!props.spFileMetadata.documentName) {
      fileName = filenameWithoutExt;
      fileNameInEdoc = props.spFileMetadata.fileName;
      extension = fileExtensionPeriod;
    } else if (docNameIncludesFileName === false) {
      fileName = props.spFileMetadata.documentName;
      fileNameInEdoc = props.spFileMetadata.documentName;
      extension = fileExtensionPeriod;
    } else {
      const DocnameReplacedwithfilename =
        props.spFileMetadata.documentName.replace(
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
      const pageOrigin = props.spFileMetadata.pageOrigin;
      const fileUrl = props.spFileMetadata.fileUrl;
      const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
      const path = pageOrigin + fileUrlWithoutDocName;
      setSuccess({ fileLink, pathBack: path, metadataSaved: true });

      if (props.spFileMetadata.action === ActionTypes.COPY) {
        window.localStorage.removeItem('spdocdata');
      } else if (props.spFileMetadata.action === ActionTypes.MOVE_AND_DELETE) {
        DeleteFile(
          props.spFileMetadata.pageOrigin,
          props.spFileMetadata.fileUrl,
          props.spFileMetadata.fileName
        );
      } else if (props.spFileMetadata.action === ActionTypes.REPLACE) {
        deletefileandreplace(
          props.spFileMetadata.pageOrigin,
          props.spFileMetadata.fileUrl,
          filenameWithoutExt,
          props.spFileMetadata.fileName,
          fileLink,
          props.spFileMetadata.contextPageAbsoluteUrl
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
        const pageOrigin = props.spFileMetadata.pageOrigin;
        const fileUrl = props.spFileMetadata.fileUrl;
        const fileUrlWithoutDocName = fileUrl.slice(
          0,
          fileUrl.lastIndexOf('/')
        );
        const path = pageOrigin + fileUrlWithoutDocName;
        setSuccess({ fileLink, pathBack: path, metadataSaved: false });
        window.localStorage.removeItem('spdocdata');
      } else {
        window.alert(`Error uploding file: ${JSON.stringify(error)}`);
        window.localStorage.removeItem('spdocdata');
      }
    }
  }

  function getRequestMetadata(
    fileDataStuff: ISPDocumentData,
    request: PostEntryWithEdocMetadataRequest
  ) {
    const Filemetadata: IPostEntryWithEdocMetadataRequest =
      fileDataStuff.metadata;
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

  async function SendtoLaserficheNoMapping(
    fileData: Blob,
    repoClient: IRepositoryApiClientExInternal,
    webClientUrl: string
  ) {
    const Filenamewithext = props.spFileMetadata.fileName;

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
      const pageOrigin = props.spFileMetadata.pageOrigin;
      const fileUrl = props.spFileMetadata.fileUrl;
      const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
      const path = pageOrigin + fileUrlWithoutDocName;

      setSuccess({ fileLink, pathBack: path, metadataSaved: true });
      window.localStorage.removeItem('spdocdata');
    } catch (error) {
      window.alert(`Error uploding file: ${JSON.stringify(error)}`);
      window.localStorage.removeItem('spdocdata');
    }
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

  function viewFile() {
    window.open(success.fileLink);
  }

  function redirect() {
    Navigation.navigate(success.pathBack, true);
  }

  return (
    <div className={styles.maindialog}>
      <lf-login
        hidden
        redirect_uri='https://lfdevm365.sharepoint.com/sites/TestSite/Shared%20Documents/Forms/AllItems.aspx'
        authorize_url_host_name='a.clouddev.laserfiche.com'
        redirect_behavior='Replace'
        client_id={clientId}
        ref={loginComponent}
      />
      <div id='overlay' className={styles.overlay} />
      <div>
        <img
          src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
          width='42'
          height='42'
        />
      </div>
      {!success && !props.missingFields && (
        <img
          style={{ marginLeft: '28%' }}
          src='/_layouts/15/images/progress.gif'
          id='imgid'
        />
      )}
      <div>
        <p className={styles.text} id='it'>
          {textInside}
        </p>
      </div>
      {success && (
        <>
          <div id='divid' className={styles.button}>
            <button
              id='divid1'
              className={styles.button1}
              onClick={async () => { await props.closeClick(true)}}
            >
              Close
            </button>
            {success.fileLink && (
              <button
                id='divid13'
                className={styles.button2}
                title='Click here to view the file in Laserfiche'
                onClick={viewFile}
              >
                Go to File
              </button>
            )}
            <button
              id='divid14'
              className={styles.button2}
              title='Click here to go back to your SharePoint library'
              onClick={redirect}
            >
              Go to Library
            </button>
          </div>
        </>
      )}
    </div>
  );
}
