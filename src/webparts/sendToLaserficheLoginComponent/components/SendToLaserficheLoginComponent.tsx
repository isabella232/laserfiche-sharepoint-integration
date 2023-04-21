import * as React from 'react';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISendToLaserficheLoginComponentState } from './ISendToLaserficheLoginComponentState';
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
} from '@laserfiche/lf-repository-api-client';
import {
  LfLoginComponent,
  LoginState,
} from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { clientId } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ActionTypes } from '../../laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { TempStorageKeys } from '../../../Utils/Enums';
import { getEntryWebAccessUrl } from '../../../Utils/Funcs';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ['lf-login']: any;
    }
  }
}

const dialog = new SendToLaserficheCustomDialog();
let filelink = '';
export default class SendToLaserficheLoginComponent extends React.Component<
  ISendToLaserficheLoginComponentProps,
  ISendToLaserficheLoginComponentState
> {
  public loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  >;
  public repoClient: IRepositoryApiClientExInternal;

  constructor(props: ISendToLaserficheLoginComponentProps) {
    super(props);
    this.loginComponent = React.createRef();
    this.loginComponent = React.createRef();

    this.state = {
      baseurl: '',
      filelink: '',
      filedata: '',
      accessToken: '',
      parentItemId: 1,
      repoId: '',
      region: this.props.devMode
        ? 'a.clouddev.laserfiche.com'
        : 'laserfiche.com',
      webClientUrl: '',
    };
  }
  public async componentDidMount(): Promise<void> {
    await SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
    );
    await SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
    this.loginComponent.current.addEventListener(
      'loginCompleted',
      this.loginCompleted
    );
    this.loginComponent.current.addEventListener(
      'logoutCompleted',
      this.logoutCompleted
    );

    const loggedOut: boolean =
      this.loginComponent.current.state === LoginState.LoggedOut;

    if (!loggedOut) {
      document.getElementById('remve').innerText =
        'You are signed in to Laserfiche';
      document.getElementById('remveHeading').innerText = 'Sign out';
      if (window.localStorage.getItem(TempStorageKeys.LContType)) {
        dialog.show();
      }
      this.getAndInitializeRepositoryClientAndServicesAsync().then(() => {
        this.GetFileData().then(async (results) => {
          this.setState({ filedata: results });
          const DocTemplate = window.localStorage.getItem(
            TempStorageKeys.DocTemplate
          );
          const LContType = window.localStorage.getItem(
            TempStorageKeys.LContType
          );
          if (LContType != 'undefined' && LContType !== null) {
            if (DocTemplate != 'None') {
              this.SendToLaserficheWithMetadata();
            } else {
              this.SendToLaserficheNoTemplate();
            }
          } else if (LContType !== null) {
            this.SendtoLaserficheNoMapping();
          } else {
            dialog.close();
          }
        });
      });
    }
  }

  public loginCompleted = async () => {
    document.getElementById('remve').innerText =
      'You are signed in to Laserfiche';
    document.getElementById('remveHeading').innerText = 'Sign out';
    if (window.localStorage.getItem(TempStorageKeys.LContType)) {
      dialog.show();
    }
    this.getAndInitializeRepositoryClientAndServicesAsync().then(() => {
      this.GetFileData().then(async (results) => {
        this.setState({ filedata: results });
        const DocTemplate = window.localStorage.getItem(
          TempStorageKeys.DocTemplate
        );
        const LContType = window.localStorage.getItem(
          TempStorageKeys.LContType
        );
        if (LContType != 'undefined' && LContType !== null) {
          if (DocTemplate != 'None') {
            this.SendToLaserficheWithMetadata();
          } else {
            this.SendToLaserficheNoTemplate();
          }
        } else if (LContType !== null) {
          this.SendtoLaserficheNoMapping();
        } else {
          dialog.close();
          //alert('Please go back to library and select a file to upload');
        }
      });
    });
  };

  //Laserfiche LF logoutCompleted
  public logoutCompleted = async () => {
    //dialog.close();
    window.location.href =
      this.props.context.pageContext.web.absoluteUrl +
      this.props.laserficheRedirectPage;
  };
  private async getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken =
      this.loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await this.ensureRepoClientInitializedAsync();

      this.setState({
        accessToken:
          this.loginComponent.current.authorization_credentials.accessToken,
        webClientUrl:
          this.loginComponent.current.account_endpoints.webClientUrl,
      });
    } else {
      // user is not logged in
    }
  }
  public async ensureRepoClientInitializedAsync(): Promise<void> {
    if (!this.repoClient) {
      const repoClientCreator = new RepositoryClientExInternal();
      this.repoClient = await repoClientCreator.createRepositoryClientAsync();
    }
  }

  public async SendToLaserficheWithMetadata() {
    const filenameWithExt = window.localStorage.getItem(
      TempStorageKeys.Filename
    );

    const fileNameSplitByDot = (filenameWithExt as string).split('.');
    const fileExtensionPeriod = fileNameSplitByDot.pop();
    const filenameWithoutExt = fileNameSplitByDot.join('.');
    const Parentid = window.localStorage.getItem(
      TempStorageKeys.Destinationfolder
    );
    const Filemetadata1 = window.localStorage.getItem(
      TempStorageKeys.Filemetadata
    );
    const Filemetadata = JSON.parse(Filemetadata1);
    const Action = window.localStorage.getItem(TempStorageKeys.Action);
    const Documentname = window.localStorage.getItem(
      TempStorageKeys.Documentname
    );
    const docfilenamecheck = Documentname.includes('FileName');
    const fileUrl = window.localStorage.getItem(TempStorageKeys.Fileurl);
    const contextPageAbsoluteUrl = window.localStorage.getItem(
      TempStorageKeys.ContextPageAbsoluteUrl
    );
    const pageOrigin = window.localStorage.getItem(TempStorageKeys.PageOrigin);
    const fieldsAlone = Filemetadata['metadata']['fields'];
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

    const edocBlob: Blob = this.state.filedata as unknown as Blob;
    const parentEntryId = Number(Parentid);
    const fieldsAndMetadata = Filemetadata; /* JSON.parse(Filemetadata); */
    // TODO make sure this matches the correct format
    const request: PostEntryWithEdocMetadataRequest =
      new PostEntryWithEdocMetadataRequest({
        template: fieldsAndMetadata['template'],
        metadata: new PutFieldValsRequest({
          fields: formattedFieldValues,
        }),
      });
    const repoId = await this.repoClient.getCurrentRepoId();
    if (Documentname === '') {
      const electronicDocument: FileParameter = {
        fileName: filenameWithExt,
        data: edocBlob,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: filenameWithoutExt,
        autoRename: true,
        electronicDocument,
        request,
        extension: fileExtensionPeriod,
      };
      try {
        const entryCreateResult: CreateEntryResult =
          await this.repoClient.entriesClient.importDocument(entryRequest);
        const Entryid = entryCreateResult.operations.entryCreate.entryId;
        const fileLink = getEntryWebAccessUrl(
          Entryid.toString(),
          repoId,
          this.state.webClientUrl,
          false
        );
        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = true;
        dialog.show();

        if (Action === ActionTypes.COPY) {
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else if (Action === ActionTypes.MOVE_AND_DELETE) {
          this.DeleteFile(pageOrigin, fileUrl, filenameWithExt);
        } else if (Action === ActionTypes.REPLACE) {
          this.deletefileandreplace(
            pageOrigin,
            fileUrl,
            filenameWithoutExt,
            filenameWithExt,
            filelink,
            contextPageAbsoluteUrl
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
            this.state.webClientUrl,
            false
          );
          await dialog.close();
          dialog.fileLink = fileLink;
          dialog.isLoading = false;
          dialog.metadataSaved = false;
          dialog.show();
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          //window.localStorage.clear();
          window.localStorage.removeItem(TempStorageKeys.LContType);
          dialog.close();
        }
      }
    } else if (docfilenamecheck === false) {
      const electronicDocument: FileParameter = {
        fileName: Documentname,
        data: edocBlob,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Documentname,
        autoRename: true,
        electronicDocument,
        request,
        extension: fileExtensionPeriod,
      };
      try {
        const entryCreateResult: CreateEntryResult =
          await this.repoClient.entriesClient.importDocument(entryRequest);
        const Entryid3 = entryCreateResult.operations.entryCreate.entryId;
        const fileLink = getEntryWebAccessUrl(
          Entryid3.toString(),
          repoId,
          this.state.webClientUrl,
          false
        );

        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = true;
        dialog.show();
        if (Action === ActionTypes.COPY) {
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else if (Action === ActionTypes.MOVE_AND_DELETE) {
          this.DeleteFile(pageOrigin, fileUrl, filenameWithExt);
        } else if (Action === ActionTypes.REPLACE) {
          this.deletefileandreplace(
            pageOrigin,
            fileUrl,
            filenameWithoutExt,
            filenameWithExt,
            filelink,
            contextPageAbsoluteUrl
          );
        } else {
          // TODO what should happen?
        }
      } catch (error) {
        if (error.operations.setFields.exceptions[0].statusCode === 409) {
          const entryidConflict2 = error.operations.entryCreate.entryId;

          const fileLink = getEntryWebAccessUrl(
            entryidConflict2.toString(),
            repoId,
            this.state.webClientUrl,
            false
          );

          await dialog.close();
          dialog.fileLink = fileLink;
          dialog.isLoading = false;
          dialog.metadataSaved = false;
          dialog.show();
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem(TempStorageKeys.LContType);
          dialog.close();
        }
      }
    } else {
      const DocnameReplacedwithfilename = Documentname.replace(
        'FileName',
        filenameWithoutExt
      );

      const electronicDocument: FileParameter = {
        fileName: DocnameReplacedwithfilename + `.${fileExtensionPeriod}`,
        data: edocBlob,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: DocnameReplacedwithfilename,
        autoRename: true,
        electronicDocument,
        request,
        extension: fileExtensionPeriod,
      };
      try {
        const entryCreateResult: CreateEntryResult =
          await this.repoClient.entriesClient.importDocument(entryRequest);
        const Entryid6 = entryCreateResult.operations.entryCreate.entryId;
        const fileLink = getEntryWebAccessUrl(
          Entryid6.toString(),
          repoId,
          this.state.webClientUrl,
          false
        );

        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = true;
        dialog.show();
        if (Action === ActionTypes.COPY) {
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else if (Action === ActionTypes.MOVE_AND_DELETE) {
          this.DeleteFile(pageOrigin, fileUrl, filenameWithExt);
        } else if (Action === ActionTypes.REPLACE) {
          this.deletefileandreplace(
            pageOrigin,
            fileUrl,
            filenameWithoutExt,
            filenameWithExt,
            filelink,
            contextPageAbsoluteUrl
          );
        } else {
          // TODO what to do here
        }
      } catch (error) {
        if (error.operations.setFields.exceptions[0].statusCode === 409) {
          const entryidConflict3 = error.operations.entryCreate.entryId;
          const fileLink = getEntryWebAccessUrl(
            entryidConflict3.toString(),
            repoId,
            this.state.webClientUrl,
            false
          );

          await dialog.close();
          dialog.fileLink = fileLink;
          dialog.isLoading = false;
          dialog.metadataSaved = false;
          dialog.show();
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem(TempStorageKeys.LContType);
          dialog.close();
        }
      }
    }
  }

  public async SendToLaserficheNoTemplate() {
    const Filenamewithext = window.localStorage.getItem(
      TempStorageKeys.Filename
    );

    const fileNameSplitByDot = (Filenamewithext as string).split('.');
    const fileExtensionPeriod = fileNameSplitByDot.pop();
    const Filenamewithoutext = fileNameSplitByDot.join('.');

    const Parentid = window.localStorage.getItem(
      TempStorageKeys.Destinationfolder
    );
    const Action = window.localStorage.getItem(TempStorageKeys.Action);
    const Documentname = window.localStorage.getItem(
      TempStorageKeys.Documentname
    );
    const docfilenamecheck = Documentname.includes('FileName');
    const Fileurl = window.localStorage.getItem(TempStorageKeys.Fileurl);
    const contextPageAbsoluteUrl = window.localStorage.getItem(
      TempStorageKeys.ContextPageAbsoluteUrl
    );
    const pageOrigin = window.localStorage.getItem(TempStorageKeys.PageOrigin);

    const edocBlob: Blob = this.state.filedata as unknown as Blob;
    const parentEntryId = Number(Parentid);
    const repoId = await this.repoClient.getCurrentRepoId();
    if (Documentname === '') {
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

      try {
        const entryCreateResult: CreateEntryResult =
          await this.repoClient.entriesClient.importDocument(entryRequest);
        const Entryid6 = entryCreateResult.operations.entryCreate.entryId;
        const fileLink = getEntryWebAccessUrl(
          Entryid6.toString(),
          repoId,
          this.state.webClientUrl,
          false
        );

        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = true;
        dialog.show();
        if (Action === ActionTypes.COPY) {
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else if (Action === ActionTypes.MOVE_AND_DELETE) {
          this.DeleteFile(pageOrigin, Fileurl, Filenamewithext);
        } else if (Action === ActionTypes.REPLACE) {
          this.deletefileandreplace(
            pageOrigin,
            Fileurl,
            Filenamewithoutext,
            Filenamewithext,
            filelink,
            contextPageAbsoluteUrl
          );
        }
      } catch (error) {
        if (error.operations.setFields.exceptions[0].statusCode === 409) {
          const entryidConflict4 = error.operations.entryCreate.entryId;
          const fileLink = getEntryWebAccessUrl(
            entryidConflict4.toString(),
            repoId,
            this.state.webClientUrl,
            false
          );
          await dialog.close();
          dialog.fileLink = fileLink;
          dialog.isLoading = false;
          dialog.metadataSaved = false;
          dialog.show();
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem(TempStorageKeys.LContType);
          dialog.close();
        }
      }
    } else if (docfilenamecheck === false) {
      const electronicDocument: FileParameter = {
        fileName: Documentname,
        data: edocBlob,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Documentname,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod,
      };

      try {
        const entryCreateResult: CreateEntryResult =
          await this.repoClient.entriesClient.importDocument(entryRequest);
        const Entryid9 = entryCreateResult.operations.entryCreate.entryId;
        const fileLink = getEntryWebAccessUrl(
          Entryid9.toString(),
          repoId,
          this.state.webClientUrl,
          false
        );

        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = true;
        dialog.show();
        if (Action === ActionTypes.COPY) {
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else if (Action === ActionTypes.MOVE_AND_DELETE) {
          this.DeleteFile(pageOrigin, Fileurl, Filenamewithext);
        } else if (Action === ActionTypes.REPLACE) {
          this.deletefileandreplace(
            pageOrigin,
            Fileurl,
            Filenamewithoutext,
            Filenamewithext,
            filelink,
            contextPageAbsoluteUrl
          );
        } else {
          // TODO
        }
      } catch (error) {
        if (error.operations.setFields.exceptions[0].statusCode === 409) {
          const entryidConflict5 = error.operations.entryCreate.entryId;
          const fileLink = getEntryWebAccessUrl(
            entryidConflict5.toString(),
            repoId,
            this.state.webClientUrl,
            false
          );

          await dialog.close();
          dialog.fileLink = fileLink;
          dialog.isLoading = false;
          dialog.metadataSaved = false;
          dialog.show();
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem(TempStorageKeys.LContType);
          dialog.close();
        }
      }
    } else {
      const DocnameReplacedwithfilename = Documentname.replace(
        'FileName',
        Filenamewithoutext
      );

      const electronicDocument: FileParameter = {
        fileName: DocnameReplacedwithfilename + `.${fileExtensionPeriod}`,
        data: edocBlob,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: DocnameReplacedwithfilename,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod,
      };

      try {
        const entryCreateResult: CreateEntryResult =
          await this.repoClient.entriesClient.importDocument(entryRequest);
        const Entryid14 = entryCreateResult.operations.entryCreate.entryId;
        const fileLink = getEntryWebAccessUrl(
          Entryid14.toString(),
          repoId,
          this.state.webClientUrl,
          false
        );

        await dialog.close();
        dialog.fileLink = fileLink;
        dialog.isLoading = false;
        dialog.metadataSaved = true;
        dialog.show();
        if (Action === ActionTypes.COPY) {
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else if (Action === ActionTypes.MOVE_AND_DELETE) {
          this.DeleteFile(pageOrigin, Fileurl, Filenamewithext);
        } else if (Action === ActionTypes.REPLACE) {
          this.deletefileandreplace(
            pageOrigin,
            Fileurl,
            Filenamewithoutext,
            Filenamewithext,
            filelink,
            contextPageAbsoluteUrl
          );
        } else {
          // TODO what to do
        }
      } catch (error) {
        if (error.operations.setFields.exceptions[0].statusCode === 409) {
          const entryidConflict6 = error.operations.entryCreate.entryId;
          const fileLink = getEntryWebAccessUrl(
            entryidConflict6.toString(),
            repoId,
            this.state.webClientUrl,
            false
          );
          await dialog.close();
          dialog.fileLink = fileLink;
          dialog.isLoading = false;
          dialog.metadataSaved = false;
          dialog.show();
          window.localStorage.removeItem(TempStorageKeys.LContType);
        } else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem(TempStorageKeys.LContType);
          dialog.close();
        }
      }
    }
  }

  public async SendtoLaserficheNoMapping() {
    const Filenamewithext = window.localStorage.getItem(
      TempStorageKeys.Filename
    );

    const fileNameSplitByDot = (Filenamewithext as string).split('.');
    const fileExtensionPeriod = fileNameSplitByDot.pop();
    const Filenamewithoutext = fileNameSplitByDot.join('.');

    const edocBlob: Blob = this.state.filedata as unknown as Blob;
    const parentEntryId = 1;

    try {
      const repoId = await this.repoClient.getCurrentRepoId();
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
        await this.repoClient.entriesClient.importDocument(entryRequest);
      const Entryid14 = entryCreateResult.operations.entryCreate.entryId;
      const fileLink = getEntryWebAccessUrl(
        Entryid14.toString(),
        repoId,
        this.state.webClientUrl,
        false
      );

      await dialog.close();
      dialog.fileLink = fileLink;
      dialog.isLoading = false;
      dialog.metadataSaved = true;
      dialog.show();
      window.localStorage.removeItem(TempStorageKeys.LContType);
    } catch (error) {
      window.alert(`Error uploding file: ${JSON.stringify(error)}`);
      window.localStorage.removeItem(TempStorageKeys.LContType);
      dialog.close();
    }
  }

  public async GetFileData() {
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

  public Redirect() {
    const pageOrigin = window.localStorage.getItem(TempStorageKeys.PageOrigin);
    const Fileurl = window.localStorage.getItem(TempStorageKeys.Fileurl);
    const Filenamewithext1 = window.localStorage.getItem(
      TempStorageKeys.Filename
    );
    const fileeee = Fileurl.replace(Filenamewithext1, '');
    const path = pageOrigin + fileeee;
    Navigation.navigate(path, true);
  }

  private Dc() {
    dialog.close();
  }

  public async DeleteFile(
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
      window.localStorage.removeItem(TempStorageKeys.LContType);
      //Perform further activity upon success, like displaying a notification
      alert('File deleted successfully');
    } else {
      window.localStorage.removeItem(TempStorageKeys.LContType);
      console.log('An error occurred. Please try again.');
    }
  }

  public async deletefileandreplace(
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
      this.GetFormDigestValue(
        fileUrl,
        filenameWithoutExt,
        filenameWithExt,
        docFilelink,
        contexPageAbsoluteUrl
      );
    } else {
      window.localStorage.removeItem(TempStorageKeys.LContType);
      console.log('An error occurred. Please try again.');
    }
  }
  //
  public async GetFormDigestValue(
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
      this.postlink(
        fileUrl,
        filenameWithoutExt,
        filenameWithExt,
        docFileLink,
        contextPageAbsoluteUrl,
        FormDigestValue
      );
    } else {
      window.localStorage.removeItem(TempStorageKeys.LContType);
      console.log('Failed');
    }
  }
  //
  public async postlink(
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
      window.localStorage.removeItem(TempStorageKeys.LContType);
      console.log('Item Inserted..!!');
      console.log(await resp.json());
    } else {
      window.localStorage.removeItem(TempStorageKeys.LContType);
      console.log('API Error');
      console.log(await resp.json());
    }
  }

  public render(): React.ReactElement {
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
            Sign In
          </span>
        </div>
        <p
          id='remve'
          style={{ textAlign: 'center', fontWeight: '600', fontSize: '20px' }}
        >
          Welcome to Laserfiche
        </p>
        {/* <p id="remvefile" style={{ "textAlign": "center", "fontWeight": "600", "fontSize": "18px","display":"none" }}>Please wait while we prepare your file to upload....</p> */}
        <div style={{ textAlign: 'center' }}>
          <lf-login
            redirect_uri={
              this.props.context.pageContext.web.absoluteUrl +
              this.props.laserficheRedirectPage
            }
            authorize_url_host_name={this.state.region}
            redirect_behavior='Replace'
            client_id={clientId}
            sign_in_text='Sign in'
            ref={this.loginComponent}
          />
        </div>
        <div>
          <div
            /* className="lf-component-container lf-right-button" */ style={{
              marginTop: '35px',
              textAlign: 'center',
            }}
          >
            <button style={{ fontWeight: '600' }} onClick={this.Redirect}>
              Go Back
            </button>
          </div>
        </div>
      </div>
    );
  }
}
