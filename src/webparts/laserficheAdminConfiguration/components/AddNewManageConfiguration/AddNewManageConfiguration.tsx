import * as React from 'react';
import * as bootstrap from 'bootstrap';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { NavLink } from 'react-router-dom';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListItem } from './IListItem';
import { IAddNewManageConfigurationProps } from './IAddNewManageConfigurationProps';
import { IAddNewManageConfigurationState } from './IAddNewManageConfigurationState';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import {
  ODataValueContextOfIListOfWTemplateInfo,
  ODataValueOfIListOfTemplateFieldInfo,
  TemplateFieldInfo,
  WTemplateInfo,
} from '@laserfiche/lf-repository-api-client';
import {
  FieldMappingError,
  ProfileConfiguration,
  SPFieldData,
} from '../EditManageConfiguration/IEditManageConfigurationState';
import {
  ConfigurationBody,
  DeleteModal,
  SharePointLaserficheColumnMatching,
} from '../ProfileConfigurationComponents';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ['lf-repository-browser']: any;
    }
  }
}

export default class AddNewManageConfiguration extends React.Component<
  IAddNewManageConfigurationProps,
  IAddNewManageConfigurationState
> {
  configNameValidation: JSX.Element | undefined;
  constructor(props: IAddNewManageConfigurationProps) {
    super(props);
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css'
    );
    this.state = {
      mappingList: [],
      laserficheTemplates: [],
      sharePointFields: [],
      laserficheFields: [],
      documentNames: [],
      loadingContent: false,
      hideContent: true,
      showFolderModal: false,
      showtokensModal: false,
      deleteModal: undefined,
      showConfirmModal: false,
      columnError: undefined,
      profileConfig: undefined,
    };
  }

  public async componentDidMount(): Promise<void> {
    this.setState({ hideContent: true });
    this.setState({ loadingContent: false });

    await SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
    );
    await SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ui-components.js'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css'
    );

    const profileConfig: ProfileConfiguration = {
      DestinationPath: '\\',
      DocumentName: 'FileName',
      EntryId: '1',
      ConfigurationName: '',
      DocumentTemplate: 'None',
      SharePointFields: [],
      LaserficheFields: [],
      Action: '',
    };
    this.setState(() => {
      return { profileConfig: profileConfig };
    });
    this.configNameValidation = undefined;

    this.GetAllSharePointSiteColumns().then((contents: any) => {
      contents.sort((a, b) => (a.DisplayName > b.DisplayName ? 1 : -1));
      this.setState({
        sharePointFields: contents,
      });
      this.GetTemplateDefinitions().then((templates: string[]) => {
        templates.sort();
        this.setState({ laserficheTemplates: templates });
        this.setState({ loadingContent: true });
        this.setState({ hideContent: false });
      });
    });
  }

  public async GetTemplateDefinitions(): Promise<string[]> {
    let array = [];

    const repoId = await this.props.repoClient.getCurrentRepoId();
    const templateInfo: WTemplateInfo[] = [];
    await this.props.repoClient.templateDefinitionsClient.getTemplateDefinitionsForEach(
      {
        callback: async (response: ODataValueContextOfIListOfWTemplateInfo) => {
          if (response.value) {
            templateInfo.push(...response.value);
          }
          return true;
        },
        repoId,
      }
    );
    array = templateInfo.map((value) => value.name);
    return array;
  }

  public async GetAllSharePointSiteColumns(): Promise<SPFieldData[]> {
    const restApiUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/fields?$filter=(Hidden ne true and Group ne '_Hidden')";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = (await res.json()).value as SPFieldData[];

      return results;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  OnChangeTemplate = (templateName: string) => {
    this.setState(() => {
      return { columnError: undefined };
    });
    this.GetLaserficheFields(templateName).then(
      (fields: TemplateFieldInfo[]) => {
        if (fields != null) {
          this.setState({ laserficheFields: fields });
          const array = [];
          for (let index = 0; index < fields.length; index++) {
            const id = (
              +new Date() + Math.floor(Math.random() * 999999)
            ).toString(36);
            const laserficheField = fields[index];
            if (laserficheField.isRequired) {
              array.push({
                id: id,
                spField: undefined,
                lfField: fields[index],
              });
            }
          }
          this.setState({ mappingList: array });
        } else {
          this.setState({ laserficheFields: [], mappingList: undefined });
        }
      }
    );
  };

  GetLaserficheFields: (templateName: string) => Promise<TemplateFieldInfo[]> =
    async (templateName: string) => {
      if (templateName != 'None') {
        const repoId = await this.props.repoClient.getCurrentRepoId();
        const apiTemplateResponse: ODataValueOfIListOfTemplateFieldInfo =
          await this.props.repoClient.templateDefinitionsClient.getTemplateFieldDefinitionsByTemplateName(
            { repoId, templateName: templateName }
          );
        const fieldsValues: TemplateFieldInfo[] = apiTemplateResponse.value;
        return fieldsValues;
      } else {
        return null;
      }
    };
  handleProfileConfigNameChange(e: any) {
    const newName = (e.target as HTMLInputElement).value;
    const profileConfig = { ...this.state.profileConfig };
    profileConfig.ConfigurationName = newName;
    this.setState(() => {
      return { profileConfig: profileConfig };
    });
  }

  SaveNewManageConfigurtaion = () => {
    this.setState(() => {
      return { columnError: undefined };
    });
    this.configNameValidation = undefined;
    const configName = this.state.profileConfig.ConfigurationName;
    let validation = true;
    if (configName == '') {
      validation = false;
      this.configNameValidation = <span>Required Field</span>;
    } else if (/[^A-Za-z0-9]/.test(configName)) {
      validation = false;
      this.configNameValidation = (
        <span>Invalid Name, only alphanumeric are allowed without space.</span>
      );
    }
    if (validation) {
      const rows = [...this.state.mappingList];
      if (
        rows.some((item) => !item.spField) &&
        this.state.profileConfig.DocumentTemplate != 'None'
      ) {
        this.setState(() => {
          return { columnError: FieldMappingError.CONTENT_TYPE };
        });
      } else if (
        rows.some((items) => !items.lfField) &&
        this.state.profileConfig.DocumentTemplate != 'None'
      ) {
        this.setState(() => {
          return { columnError: FieldMappingError.CONTENT_TYPE };
        });
      } else {
        if (validation) {
          this.setState(() => {
            return { columnError: undefined };
          });
          const configName = this.state.profileConfig.ConfigurationName;
          const documentName = this.state.profileConfig.DocumentName;
          const docTemp = this.state.profileConfig.DocumentTemplate;
          const destPath = this.state.profileConfig.DestinationPath;
          const entryId = this.state.profileConfig.EntryId;
          const action = document.getElementById('action')['value'];
          const sharepointFields = [];
          const laserficheFields = [];
          if (docTemp != 'None') {
            for (let i = 0; i < rows.length; i++) {
              sharepointFields.push(rows[i].spField);
              laserficheFields.push(rows[i].lfField);
            }
          }

          const jsonData = [
            {
              ConfigurationName: configName,
              DocumentName: documentName,
              DocumentTemplate: docTemp,
              DestinationPath: destPath,
              EntryId: entryId,
              Action: action,
              SharePointFields: sharepointFields,
              LaserficheFields: laserficheFields,
            },
          ];
          this.GetItemIdByTitle().then((results: IListItem[]) => {
            if (results != null) {
              const itemId = results[0].Id;
              const jsonValue: ProfileConfiguration[] = JSON.parse(
                results[0].JsonValue
              );
              if (jsonValue.length > 0) {
                let entryExists = false;
                for (let i = 0; i < jsonValue.length; i++) {
                  if (
                    jsonValue[i].ConfigurationName ==
                    this.state.profileConfig?.ConfigurationName
                  ) {
                    this.configNameValidation = (
                      <span>
                        Profile with this name already exists, please provide
                        different name
                      </span>
                    );
                    entryExists = true;
                    break;
                  }
                }
                if (entryExists == false) {
                  const restApiUrl: string =
                    this.props.context.pageContext.web.absoluteUrl +
                    "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
                    itemId +
                    ')';
                  const newJsonValue = jsonValue.concat({
                    ConfigurationName: configName,
                    DocumentName: documentName,
                    DocumentTemplate: docTemp,
                    DestinationPath: destPath,
                    EntryId: entryId,
                    Action: action,
                    SharePointFields: sharepointFields,
                    LaserficheFields: laserficheFields,
                  });
                  const jsonObject = JSON.stringify(newJsonValue);
                  const body: string = JSON.stringify({
                    Title: 'ManageConfigurations',
                    JsonValue: jsonObject,
                  });
                  const options: ISPHttpClientOptions = {
                    headers: {
                      Accept: 'application/json;odata=nometadata',
                      'content-type': 'application/json;odata=nometadata',
                      'odata-version': '',
                      'IF-MATCH': '*',
                      'X-HTTP-Method': 'MERGE',
                    },
                    body: body,
                  };
                  this.props.context.spHttpClient
                    .post(restApiUrl, SPHttpClient.configurations.v1, options)
                    .then((): void => {
                      this.setState(() => {
                        return { showConfirmModal: true };
                      });
                    });
                }
              } else {
                const jsonObj = JSON.stringify(jsonData);
                const restApiUrl: string =
                  this.props.context.pageContext.web.absoluteUrl +
                  "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
                  itemId +
                  ')';
                const body: string = JSON.stringify({
                  Title: 'ManageConfigurations',
                  JsonValue: jsonObj,
                });
                const options: ISPHttpClientOptions = {
                  headers: {
                    Accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;odata=nometadata',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                  },
                  body: body,
                };
                this.props.context.spHttpClient
                  .post(restApiUrl, SPHttpClient.configurations.v1, options)
                  .then((): void => {
                    this.setState(() => {
                      return { showConfirmModal: true };
                    });
                  });
              }
            } else {
              this.SaveNewConfiguration(jsonData);
            }
          });
        }
      }
    }
  }

  public SaveNewConfiguration(jsonObject) {
    const jsonData = JSON.stringify(jsonObject);
    const restApiUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/items";
    const body: string = JSON.stringify({
      Title: 'ManageConfigurations',
      JsonValue: jsonData,
    });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'odata-version': '',
      },
      body: body,
    };
    this.props.context.spHttpClient
      .post(restApiUrl, SPHttpClient.configurations.v1, options)
      .then((): void => {
        this.setState(() => {
          return { showConfirmModal: true };
        });
      });
  }

  public async GetItemIdByTitle(): Promise<IListItem[]> {
    const array: IListItem[] = [];
    const restApiUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/Items?$select=Id,Title,JsonValue&$filter=Title eq 'ManageConfigurations'";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
        },
      });
      const results = await res.json();
      if (results.value.length > 0) {
        return results.value as IListItem[];
      } else {
        return null;
      }
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  AddNewMappingFields = () => {
    const templatename = this.state.profileConfig.DocumentTemplate;
    if (templatename != 'None') {
      const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(
        36
      );
      const item = {
        id: id,
        spField: undefined,
        lfField: undefined,
      };
      this.setState({
        mappingList: [...this.state.mappingList, item],
      });
    } else {
      this.setState(() => {
        return { columnError: FieldMappingError.SELECT_TEMPLATE };
      });
    }
  };

  public RemoveSpecificMapping = (idx) => {
    this.setState(() => {
      return { columnError: undefined };
    });
    const del = (
      <DeleteModal
        configurationName='the field mapping'
        onCancel={this.CloseModalUp}
        onConfirmDelete={() => this.DeleteMapping(idx)}
      ></DeleteModal>
    );
    this.setState(() => {
      return { deleteModal: del };
    });
  };
  public DeleteMapping(id: number) {
    const rows = [...this.state.mappingList];
    rows.splice(id, 1);
    this.setState({ mappingList: rows });
    this.setState(() => {
      return { deleteModal: undefined };
    });
  }

  public handleChange = (e) => {
    const targetElement = e.target as HTMLSelectElement;
    const item = {
      id: targetElement.id,
      name: targetElement.name,
      value: targetElement.value,
    };
    const rowsArray = [...this.state.mappingList];
    const currentRow = rowsArray.findIndex((row) => {
      return item.id == row.id;
    });
    if (item.name === 'LaserficheField') {
      const lfField = this.state.laserficheFields.find(
        (data) => data.id.toString() === item.value
      );
      if (lfField) {
        rowsArray[currentRow].lfField = lfField;
      }
    } else if (item.name === 'SharePointField') {
      const spField = this.state.sharePointFields.find(
        (data) => data.InternalName === item.value
      );
      if (spField) {
        rowsArray[currentRow].spField = spField;
      }
    }
    this.setState({ mappingList: rowsArray });
  };

  public CloseModalUp() {
    this.setState(() => {
      return { deleteModal: undefined };
    });
  }

  public SelectDocumentToken() {
    this.setState(() => {
      return { showtokensModal: true };
    });
  }

  public CloseFolderModalUp() {
    this.setState(() => {
      return { showFolderModal: false };
    });
  }

  public CloseTokenModalUp() {
    this.setState(() => {
      return { showtokensModal: false };
    });
  }

  public ConfirmButton() {
    history.back();
    this.setState(() => {
      return { showConfirmModal: false };
    });
  }

  handleProfileConfigUpdate = (profileConfig: ProfileConfiguration) => {
    this.setState(() => {
      return { profileConfig: profileConfig };
    });
  };

  public render(): React.ReactElement {
    const laserficheTemplate = this.state.laserficheTemplates.map((item) => (
      <option value={item}>{item}</option>
    ));
    const header = (
      <div>
        <h6 className='mb-0'>Add New Profile</h6>
      </div>
    );
    const extraConfiguration = (
      <>
        <div className='form-group row'>
          <label
            htmlFor='txt0'
            className='col-sm-2 col-form-label'
            style={{ width: '165px' }}
          >
            Profile Name <span style={{ color: 'red' }}>*</span>
          </label>
          <div className='col-sm-6'>
            <input
              type='text'
              className='form-control'
              id='configurationName'
              onChange={(e) => this.handleProfileConfigNameChange(e)}
              placeholder='Profile Name'
            />
            <div
              id='configurationExists'
              hidden={!this.configNameValidation}
              style={{ color: 'red' }}
            >
              {this.configNameValidation}
            </div>
          </div>
        </div>
      </>
    );
    return (
      <div>
        <div
          className='container-fluid p-3'
          style={{ maxWidth: '85%', marginLeft: '-26px' }}
        >
          <main className='bg-white shadow-sm'>
            <div
              className='addPageSpinloader'
              hidden={this.state.loadingContent}
            >
              {!this.state.loadingContent && (
                <Spinner size={SpinnerSize.large} label='loading' />
              )}
            </div>
            <div className='p-3' hidden={this.state.hideContent}>
              <div className='card rounded-0'>
                <div className='card-header d-flex justify-content-between'>
                  {header}
                </div>
                <div className='card-body'>
                  {extraConfiguration}
                  <ConfigurationBody
                    laserficheTemplate={laserficheTemplate}
                    repoClient={this.props.repoClient}
                    loggedIn={this.props.loggedIn}
                    handleTemplateChange={this.OnChangeTemplate}
                    profileConfig={this.state.profileConfig}
                    handleProfileConfigUpdate={this.handleProfileConfigUpdate}
                  ></ConfigurationBody>
                </div>
                <h6 className='card-header border-top'>
                  Mappings from SharePoint Column to Laserfiche Field Values
                </h6>
                <div className='card-body'>
                  <SharePointLaserficheColumnMatching
                    sharePointFields={this.state.sharePointFields}
                    laserficheFields={this.state.laserficheFields}
                    mappingList={this.state.mappingList}
                    handleChange={(e) => this.handleChange(e)}
                    RemoveSpecificMapping={this.RemoveSpecificMapping}
                    AddNewMappingFields={this.AddNewMappingFields}
                    ColumnMatchingError={this.state.columnError}
                  ></SharePointLaserficheColumnMatching>
                </div>
                <div className='card-footer bg-transparent'>
                  {this.props.loggedIn && (
                    <NavLink id='navid' to='/ManageConfigurationsPage'>
                      <a className='btn btn-primary pl-5 pr-5 float-right ml-2'>
                        Back
                      </a>
                    </NavLink>
                  )}
                  <a
                    href='javascript:;'
                    className='btn btn-primary pl-5 pr-5 float-right ml-2'
                    onClick={() => this.SaveNewManageConfigurtaion()}
                  >
                    Save
                  </a>
                </div>
              </div>
            </div>
          </main>
        </div>
        <div
          className='modal'
          id='deleteModal'
          hidden={!this.state.deleteModal}
          data-backdrop='static'
          data-keyboard='false'
        >
          {this.state.deleteModal}
        </div>
        <div
          className='modal'
          data-backdrop='static'
          data-keyboard='false'
          id='ConfirmModal'
          hidden={!this.state.showConfirmModal}
        >
          <div className='modal-dialog modal-dialog-centered'>
            <div className='modal-content'>
              <div className='modal-body'>Profile Added</div>
              <div className='modal-footer'>
                <button
                  type='button'
                  className='btn btn-primary btn-sm'
                  data-dismiss='modal'
                  onClick={() => this.ConfirmButton()}
                >
                  OK
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
