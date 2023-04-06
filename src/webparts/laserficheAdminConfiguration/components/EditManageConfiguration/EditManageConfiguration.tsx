import * as React from 'react';
import * as bootstrap from 'bootstrap';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { NavLink } from 'react-router-dom';
import { IEditManageConfigurationProps } from './IEditManageConfigurationProps';
import {
  FieldMappingError,
  IEditManageConfigurationState,
  MappedFields,
  ProfileConfiguration,
  SPFieldData,
} from './IEditManageConfigurationState';
import { IListItem } from './IListItem';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import {
  ODataValueContextOfIListOfWTemplateInfo,
  ODataValueOfIListOfTemplateFieldInfo,
  WTemplateInfo,
  EntryType,
  TemplateFieldInfo,
} from '@laserfiche/lf-repository-api-client';
import {
  LfRepoTreeNode,
  LfRepoTreeNodeService,
} from '@laserfiche/lf-ui-components-services';
import { LfRepositoryBrowserComponent } from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';
import { NgElement, WithProperties } from '@angular/elements';
import { useState } from 'react';
import { clientId } from '../../../constants';
import {
  DeleteModal,
  ProfileHeader,
  ConfigurationBody,
  SharePointLaserficheColumnMatching,
} from '../ProfileConfigurationComponents';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ['lf-login']: any;
      ['lf-repository-browser']: any;
    }
  }
}

export default class EditManageConfiguration extends React.Component<
  IEditManageConfigurationProps,
  IEditManageConfigurationState
> {
  constructor(props: IEditManageConfigurationProps) {
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
    const configurationName = this.props.match.params.name;
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

    this.setState(() => {
      return {
        showFolderModal: false,
        showtokensModal: false,
        showConfirmModal: false,
      };
    });

    this.GetItemIdByTitle().then((results: IListItem[]) => {
      this.GetTemplateDefinitions().then((templates: string[]) => {
        templates.sort();
        this.setState({ laserficheTemplates: templates });
        if (results != null) {
          const profileConfigs = JSON.parse(results[0].JsonValue);
          if (profileConfigs.length > 0) {
            for (let i = 0; i < profileConfigs.length; i++) {
              if (profileConfigs[i].ConfigurationName == configurationName) {
                const selectedConfig: ProfileConfiguration = profileConfigs[i];
                this.setState(() => {
                  return { profileConfig: selectedConfig };
                });
                this.MappingFields(
                  selectedConfig.DocumentTemplate,
                  selectedConfig.SharePointFields,
                  selectedConfig.LaserficheFields
                );
              }
            }
          }
        }
      });
    });

    this.GetAllSharePointSiteColumns().then((contents: SPFieldData[]) => {
      contents.sort((a, b) => (a.Title > b.Title ? 1 : -1));
      this.setState({
        sharePointFields: contents,
      });
    });
  }

  public async GetAllSharePointSiteColumns(): Promise<SPFieldData[]> {
    const restApiUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/fields?$filter=(Hidden ne true and Group ne '_Hidden')";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
      });
      const results = (await res.json()).value as SPFieldData[];

      return results;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  public async GetTemplateDefinitions(): Promise<string[]> {
    let array = [];
    const templateInfo: WTemplateInfo[] = [];
    const repoId = await this.props.repoClient.getCurrentRepoId();
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

  public async GetItemIdByTitle(): Promise<IListItem[]> {
    const restApiUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/Items?$select=Id,Title,JsonValue&$filter=Title eq 'ManageConfigurations'";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
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

  public async MappingFields(
    DocumentTemplate: string,
    SharePointFields: SPFieldData[],
    LaserficheFields: TemplateFieldInfo[]
  ) {
    if (DocumentTemplate != 'None') {
      const repoId = await this.props.repoClient.getCurrentRepoId();
      const apiTemplateResponse: ODataValueOfIListOfTemplateFieldInfo =
        await this.props.repoClient.templateDefinitionsClient.getTemplateFieldDefinitionsByTemplateName(
          { repoId, templateName: DocumentTemplate }
        );
      const fieldsValuesForSelectedTemplate: TemplateFieldInfo[] =
        apiTemplateResponse.value;
      this.setState({ laserficheFields: fieldsValuesForSelectedTemplate });
      const mappedFieldArray: MappedFields[] = [];
      for (let index = 0; index < SharePointFields.length; index++) {
        const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(
          36
        );
        if (DocumentTemplate != 'None') {
          const laserfciheItems = fieldsValuesForSelectedTemplate;
          const laserficheValue = LaserficheFields[index].name;
          const findMatchingItem = laserfciheItems.find((item) => {
            const fieldName = item.name;
            return fieldName === laserficheValue;
          });
          mappedFieldArray.push({
            id: id,
            spField: SharePointFields[index],
            lfField: findMatchingItem,
          });
        }
      }
      this.setState({ mappingList: mappedFieldArray });
      const rows = [...this.state.mappingList];
      const currentlyMappedFields = [];
      for (const mappedField of rows) {
        if (mappedField.lfField?.id) {
          currentlyMappedFields.push(mappedField.lfField?.id);
        }
      }
      const newArray: MappedFields[] = [];
      for (const lfField of LaserficheFields) {
        const requiredField = lfField.id;
        if (lfField.isRequired) {
          if (currentlyMappedFields.indexOf(requiredField) != -1) {
            // do nothing
          } else {
            const id1 = (
              +new Date() + Math.floor(Math.random() * 999999)
            ).toString(36);
            newArray.push({
              id: id1,
              spField: undefined,
              lfField: lfField,
            });
          }
        }
      }
      this.setState({
        mappingList: this.state.mappingList.concat(newArray),
      });
    } else {
      this.setState({ laserficheFields: [] });
    }
    this.setState({ loadingContent: true });
    this.setState({ hideContent: false });
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

  public CloseModalUp() {
    this.setState(() => {
      return { deleteModal: undefined };
    });
  }

  SaveNewManageConfigurtaion = () => {
    this.setState(() => {
      return { columnError: undefined };
    });
    let validation = true;
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
        this.GetItemIdByTitle().then((results: IListItem[]) => {
          if (results != null) {
            const itemId = results[0].Id;
            const jsonValue = JSON.parse(results[0].JsonValue);
            if (jsonValue.length > 0) {
              for (let i = 0; i < jsonValue.length; i++) {
                if (
                  jsonValue[i].ConfigurationName == this.props.match.params.name
                ) {
                  jsonValue[i].DocumentName = documentName;
                  jsonValue[i].DocumentTemplate = docTemp;
                  jsonValue[i].DestinationPath = destPath;
                  jsonValue[i].EntryId = entryId;
                  jsonValue[i].Action = action;
                  jsonValue[i].SharePointFields = sharepointFields;
                  jsonValue[i].LaserficheFields = laserficheFields;
                  break;
                }
              }
              const restApiUrl: string =
                this.props.context.pageContext.web.absoluteUrl +
                "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
                itemId +
                ')';
              const newJsonValue = [...jsonValue];
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
          }
        });
      }
    }
  }
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

  public ConfirmButton() {
    history.back();
    this.setState(() => {
      return { showConfirmModal: false };
    });
  }

  public SelectDocumentToken() {
    this.setState(() => {
      return { showtokensModal: true };
    });
  }

  public CloseTokenModalUp() {
    this.setState(() => {
      return { showtokensModal: false };
    });
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

  handleProfileConfigUpdate = (profileConfig: ProfileConfiguration) => {
    this.setState(() => {
      return { profileConfig: profileConfig };
    });
  };

  public render(): React.ReactElement {
    const laserficheTemplate = this.state.laserficheTemplates.map((item) => (
      <option value={item}>{item}</option>
    ));
    const header = (<div>
      <ProfileHeader configurationName='Test config name'></ProfileHeader>
    </div>);
    return (
      <div>
        <div style={{ display: 'none' }}>
          <lf-login
            redirect_uri={
              this.props.context.pageContext.web.absoluteUrl +
              this.props.laserficheRedirectPage
            }
            redirect_behavior='Replace'
            authorize_url_host_name='a.clouddev.laserfiche.com'
            client_id={clientId}
          />
        </div>
        <div
          className='container-fluid p-3'
          style={{ maxWidth: '85%', marginLeft: '-26px' }}
        >
          <main className='bg-white shadow-sm'>
            <div
              className='editPageSpinloader'
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
                  <NavLink id='navid' to='/ManageConfigurationsPage'>
                    <a className='btn btn-primary pl-5 pr-5 float-right ml-2'>
                      Back
                    </a>
                  </NavLink>
                  <a
                    href='javascript:;'
                    className='btn btn-primary pl-5 pr-5 float-right ml-2'
                    onClick={this.SaveNewManageConfigurtaion}
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
          id='ConfirmModal'
          hidden={!this.state.showConfirmModal}
          data-backdrop='static'
          data-keyboard='false'
        >
          <div className='modal-dialog modal-dialog-centered'>
            <div className='modal-content'>
              <div className='modal-body'>Profile Updated</div>
              <div className='modal-footer'>
                <button
                  type='button'
                  className='btn btn-primary btn-sm'
                  data-dismiss='modal'
                  onClick={this.ConfirmButton}
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
