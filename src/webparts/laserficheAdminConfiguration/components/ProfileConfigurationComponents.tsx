import { NgElement, WithProperties } from '@angular/elements';
import {
  EntryType,
  TemplateFieldInfo,
} from '@laserfiche/lf-repository-api-client';
import {
  LfRepoTreeNode,
  LfRepoTreeNodeService,
} from '@laserfiche/lf-ui-components-services';
import { LfRepositoryBrowserComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { useState } from 'react';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import {
  SPFieldData,
  MappedFields,
  FieldMappingError,
  ProfileConfiguration,
} from './EditManageConfiguration/IEditManageConfigurationState';

export function ProfileHeader(props: { configurationName: string }) {
  return (
    <h6 className='mb-0'>
      Profile :{' '}
      <span className='h5' id='configurationTitle'>
        {props.configurationName}
      </span>
    </h6>
  );
}

export function ConfigurationBody(props: {
  laserficheTemplate: JSX.Element[];
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
  profileConfig: ProfileConfiguration;
  handleProfileConfigUpdate: (config: ProfileConfiguration) => void;
  handleTemplateChange: (templateName: string) => void;
}) {
  const [showFolderModal, setShowFolderModal] = useState(false);
  const selectedEntryNodePath = props.profileConfig?.DestinationPath;

  const onSelectFolder = async (selectedNode: LfRepoTreeNode | undefined) => {
    if (!props.repoClient) {
      throw new Error('Repo Client is undefined.');
    }
    const config = { ...props.profileConfig };
    config.DestinationPath = selectedNode.path;
    config.EntryId = selectedNode.id;
    props.handleProfileConfigUpdate(config);
    setShowFolderModal(false);
  };

  const handleTemplateChange = (event) => {
    const value = (event.target as HTMLSelectElement).value;
    const templatename = value;
    const profileConfig = { ...props.profileConfig };
    profileConfig.DocumentTemplate = templatename;
    props.handleProfileConfigUpdate(profileConfig);
    props.handleTemplateChange(templatename);
  };

  function CloseFolderModalUp() {
    setShowFolderModal(false);
  }
  
  async function OpenFoldersModal() {
    setShowFolderModal(true);
  }
  return (
    <>
      <div className='form-group row'>
        <DocumentName
          documentName={props.profileConfig?.DocumentName}
        ></DocumentName>
      </div>
      <div className='form-group row'>
        <TemplateSelector
          laserficheTemplate={props.laserficheTemplate}
          selectedTemplate={props.profileConfig?.DocumentTemplate}
          repoClient={props.repoClient}
          onChangeTemplate={handleTemplateChange}
        ></TemplateSelector>
      </div>
      <div className='form-group row'>
        <label htmlFor='txt3' className='col-sm-2 col-form-label'>
          Laserfiche Destination
        </label>
        <div className='col-sm-6'>
          <input
            type='text'
            className='form-control'
            id='destinationPath'
            placeholder='(Path in Laserfiche) Example: \folder\subfolder'
            disabled
            value={props.profileConfig?.DestinationPath}
          />
          <div>
            <span>Use the Browse button to select a path</span>
          </div>
        </div>
        <div className='col-sm-2' id='folderModal' style={{ marginTop: '2px' }}>
          <a
            href='javascript:;'
            className='btn btn-primary btn-sm'
            data-toggle='modal'
            data-target='#folderModal'
            onClick={OpenFoldersModal}
          >
            Browse
          </a>
        </div>
      </div>
      <div className='form-group row'>
        <label htmlFor='dwl4' className='col-sm-2 col-form-label'>
          After import
        </label>
        <div className='col-sm-6'>
          <select className='custom-select' id='action'>
            <option value={'Copy'}>
              Leave a copy of the file in SharePoint
            </option>
            <option value={'Replace'}>
              Replace SharePoint file with a link to the document in Laserfiche
            </option>
            <option value={'Move and Delete'}>Delete SharePoint file</option>
          </select>
        </div>
        <div className='col-sm-2'>
          {/* <div className="custom-control custom-checkbox mt-2" style={{ paddingLeft: "3px !important", "marginLeft": "-23px" }}>
                          <a data-toggle="tooltip" style={{ "color": "#0062cc" }}><span className="fa fa-question-circle fa-2"></span></a>
                        </div> */}
        </div>
      </div>
      <div
        className='modal'
        id='folderModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showFolderModal}
      >
        {showFolderModal && (
          <RepositoryBrowserModal
            repoClient={props.repoClient}
            CloseFolderBrowserUp={CloseFolderModalUp}
            selectedEntryNodePath={selectedEntryNodePath}
            SelectFolder={onSelectFolder}
          ></RepositoryBrowserModal>
        )}
      </div>
    </>
  );
}

export function RepositoryBrowserModal(props: {
  CloseFolderBrowserUp: () => void;
  SelectFolder: (node: LfRepoTreeNode | undefined) => void;
  selectedEntryNodePath: string;
  repoClient: IRepositoryApiClientExInternal;
}) {
  const [shouldShowOpen, setShouldShowOpen] = useState(false);
  const [shouldShowSelect, setShouldShowSelect] = useState(false);
  const [shouldDisableSelect, setShouldDisableSelect] = useState(false);

  const [entrySelected, setEntrySelected] = useState<
    LfRepoTreeNode | undefined
  >(undefined);
  const onEntrySelected = (event: CustomEvent<LfRepoTreeNode[]>) => {
    const treeNodesSelected: LfRepoTreeNode[] = event.detail;
    const selectedNode =
      treeNodesSelected?.length > 0 ? treeNodesSelected[0] : undefined;
    setEntrySelected(selectedNode);
    setShouldShowOpen(selectedNode && selectedNode.isContainer);
    setShouldShowSelect(
      !selectedNode && !!repositoryBrowser?.current?.currentFolder
    );
    setShouldDisableSelect(getShouldDisableSelect());
  };
  const repositoryBrowser: React.RefObject<
    NgElement & WithProperties<LfRepositoryBrowserComponent>
  > = React.useRef();
  let lfRepoTreeService;
  React.useEffect(() => {
    if (props.repoClient) {
      lfRepoTreeService = new LfRepoTreeNodeService(props.repoClient);
      lfRepoTreeService.viewableEntryTypes = [
        EntryType.Folder,
        EntryType.Shortcut,
      ];
      initializeTreeAsync();
    }
  }, [props.repoClient]);
  const isNodeSelectable = (node: LfRepoTreeNode) => {
    if (node?.entryType == EntryType.Folder) {
      return true;
    } else if (
      node?.entryType == EntryType.Shortcut &&
      node?.targetType == EntryType.Folder
    ) {
      return true;
    } else {
      return false;
    }
  };

  async function initializeTreeAsync() {
    if (!props.repoClient) {
      throw new Error('RepoId is undefined');
    }
    repositoryBrowser.current?.addEventListener(
      'entrySelected',
      onEntrySelected
    );
    let focusedNode: LfRepoTreeNode | undefined;
    if (props.selectedEntryNodePath) {
      const repoId = await props.repoClient.getCurrentRepoId();
      const focusedNodeByPath =
        await props.repoClient.entriesClient.getEntryByPath({
          repoId: repoId,
          fullPath: props.selectedEntryNodePath,
        });
      const repoName = await props.repoClient.getCurrentRepoName();
      const focusedNodeEntry = focusedNodeByPath?.entry;
      if (focusedNodeEntry) {
        focusedNode = lfRepoTreeService?.createLfRepoTreeNode(
          focusedNodeEntry,
          repoName
        );
      }
    }
    if (lfRepoTreeService) {
      await repositoryBrowser?.current?.initAsync(
        lfRepoTreeService,
        focusedNode
      );
    } else {
      console.debug(
        'Unable to initialize tree, lfRepoTreeService is undefined'
      );
    }
  }
  function getShouldShowSelect(): boolean {
    const showSelect =
      !entrySelected && !!repositoryBrowser?.current?.currentFolder;
    return showSelect;
  }

  function getShouldShowOpen(): boolean {
    return !!entrySelected;
  }

  function getShouldDisableSelect(): boolean {
    return !isNodeSelectable(
      repositoryBrowser?.current?.currentFolder as LfRepoTreeNode
    );
  }

  const onSelectFolder = () => {
    props.SelectFolder(
      repositoryBrowser?.current?.currentFolder as LfRepoTreeNode
    );
  };

  const onOpenNode = async () => {
    await repositoryBrowser?.current?.openSelectedNodesAsync();
    setShouldShowOpen(getShouldShowOpen());
    setShouldShowSelect(getShouldShowSelect());
  };

  return (
    <>
      <div className='modal-dialog modal-dialog-centered'>
        <div className='modal-content'>
          <div className='modal-header'>
            <h5 className='modal-title' id='ModalLabel'>
              Select folder for saving to Laserfiche
            </h5>
            <button
              type='button'
              className='close'
              data-dismiss='modal'
              aria-label='Close'
              onClick={props.CloseFolderBrowserUp}
            >
              <span aria-hidden='true'>&times;</span>
            </button>
          </div>
          <div className='modal-body'>
            <div
              className='lf-folder-browser-sample-container'
              style={{ height: '400px' }}
            >
              {/* <lf-folder-browser ref={this.folderbrowser} ok_button_text="Okay" cancel_button_text="Cancel"></lf-folder-browser> */}
              <div className='repository-browser'>
                <lf-repository-browser
                  ref={repositoryBrowser}
                  ok_button_text='Okay'
                  cancel_button_text='Cancel'
                  multiple='false'
                  style={{ height: '420px' }}
                  isSelectable={isNodeSelectable}
                />
                <div className='repository-browser-button-containers'>
                  <span>
                    <button
                      className='lf-button primary-button'
                      onClick={onOpenNode}
                      hidden={!shouldShowOpen}
                    >
                      OPEN
                    </button>
                    <button
                      className='lf-button primary-button'
                      onClick={onSelectFolder}
                      hidden={!shouldShowSelect}
                      disabled={shouldDisableSelect}
                    >
                      Select
                    </button>
                    <button
                      className='sec-button lf-button margin-left-button'
                      onClick={props.CloseFolderBrowserUp}
                    >
                      CANCEL
                    </button>
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

export function DocumentName(props: { documentName: string }) {
  return (
    <>
      <label htmlFor='txt1' className='col-sm-2 col-form-label'>
        Document Name
      </label>
      <div className='col-sm-6'>
        <input
          type='text'
          className='form-control'
          id='documentName'
          placeholder='Document Name'
          disabled
          value={props.documentName}
        />
      </div>
    </>
  );
}

export function TemplateSelector(props: {
  laserficheTemplate: JSX.Element[];
  selectedTemplate: any;
  repoClient: IRepositoryApiClientExInternal;
  onChangeTemplate: (event: any) => void;
}) {
  return (
    <>
      <label htmlFor='dwl2' className='col-sm-2 col-form-label'>
        Laserfiche Template
      </label>
      <div className='col-sm-6'>
        <select
          className='custom-select'
          id='documentTemplate'
          onChange={(e) => props.onChangeTemplate(e)}
          value={props.selectedTemplate}
        >
          <option>None</option>
          {props.laserficheTemplate}
        </select>
      </div>
    </>
  );
}
export function SharePointLaserficheColumnMatching(props: {
  sharePointFields: SPFieldData[];
  laserficheFields: TemplateFieldInfo[];
  mappingList: MappedFields[];
  handleChange: (e: any) => void;
  RemoveSpecificMapping: (index: number) => void;
  AddNewMappingFields: () => void;
  ColumnMatchingError: FieldMappingError | undefined;
}) {
  const displayError = getErrorMessage(props.ColumnMatchingError);

  function getErrorMessage(error: FieldMappingError): JSX.Element | undefined {
    let errorMessage: string | undefined;
    switch (error) {
      case FieldMappingError.CONTENT_TYPE:
        errorMessage =
          'Cannot save configuration. Please select a Content Type';
        break;
      case FieldMappingError.SELECT_TEMPLATE:
        errorMessage =
          'Please select any template from Laserfiche Template to add new mapping';
        break;
    }
    if (errorMessage) {
      return (
        <div style={{ color: 'red' }}>
          <span>{errorMessage}</span>
        </div>
      );
    } else {
      return undefined;
    }
  }

  const spFields = props.sharePointFields.slice()?.map((field) => (
    <option value={field.InternalName}>
      {field.Title} ({field.TypeAsString})
    </option>
  ));
  const laserficheFields = props.laserficheFields.map((items) => {
    return (
      <option value={items.id}>
        {items.name} ({items.fieldType})
      </option>
    );
  });
  // TODO also filter for ones that are already mapped
  const optionalFields = props.laserficheFields.filter((field) => !field.isRequired).map((items) => {
    return (
      <option value={items.id}>
        {items.name} ({items.fieldType})
      </option>
    );
  });
  const handleFieldChange = (e) => {
    props.handleChange(e);
  };
  const mappedList = props.mappingList.map((fieldMapping, index) => {
    const errorMessageMapping: JSX.Element | undefined =
      getMappingErrorMessage(fieldMapping);
    return (
      <tr id={index.toString()} key={index}>
        <td>
          <select
            name='SharePointField'
            className='custom-select'
            value={fieldMapping.spField?.InternalName ?? 'Select'}
            id={fieldMapping.id}
            onChange={(e) => handleFieldChange(e)}
          >
            <option>Select</option>
            {spFields}
          </select>
        </td>
        <td>
          <select
            name='LaserficheField'
            className='custom-select'
            value={fieldMapping.lfField?.id ?? 'Select'}
            id={fieldMapping.id}
            disabled={fieldMapping.lfField?.isRequired}
            onChange={(e) => handleFieldChange(e)}
          >
            <option>Select</option>
            {fieldMapping.lfField?.isRequired ? laserficheFields : optionalFields}
          </select>
        </td>
        <td>
          {fieldMapping.lfField?.isRequired ? (
            <span style={{ fontSize: '13px', color: 'red' }}>
              *Required field in Laserfiche
            </span>
          ) : (
            <a
              href='javascript:;'
              className='ml-3'
              onClick={() => props.RemoveSpecificMapping(index)}
            >
              <span className='material-icons'>delete</span>
            </a>
          )}
          {errorMessageMapping}
        </td>
      </tr>
    );
  });

  return (
    <>
      <table className='table table-sm'>
        <thead>
          <tr>
            <th className='text-center' style={{ width: '39%' }}>
              SharePoint Column
            </th>
            <th className='text-center' style={{ width: '38%' }}>
              Laserfiche Field
            </th>
          </tr>
        </thead>
        <tbody id='tableEditBodyId'>{mappedList}</tbody>
      </table>
      {displayError}
      <a
        onClick={props.AddNewMappingFields}
        className='btn btn-primary pl-5 pr-5 float-right ml-2'
      >
        Add Field
      </a>
    </>
  );
}

export function DeleteModal(props: {
  configurationName: string;
  onConfirmDelete: () => void;
  onCancel: () => void;
}) {
  return (
    <div className='modal-dialog modal-dialog-centered'>
      <div className='modal-content'>
        <div className='modal-header'>
          <h5 className='modal-title' id='ModalLabel'>
            Delete Confirmation
          </h5>
          <button
            type='button'
            className='close'
            data-dismiss='modal'
            aria-label='Close'
            onClick={props.onCancel}
          >
            <span aria-hidden='true'>&times;</span>
          </button>
        </div>
        <div className='modal-body'>
          Do you want to permanently delete &quot;
          {props.configurationName}&quot;?
        </div>
        <div className='modal-footer'>
          <button
            type='button'
            className='btn btn-primary btn-sm'
            data-dismiss='modal'
            onClick={props.onConfirmDelete}
          >
            OK
          </button>
          <button
            type='button'
            className='btn btn-secondary btn-sm'
            data-dismiss='modal'
            onClick={props.onCancel}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}
function getMappingErrorMessage(
  mappedField: MappedFields
): JSX.Element | undefined {
  let hasMismatch: boolean = false;
  if (mappedField.lfField && mappedField.spField) {
    const spFieldtype = mappedField.spField.TypeAsString;
    const lfFieldtype = mappedField.lfField.fieldType;
    if (
      lfFieldtype == 'DateTime' ||
      lfFieldtype == 'Date' ||
      lfFieldtype == 'Time'
    ) {
      if (spFieldtype != 'DateTime') {
        hasMismatch = true;
      }
    } else if (lfFieldtype == 'LongInteger' || lfFieldtype == 'ShortInteger') {
      if (spFieldtype != 'Number') {
        hasMismatch = true;
      }
    } else if (lfFieldtype == 'Number') {
      if (spFieldtype != 'Number' && spFieldtype != 'Currency') {
        hasMismatch = true;
      }
    } else if (lfFieldtype == 'List') {
      if (spFieldtype != 'Choice') {
        hasMismatch = true;
      }
    }

    if (hasMismatch) {
      return (
        <span
          style={{
            display: 'none',
            color: 'red',
            fontSize: '13px',
            marginLeft: '10px',
          }}
          title={`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`}
        >
          SharePoint field type of ${spFieldtype} cannot be mapped with
          Laserfiche field type of ${lfFieldtype}
          <span className='material-icons'>warning</span>Data types mismatch
        </span>
      );
    }
    else {
      return undefined;
    }
  }
  return undefined;
}
