import { NgElement, WithProperties } from '@angular/elements';
import {
  EntryType,
  PostEntryChildrenRequest,
  PostEntryChildrenEntryType,
  Entry,
  FieldToUpdate,
  ValueToUpdate,
  PostEntryWithEdocMetadataRequest,
  PutFieldValsRequest,
  FileParameter,
  CreateEntryResult,
} from '@laserfiche/lf-repository-api-client';
import {
  LfRepoTreeNodeService,
  LfFieldsService,
  LfRepoTreeNode,
} from '@laserfiche/lf-ui-components-services';
import {
  ColumnDef,
  LfFieldContainerComponent,
  LfRepositoryBrowserComponent,
} from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';

const cols: ColumnDef[] = [
  {
    id: 'creationTime',
    displayName: 'Creation Date',
    defaultWidth: '100px',
    resizable: true,
  },
  {
    id: 'lastModifiedTime',
    displayName: 'Last Modified Date',
    defaultWidth: '100px',
    resizable: true,
  },
  { id: 'pageCount', displayName: 'Page', defaultWidth: '100px' },
  {
    id: 'templateName',
    displayName: 'Template Name',
    defaultWidth: '100px',
  },
];
export default function RepositoryViewComponent(props: {
  repoClient: IRepositoryApiClientExInternal;
  webPartTitle: string;
  webClientUrl: string;
  loggedIn: boolean;
}) {
  const repositoryBrowser: React.RefObject<
    NgElement & WithProperties<LfRepositoryBrowserComponent>
  > = React.useRef<NgElement & WithProperties<LfRepositoryBrowserComponent>>();
  let lfRepoTreeService: LfRepoTreeNodeService;
  const [parentItem, setParentItem] = React.useState<
    LfRepoTreeNode | undefined
  >(undefined);
  const [selectedItem, setSelectedItem] = React.useState<
    LfRepoTreeNode | undefined
  >(undefined);

  React.useEffect(() => {
    if (props.repoClient) {
      initializeTreeAsync();
    }
    // setwebClientUrl(loginComponent.current.account_endpoints.webClientUrl);
  }, [props.repoClient, props.loggedIn]);

  const onEntrySelected = (
    event: CustomEvent<LfRepoTreeNode[] | undefined>
  ) => {
    const selectedNode = event.detail ? event.detail[0] : undefined;
    setSelectedItem(selectedNode);
  };

  const onEntryOpened = (event: CustomEvent<LfRepoTreeNode[] | undefined>) => {
    const openedNode = event.detail ? event.detail[0] : undefined;
    setParentItem(openedNode);
  };

  const initializeTreeAsync = async () => {
    const repoBrowser = repositoryBrowser.current;
    lfRepoTreeService = new LfRepoTreeNodeService(props.repoClient);
    lfRepoTreeService.viewableEntryTypes = [
      EntryType.Folder,
      EntryType.Shortcut,
      EntryType.Document,
    ];
    repoBrowser?.addEventListener('entrySelected', onEntrySelected);
    repoBrowser?.addEventListener('entryDblClicked', onEntryOpened);
    let focusedNode: LfRepoTreeNode | undefined;
    if (lfRepoTreeService) {
      lfRepoTreeService.columnIds = [
        'creationTime',
        'lastModifiedTime',
        'pageCount',
        'templateName',
      ];
      await repoBrowser?.initAsync(lfRepoTreeService, focusedNode);
      // TODO columns not working right?
      repoBrowser?.setColumnsToDisplay(cols);
      await repoBrowser?.refreshAsync();
    } else {
      console.debug(
        'Unable to initialize tree, lfRepoTreeService is undefined'
      );
    }
  };

  const isNodeSelectable = (node: LfRepoTreeNode) => {
    if (
      node?.entryType == EntryType.Folder ||
      node?.entryType === EntryType.Document
    ) {
      return true;
    } else if (
      (node?.entryType == EntryType.Shortcut &&
        node?.targetType == EntryType.Folder) ||
      (node?.entryType == EntryType.Shortcut &&
        node?.targetType == EntryType.Document)
    ) {
      return true;
    } else {
      return false;
    }
  };
  return (
    <>
      <div>
        <main className='bg-white shadow-sm'>
          <nav className='navbar navbar-dark bg-white flex-md-nowrap'>
            <a className='navbar-brand pl-0' href='#'>
              <img
                src={require('./../../../Assets/Images/laserfiche-logo.png')}
              />{' '}
              {props.webPartTitle}
            </a>
          </nav>
          {props.loggedIn && (
            <>
              <RepositoryBrowserToolbar
                repoClient={props.repoClient}
                selectedItem={selectedItem}
                parentItem={parentItem}
                loggedIn={props.loggedIn}
                webClientUrl={props.webClientUrl}
              ></RepositoryBrowserToolbar>
              <div
                className='lf-folder-browser-sample-container'
                style={{ height: '400px' }}
              >
                <div className='repository-browser'>
                  <lf-repository-browser
                    ref={repositoryBrowser}
                    ok_button_text='Okay'
                    cancel_button_text='Cancel'
                    multiple='false'
                    style={{ height: '420px' }}
                    isSelectable={isNodeSelectable}
                  />
                </div>
              </div>
            </>
          )}
        </main>
      </div>
    </>
  );
}

function RepositoryBrowserToolbar(props: {
  repoClient: IRepositoryApiClientExInternal;
  webClientUrl: string;
  selectedItem: LfRepoTreeNode;
  parentItem: LfRepoTreeNode;
  loggedIn: boolean;
}) {
  const fieldContainer: React.RefObject<
    NgElement & WithProperties<LfFieldContainerComponent>
  > = React.useRef();
  let lfFieldsService: LfFieldsService;

  React.useEffect(() => {
    if (props.repoClient) {
      initializeFieldContainerAsync();
    }
    // setwebClientUrl(loginComponent.current.account_endpoints.webClientUrl);
  }, [props.repoClient, props.loggedIn]);

  const [uploadProgressBar, setuploadProgressBar] = React.useState(false);
  const [fileUploadPercentage, setfileUploadPercentage] = React.useState(5);
  const [showUploadModal, setshowUploadModal] = React.useState(false);
  const [showCreateModal, setshowCreateModal] = React.useState(false);
  const [showAlertModal, setshowAlertModal] = React.useState(false);

  //Open New folder Modal Popup
  const OpenNewFolderModal = () => {
    $('#folderValidation').hide();
    $('#folderExists').hide();
    $('#folderNameValidation').hide();
    $('#folderName').val('');
    setshowCreateModal(true);
  };
  //Open Import file Modal Popup
  const OpenImportFileModal = () => {
    $('#fileValidation').hide();
    $('#fileSizeValidation').hide();
    $('#fileNameValidation').hide();
    $('#fileNameWithBacklash').hide();
    $('#importFileName').text('Choose file');
    $('#importFile').val('');
    $('#uploadFileID').val('');
    setshowUploadModal(true);
    $('#uploadModal .modal-footer').show();
    $('.progress').css('display', 'none');
  };
  //Open file button functinality to open files/folder in repository from the command bar
  const OpenFileOrFolder = async () => {
    const repoId = await props.repoClient.getCurrentRepoId();

    if (props.selectedItem && props.selectedItem.entryType !== EntryType.Folder) {
      if (props.selectedItem.id) {
        // assign the first repoId for now, in production there is only one repository
        window.open(
          props.webClientUrl +
            '/DocView.aspx?db=' +
            repoId +
            '&docid=' +
            props.selectedItem.id
        );
      } else {
        setshowAlertModal(true);
      }
    } else {
      if (props.selectedItem?.id) {
        // assign the first repoId for now, in production there is only one repository
        window.open(
          props.webClientUrl +
            '/browse.aspx?repo=' +
            repoId +
            '#?id=' +
            props.selectedItem.id
        );
      } else {
        setshowAlertModal(true);
      }
    }
  };
  const onDialogOpened = () => {
    $('div.adhoc-modal').css('height', '450px');
  };
  const initializeFieldContainerAsync = async () => {
    fieldContainer.current.addEventListener('dialogOpened', onDialogOpened);

    lfFieldsService = new LfFieldsService(props.repoClient);
    await fieldContainer.current.initAsync(lfFieldsService);
  };

  //Close New folder Modal Popup
  const CloseNewFolderModal = () => {
    $('#folderValidation').hide();
    $('#folderExists').hide();
    $('#folderNameValidation').hide();
    $('#folderName').val('');
    setshowCreateModal(false);
  };

  //Create New Folder in Repository
  const CreateNewFolder = async (folderName) => {
    if ($('#folderName').val() != '') {
      if (/[^ A-Za-z0-9]/.test(folderName)) {
        $('#folderValidation').hide();
        $('#folderExists').hide();
        $('#folderNameValidation').show();
      } else {
        $('#folderValidation').hide();
        $('#folderExists').hide();
        $('#folderNameValidation').hide();

        const repoId = await props.repoClient.getCurrentRepoId();
        const postEntryChildrenRequest: PostEntryChildrenRequest =
          new PostEntryChildrenRequest({
            entryType: PostEntryChildrenEntryType.Folder,
            name: folderName,
          });
        const requestParameters = {
          repoId,
          entryId: Number.parseInt(props.parentItem.id, 10),
          request: postEntryChildrenRequest,
        };
        try {
          const array = [];
          const newFolderEntry: Entry =
            await props.repoClient.entriesClient.createOrCopyEntry(
              requestParameters
            );

          array.push(newFolderEntry);
          setshowCreateModal(false);
          $('#folderName').val('');
        } catch {
          $('#folderExists').show();
        }
      }
    } else {
      $('#folderValidation').show();
    }
  };

  //Close Import File Modal Popup
  const CloseImportFileModal = () => {
    fieldContainer.current.clearAsync();
    $('#importFileName').text('Choose file');
    $('#importFile').val('');
    $('#uploadFileID').val('');
    setshowUploadModal(false);
    $('#fileValidation').hide();
    $('#fileSizeValidation').hide();
    $('#fileNameValidation').hide();
    $('#fileNameWithBacklash').hide();
    $('#uploadModal .modal-footer').show();
  };

  //Import file in Repository
  const ImportFileToRepository = async () => {
    const fileData = document.getElementById('importFile')['files'][0];
    let renameFileName;
    const repoId = await props.repoClient.getCurrentRepoId();
    //Checking file has been uploaded or not
    if (fileData != undefined) {
      const fileDataSize =
        document.getElementById('importFile')['files'][0].size;
      //Checking file size is not exceeding 100mb
      if (fileDataSize < 100000000) {
        //Checking file name is valid or not
        if (
          $('#importFileName').text() !=
          '.' +
            document
              .getElementById('importFile')
              ['value'].split('\\')
              .pop()
              .split('.')[1]
        ) {
          //Checking if user want to change the uploaded file name
          if (fileData.name != $('#importFileName').text()) {
            renameFileName = new File([fileData], $('#importFileName').text());
          } else {
            renameFileName = document.getElementById('importFile')['files'][0];
          }
          $('#fileValidation').hide();
          $('#fileSizeValidation').hide();
          $('#fileNameValidation').hide();
          $('#fileNameWithBacklash').hide();
          const fileContainsBacklash = renameFileName.name.includes('\\')
            ? 'Yes'
            : 'No';
          //Checking if filename contains backlash
          if (fileContainsBacklash === 'No') {
            const fieldValidation = fieldContainer.current.forceValidation();
            //Checking field validation
            if (fieldValidation == true) {
              $('#uploadModal .modal-footer').hide();
              const fieldValues = fieldContainer.current.getFieldValues();
              const formattedFieldValues:
                | {
                    [key: string]: FieldToUpdate;
                  }
                | undefined = {};

              for (const key in fieldValues) {
                const value = fieldValues[key];
                formattedFieldValues[key] = new FieldToUpdate({
                  ...value,
                  values: value.values.map((val) => new ValueToUpdate(val)),
                });
              }

              const templateValue = getTemplateName();
              let templateName;
              if (templateValue != undefined) {
                templateName = templateValue;
              }
              $('.progress').css('display', 'block');

              setuploadProgressBar(!uploadProgressBar);
              setfileUploadPercentage(100);
              const fieldsmetadata: PostEntryWithEdocMetadataRequest =
                new PostEntryWithEdocMetadataRequest({
                  template: templateName,
                  metadata: new PutFieldValsRequest({
                    fields: formattedFieldValues,
                  }),
                });
              //const fileNameSplitByDot = (renameFileName.name as string).split(".");
              const fileNameWithExt = renameFileName.name as string;
              const fileNameSplitByDot = fileNameWithExt.split('.');
              const fileextensionperiod = fileNameSplitByDot.pop();
              const fileNameNoPeriod = fileNameSplitByDot.join('.');
              const parentEntryId = props.parentItem.id;

              const file: FileParameter = {
                data: fileData,
                fileName: fileNameWithExt,
              };
              const requestParameters = {
                repoId,
                parentEntryId: Number.parseInt(parentEntryId, 10),
                electronicDocument: file,
                autoRename: true,
                fileName: fileNameNoPeriod,
                request: fieldsmetadata,
                extension: fileextensionperiod,
              };

              try {
                const entryCreateResult: CreateEntryResult =
                  await props.repoClient.entriesClient.importDocument(
                    requestParameters
                  );
                const result = entryCreateResult.documentLink;
                const parentId = parseInt(result.split('Entries/')[1]);
                const entryResult = [];
                const entry: Entry =
                  await props.repoClient.entriesClient.getEntry({
                    repoId,
                    entryId: parentId,
                    select:
                      'name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id',
                  });
                entryResult.push(entry);
                setshowUploadModal(false);
              } catch (error) {
                window.alert('Error uploding file:' + JSON.stringify(error));
              }
            } else {
              fieldContainer.current.forceValidation();
            }
          } else {
            $('#fileNameWithBacklash').show();
          }
        } else {
          $('#fileNameValidation').show();
        }
      } else {
        $('#fileSizeValidation').show();
      }
    } else {
      $('#fileValidation').show();
    }
  };
  function getTemplateName() {
    const templateValue = fieldContainer.current.getTemplateValue();
    if (templateValue) {
      return templateValue.name;
    }
    return undefined;
  }

  //Set the input file Name
  function SetImportFileName() {
    let fileNamee = '';
    const fileSize = document.getElementById('importFile')['files'][0].size;
    const filenameLength = document
      .getElementById('importFile')
      ['value'].split('\\')
      .pop()
      .split('.').length;
    for (let j = 0; j < filenameLength - 1; j++) {
      const fileSplitValue = document
        .getElementById('importFile')
        ['value'].split('\\')
        .pop()
        .split('.')[j];
      fileNamee += fileSplitValue + '.';
    }
    if (fileSize < 100000000) {
      $('#importFileName').text(
        document.getElementById('importFile')['value'].split('\\').pop()
      );
      //$('#uploadFileID').val(document.getElementById('importFile')["value"].split('\\').pop().split(".")[0]);
      $('#uploadFileID').val(fileNamee.slice(0, -1));
      $('#fileValidation').hide();
      $('#fileSizeValidation').hide();
      $('#fileNameValidation').hide();
      $('#fileNameWithBacklash').hide();
    } else {
      $('#importFileName').text(
        document.getElementById('importFile')['value'].split('\\').pop()
      );
      //$('#uploadFileID').val(document.getElementById('importFile')["value"].split('\\').pop().split(".")[0]);
      $('#uploadFileID').val(fileNamee.slice(0, -1));
      $('#fileSizeValidation').show();
    }
  }

  const SetNewFileName = () => {
    let fileNamee = '';
    const filenameLength = document
      .getElementById('importFile')
      ['value'].split('\\')
      .pop()
      .split('.').length;
    const fileExtension = document
      .getElementById('importFile')
      ['value'].split('\\')
      .pop()
      .split('.')[filenameLength - 1];
    for (let k = 0; k < filenameLength - 1; k++) {
      const fileSplitValue = document
        .getElementById('importFile')
        ['value'].split('\\')
        .pop()
        .split('.')[k];
      fileNamee += fileSplitValue + '.';
    }
    const importFileName = fileNamee.slice(0, -1);
    //let importFileName = document.getElementById('importFile')["value"].split('\\').pop().split(".")[0];
    const fileChangeName = $('#uploadFileID').val();
    if (importFileName != fileChangeName) {
      $('#importFileName').text(
        fileChangeName +
          '.' +
          fileExtension /* document.getElementById('importFile')["value"].split('\\').pop().split(".")[1] */
      );
    }
  };

  const ConfirmAlertButton = () => {
    setshowAlertModal(false);
  };

  return (
    <>
      <div className='p-3' id='mainWebpartContent'>
        <div className='d-flex justify-content-between border p-2 file-option'>
          <a
            href='javascript:;'
            className='mr-3'
            title='Open File'
            onClick={OpenFileOrFolder}
          >
            <span className='material-icons'>description</span>
          </a>
          <span>
            <a
              href='javascript:;'
              className='mr-3'
              title='Upload File'
              onClick={OpenImportFileModal}
            >
              <span className='material-icons'>upload</span>
            </a>
          </span>
          <span>
            <a
              href='javascript:;'
              className='mr-3'
              title='Create Folder'
              onClick={OpenNewFolderModal}
            >
              <span className='material-icons'>create_new_folder</span>
            </a>
          </span>
        </div>
      </div>
      <div
        className='modal'
        id='uploadModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showUploadModal}
      >
        <div className='modal-dialog modal-dialog-scrollable modal-lg'>
          <div className='modal-content' style={{ width: '724px' }}>
            <div className='modal-header'>
              <h5 className='modal-title' id='ModalLabel'>
                Upload File
              </h5>
              <div
                className='progress'
                style={{
                  display: uploadProgressBar ? '' : 'none',
                  width: '100%',
                }}
              >
                <div
                  className='progress-bar progress-bar-striped active'
                  style={{
                    width: fileUploadPercentage + '%',
                    backgroundColor: 'orange',
                  }}
                >
                  Uploading
                </div>
              </div>
            </div>
            <div className='modal-body' style={{ height: '600px' }}>
              <div className='input-group mb-3'>
                <div className='custom-file'>
                  <input
                    type='file'
                    className='custom-file-input'
                    id='importFile'
                    onChange={SetImportFileName}
                    aria-describedby='inputGroupFileAddon04'
                    placeholder='Choose file'
                  />
                  <label className='custom-file-label' id='importFileName'>
                    Choose file
                  </label>
                </div>
              </div>
              <div id='fileValidation' style={{ color: 'red' }}>
                <span>Please select the file to upload</span>
              </div>
              <div id='fileSizeValidation' style={{ color: 'red' }}>
                <span>Please select a file below 100MB size</span>
              </div>
              <div id='fileNameValidation' style={{ color: 'red' }}>
                <span>Please provide proper name of the file</span>
              </div>
              <div id='fileNameWithBacklash' style={{ color: 'red' }}>
                <span>
                  Please provide proper name of the file without backslash
                </span>
              </div>
              <div className='form-group row mb-3'>
                <label className='col-sm-2 col-form-label'>Name</label>
                <div className='col-sm-10'>
                  <input
                    type='text'
                    className='form-control'
                    id='uploadFileID'
                    onChange={SetNewFileName}
                  />
                </div>
              </div>
              <div className='card'>
                <div className='lf-component-container'>
                  <lf-field-container
                    collapsible='true'
                    startCollapsed='true'
                    ref={fieldContainer}
                  />
                </div>
              </div>
            </div>
            <div className='modal-footer'>
              <button
                type='button'
                className='btn btn-primary btn-sm'
                onClick={ImportFileToRepository}
              >
                OK
              </button>
              <button
                type='button'
                className='btn btn-secondary btn-sm'
                onClick={CloseImportFileModal}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      </div>
      <div
        className='modal'
        id='createModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showCreateModal}
      >
        <div className='modal-dialog'>
          <div className='modal-content'>
            <div className='modal-header'>
              <h5 className='modal-title' id='ModalLabel'>
                Create Folder
              </h5>
              <button
                type='button'
                className='close'
                data-dismiss='modal'
                aria-label='Close'
                onClick={CloseNewFolderModal}
              >
                <span aria-hidden='true'>&times;</span>
              </button>
            </div>
            <div className='modal-body'>
              <div className='form-group'>
                <label>Folder Name</label>
                <input
                  type='text'
                  className='form-control'
                  id='folderName'
                  placeholder='Name'
                  ref={(input) => input && input.focus()}
                />
              </div>
              <div id='folderValidation' style={{ color: 'red' }}>
                <span>Please provide folder name</span>
              </div>
              <div id='folderNameValidation' style={{ color: 'red' }}>
                <span>Invalid Name, only alphanumeric are allowed.</span>
              </div>
              <div id='folderExists' style={{ color: 'red' }}>
                <span>Object already exists</span>
              </div>
            </div>
            <div className='modal-footer'>
              <button
                type='button'
                className='btn btn-primary btn-sm'
                data-dismiss='modal'
                onClick={() => CreateNewFolder($('#folderName').val())}
              >
                Submit
              </button>
              <button
                type='button'
                className='btn btn-secondary btn-sm'
                data-dismiss='modal'
                onClick={CloseNewFolderModal}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      </div>
      <div
        className='modal'
        id='AlertModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showAlertModal}
      >
        <div className='modal-dialog'>
          <div className='modal-content'>
            <div className='modal-body'>Please select file/folder to open</div>
            <div className='modal-footer'>
              <button
                type='button'
                className='btn btn-primary btn-sm'
                data-dismiss='modal'
                onClick={ConfirmAlertButton}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
