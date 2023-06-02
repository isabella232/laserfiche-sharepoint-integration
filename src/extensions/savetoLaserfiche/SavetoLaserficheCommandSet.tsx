import { Log } from '@microsoft/sp-core-library';
import SaveToLaserficheCustomDialog from './CustomDialog';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as React from 'react';
import { Navigation } from 'spfx-navigation';
import { NgElement, WithProperties } from '@angular/elements';
import { LfFieldContainerComponent } from '@laserfiche/types-lf-ui-components';
import { IListItem } from '../../webparts/laserficheAdminConfiguration/components/IListItem';
import {
  ActionTypes,
  ProfileConfiguration,
  SPProfileConfigurationData,
} from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { PathUtils } from '@laserfiche/lf-js-utils';
import {
  FieldToUpdate,
  IPostEntryWithEdocMetadataRequest,
  IPutFieldValsRequest,
  IValueToUpdate,
  PutFieldValsRequest,
  TemplateFieldInfo,
  ValueToUpdate,
  WFieldType,
} from '@laserfiche/lf-repository-api-client';
import { ISPDocumentData } from '../../Utils/Types';
import { CreateConfigurations } from '../../Utils/CreateConfigurations';
import {
  ADMIN_CONFIGURATION_LIST,
  MANAGE_CONFIGURATIONS,
  MANAGE_MAPPING,
} from '../../webparts/constants';
import { getSPListURL } from '../../Utils/Funcs';

interface ProfileMappingConfiguration {
  id: string;
  SharePointContentType: string;
  LaserficheContentType: string;
  toggle: boolean;
}

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISendToLfCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

enum SpWebPartNames {
  'LaserficheSpAdministration' = 'LaserficheSpAdministration',
  'LaserficheSpSignIn' = 'LaserficheSpSignIn',
}

const LOG_SOURCE = 'SendToLfCommandSet';
const dialog: SaveToLaserficheCustomDialog = new SaveToLaserficheCustomDialog();
const Redirectpagelink = '/SitePages/LaserficheSpSignIn.aspx';

export default class SendToLfCommandSet extends BaseListViewCommandSet<ISendToLfCommandSetProperties> {
  fieldContainer: React.RefObject<
    NgElement & WithProperties<LfFieldContainerComponent>
  >;
  spFieldNameDefs: {
    InternalName: string;
    Title: string;
    StaticName: string;
  }[] = [];
  allFieldValueStore: object;
  hasSignInPage = false;
  hasAdminPage = false;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SendToLfCommandSet');
    this.fieldContainer = React.createRef();
    window.localStorage.removeItem('spdocdata');
    CreateConfigurations.ensureAdminConfigListCreated(this.context);
    CreateConfigurations.ensureDocumentConfigListCreated(this.context);
    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible =
        event.selectedRows.length === 1 &&
        event.selectedRows[0].getValueByName('ContentType') !== 'Folder';
    }
  }

  public async onExecute(
    event: IListViewCommandSetExecuteEventParameters
  ): Promise<void> {
    const libraryUrl = this.context.pageContext.list.title;
    const allfieldsvalues: RowAccessor = event.selectedRows[0];
    const fileId = allfieldsvalues.getValueByName('ID');
    const fileSize = allfieldsvalues.getValueByName('File_x0020_Size');
    const fileUrl = allfieldsvalues.getValueByName('FileRef');
    const fileName = allfieldsvalues.getValueByName('FileLeafRef');
    const filecontenttypename = allfieldsvalues.getValueByName('ContentType');
    const isCheckedOut = allfieldsvalues.getValueByName('CheckoutUser');

    await this.GetAllFieldsProperties(libraryUrl);
    await this.GetAllFieldsValues(libraryUrl, fileId);
    await this.pageConfigurationCheck();

    const fileExtensionOnly = PathUtils.getFileExtension(fileName);
    const fileNoName = PathUtils.removeFileExtension(fileName);

    const pageOrigin = window.location.origin;
    if (filecontenttypename === 'Folder') {
      alert('Cannot Send a Folder To Laserfiche');
    } else if (!fileNoName || fileNoName.length === 0) {
      alert(
        'Please add a filename to the selected file before trying to save to Laserfiche.'
      );
    } else if (fileExtensionOnly === 'url') {
      alert('Cannot send the .url file to Laserfiche');
    } else if (isCheckedOut?.length > 0) {
      alert(
        'The selected file is checked out. Please discard the checkout or check the file back in before trying to save to Laserfiche.'
      );
    } else if (fileSize > 100000000) {
      alert('Please select a file below 100MB size');
    } else if (!this.hasSignInPage) {
      alert(
        'Missing "LaserficheSpSignIn" SharePoint page. Please refer to the admin guide and complete configuration steps exactly as described.'
      );
    } else if (!this.hasAdminPage) {
      alert(
        'Missing "LaserficheSpAdministration" SharePoint page. Please refer to the admin guide and complete configuration steps exactly as described.'
      );
    } else {
      this.getAdminData(fileName, filecontenttypename, fileUrl, pageOrigin);
    }
  }

  public async pageConfigurationCheck() {
    try {
      const res = await fetch(
        `${getSPListURL(this.context, 'Site Pages')}/items`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const sitePages = await res.json();
      console.log(sitePages);
      for (let o = 0; o < sitePages.value.length; o++) {
        const pageName = sitePages['value'][o]['Title'];
        if (pageName === SpWebPartNames.LaserficheSpSignIn) {
          this.hasSignInPage = true;
        } else if (pageName === SpWebPartNames.LaserficheSpAdministration) {
          this.hasAdminPage = true;
        }
      }
    } catch (error) {
      // TODO
    }
  }

  public async GetAllFieldsProperties(libraryUrl) {
    try {
      const res = await fetch(
        `${getSPListURL(
          this.context,
          libraryUrl
        )}/Fields?$filter=Group ne '_Hidden'`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const results = await res.json();
      this.spFieldNameDefs = JSON.parse(results.value);
      console.log(this.spFieldNameDefs);
      return results;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  public async GetAllFieldsValues(libraryUrl, fileId) {
    try {
      const res = await fetch(
        `${getSPListURL(
          this.context,
          libraryUrl
        )}/items(${fileId})/FieldValuesForEdit`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const results = await res.json();
      this.allFieldValueStore = results;
      console.log(this.allFieldValueStore);
      return this.allFieldValueStore;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  public getAdminData(
    fileName: string,
    filecontenttypename: string,
    fileUrl: string,
    pageOrigin: string
  ) {
    dialog.textInside = <span>Saving your document to Laserfiche</span>;
    dialog.isLoading = true;
    dialog.show();
    const contextPageAbsoluteUrl = this.context.pageContext.web.absoluteUrl;

    this.context.spHttpClient
      .get(
        `${getSPListURL(
          this.context,
          ADMIN_CONFIGURATION_LIST
        )}/items?$filter=Title eq '${MANAGE_MAPPING}'&$top=1`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json',
          },
        }
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(async (response) => {
        let matchingMapping = undefined;
        if (response.value?.length > 0) {
          const SPListItem: IListItem = response.value[0];
          const manageMappingDetails: ProfileMappingConfiguration[] =
            JSON.parse(SPListItem.JsonValue);
          matchingMapping = manageMappingDetails.find(
            (el) => el.SharePointContentType === filecontenttypename
          );
        }

        if (!matchingMapping) {
          this.sendToLaserficheWithNoMetadata(
            fileName,
            fileUrl,
            contextPageAbsoluteUrl,
            pageOrigin
          );
        } else {
          const laserficheProfile = matchingMapping.LaserficheContentType;

          this.context.spHttpClient
            .get(
              `${getSPListURL(
                this.context,
                ADMIN_CONFIGURATION_LIST
              )}/items?$filter=Title eq '${MANAGE_CONFIGURATIONS}'&$top=1`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: 'application/json',
                },
              }
            )
            .then((response1: SPHttpClientResponse) => {
              return response1.json();
            })
            .then(async (response1) => {
              const allConfigs: ProfileConfiguration[] = JSON.parse(
                response1.value[0]['JsonValue']
              );
              const matchingLFConfig = allConfigs.find(
                (lfConfig) => lfConfig.ConfigurationName === laserficheProfile
              );
              if (matchingLFConfig.selectedTemplateName?.length > 0) {
                const metadata: IPostEntryWithEdocMetadataRequest = {
                  template: matchingLFConfig.selectedTemplateName,
                };
                const missingRequiredFields: SPProfileConfigurationData[] = [];
                const fields: { [key: string]: FieldToUpdate } = {};
                this.formatMetadata(
                  matchingLFConfig,
                  missingRequiredFields,
                  fields
                );

                if (missingRequiredFields.length === 0) {
                  this.sendToLaserficheWithMetadata(
                    fields,
                    metadata,
                    matchingLFConfig,
                    contextPageAbsoluteUrl,
                    fileUrl,
                    fileName,
                    pageOrigin,
                    laserficheProfile
                  );
                } else {
                  await dialog.close();
                  const listFields = missingRequiredFields.map((field) => (
                    <div key={field.Title}>- {field.Title}</div>
                  ));
                  dialog.textInside = (
                    <span>
                      The following SharePoint field values are blank and are
                      mapped to required Laserfiche fields:
                      {listFields}Please fill out these required fields and try
                      again.
                    </span>
                  );
                  dialog.isLoading = false;
                  dialog.show();
                  this.spFieldNameDefs = [];
                  this.allFieldValueStore = {};
                }
              } else {
                const fileData: ISPDocumentData = {
                  action: matchingLFConfig.Action,
                  fileName,
                  fileUrl,
                  documentName: matchingLFConfig.DocumentName,
                  entryId: matchingLFConfig.selectedFolder.id,
                  contextPageAbsoluteUrl,
                  pageOrigin,
                  lfProfile: laserficheProfile,
                };
                window.localStorage.setItem(
                  'spdocdata',
                  JSON.stringify(fileData)
                );
                Navigation.navigate(
                  contextPageAbsoluteUrl + Redirectpagelink,
                  true
                );
              }
            });
        }
      });
  }

  private formatMetadata(
    matchingLFConfig: ProfileConfiguration,
    missingRequiredFields: SPProfileConfigurationData[],
    fields: { [key: string]: FieldToUpdate }
  ) {
    for (const mapping of matchingLFConfig.mappedFields) {
      const spFieldName = mapping.spField.Title;
      let spDocFieldValue = this.allFieldValueStore[spFieldName];

      if (spDocFieldValue != undefined || spDocFieldValue != null) {
        const lfField = mapping.lfField;

        spDocFieldValue = forceTruncateToFieldTypeLength(
          lfField,
          spDocFieldValue
        );
        spDocFieldValue = spDocFieldValue.replace(/[\\]/g, `\\\\`);
        spDocFieldValue = spDocFieldValue.replace(/["]/g, `\\"`);

        if (
          lfField.isRequired &&
          (!spDocFieldValue || spDocFieldValue.length === 0)
        ) {
          missingRequiredFields.push(mapping.spField);
        }

        const valueToUpdate: IValueToUpdate = {
          value: spDocFieldValue,
          position: 1,
        };
        const newValueToUpdate = new ValueToUpdate(valueToUpdate);
        fields[lfField.name] = new FieldToUpdate({
          values: [newValueToUpdate],
        });
      } else {
        if (mapping.lfField.isRequired) {
          missingRequiredFields.push(mapping.spField);
        }
      }
    }
  }

  private sendToLaserficheWithMetadata(
    fields: { [key: string]: FieldToUpdate },
    metadata: IPostEntryWithEdocMetadataRequest,
    matchingLFConfig: ProfileConfiguration,
    contextPageAbsoluteUrl: string,
    fileUrl: string,
    fileName: string,
    pageOrigin: string,
    laserficheProfile: string
  ) {
    const metadataFields: IPutFieldValsRequest = {
      fields,
    };
    metadata.metadata = new PutFieldValsRequest(metadataFields);
    const fileData: ISPDocumentData = {
      action: matchingLFConfig.Action,
      contextPageAbsoluteUrl,
      documentName: matchingLFConfig.DocumentName,
      templateName: matchingLFConfig.selectedTemplateName,
      entryId: matchingLFConfig.selectedFolder.id,
      fileUrl,
      fileName,
      metadata,
      pageOrigin,
      lfProfile: laserficheProfile,
    };
    window.localStorage.setItem('spdocdata', JSON.stringify(fileData));
    Navigation.navigate(contextPageAbsoluteUrl + Redirectpagelink, true);
  }

  private sendToLaserficheWithNoMetadata(
    fileName: string,
    fileUrl: string,
    contextPageAbsoluteUrl: string,
    pageOrigin: string
  ) {
    const fileData: ISPDocumentData = {
      fileName,
      documentName: fileName,
      fileUrl,
      entryId: '1',
      contextPageAbsoluteUrl,
      pageOrigin,
      action: ActionTypes.COPY,
    };
    window.localStorage.setItem('spdocdata', JSON.stringify(fileData));

    Navigation.navigate(contextPageAbsoluteUrl + Redirectpagelink, true);
  }
}

function forceTruncateToFieldTypeLength(
  lfField: TemplateFieldInfo,
  spDocFieldValue: string
) {
  if (lfField.length != 0) {
    if (spDocFieldValue.length > lfField.length) {
      // automatically trims length to match constraint
      spDocFieldValue = spDocFieldValue.slice(0, lfField.length);
    }
  } else if (lfField.fieldType === WFieldType.ShortInteger) {
    const extractOnlynumbers = spDocFieldValue.replace(/[^0-9]/g, '');
    const extractOnlynumberslength = extractOnlynumbers.length;
    if (extractOnlynumberslength > 5) {
      spDocFieldValue = extractOnlynumbers.slice(0, 5);
    } else {
      spDocFieldValue = extractOnlynumbers;
    }
  } else if (lfField.fieldType === WFieldType.LongInteger) {
    const extractOnlynumbersLonginteger = spDocFieldValue.replace(
      /[^0-9]/g,
      ''
    );
    const extractOnlynumbersLongintegerlength =
      extractOnlynumbersLonginteger.length;
    if (extractOnlynumbersLongintegerlength > 10) {
      spDocFieldValue = extractOnlynumbersLonginteger.slice(0, 10);
    } else {
      spDocFieldValue = extractOnlynumbersLonginteger;
    }
  } else if (lfField.fieldType === WFieldType.Number) {
    const valueOnlyNumbers = spDocFieldValue.replace(/[^0-9.]/g, '');
    const valueOnlyNumberssplit = valueOnlyNumbers.split('.');
    if (valueOnlyNumberssplit.length === 1) {
      const valueOnlyNumbersLimitcheck = valueOnlyNumbers.split('.')[0];
      if (valueOnlyNumbersLimitcheck.length > 13) {
        spDocFieldValue = valueOnlyNumbersLimitcheck.slice(0, 13);
      } else {
        spDocFieldValue = valueOnlyNumbers;
      }
    } else {
      const valueOnlyNumbersbfrPeriod = valueOnlyNumbers.split('.')[0];
      const valueOnlyNumbersafrPeriod = valueOnlyNumbers.split('.')[1];
      if (
        valueOnlyNumbersbfrPeriod.length <= 13 &&
        valueOnlyNumbersafrPeriod.length <= 5
      ) {
        spDocFieldValue = valueOnlyNumbers;
      } else {
        const valueOnlyNumbersbfrPeriod1 = valueOnlyNumbersbfrPeriod.slice(
          0,
          13
        );
        const valueOnlyNumbersafrPeriod1 = valueOnlyNumbersafrPeriod.slice(
          0,
          5
        );
        spDocFieldValue =
          valueOnlyNumbersbfrPeriod1 + '.' + valueOnlyNumbersafrPeriod1;
      }
    }
  }
  return spDocFieldValue;
}
