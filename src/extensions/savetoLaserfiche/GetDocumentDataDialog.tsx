import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import styles from './SendToLaserFiche.module.scss';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import {
  ADMIN_CONFIGURATION_LIST,
  MANAGE_CONFIGURATIONS,
  MANAGE_MAPPING,
  SP_LOCAL_STORAGE_KEY,
} from '../../webparts/constants';
import { ISPDocumentData } from '../../Utils/Types';
import { Navigation } from 'spfx-navigation';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  ActionTypes,
  ProfileConfiguration,
  SPProfileConfigurationData,
} from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import {
  IPostEntryWithEdocMetadataRequest,
  FieldToUpdate,
  IPutFieldValsRequest,
  PutFieldValsRequest,
  IValueToUpdate,
  TemplateFieldInfo,
  ValueToUpdate,
  WFieldType,
} from '@laserfiche/lf-repository-api-client';
import { IListItem } from '../../webparts/laserficheAdminConfiguration/components/IListItem';
import { getSPListURL } from '../../Utils/Funcs';
import SaveToLaserficheCustomDialog from './SaveToLaserficheDialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import LoadingDialog from './CommonDialogs';

interface ProfileMappingConfiguration {
  id: string;
  SharePointContentType: string;
  LaserficheContentType: string;
  toggle: boolean;
}

const signInPageRoute = '/SitePages/LaserficheSpSignIn.aspx';

export class GetDocumentDataCustomDialog extends BaseDialog {
  successful = false;

  handleCloseClickAsync = async (success: boolean) => {
    this.successful = success;
    await this.close();
  };

  constructor(
    private fileInfo: {
      fileName: string;
      spContentType: string;
      spFileUrl: string;
      fileId: string;
    },
    private context: BaseComponentContext
  ) {
    super();
  }

  showNextDialog = (data: ISPDocumentData) => {
    const saveToLfDialog = new SaveToLaserficheCustomDialog(data);
    this.secondaryDialogProvider.show(saveToLfDialog).then(() => {
      if (!saveToLfDialog.successful) {
        Navigation.navigate(
          this.context.pageContext.web.absoluteUrl + signInPageRoute,
          true
        );
      }
      this.handleCloseClickAsync(saveToLfDialog.successful);
    });
  };

  public render(): void {
    const element: React.ReactElement = (
      <GetDocumentDialogData
        closeClick={this.handleCloseClickAsync}
        spFileInfo={this.fileInfo}
        context={this.context}
        showSaveToDialog={this.showNextDialog}
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

const FOLLOWING_SP_FIELDS_BLANK_MAPPED_TO_REQUIRED_LF_FIELDS =
  'The following SharePoint field values are blank and are mapped to required Laserfiche fields:';
const PLEASE_FILL_OUT_REQUIRED_FIELDS_TRY_AGAIN =
  'Please fill out these required fields and try again.';

function GetDocumentDialogData(props: {
  closeClick: (success: boolean) => Promise<void>;
  showSaveToDialog: (fileData: ISPDocumentData) => void;
  spFileInfo: {
    fileName: string;
    spContentType: string;
    spFileUrl: string;
    fileId: string;
  };
  context: BaseComponentContext;
}) {
  const [missingFields, setMissingFields] = React.useState<
    undefined | SPProfileConfigurationData[]
  >(undefined);

  if (missingFields) {
    <span>
      {FOLLOWING_SP_FIELDS_BLANK_MAPPED_TO_REQUIRED_LF_FIELDS}
      {missingFields}
      {PLEASE_FILL_OUT_REQUIRED_FIELDS_TRY_AGAIN}
    </span>;
  }

  const listFields = missingFields?.map((field) => (
    <div key={field.Title}>- {field.Title}</div>
  ));

  React.useEffect(() => {
    const libraryUrl = props.context.pageContext.list.title;
    GetAllFieldsValues(libraryUrl, props.spFileInfo.fileId).then(
      (allSpFieldValues: object) => {
        getDocumentDataAsync(allSpFieldValues);
      }
    );
  }, []);

  async function getDocumentDataAsync(allSpFieldValues: object) {
    const response: SPHttpClientResponse = await props.context.spHttpClient.get(
      `${getSPListURL(
        props.context,
        ADMIN_CONFIGURATION_LIST
      )}/items?$filter=Title eq '${MANAGE_MAPPING}'&$top=1`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
        },
      }
    );

    const itemsWithTitleManageMapping = await response.json();
    let matchingMapping = undefined;
    if (itemsWithTitleManageMapping.value?.length > 0) {
      const manageMappingListItem: IListItem =
        itemsWithTitleManageMapping.value[0];
      const manageMappingDetails: ProfileMappingConfiguration[] = JSON.parse(
        manageMappingListItem.JsonValue
      );
      matchingMapping = manageMappingDetails.find(
        (el) => el.SharePointContentType === props.spFileInfo.spContentType
      );
    }

    if (!matchingMapping) {
      sendToLaserficheWithNoMetadata();
    } else {
      const laserficheProfile = matchingMapping.LaserficheContentType;

      const adminConfigList = await props.context.spHttpClient.get(
        `${getSPListURL(
          props.context,
          ADMIN_CONFIGURATION_LIST
        )}/items?$filter=Title eq '${MANAGE_CONFIGURATIONS}'&$top=1`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json',
          },
        }
      );
      const adminConfigListJson = await adminConfigList.json();

      const allConfigs: ProfileConfiguration[] = JSON.parse(
        adminConfigListJson.value[0]['JsonValue']
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
        formatMetadata(
          matchingLFConfig,
          missingRequiredFields,
          allSpFieldValues,
          fields
        );

        if (missingRequiredFields.length === 0) {
          sendToLaserficheWithMetadata(
            fields,
            metadata,
            matchingLFConfig,
            laserficheProfile
          );
        } else {
          setMissingFields(missingRequiredFields);
        }
      } else {
        const fileData: ISPDocumentData = {
          action: matchingLFConfig.Action,
          fileName: props.spFileInfo.fileName,
          fileUrl: props.spFileInfo.spFileUrl,
          documentName: matchingLFConfig.DocumentName,
          entryId: matchingLFConfig.selectedFolder.id,
          contextPageAbsoluteUrl: props.context.pageContext.web.absoluteUrl,
          lfProfile: laserficheProfile,
        };
        window.localStorage.setItem(SP_LOCAL_STORAGE_KEY, JSON.stringify(fileData));

        props.showSaveToDialog(fileData);
      }
    }
  }

  async function GetAllFieldsValues(
    libraryUrl: string,
    fileId: string
  ): Promise<object> {
    try {
      const res = await props.context.spHttpClient.get(
        `${getSPListURL(
          props.context,
          libraryUrl
        )}/items(${fileId})/FieldValuesForEdit`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );

      const allSpFieldValues = await res.json();
      return allSpFieldValues;
    } catch (error) {
      console.log('error ocurred' + error);
    }
  }

  async function sendToLaserficheWithMetadata(
    fields: { [key: string]: FieldToUpdate },
    metadata: IPostEntryWithEdocMetadataRequest,
    matchingLFConfig: ProfileConfiguration,
    laserficheProfileName: string
  ) {
    const metadataFields: IPutFieldValsRequest = {
      fields,
    };
    metadata.metadata = new PutFieldValsRequest(metadataFields);

    const fileData: ISPDocumentData = {
      action: matchingLFConfig.Action,
      contextPageAbsoluteUrl: props.context.pageContext.web.absoluteUrl,
      documentName: matchingLFConfig.DocumentName,
      templateName: matchingLFConfig.selectedTemplateName,
      entryId: matchingLFConfig.selectedFolder.id,
      fileUrl: props.spFileInfo.spFileUrl,
      fileName: props.spFileInfo.fileName,
      metadata,
      lfProfile: laserficheProfileName,
    };
    window.localStorage.setItem(SP_LOCAL_STORAGE_KEY, JSON.stringify(fileData));

    props.showSaveToDialog(fileData);
  }

  async function sendToLaserficheWithNoMetadata() {
    const fileData: ISPDocumentData = {
      fileName: props.spFileInfo.fileName,
      documentName: props.spFileInfo.fileName,
      fileUrl: props.spFileInfo.spFileUrl,
      entryId: '1',
      contextPageAbsoluteUrl: props.context.pageContext.web.absoluteUrl,
      action: ActionTypes.COPY,
    };
    window.localStorage.setItem(SP_LOCAL_STORAGE_KEY, JSON.stringify(fileData));

    props.showSaveToDialog(fileData);
  }

  function formatMetadata(
    matchingLFConfig: ProfileConfiguration,
    missingRequiredFields: SPProfileConfigurationData[],
    allSpFieldValues: object,
    fields: { [key: string]: FieldToUpdate }
  ) {
    for (const mapping of matchingLFConfig.mappedFields) {
      const spFieldName = mapping.spField.Title;
      let spDocFieldValue = allSpFieldValues[spFieldName];

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
      const extractOnlyNumbers = spDocFieldValue.replace(/[^0-9]/g, '');
      const valueAsNumber = Number.parseInt(extractOnlyNumbers, 10);
      if (valueAsNumber > 64999 || valueAsNumber < 0) {
        // TODO invalid field -- should it truncate???
        spDocFieldValue = '';
      } else {
        spDocFieldValue = extractOnlyNumbers;
      }
    } else if (lfField.fieldType === WFieldType.LongInteger) {
      const extractOnlyNumbers = spDocFieldValue.replace(/[^0-9]/g, '');
      const valueAsNumber = Number.parseInt(extractOnlyNumbers, 10);
      if (valueAsNumber > 3999999999 || valueAsNumber < 0) {
        // TODO invalid field -- should it truncate???
        spDocFieldValue = '';
      } else {
        spDocFieldValue = extractOnlyNumbers;
      }
    } else if (lfField.fieldType === WFieldType.Number) {
      const valueRemoveNonNumbers = spDocFieldValue.replace(/[^0-9.]/g, '');
      if (!isNaN(Number.parseFloat(valueRemoveNonNumbers))) {
        spDocFieldValue = valueRemoveNonNumbers;
      } else {
        // TODO invalid field -- should it truncate???
        spDocFieldValue = '';
      }
    }
    return spDocFieldValue;
  }

  return (
    <div className={styles.maindialog}>
      <div id='overlay' className={styles.overlay} />
      <div>
        <img
          src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
          width='42'
          height='42'
        />
        {!(missingFields?.length > 0) && <LoadingDialog />}
        {missingFields?.length > 0 && (
          <MissingFieldsDialog missingFields={listFields} />
        )}
      </div>
    </div>
  );
}

function MissingFieldsDialog(props: { missingFields: JSX.Element[] }) {
  const textInside = (
    <span>
      The following SharePoint field values are blank and are mapped to required
      Laserfiche fields:
      {props.missingFields}Please fill out these required fields and try again.
    </span>
  );

  return (
    <div>
      <p className={styles.text}>{textInside}</p>
    </div>
  );
}
