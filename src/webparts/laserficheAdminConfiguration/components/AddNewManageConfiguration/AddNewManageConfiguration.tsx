import * as React from 'react';
import { IAddNewManageConfigurationProps } from './IAddNewManageConfigurationProps';
import ManageConfiguration from '../ManageConfigurationComponent';
import { useState } from 'react';
import { ActionTypes, ProfileConfiguration, validateNewConfiguration } from '../ProfileConfigurationComponents';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListItem } from '../IListItem';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-repository-browser']: any;
    }
  }
}

const initialConfig: ProfileConfiguration = {
  selectedFolder: undefined,
  DocumentName: 'FileName',
  ConfigurationName: '',
  selectedTemplateName: undefined,
  mappedFields: [],
  Action: ActionTypes.COPY,
};

export default function AddNewManageConfiguration(
  props: IAddNewManageConfigurationProps
) {
  const [profileConfig, setProfileConfig] = useState(initialConfig);
  const [validate, setValidate] = useState(false);
  const [configNameError, setConfigNameError] = useState(undefined);
  const handleProfileConfigUpdate = (profileConfig: ProfileConfiguration) => {
    setValidate(false);
    setProfileConfig(profileConfig);
  };
  function handleProfileConfigNameChange(e: React.ChangeEvent) {
    const newName = (e.target as HTMLInputElement).value;
    const profileConfiguration = { ...profileConfig };
    profileConfiguration.ConfigurationName = newName;
    setValidate(false);
    setProfileConfig(profileConfiguration);
  }

  async function saveSPConfigurations(
    Id: string,
    configsToSave: ProfileConfiguration[]
  ) {
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
      Id +
      ')';
    const body: string = JSON.stringify({
      Title: 'ManageConfigurations',
      JsonValue: JSON.stringify(configsToSave),
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
    const response = await props.context.spHttpClient.post(
      restApiUrl,
      SPHttpClient.configurations.v1,
      options
    );
    if (response.ok) {
      return true;
    } else {
      return false;
    }
    // TODO should this really throw?
  }

  async function SaveNewManageConfiguration() {
    setValidate(true);
    setConfigNameError(undefined);
    const validate = validateNewConfiguration(profileConfig);
    if (validate) {
      const manageConfigurationConfig: IListItem[] = await GetItemIdByTitle();
      if (manageConfigurationConfig != null) {
        const configWithCurrentName = manageConfigurationConfig[0];
        const savedProfileConfigurations: ProfileConfiguration[] =
          JSON.parse(configWithCurrentName.JsonValue) ?? [];
        const profileExists = savedProfileConfigurations.find(
          (config) =>
            config.ConfigurationName === profileConfig.ConfigurationName
        );
        if (!profileExists) {
          const allConfigurations =
            savedProfileConfigurations.concat(profileConfig);
          const succeeeded = await saveSPConfigurations(
            configWithCurrentName.Id,
            allConfigurations
          );
          return succeeeded;
        } else {
          setConfigNameError(<span>
            Profile with this name already exists, please provide
            different name
          </span>)
        }
      } else {
        const suceeded = await SaveNewPageConfiguration();
        return suceeded;
      }
    }
    return false;
  }

  async function GetItemIdByTitle(): Promise<IListItem[]> {
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl +
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

  async function SaveNewPageConfiguration() {
    const profileConfigAsString = JSON.stringify(profileConfig);
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/items";
    const body: string = JSON.stringify({
      Title: 'ManageConfigurations',
      JsonValue: profileConfigAsString,
    });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'odata-version': '',
      },
      body: body,
    };
    const response = await props.context.spHttpClient.post(
      restApiUrl,
      SPHttpClient.configurations.v1,
      options
    );
    if (response.ok) {
      return true;
    } else {
      return false;
    }
  }

  let configNameValidation: JSX.Element | undefined;
  if (validate) {
    if(configNameError) {
      configNameValidation = configNameError;
    }
    else if (profileConfig.ConfigurationName == '') {
      configNameValidation = (
        <span>Please specify a name for this configuration</span>
      );
    }
  }

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
            onChange={handleProfileConfigNameChange}
            placeholder='Profile Name'
          />
          <div
            id='configurationExists'
            hidden={!configNameValidation}
            style={{ color: 'red' }}
          >
            {configNameValidation}
          </div>
        </div>
      </div>
    </>
  );
  return (
    <ManageConfiguration
      header={header}
      extraConfiguration={extraConfiguration}
      repoClient={props.repoClient}
      loggedIn={props.loggedIn}
      profileConfig={profileConfig}
      loadingContent={true}
      createNew={true}
      context={props.context}
      handleProfileConfigUpdate={handleProfileConfigUpdate}
      saveConfiguration={SaveNewManageConfiguration}
      validate={validate}
    />
  );
}
