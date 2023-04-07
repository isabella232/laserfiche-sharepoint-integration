import * as React from 'react';
import * as bootstrap from 'bootstrap';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IEditManageConfigurationProps } from './IEditManageConfigurationProps';
import { IListItem } from '../IListItem';

import { useEffect, useState } from 'react';
import {
  ProfileHeader,
  validateNewConfiguration,
} from '../ProfileConfigurationComponents';
import ManageConfiguration from '../ManageConfigurationComponent';
import { ProfileConfiguration } from '../ProfileConfigurationComponents';
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

export default function EditManageConfiguration(
  props: IEditManageConfigurationProps
) {
  const [profileConfig, setProfileConfig] = useState<
    ProfileConfiguration | undefined
  >(undefined);

  const [validate, setValidate] = useState(false);
  const handleProfileConfigUpdate = (profileConfig: ProfileConfiguration) => {
    setValidate(false);
    setProfileConfig(profileConfig);
  };

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

  useEffect(() => {
    GetItemIdByTitle().then((results: IListItem[]) => {
      const configurationName = props.match.params.name;
      if (results != null) {
        const profileConfigs = JSON.parse(results[0].JsonValue);
        if (profileConfigs.length > 0) {
          for (let i = 0; i < profileConfigs.length; i++) {
            if (profileConfigs[i].ConfigurationName == configurationName) {
              const selectedConfig: ProfileConfiguration = profileConfigs[i];
              setProfileConfig(selectedConfig);
            }
          }
        }
      }
    });
  }, []);

  async function SaveEditExisitingConfiguration() {
    setValidate(true);
    const validate = validateNewConfiguration(profileConfig);
    if (validate) {
      const manageConfigurationConfig: IListItem[] = await GetItemIdByTitle();
      if (manageConfigurationConfig != null) {
        const configWithCurrentName = manageConfigurationConfig[0];
        const savedProfileConfigurations: ProfileConfiguration[] = JSON.parse(
          configWithCurrentName.JsonValue
        );
        const profileIndex = savedProfileConfigurations.findIndex(
          (config) => config.ConfigurationName === profileConfig.ConfigurationName
        );
        if (profileIndex !== -1) {
          savedProfileConfigurations[profileIndex] = profileConfig;
          const configsToSave = savedProfileConfigurations;
          const succeeded = await saveSPConfigurations(
            configWithCurrentName.Id,
            configsToSave
          );
          return succeeded;
        } else {
          // error this config should exist
        }
      }
    }
    return false;
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

  const header = (
    <div>
      <ProfileHeader
        configurationName={profileConfig?.ConfigurationName}
      ></ProfileHeader>
    </div>
  );
  return profileConfig ? (
    <ManageConfiguration
      header={header}
      repoClient={props.repoClient}
      loggedIn={props.loggedIn}
      profileConfig={profileConfig}
      loadingContent={true}
      createNew={false}
      context={props.context}
      handleProfileConfigUpdate={handleProfileConfigUpdate}
      saveConfiguration={SaveEditExisitingConfiguration}
      validate={validate}
    ></ManageConfiguration>
  ) : (
    <span>'Nothing to see'</span>
  );
}
