import * as React from 'react';
import { NavLink } from 'react-router-dom';
import { useEffect } from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';
import { IAdminPageProps } from './IAdminPageProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

export default function AdminMainPage(props: IAdminPageProps) {
  useEffect(() => {
    CreateConfigurations.CreateAdminConfigList(props.context);
    CreateConfigurations.CreateDocumentConfigList(props.context);
  }, []);

  const linkData: LinkInfo[] = [
    { route: '/HomePage', name: 'About' },
    { route: '/ManageConfigurationsPage', name: 'Profiles' },
    { route: '/ManageMappingsPage', name: 'Profile Mapping' },
  ];

  return (
    <div style={{ borderBottom: '3px solid #CE7A14', width: '80%' }}>
      <div>
        <span
          style={{
            marginRight: '450px',
            fontSize: '18px',
            fontWeight: '500',
          }}
        >
          Profile Editor
        </span>
        {props.loggedIn && <Links linkData={linkData} />}
      </div>
    </div>
  );
}

interface LinkInfo {
  route: string;
  name: string;
}

function Links(props: { linkData: LinkInfo[] }) {
  const linkEls = props.linkData.map((link: LinkInfo) => (
    <span key={link.name}>
      <NavLink
        to={link.route}
        activeStyle={{ fontWeight: 'bold', color: 'red' }}
        style={{
          marginRight: '25px',
          fontWeight: '500',
          fontSize: '15px',
        }}
      >
        {link.name}
      </NavLink>
    </span>
  ));
  return <div>{linkEls}</div>;
}

class CreateConfigurations {
  public static CreateDocumentConfigList(context: WebPartContext) {
    const listUrl: string =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('DocumentNameConfigList')";
    context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          return;
        }
        if (response.status === 404) {
          const url: string =
            context.pageContext.web.absoluteUrl + '/_api/web/lists';
          const listDefinition = {
            Title: 'DocumentNameConfigList',
            Description: 'My description',
            BaseTemplate: 100,
          };
          const spHttpClientOptions: ISPHttpClientOptions = {
            body: JSON.stringify(listDefinition),
          };
          context.spHttpClient
            .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((responses: SPHttpClientResponse) => {
              return responses.json();
            })
            .then((responses: { value: [] }): void => {
              console.log(responses);
              const documentlist = responses['Title'];
              this.AddItemsInDocumentConfigList(context, documentlist);
            });
        }
      });
  }

  private static AddItemsInDocumentConfigList(context: WebPartContext, documentList: string) {
    const arary = [
      '%(count)',
      '%(date)',
      '%(datetime)',
      '%(gcount)',
      '%(id)',
      '%(name)',
      '%(parent)',
      '%(parentid)',
      '%(parentname)',
      '%(parentuuid)',
      '%(time)',
      '%(username)',
      '%(usersid)',
      '%(uuid)',
    ];
    for (let i = 0; i < arary.length; i++) {
      const restApiUrl: string =
        context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getByTitle('" +
        documentList +
        "')/items";
      const body: string = JSON.stringify({ Title: arary[i] });
      const options: ISPHttpClientOptions = {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body: body,
      };
      context.spHttpClient.post(
        restApiUrl,
        SPHttpClient.configurations.v1,
        options
      );
    }
  }

  public static CreateAdminConfigList(context: WebPartContext) {
    const listUrl: string =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('AdminConfigurationList')";
    context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          return;
        }
        if (response.status === 404) {
          const url: string =
            context.pageContext.web.absoluteUrl + '/_api/web/lists';
          const listDefinition = {
            Title: 'AdminConfigurationList',
            Description: 'My description',
            BaseTemplate: 100,
          };
          const spHttpClientOptions: ISPHttpClientOptions = {
            body: JSON.stringify(listDefinition),
          };
          context.spHttpClient
            .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((responses: SPHttpClientResponse) => {
              return responses.json();
            })
            .then(async (responses: { value: [] }): Promise<void> => {
              const listtitle = responses['Title'];
              await this.GetFormDigestValue(context, listtitle);
            });
        }
      });
  }

  private static async GetFormDigestValue(context: WebPartContext, listTitle: string) {
    try {
      const res = await fetch(
        context.pageContext.web.absoluteUrl + '/_api/contextinfo',
        {
          method: 'POST',
          headers: { accept: 'application/json;odata=verbose' },
        }
      );
      const contextInfo = await res.json();
      const FormDigestValue =
        contextInfo.GetContextWebInformation.FormDigestValue;
      this.CreateColumns(context, listTitle, FormDigestValue);
    } catch {
      // TODO handle
    }
  }

  private static async CreateColumns(context: WebPartContext, listTitle: string, formDigestValue: string) {
    const siteUrl: string =
      context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/fields";
    try {
      await fetch(siteUrl, {
        method: 'POST',
        body: JSON.stringify({
          __metadata: { type: 'SP.Field' },
          Title: 'JsonValue',
          FieldTypeKind: 3,
        }),
        headers: {
          accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
          'X-RequestDigest': formDigestValue,
        },
      });
      console.log('Fields created');
    } catch {
      console.log('Error!');
    }
  }
}
