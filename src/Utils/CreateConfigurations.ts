import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';
import {
  ADMIN_CONFIGURATION_LIST,
  DOCUMENT_NAME_CONFIG_LIST,
} from '../webparts/constants';
import { getSPListURL } from './Funcs';
import { BaseComponentContext } from '@microsoft/sp-component-base';

const documentNameTokens = [
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

export class CreateConfigurations {
  public static async ensureDocumentConfigListCreatedAsync(
    context: BaseComponentContext
  ) {
    const listUrl: string = getSPListURL(context, DOCUMENT_NAME_CONFIG_LIST);
    const response = await context.spHttpClient.get(
      listUrl,
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      return;
    }
    if (response.status === 404) {
      await CreateConfigurations.createDocumentConfigNameListAsync(context);
    }
  }

  private static async createDocumentConfigNameListAsync(
    context: BaseComponentContext
  ) {
    const url: string = context.pageContext.web.absoluteUrl + '/_api/web/lists';
    const listDefinition = {
      Title: DOCUMENT_NAME_CONFIG_LIST,
      Description: 'My description',
      BaseTemplate: 100,
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(listDefinition),
    };
    const response = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      spHttpClientOptions
    );
    const documentConfigList = await response.json();
    const documentListTitle = documentConfigList['Title'];
    await this.addItemsInDocumentConfigListAsync(context, documentListTitle);
  }

  private static async addItemsInDocumentConfigListAsync(
    context: BaseComponentContext,
    documentList: string
  ) {
    for (const tokenName of documentNameTokens) {
      const restApiUrl: string = getSPListURL(context, documentList) + '/items';
      const body: string = JSON.stringify({ Title: tokenName });
      const options: ISPHttpClientOptions = {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body,
      };
      await context.spHttpClient.post(
        restApiUrl,
        SPHttpClient.configurations.v1,
        options
      );
    }
  }

  public static async ensureAdminConfigListCreatedAsync(
    context: BaseComponentContext
  ) {
    const listUrl: string = getSPListURL(context, ADMIN_CONFIGURATION_LIST);
    const response = await context.spHttpClient.get(
      listUrl,
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      return;
    }
    if (response.status === 404) {
      await CreateConfigurations.createAdminConfigListAsync(context);
    }
  }

  private static async createAdminConfigListAsync(
    context: BaseComponentContext
  ) {
    const url: string = context.pageContext.web.absoluteUrl + '/_api/web/lists';
    const listDefinition = {
      Title: ADMIN_CONFIGURATION_LIST,
      Description: 'My description',
      BaseTemplate: 100,
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(listDefinition),
    };
    const responses: SPHttpClientResponse = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      spHttpClientOptions
    );
    const adminConfigList = await responses.json();
    const listTitle = adminConfigList['Title'];
    const formDigestValue = await this.getFormDigestValueAsync(context);
    await this.createColumnsAsync(context, listTitle, formDigestValue);
  }

  private static async getFormDigestValueAsync(context: BaseComponentContext) {
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
        contextInfo.d.GetContextWebInformation.FormDigestValue;
      return FormDigestValue;
    } catch {
      // TODO handle
    }
  }

  private static async createColumnsAsync(
    context: BaseComponentContext,
    listTitle: string,
    formDigestValue: string
  ) {
    const siteUrl: string = getSPListURL(context, listTitle) + '/fields';
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
