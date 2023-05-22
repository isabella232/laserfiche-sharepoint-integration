import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export class CreateConfigurations {
  public static CreateDocumentConfigList(context: WebPartContext | ListViewCommandSetContext) {
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

  private static AddItemsInDocumentConfigList(
    context: WebPartContext | ListViewCommandSetContext,
    documentList: string
  ) {
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

  public static CreateAdminConfigList(context: WebPartContext | ListViewCommandSetContext) {
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

  private static async GetFormDigestValue(
    context: WebPartContext | ListViewCommandSetContext,
    listTitle: string
  ) {
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
      this.CreateColumns(context, listTitle, FormDigestValue);
    } catch {
      // TODO handle
    }
  }

  private static async CreateColumns(
    context: WebPartContext | ListViewCommandSetContext,
    listTitle: string,
    formDigestValue: string
  ) {
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
