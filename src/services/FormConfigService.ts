import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IFormConfig {
  id: number;
  title: string;
  formKey: string;
  listUrl: string;
  listId: string;
  formJson?: string;
}

export async function getActiveForms(
  siteUrl: string,
  spHttpClient: SPHttpClient
): Promise<IFormConfig[]> {
  const requestUrl =
    `${siteUrl}/_api/web/lists/getbytitle('FRCC Forms Configuration')/items` +
    `?$select=Id,Title,FormKey,ListUrl,ListId,FormJson,IsActive` +
    `&$filter=IsActive eq 'Yes'` +
    `&$orderby=Title asc`;

  const response: SPHttpClientResponse = await spHttpClient.get(
    requestUrl,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to load forms configuration: ${response.status}. Details: ${errorText}`);
  }

  const data = await response.json();

  return data.value.map((item: {
    Id: number;
    Title: string;
    FormKey: string;
    ListUrl: string;
    ListId: string;
    FormJson?: string;
  }): IFormConfig => ({
    id: item.Id,
    title: item.Title,
    formKey: item.FormKey,
    listUrl: item.ListUrl,
    listId: item.ListId,
    formJson: item.FormJson
  }));
}

export async function updateFormJson(
  siteUrl: string,
  spHttpClient: SPHttpClient,
  itemId: number,
  formJson: string
): Promise<void> {
  const requestUrl =
    `${siteUrl}/_api/web/lists/getbytitle('FRCC Forms Configuration')/items(${itemId})`;

  const response: SPHttpClientResponse = await spHttpClient.post(
    requestUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
        'odata-version': ''
      },
      body: JSON.stringify({
        FormJson: formJson
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to update FormJson: ${response.status}. Details: ${errorText}`);
  }
}