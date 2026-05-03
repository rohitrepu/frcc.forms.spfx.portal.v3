import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IField {
  title: string;
  internalName: string;
  typeAsString: string;
  required: boolean;
  choices?: string[];
  lookupList?: string;
  lookupField?: string;
}

export interface ILookupOption {
  id: number;
  title: string;
}

interface ISharePointField {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Required: boolean;
  Choices?: string[];
  LookupList?: string;
  LookupField?: string;
}

export async function getListFields(
  siteUrl: string,
  listId: string,
  spHttpClient: SPHttpClient
): Promise<IField[]> {
  const requestUrl =
    `${siteUrl}/_api/web/lists(guid'${listId}')/fields` +
    `?$select=Title,InternalName,TypeAsString,Required,Hidden,ReadOnlyField,Choices,LookupList,LookupField` +
    `&$filter=Hidden eq false and ReadOnlyField eq false`;

  const response: SPHttpClientResponse = await spHttpClient.get(
    requestUrl,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to load list schema. Status: ${response.status}. ListId: ${listId}. Details: ${errorText}`);
  }

  const data = await response.json();

  return data.value
    .filter((field: ISharePointField) =>
      field.InternalName !== 'Attachments' &&
      field.InternalName !== 'ContentType'
    )
    .map((field: ISharePointField): IField => ({
      title: field.Title,
      internalName: field.InternalName,
      typeAsString: field.TypeAsString,
      required: field.Required,
      choices: field.Choices,
      lookupList: field.LookupList,
      lookupField: field.LookupField
    }));
}

export async function getLookupOptions(
  siteUrl: string,
  lookupListId: string,
  lookupField: string,
  spHttpClient: SPHttpClient
): Promise<ILookupOption[]> {
  const safeField = lookupField || 'Title';

  const requestUrl =
    `${siteUrl}/_api/web/lists(guid'${lookupListId}')/items` +
    `?$select=Id,${safeField}&$top=5000&$orderby=${safeField} asc`;

  const response: SPHttpClientResponse = await spHttpClient.get(
    requestUrl,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to load lookup options. Status: ${response.status}. LookupListId: ${lookupListId}. Details: ${errorText}`);
  }

  const data = await response.json();

  return data.value.map((item: { Id: number; [key: string]: string | number }) => ({
    id: item.Id,
    title: String(item[safeField] || item.Id)
  }));
}


export interface IGeneratedFormConfig {
  id: string;
  title: string;
  listId: string;
  listUrl: string;
  layout: string;
  sections: {
    title: string;
    fields: string[];
  }[];
}

export function generateFormConfigFromFields(
  formKey: string,
  title: string,
  listId: string,
  listUrl: string,
  fields: IField[]
): IGeneratedFormConfig {
  return {
    id: formKey,
    title: title,
    listId: listId,
    listUrl: listUrl,
    layout: 'sectioned',
    sections: [
      {
        title: 'Main Information',
        fields: fields.map(field => field.internalName)
      }
    ]
  };
}