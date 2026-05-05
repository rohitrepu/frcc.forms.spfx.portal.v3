import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import styles from './FrccFormsPortal.module.scss';
import type { IFrccFormsPortalProps } from './IFrccFormsPortalProps';
import { getActiveForms, updateFormJson, IFormConfig } from '../../../services/FormConfigService';
import { getListFields, getLookupOptions, IField, ILookupOption } from '../../../services/SchemaService';
import { getTheme } from '../../../services/ThemeService';
import { generateFormLayoutFromFields } from '../../../services/FormSchemaLayoutService';

import prerequisiteOverride from '../../../config/forms/prerequisite-placement-override.form.json';
import courseSubstitution from '../../../config/forms/course-substitution.form.json';
import studentWaiver from '../../../config/forms/student-waiver.form.json';

const WORKBENCH_STYLE_ID = 'frcc-workbench-fix';

type FormValue = string | number | boolean | string[] | undefined;
type SectionLayout = 'oneColumn' | 'twoColumn';
type TransformType = 'trim' | 'uppercase' | 'lowercase' | 'upper' | 'lower';

interface IFormPermissionState {
  canView: boolean;
  canAdd: boolean;
  canEdit: boolean;
}

interface ISectionConfig {
  title: string;
  description?: string;
  layout?: SectionLayout;
  fields: string[];
}

interface IFieldUiConfig {
  placeholder?: string;
  width?: 'half' | 'full';
  readOnly?: boolean;
}

interface IValidationRule {
  required?: boolean;
  message?: string;
  warning?: boolean;
  pattern?: string;
  minLength?: number;
  maxLength?: number;
}

interface IFormJsonConfig {
  hiddenFields?: string[];
  fieldOrder?: string[];
  labels?: { [internalName: string]: string };
  helpText?: { [internalName: string]: string };
  sections?: ISectionConfig[];
  layout?: {
    rows?: string[][];
  };
  fields?: {
    [internalName: string]: IFieldUiConfig;
  };
  validation?: {
    [internalName: string]: IValidationRule;
  };
  defaults?: {
    [internalName: string]: FormValue;
  };
  transform?: {
    [internalName: string]: TransformType;
  };
  theme?: {
    accentColor?: string;
  };
}

const LOCAL_FORM_CONFIGS: { [key: string]: IFormJsonConfig } = {
  prerequisiteOverride: prerequisiteOverride as IFormJsonConfig,
  courseSubstitution: courseSubstitution as IFormJsonConfig,
  studentWaiver: studentWaiver as IFormJsonConfig
};

function applyWorkbenchFix(): void {
  if (window.location.href.indexOf('/_layouts/15/workbench.aspx') === -1) return;
  if (document.getElementById(WORKBENCH_STYLE_ID)) return;

  const style = document.createElement('style');
  style.id = WORKBENCH_STYLE_ID;

  style.innerHTML = `
    body,
    #workbenchPageContent,
    .CanvasZone,
    .CanvasSection,
    .ControlZone,
    .SPCanvas,
    .SPCanvas-canvas {
      max-width: none !important;
      width: 100% !important;
      margin: 0 !important;
    }

    .CanvasZone {
      padding: 0 !important;
    }

    iframe {
      width: 100% !important;
      max-width: none !important;
    }

    body {
      overflow-x: hidden !important;
    }
  `;

  document.head.appendChild(style);
}

function parseFormJson(formJson?: string, formKey?: string): IFormJsonConfig {
  if (formJson) {
    try {
      return JSON.parse(formJson) as IFormJsonConfig;
    } catch (error: unknown) {
      console.warn('Invalid SharePoint FormJson. Falling back to local JSON.', error);
    }
  }

  if (formKey && LOCAL_FORM_CONFIGS[formKey]) {
    return LOCAL_FORM_CONFIGS[formKey];
  }

  return {};
}

function isPdfForm(listUrl?: string): boolean {
  if (!listUrl) return false;

  const lowerUrl = listUrl.toLowerCase();

  return (
    lowerUrl.indexOf('.pdf') !== -1 ||
    lowerUrl.indexOf('/forms/allitems.aspx') !== -1 ||
    lowerUrl.indexOf('/forms/all%20forms.aspx') !== -1
  );
}

function isEmptyValue(value: FormValue): boolean {
  if (value === undefined || value === null || value === '') return true;
  if (Array.isArray(value) && value.length === 0) return true;
  return false;
}

function applyTransform(value: FormValue, transform?: TransformType): FormValue {
  if (typeof value !== 'string') return value;

  if (transform === 'trim') return value.trim();
  if (transform === 'uppercase' || transform === 'upper') return value.toUpperCase();
  if (transform === 'lowercase' || transform === 'lower') return value.toLowerCase();

  return value;
}

export default function FrccFormsPortal(
  props: IFrccFormsPortalProps
): React.ReactElement<IFrccFormsPortalProps> {
  const [forms, setForms] = React.useState<IFormConfig[]>([]);
  const [fields, setFields] = React.useState<IField[]>([]);
  const [lookupOptions, setLookupOptions] = React.useState<{ [fieldName: string]: ILookupOption[] }>({});
  const [formData, setFormData] = React.useState<{ [fieldName: string]: FormValue }>({});
  const [validationErrors, setValidationErrors] = React.useState<{ [fieldName: string]: string }>({});
  const [validationWarnings, setValidationWarnings] = React.useState<{ [fieldName: string]: string }>({});
  const [searchText, setSearchText] = React.useState<string>('');
  const [selectedForm, setSelectedForm] = React.useState<IFormConfig | undefined>(undefined);
  const [isLoadingForms, setIsLoadingForms] = React.useState<boolean>(true);
  const [isLoadingFields, setIsLoadingFields] = React.useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [isGeneratingJson, setIsGeneratingJson] = React.useState<boolean>(false);
  const [isLayoutAdmin, setIsLayoutAdmin] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const [successMessage, setSuccessMessage] = React.useState<string>('');
  const [formPermissionMap, setFormPermissionMap] = React.useState<{ [formId: number]: IFormPermissionState }>({});
  const accessCheckCacheRef = React.useRef<{ [cacheKey: string]: IFormPermissionState }>({});

  const setAccessCheckCache = (
    cacheKey: string,
    permissions: IFormPermissionState
  ): void => {
    accessCheckCacheRef.current = {
      ...accessCheckCacheRef.current,
      [cacheKey]: permissions
    };
  };

  const theme = getTheme();

  const config = parseFormJson(selectedForm?.formJson, selectedForm?.formKey);

  const autoGeneratedConfig = React.useMemo(() => {
    if (selectedForm?.formJson) {
      return config;
    }

    return generateFormLayoutFromFields(fields);
  }, [selectedForm?.formJson, fields]);

  const activeConfig = autoGeneratedConfig;

  const accentColor =
    activeConfig.theme && activeConfig.theme.accentColor
      ? activeConfig.theme.accentColor
      : '#005a9e';

  const inputStyle: React.CSSProperties = {
    width: '100%',
    padding: theme.input.padding,
    borderRadius: theme.input.borderRadius,
    border: `1px solid ${theme.input.borderColor}`,
    fontSize: theme.input.fontSize,
    background: theme.input.background
  };

  const labelStyle: React.CSSProperties = {
    display: 'block',
    fontSize: theme.text.labelSize,
    fontWeight: 600,
    marginBottom: '6px',
    color: theme.text.primary
  };

  const cardStyle: React.CSSProperties = {
    background: theme.card.background,
    border: `1px solid ${theme.card.borderColor}`,
    borderRadius: theme.card.borderRadius,
    padding: theme.card.padding,
    boxShadow: theme.card.shadow
  };

  const primaryButtonStyle: React.CSSProperties = {
    background: accentColor,
    borderColor: accentColor,
    color: theme.button.primaryText,
    borderRadius: theme.button.borderRadius,
    padding: theme.button.padding,
    fontWeight: theme.button.fontWeight
  };

  const sectionHeaderStyle: React.CSSProperties = {
    marginTop: 0,
    padding: theme.section.padding,
    background: accentColor,
    color: theme.section.headerText,
    borderRadius: theme.section.borderRadius,
    fontSize: '16px',
    fontWeight: 600
  };

  const modernJsonHeaders: HeadersInit = {
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
    'odata-version': ''
  };

  const isSharePointUrl = (url?: string): boolean => {
    if (!url) return false;

    const lowerUrl = url.toLowerCase();

    return (
      lowerUrl.indexOf('https://cccs.sharepoint.com/') === 0 ||
      lowerUrl.indexOf(props.context.pageContext.web.absoluteUrl.toLowerCase()) === 0
    );
  };

  const noListPermission: IFormPermissionState = {
    canView: false,
    canAdd: false,
    canEdit: false
  };

  const readOnlyPermission: IFormPermissionState = {
    canView: true,
    canAdd: false,
    canEdit: false
  };

  const externalFormPermission: IFormPermissionState = {
    canView: true,
    canAdd: false,
    canEdit: false
  };

  const getCurrentUserListPermissions = async (listId?: string): Promise<IFormPermissionState> => {
    if (!listId || listId.trim() === '') return noListPermission;

    const cacheKey = `list:${listId.toLowerCase()}`;

    if (accessCheckCacheRef.current[cacheKey] !== undefined) {
      return accessCheckCacheRef.current[cacheKey];
    }

    try {
      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/EffectiveBasePermissions`;

      const response = await props.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        setAccessCheckCache(cacheKey, noListPermission);
        return noListPermission;
      }

      const data = await response.json();
      const permissions = new SPPermission(data);

      const result: IFormPermissionState = {
        canView: permissions.hasPermission(SPPermission.viewListItems),
        canAdd: permissions.hasPermission(SPPermission.addListItems),
        canEdit: permissions.hasPermission(SPPermission.editListItems)
      };

      setAccessCheckCache(cacheKey, result);
      return result;
    } catch (error: unknown) {
      console.warn(`Permission check failed for ListId ${listId}. Hiding form from portal.`, error);
      setAccessCheckCache(cacheKey, noListPermission);
      return noListPermission;
    }
  };

  const getCurrentUserSharePointUrlPermissions = async (url?: string): Promise<IFormPermissionState> => {
    if (!url || !isSharePointUrl(url)) return externalFormPermission;

    const cacheKey = `url:${url.toLowerCase()}`;

    if (accessCheckCacheRef.current[cacheKey] !== undefined) {
      return accessCheckCacheRef.current[cacheKey];
    }

    try {
      const response = await props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      const result = response.ok ? readOnlyPermission : noListPermission;
      setAccessCheckCache(cacheKey, result);
      return result;
    } catch (error: unknown) {
      console.warn(`Access check failed for URL ${url}. Hiding form from portal.`, error);
      setAccessCheckCache(cacheKey, noListPermission);
      return noListPermission;
    }
  };

  const getCurrentUserFormPermissions = async (form: IFormConfig): Promise<IFormPermissionState> => {
    if (form.listId) {
      return getCurrentUserListPermissions(form.listId);
    }

    if (isSharePointUrl(form.listUrl)) {
      return getCurrentUserSharePointUrlPermissions(form.listUrl);
    }

    return externalFormPermission;
  };

  const checkCurrentUserLayoutAdminAccess = async (): Promise<void> => {
    try {
      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('FRCC Forms Configuration')/EffectiveBasePermissions`;

      const response = await props.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        setIsLayoutAdmin(false);
        return;
      }

      const data = await response.json();
      const permissions = new SPPermission(data);

      setIsLayoutAdmin(permissions.hasPermission(SPPermission.editListItems));
    } catch (error: unknown) {
      console.warn('Layout admin check failed. Hiding layout management controls.', error);
      setIsLayoutAdmin(false);
    }
  };

  const filterFormsByCurrentUserPermissions = async (
    formConfigs: IFormConfig[]
  ): Promise<{ visibleForms: IFormConfig[]; permissionMap: { [formId: number]: IFormPermissionState } }> => {
    const permissionResults = await Promise.all(
      formConfigs.map(async form => {
        const permissions = await getCurrentUserFormPermissions(form);
        return { form, permissions };
      })
    );

    const nextPermissionMap: { [formId: number]: IFormPermissionState } = {};
    const visibleForms: IFormConfig[] = [];

    permissionResults.forEach(result => {
      nextPermissionMap[result.form.id] = result.permissions;

      if (result.permissions.canView) {
        visibleForms.push(result.form);
      }
    });

    return { visibleForms, permissionMap: nextPermissionMap };
  };

  React.useEffect(() => {
    applyWorkbenchFix();

    async function loadForms(): Promise<void> {
      try {
        await checkCurrentUserLayoutAdminAccess();

        const results = await getActiveForms(
          props.context.pageContext.web.absoluteUrl,
          props.context.spHttpClient
        );

        const accessFilteredResult = await filterFormsByCurrentUserPermissions(results);

        setFormPermissionMap(accessFilteredResult.permissionMap);
        setForms(accessFilteredResult.visibleForms);
        setSelectedForm(accessFilteredResult.visibleForms[0]);
      } catch (error: unknown) {
        setErrorMessage(error instanceof Error ? error.message : 'Failed to load forms.');
      } finally {
        setIsLoadingForms(false);
      }
    }

    loadForms().catch((error: unknown) => {
      setErrorMessage(error instanceof Error ? error.message : 'Failed to load forms.');
      setIsLoadingForms(false);
    });
  }, [props.context.pageContext.web.absoluteUrl, props.context.spHttpClient]);

  React.useEffect(() => {
    async function loadFields(): Promise<void> {
      if (!selectedForm) {
        setFields([]);
        return;
      }

      if (isPdfForm(selectedForm.listUrl)) {
        setFields([]);
        setErrorMessage('');
        setSuccessMessage('');
        setFormData({});
        setValidationErrors({});
        setValidationWarnings({});
        setLookupOptions({});
        setIsLoadingFields(false);
        return;
      }

      if (!selectedForm.listId) {
        setFields([]);
        setErrorMessage(`Missing ListId for ${selectedForm.title}. Add ListId in FRCC Forms Configuration.`);
        return;
      }

      try {
        setIsLoadingFields(true);
        setErrorMessage('');
        setSuccessMessage('');
        setFormData({});
        setValidationErrors({});
        setValidationWarnings({});
        setLookupOptions({});

        const siteUrl = props.context.pageContext.web.absoluteUrl;

        const result = await getListFields(
          siteUrl,
          selectedForm.listId,
          props.context.spHttpClient
        );

        setFields(result);

        const configForDefaults = parseFormJson(selectedForm.formJson, selectedForm.formKey);

        if (configForDefaults.defaults) {
          setFormData(configForDefaults.defaults);
        }

        const lookupMap: { [fieldName: string]: ILookupOption[] } = {};

        for (const field of result) {
          if (field.typeAsString === 'Lookup' && field.lookupList) {
            lookupMap[field.internalName] = await getLookupOptions(
              siteUrl,
              field.lookupList,
              field.lookupField || 'Title',
              props.context.spHttpClient
            );
          }
        }

        setLookupOptions(lookupMap);
      } catch (error: unknown) {
        setFields([]);
        setErrorMessage(error instanceof Error ? error.message : 'Failed to load list schema.');
      } finally {
        setIsLoadingFields(false);
      }
    }

    loadFields().catch((error: unknown) => {
      setFields([]);
      setErrorMessage(error instanceof Error ? error.message : 'Failed to load list schema.');
      setIsLoadingFields(false);
    });
  }, [selectedForm, props.context.pageContext.web.absoluteUrl, props.context.spHttpClient]);

  const getConfig = (): IFormJsonConfig => activeConfig;

  const handleChange = (fieldName: string, value: FormValue): void => {
    const configForChange = getConfig();
    const transformedValue = applyTransform(value, configForChange.transform ? configForChange.transform[fieldName] : undefined);

    setFormData(previous => ({
      ...previous,
      [fieldName]: transformedValue
    }));

    setValidationErrors(previous => ({
      ...previous,
      [fieldName]: ''
    }));

    setValidationWarnings(previous => ({
      ...previous,
      [fieldName]: ''
    }));
  };

  const resolveUserId = async (email: string): Promise<number> => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const normalizedEmail = email.trim().toLowerCase();

    const response: SPHttpClientResponse = await props.context.spHttpClient.post(
      `${siteUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: modernJsonHeaders,
        body: JSON.stringify({
          logonName: `i:0#.f|membership|${normalizedEmail}`
        })
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to resolve user "${normalizedEmail}". Status: ${response.status}. Details: ${errorText}`
      );
    }

    const data = await response.json();
    return data.Id as number;
  };

  const getRenderableFields = (): IField[] => {
    const configForFields = getConfig();
    const hiddenFields = configForFields.hiddenFields || [];
    const fieldOrder = configForFields.fieldOrder || [];

    const visibleFields = fields.filter(field =>
      hiddenFields.indexOf(field.internalName) === -1
    );

    if (fieldOrder.length === 0) {
      return visibleFields;
    }

    const ordered: IField[] = [];

    fieldOrder.forEach(fieldName => {
      const match = visibleFields.filter(field => field.internalName === fieldName)[0];
      if (match) ordered.push(match);
    });

    visibleFields.forEach(field => {
      if (fieldOrder.indexOf(field.internalName) === -1) {
        ordered.push(field);
      }
    });

    return ordered;
  };

  const validateRule = (
    field: IField,
    value: FormValue,
    rule?: IValidationRule
  ): string => {
    const required = rule && rule.required !== undefined
      ? rule.required
      : field.required;

    if (required && isEmptyValue(value)) {
      return rule && rule.message ? rule.message : `${field.title} is required.`;
    }

    if (typeof value === 'string') {
      if (rule && rule.minLength !== undefined && value.length < rule.minLength) {
        return rule.message || `${field.title} must be at least ${rule.minLength} characters.`;
      }

      if (rule && rule.maxLength !== undefined && value.length > rule.maxLength) {
        return rule.message || `${field.title} must be ${rule.maxLength} characters or fewer.`;
      }

      if (rule && rule.pattern) {
        const regex = new RegExp(rule.pattern);

        if (value && !regex.test(value)) {
          return rule.message || `${field.title} format is invalid.`;
        }
      }
    }

    return '';
  };

  const validateForm = (renderableFields: IField[]): boolean => {
    const configForValidation = getConfig();
    const errors: { [fieldName: string]: string } = {};
    const warnings: { [fieldName: string]: string } = {};

    renderableFields.forEach(field => {
      const rule = configForValidation.validation ? configForValidation.validation[field.internalName] : undefined;
      const value = formData[field.internalName];
      const message = validateRule(field, value, rule);

      if (!message) return;

      if (rule && rule.warning) {
        warnings[field.internalName] = message;
      } else {
        errors[field.internalName] = message;
      }
    });

    setValidationErrors(errors);
    setValidationWarnings(warnings);

    return Object.keys(errors).length === 0;
  };

  const generateFormJson = async (): Promise<void> => {
    if (!selectedForm) return;

    if (!isLayoutAdmin) {
      setErrorMessage('You do not have permission to generate or refresh layout configuration.');
      return;
    }

    if (isPdfForm(selectedForm.listUrl)) {
      setErrorMessage('PDF forms do not need generated FormJson.');
      return;
    }

    try {
      setIsGeneratingJson(true);
      setErrorMessage('');
      setSuccessMessage('');

      const generatedJson: IFormJsonConfig = generateFormLayoutFromFields(fields);
      const formattedJson = JSON.stringify(generatedJson, null, 2);

      await updateFormJson(
        props.context.pageContext.web.absoluteUrl,
        props.context.spHttpClient,
        selectedForm.id,
        formattedJson
      );

      setSelectedForm({
        ...selectedForm,
        formJson: formattedJson
      });

      setForms(previousForms =>
        previousForms.map(form =>
          form.id === selectedForm.id
            ? { ...form, formJson: formattedJson }
            : form
        )
      );

      setSuccessMessage('Enterprise FormJson generated and saved successfully.');
    } catch (error: unknown) {
      setErrorMessage(error instanceof Error ? error.message : 'Failed to generate FormJson.');
    } finally {
      setIsGeneratingJson(false);
    }
  };

  const handleSubmit = async (): Promise<void> => {
    if (!selectedForm) return;

    setErrorMessage('');
    setSuccessMessage('');

    if (!selectedForm.listId) {
      setErrorMessage(`Missing ListId for ${selectedForm.title}. Add ListId in FRCC Forms Configuration.`);
      return;
    }

    const selectedPermissions = formPermissionMap[selectedForm.id];

    if (!selectedPermissions || !selectedPermissions.canAdd) {
      setErrorMessage('You have read-only access to this form. You can view existing requests, but you cannot submit a new request.');
      return;
    }

    const renderableFieldsForSubmit = getRenderableFields();

    if (!validateForm(renderableFieldsForSubmit)) {
      setErrorMessage('Please complete the required fields.');
      return;
    }

    try {
      setIsSubmitting(true);

      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists(guid'${selectedForm.listId}')/items`;

      const payload: { [fieldName: string]: string | number | boolean | { results: string[] } } = {};

      for (const field of renderableFieldsForSubmit) {
        const configForSubmit = getConfig();
        const value = applyTransform(
          formData[field.internalName],
          configForSubmit.transform ? configForSubmit.transform[field.internalName] : undefined
        );

        if (isEmptyValue(value)) {
          continue;
        }

        if (field.typeAsString === 'User') {
          const userId = await resolveUserId(String(value));
          payload[`${field.internalName}Id`] = userId;
        } else if (field.typeAsString === 'Lookup') {
          payload[`${field.internalName}Id`] = Number(value);
        } else if (field.typeAsString === 'MultiChoice' && Array.isArray(value)) {
          payload[field.internalName] = { results: value };
        } else if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
          payload[field.internalName] = value;
        }
      }

      const response: SPHttpClientResponse = await props.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: modernJsonHeaders,
          body: JSON.stringify(payload)
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Submit failed. Status: ${response.status}. Endpoint: ${endpoint}. Payload: ${JSON.stringify(payload)}. Details: ${errorText}`
        );
      }

      setSuccessMessage('Form submitted successfully.');
      setFormData({});
      setValidationErrors({});
      setValidationWarnings({});
    } catch (error: unknown) {
      setErrorMessage(error instanceof Error ? error.message : 'Unexpected submit error.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const filteredForms = forms.filter(form =>
    form.title.toLowerCase().indexOf(searchText.toLowerCase()) !== -1
  );

  const renderableFields = getRenderableFields();
  const selectedFormIsPdf = isPdfForm(selectedForm?.listUrl);
  const selectedFormPermissions = selectedForm ? formPermissionMap[selectedForm.id] : undefined;
  const selectedCanSubmit = selectedFormPermissions ? selectedFormPermissions.canAdd : false;

  const renderField = (field: IField): React.ReactElement => {
    const fieldUiConfig = activeConfig.fields ? activeConfig.fields[field.internalName] : undefined;

    const label = activeConfig.labels && activeConfig.labels[field.internalName]
      ? activeConfig.labels[field.internalName]
      : field.title;

    const help = activeConfig.helpText ? activeConfig.helpText[field.internalName] : undefined;
    const value = formData[field.internalName];
    const readOnly = fieldUiConfig && fieldUiConfig.readOnly === true;

    return (
      <div key={field.internalName} style={{ marginBottom: '12px' }}>
        <label style={labelStyle}>
          {label}
          {field.required || (activeConfig.validation && activeConfig.validation[field.internalName] && activeConfig.validation[field.internalName].required) ? ' *' : ''}
        </label>

        {field.typeAsString === 'Text' && (
          <input type="text" placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'Note' && (
          <textarea placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'Choice' && (
          <select disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)}>
            <option value="">Select...</option>
            {field.choices?.map((choice: string) => (
              <option key={choice} value={choice}>{choice}</option>
            ))}
          </select>
        )}

        {field.typeAsString === 'MultiChoice' && (
          <select multiple disabled={readOnly} style={{ ...inputStyle, minHeight: '80px' }} value={(value as string[]) || []} onChange={(event) => {
            const selectedValues = Array.from(event.target.selectedOptions).map(option => option.value);
            handleChange(field.internalName, selectedValues);
          }}>
            {field.choices?.map((choice: string) => (
              <option key={choice} value={choice}>{choice}</option>
            ))}
          </select>
        )}

        {(field.typeAsString === 'Number' || field.typeAsString === 'Currency') && (
          <input type="number" placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={inputStyle} value={value === undefined ? '' : String(value)} onChange={(event) => {
            const nextValue = event.target.value;
            handleChange(field.internalName, nextValue === '' ? undefined : Number(nextValue));
          }} />
        )}

        {field.typeAsString === 'Boolean' && (
          <input type="checkbox" disabled={readOnly} checked={Boolean(value)} onChange={(event) => handleChange(field.internalName, event.target.checked)} />
        )}

        {field.typeAsString === 'DateTime' && (
          <input type="date" disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'User' && (
          <input type="email" placeholder={fieldUiConfig && fieldUiConfig.placeholder ? fieldUiConfig.placeholder : 'Enter user email'} disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'Lookup' && (
          <select disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value === '' ? undefined : Number(event.target.value))}>
            <option value="">Select...</option>
            {(lookupOptions[field.internalName] || []).map((option: ILookupOption) => (
              <option key={option.id} value={option.id}>{option.title}</option>
            ))}
          </select>
        )}

        {field.typeAsString === 'URL' && (
          <input type="url" placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={inputStyle} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {['Text', 'Note', 'Choice', 'MultiChoice', 'Number', 'Currency', 'Boolean', 'DateTime', 'User', 'Lookup', 'URL'].indexOf(field.typeAsString) === -1 && (
          <input type="text" disabled placeholder={`Unsupported: ${field.typeAsString}`} style={inputStyle} />
        )}

        {help && (
          <div style={{ fontSize: '12px', color: theme.text.secondary }}>{help}</div>
        )}

        {validationWarnings[field.internalName] && (
          <div style={{ color: '#8a6d00', fontSize: '12px' }}>
            {validationWarnings[field.internalName]}
          </div>
        )}

        {validationErrors[field.internalName] && (
          <div style={{ color: '#a80000', fontSize: '12px' }}>
            {validationErrors[field.internalName]}
          </div>
        )}
      </div>
    );
  };

  const getFieldWrapperClass = (field: IField): string => {
    const fieldUiConfig = activeConfig.fields ? activeConfig.fields[field.internalName] : undefined;

    if (fieldUiConfig && fieldUiConfig.width === 'full') return styles.sectionFull;
    if (field.typeAsString === 'Note' || field.typeAsString === 'User') return styles.sectionFull;

    return '';
  };

  const renderPdfPanel = (): React.ReactElement | null => {
    if (!selectedForm || !selectedForm.listUrl) return null;

    return (
      <div style={{
        maxWidth: '760px',
        ...cardStyle
      }}>
        <h3 style={{
          marginTop: 0,
          marginBottom: '12px',
          color: theme.text.primary
        }}>
          PDF Form
        </h3>

        <p style={{
          fontSize: '14px',
          lineHeight: '1.6',
          color: theme.text.primary,
          marginBottom: '18px'
        }}>
          This form is provided as a PDF instead of an online SharePoint form. Open the PDF in a new tab to download it, fill it out, print it, or save a completed copy.
        </p>

        <button
          type="button"
          onClick={() => window.open(selectedForm.listUrl, '_blank', 'noopener,noreferrer')}
          className={styles.submitButton}
          style={primaryButtonStyle}
        >
          Open PDF Form
        </button>

        <div style={{
          marginTop: '18px',
          fontSize: '12px',
          color: theme.text.secondary,
          wordBreak: 'break-all'
        }}>
          {selectedForm.listUrl}
        </div>
      </div>
    );
  };

  const renderFormHubPanel = (): React.ReactElement | null => {
    if (!selectedForm || selectedFormIsPdf) return null;

    return (
      <div style={{
        maxWidth: '900px',
        ...cardStyle,
        marginBottom: '26px'
      }}>
        <h3 style={{
          marginTop: 0,
          marginBottom: '8px',
          color: '#1f2937'
        }}>
          Form Hub
        </h3>

        <p style={{
          fontSize: '14px',
          lineHeight: '1.6',
          color: theme.text.primary,
          marginBottom: '16px'
        }}>
          {selectedCanSubmit
            ? 'Start a new request here, or open the SharePoint list view to review, search, edit, or display existing submissions.'
            : 'You have read-only access to this form. You can open the SharePoint list view to review existing submissions, but you cannot submit a new request from the portal.'}
        </p>

        {!selectedCanSubmit && (
          <div style={{
            marginBottom: '16px',
            padding: '10px 12px',
            border: '1px solid #f1c232',
            borderRadius: '6px',
            background: '#fff8dc',
            color: '#5f4b00',
            fontSize: '13px',
            lineHeight: '1.5'
          }}>
            Read-only access: submission is disabled because your account does not have Add Items permission on this list.
          </div>
        )}

        <div style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          {selectedCanSubmit && (
            <button
              type="button"
              className={styles.submitButton}
              style={primaryButtonStyle}
              onClick={() => {
                setSuccessMessage('');
                setErrorMessage('');
                setFormData({});
                setValidationErrors({});
                setValidationWarnings({});
                window.scrollTo({ top: 0, behavior: 'smooth' });
              }}
            >
              Start New Request
            </button>
          )}

          {selectedForm.listUrl && (
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => window.open(selectedForm.listUrl, '_blank', 'noopener,noreferrer')}
            >
              View Existing Requests
            </button>
          )}
        </div>
      </div>
    );
  };

  const renderRowsLayout = (): React.ReactElement | null => {
    if (!activeConfig.layout || !activeConfig.layout.rows || activeConfig.layout.rows.length === 0) return null;

    const renderedFieldNames: string[] = [];

    const rowElements = activeConfig.layout.rows.map((row, index) => {
      const rowFields = row
        .map(fieldName => renderableFields.filter(field => field.internalName === fieldName)[0])
        .filter((field): field is IField => field !== undefined);

      rowFields.forEach(field => renderedFieldNames.push(field.internalName));

      return (
        <div key={`layout-row-${index}`} style={{
          display: 'grid',
          gridTemplateColumns: rowFields.length > 1 ? `repeat(${rowFields.length}, minmax(250px, 1fr))` : '1fr',
          gap: '18px 28px',
          marginBottom: '18px'
        }}>
          {rowFields.map(field => (
            <div key={field.internalName}>{renderField(field)}</div>
          ))}
        </div>
      );
    });

    const remainingFields = renderableFields.filter(field =>
      renderedFieldNames.indexOf(field.internalName) === -1
    );

    if (remainingFields.length > 0) {
      rowElements.push(
        <div key="layout-row-remaining" className={styles.sectionGrid}>
          {remainingFields.map(field => (
            <div key={field.internalName} className={getFieldWrapperClass(field)}>
              {renderField(field)}
            </div>
          ))}
        </div>
      );
    }

    return <div>{rowElements}</div>;
  };

  const renderSections = (): React.ReactElement[] => {
    const sections = activeConfig.sections || [];

    if (sections.length === 0) {
      return [
        <div key="default-section">
          {renderRowsLayout() || (
            <div className={styles.sectionGrid}>
              {renderableFields.map(field => (
                <div key={field.internalName} className={getFieldWrapperClass(field)}>
                  {renderField(field)}
                </div>
              ))}
            </div>
          )}
        </div>
      ];
    }

    const renderedFieldNames: string[] = [];

    const sectionElements = sections.map(section => {
      const sectionFields = section.fields
        .map(fieldName => renderableFields.filter(field => field.internalName === fieldName)[0])
        .filter((field): field is IField => field !== undefined);

      sectionFields.forEach(field => renderedFieldNames.push(field.internalName));

      return (
        <div key={section.title} style={{ marginBottom: '28px' }}>
          <h3 style={{
            ...sectionHeaderStyle,
            marginBottom: section.description ? '10px' : '18px'
          }}>
            {section.title}
          </h3>

          {section.description && (
            <div style={{ fontSize: '13px', color: theme.text.secondary, marginBottom: '16px' }}>
              {section.description}
            </div>
          )}

          <div className={section.layout === 'oneColumn' ? '' : styles.sectionGrid}>
            {sectionFields.map(field => (
              <div key={field.internalName} className={section.layout === 'oneColumn' ? styles.sectionFull : getFieldWrapperClass(field)}>
                {renderField(field)}
              </div>
            ))}
          </div>
        </div>
      );
    });

    const remainingFields = renderableFields.filter(field =>
      renderedFieldNames.indexOf(field.internalName) === -1
    );

    if (remainingFields.length > 0) {
      sectionElements.push(
        <div key="other-fields" style={{ marginBottom: '28px' }}>
          <h3 style={{
            ...sectionHeaderStyle,
            marginBottom: '18px'
          }}>
            Other Information
          </h3>

          <div className={styles.sectionGrid}>
            {remainingFields.map(field => (
              <div key={field.internalName} className={getFieldWrapperClass(field)}>
                {renderField(field)}
              </div>
            ))}
          </div>
        </div>
      );
    }

    return sectionElements;
  };

  return (
    <div
      className={styles.portalLayout}
      style={{
        background: theme.app.backgroundColor,
        fontFamily: theme.app.fontFamily,
        minHeight: '100vh'
      }}
    >
      <div className={styles.leftNav}>
        <h3>FRCC Forms</h3>

        <input
          type="text"
          placeholder={`Search ${forms.length} forms...`}
          value={searchText}
          onChange={(event) => setSearchText(event.target.value)}
          className={styles.searchBox}
        />

        {searchText && (
          <div className={styles.searchMeta}>
            {filteredForms.length} result{filteredForms.length === 1 ? '' : 's'}
          </div>
        )}

        {isLoadingForms && <div>Loading forms...</div>}

        {!isLoadingForms && errorMessage && forms.length === 0 && (
          <div>{errorMessage}</div>
        )}

        {!isLoadingForms && !errorMessage && filteredForms.length === 0 && (
          <div>No active forms found.</div>
        )}

        {!isLoadingForms && filteredForms.map(form => (
          <button
            key={form.id}
            onClick={() => setSelectedForm(form)}
            className={selectedForm?.id === form.id ? styles.activeFormButton : styles.formButton}
          >
            {form.title}
          </button>
        ))}
      </div>

      <div className={styles.formArea}>
        <div className={styles.formHeader}>
          <h2>{selectedForm?.title || 'Select a form'}</h2>
          <p>{selectedForm?.listUrl || ''}</p>
        </div>

        <div className={styles.formPlaceholder}>
          {isLoadingFields && <div>Loading form fields...</div>}

          {!isLoadingFields && errorMessage && selectedForm && !selectedFormIsPdf && (
            <div style={{ color: '#a80000', marginBottom: '12px' }}>
              {errorMessage}
            </div>
          )}

          {!isLoadingFields && successMessage && (
            <div style={{ color: '#107c10', marginBottom: '12px' }}>
              {successMessage}
            </div>
          )}

          {!isLoadingFields && !selectedForm && (
            <div>Select a form from the left.</div>
          )}

          {!isLoadingFields && selectedForm && selectedFormIsPdf && renderPdfPanel()}

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && renderableFields.length > 0 && renderFormHubPanel()}

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && renderableFields.length === 0 && (
            <div>No editable fields found for this list.</div>
          )}

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && selectedCanSubmit && renderableFields.length > 0 && renderSections()}

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && !selectedCanSubmit && renderableFields.length > 0 && (
            <div style={{
              ...cardStyle,
              maxWidth: '900px',
              marginBottom: '26px'
            }}>
              <h3 style={{ marginTop: 0, color: theme.text.primary }}>Read-only access</h3>
              <p style={{ fontSize: '14px', lineHeight: '1.6', color: theme.text.primary, marginBottom: 0 }}>
                This form is visible because you have Read access to the list. New request fields and submission controls are hidden because you do not have Add Items permission. Use View Existing Requests above to review the list.
              </p>
            </div>
          )}

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && (selectedCanSubmit || isLayoutAdmin) && renderableFields.length > 0 && (
            <div className={styles.submitBar}>
              {isLayoutAdmin && (
                <button
                  type="button"
                  onClick={() => {
                    generateFormJson().catch((error: unknown) => {
                      setErrorMessage(error instanceof Error ? error.message : 'Failed to generate FormJson.');
                    });
                  }}
                  disabled={isSubmitting || isLoadingFields || isGeneratingJson}
                  className={styles.secondaryButton}
                >
                  {isGeneratingJson ? 'Generating...' : 'Generate / Refresh JSON'}
                </button>
              )}

              {selectedCanSubmit && (
                <button
                  type="button"
                  onClick={() => {
                    handleSubmit().catch((error: unknown) => {
                      setErrorMessage(error instanceof Error ? error.message : 'Unexpected submit error.');
                    });
                  }}
                  disabled={isSubmitting || isGeneratingJson}
                  className={styles.submitButton}
                  style={primaryButtonStyle}
                >
                  {isSubmitting ? 'Submitting...' : 'Submit Request'}
                </button>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}