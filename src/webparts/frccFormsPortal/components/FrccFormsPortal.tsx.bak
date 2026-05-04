import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './FrccFormsPortal.module.scss';
import type { IFrccFormsPortalProps } from './IFrccFormsPortalProps';
import { getActiveForms, updateFormJson, IFormConfig } from '../../../services/FormConfigService';
import { getListFields, getLookupOptions, IField, ILookupOption } from '../../../services/SchemaService';

import prerequisiteOverride from '../../../config/forms/prerequisite-placement-override.form.json';
import courseSubstitution from '../../../config/forms/course-substitution.form.json';
import studentWaiver from '../../../config/forms/student-waiver.form.json';

const WORKBENCH_STYLE_ID = 'frcc-workbench-fix';

type FormValue = string | number | boolean | string[] | undefined;
type SectionLayout = 'oneColumn' | 'twoColumn';
type TransformType = 'trim' | 'uppercase' | 'lowercase' | 'upper' | 'lower';

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
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const [successMessage, setSuccessMessage] = React.useState<string>('');

  const modernJsonHeaders: HeadersInit = {
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata;charset=utf-8',
    'odata-version': ''
  };

  React.useEffect(() => {
    applyWorkbenchFix();

    async function loadForms(): Promise<void> {
      try {
        const results = await getActiveForms(
          props.context.pageContext.web.absoluteUrl,
          props.context.spHttpClient
        );

        setForms(results);
        setSelectedForm(results[0]);
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

  const getConfig = (): IFormJsonConfig =>
    parseFormJson(selectedForm?.formJson, selectedForm?.formKey);

  const handleChange = (fieldName: string, value: FormValue): void => {
    const config = getConfig();
    const transformedValue = applyTransform(value, config.transform ? config.transform[fieldName] : undefined);

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
    const config = getConfig();
    const hiddenFields = config.hiddenFields || [];
    const fieldOrder = config.fieldOrder || [];

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
    const config = getConfig();
    const errors: { [fieldName: string]: string } = {};
    const warnings: { [fieldName: string]: string } = {};

    renderableFields.forEach(field => {
      const rule = config.validation ? config.validation[field.internalName] : undefined;
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

    if (isPdfForm(selectedForm.listUrl)) {
      setErrorMessage('PDF forms do not need generated FormJson.');
      return;
    }

    try {
      setIsGeneratingJson(true);
      setErrorMessage('');
      setSuccessMessage('');

      const hiddenFields = ['Status'];

      const visibleFields = fields.filter(field =>
        hiddenFields.indexOf(field.internalName) === -1
      );

      const requestFields = visibleFields
        .filter(field =>
          field.internalName.toLowerCase().indexOf('request') !== -1 ||
          field.internalName.toLowerCase().indexOf('term') !== -1 ||
          field.internalName.toLowerCase().indexOf('year') !== -1
        )
        .map(field => field.internalName);

      const studentFields = visibleFields
        .filter(field =>
          field.internalName.toLowerCase().indexOf('student') !== -1 ||
          field.internalName.toLowerCase().indexOf('campus') !== -1 ||
          field.internalName.toLowerCase().indexOf('email') !== -1
        )
        .map(field => field.internalName);

      const courseFields = visibleFields
        .filter(field =>
          field.internalName.toLowerCase().indexOf('course') !== -1 ||
          field.internalName.toLowerCase().indexOf('prereq') !== -1 ||
          field.internalName.toLowerCase().indexOf('placement') !== -1 ||
          field.internalName.toLowerCase().indexOf('cutscore') !== -1
        )
        .map(field => field.internalName);

      const reviewFields = visibleFields
        .filter(field =>
          field.internalName.toLowerCase().indexOf('justification') !== -1 ||
          field.internalName.toLowerCase().indexOf('chair') !== -1 ||
          field.internalName.toLowerCase().indexOf('notes') !== -1 ||
          field.internalName.toLowerCase().indexOf('review') !== -1 ||
          field.typeAsString === 'User' ||
          field.typeAsString === 'Note'
        )
        .map(field => field.internalName);

      const alreadyGrouped = requestFields
        .concat(studentFields)
        .concat(courseFields)
        .concat(reviewFields);

      const otherFields = visibleFields
        .filter(field => alreadyGrouped.indexOf(field.internalName) === -1)
        .map(field => field.internalName);

      const sections: ISectionConfig[] = [
        {
          title: 'Request Information',
          description: 'Basic request details.',
          layout: 'twoColumn' as const,
          fields: requestFields
        },
        {
          title: 'Student Information',
          description: 'Student and contact information.',
          layout: 'twoColumn' as const,
          fields: studentFields
        },
        {
          title: 'Course Details',
          description: 'Course, prerequisite, placement, and related details.',
          layout: 'twoColumn' as const,
          fields: courseFields
        },
        {
          title: 'Review Information',
          description: 'Review notes, approvals, and internal follow-up fields.',
          layout: 'oneColumn' as const,
          fields: reviewFields
        }
      ].filter(section => section.fields.length > 0);

      if (otherFields.length > 0) {
        sections.push({
          title: 'Other Information',
          description: 'Additional fields from the SharePoint list.',
          layout: 'twoColumn' as const,
          fields: otherFields
        });
      }

      const labels = visibleFields.reduce((accumulator: { [key: string]: string }, field: IField) => {
        accumulator[field.internalName] = field.title
          .replace(/([a-z])([A-Z])/g, '$1 $2')
          .replace(/ID/g, 'ID')
          .replace(/Email/g, 'Email')
          .replace(/Prereq/g, 'Prerequisite')
          .replace(/CutScore/g, 'Cut Score');

        return accumulator;
      }, {});

      const helpText: { [key: string]: string } = {};
      const fieldConfig: { [key: string]: IFieldUiConfig } = {};
      const validation: { [key: string]: IValidationRule } = {};
      const transform: { [key: string]: TransformType } = {};

      visibleFields.forEach(field => {
        const lowerName = field.internalName.toLowerCase();

        fieldConfig[field.internalName] = {
          width: field.typeAsString === 'Note' || field.typeAsString === 'User' ? 'full' : 'half'
        };

        if (field.required) {
          validation[field.internalName] = {
            required: true,
            message: `${field.title} is required.`
          };
        }

        if (lowerName.indexOf('email') !== -1) {
          helpText[field.internalName] = 'Enter the email address that should receive updates.';
          fieldConfig[field.internalName].placeholder = 'name@frontrange.edu';
          transform[field.internalName] = 'lowercase';
        }

        if (lowerName.indexOf('name') !== -1 || lowerName.indexOf('title') !== -1) {
          transform[field.internalName] = 'trim';
        }

        if (lowerName.indexOf('justification') !== -1) {
          helpText[field.internalName] = 'Briefly explain why this request is needed.';
          fieldConfig[field.internalName].width = 'full';
        }

        if (field.typeAsString === 'User') {
          helpText[field.internalName] = 'Enter the person’s email address. This will be upgraded to a People Picker later.';
          fieldConfig[field.internalName].placeholder = 'name@frontrange.edu';
          fieldConfig[field.internalName].width = 'full';
        }

        if (field.typeAsString === 'Lookup') {
          helpText[field.internalName] = 'Select the appropriate value from the list.';
        }

        if (field.typeAsString === 'Note') {
          fieldConfig[field.internalName].width = 'full';
        }
      });

      const generatedJson: IFormJsonConfig = {
        hiddenFields,
        sections,
        labels,
        helpText,
        fields: fieldConfig,
        validation,
        transform,
        theme: {
          accentColor: '#005a9e'
        }
      };

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

    const renderableFields = getRenderableFields();

    if (!validateForm(renderableFields)) {
      setErrorMessage('Please complete the required fields.');
      return;
    }

    try {
      setIsSubmitting(true);

      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists(guid'${selectedForm.listId}')/items`;

      const payload: { [fieldName: string]: string | number | boolean | { results: string[] } } = {};

      for (const field of renderableFields) {
        const config = getConfig();
        const value = applyTransform(
          formData[field.internalName],
          config.transform ? config.transform[field.internalName] : undefined
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

  const config = getConfig();
  const renderableFields = getRenderableFields();
  const accentColor = config.theme && config.theme.accentColor ? config.theme.accentColor : '#005a9e';
  const selectedFormIsPdf = isPdfForm(selectedForm?.listUrl);

  const renderField = (field: IField): React.ReactElement => {
    const fieldUiConfig = config.fields ? config.fields[field.internalName] : undefined;

    const label = config.labels && config.labels[field.internalName]
      ? config.labels[field.internalName]
      : field.title;

    const help = config.helpText ? config.helpText[field.internalName] : undefined;
    const value = formData[field.internalName];
    const readOnly = fieldUiConfig && fieldUiConfig.readOnly === true;

    return (
      <div key={field.internalName} style={{ marginBottom: '12px' }}>
        <label>
          {label}
          {field.required || (config.validation && config.validation[field.internalName] && config.validation[field.internalName].required) ? ' *' : ''}
        </label>

        {field.typeAsString === 'Text' && (
          <input type="text" placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'Note' && (
          <textarea placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'Choice' && (
          <select disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)}>
            <option value="">Select...</option>
            {field.choices?.map((choice: string) => (
              <option key={choice} value={choice}>{choice}</option>
            ))}
          </select>
        )}

        {field.typeAsString === 'MultiChoice' && (
          <select multiple disabled={readOnly} style={{ width: '100%', minHeight: '80px' }} value={(value as string[]) || []} onChange={(event) => {
            const selectedValues = Array.from(event.target.selectedOptions).map(option => option.value);
            handleChange(field.internalName, selectedValues);
          }}>
            {field.choices?.map((choice: string) => (
              <option key={choice} value={choice}>{choice}</option>
            ))}
          </select>
        )}

        {(field.typeAsString === 'Number' || field.typeAsString === 'Currency') && (
          <input type="number" placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={{ width: '100%' }} value={value === undefined ? '' : String(value)} onChange={(event) => {
            const nextValue = event.target.value;
            handleChange(field.internalName, nextValue === '' ? undefined : Number(nextValue));
          }} />
        )}

        {field.typeAsString === 'Boolean' && (
          <input type="checkbox" disabled={readOnly} checked={Boolean(value)} onChange={(event) => handleChange(field.internalName, event.target.checked)} />
        )}

        {field.typeAsString === 'DateTime' && (
          <input type="date" disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'User' && (
          <input type="email" placeholder={fieldUiConfig && fieldUiConfig.placeholder ? fieldUiConfig.placeholder : 'Enter user email'} disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {field.typeAsString === 'Lookup' && (
          <select disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value === '' ? undefined : Number(event.target.value))}>
            <option value="">Select...</option>
            {(lookupOptions[field.internalName] || []).map((option: ILookupOption) => (
              <option key={option.id} value={option.id}>{option.title}</option>
            ))}
          </select>
        )}

        {field.typeAsString === 'URL' && (
          <input type="url" placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined} disabled={readOnly} style={{ width: '100%' }} value={String(value || '')} onChange={(event) => handleChange(field.internalName, event.target.value)} />
        )}

        {['Text', 'Note', 'Choice', 'MultiChoice', 'Number', 'Currency', 'Boolean', 'DateTime', 'User', 'Lookup', 'URL'].indexOf(field.typeAsString) === -1 && (
          <input type="text" disabled placeholder={`Unsupported: ${field.typeAsString}`} style={{ width: '100%' }} />
        )}

        {help && (
          <div style={{ fontSize: '12px', color: '#666' }}>{help}</div>
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
    const fieldUiConfig = config.fields ? config.fields[field.internalName] : undefined;

    if (fieldUiConfig && fieldUiConfig.width === 'full') return styles.sectionFull;
    if (field.typeAsString === 'Note' || field.typeAsString === 'User') return styles.sectionFull;

    return '';
  };

  const renderPdfPanel = (): React.ReactElement | null => {
    if (!selectedForm || !selectedForm.listUrl) return null;

    return (
      <div style={{
        maxWidth: '760px',
        border: '1px solid #e1dfdd',
        borderRadius: '4px',
        padding: '28px',
        background: '#ffffff'
      }}>
        <h3 style={{
          marginTop: 0,
          marginBottom: '12px',
          color: '#1f2937'
        }}>
          PDF Form
        </h3>

        <p style={{
          fontSize: '14px',
          lineHeight: '1.6',
          color: '#323130',
          marginBottom: '18px'
        }}>
          This form is provided as a PDF instead of an online SharePoint form. Open the PDF in a new tab to download it, fill it out, print it, or save a completed copy.
        </p>

        <button
          type="button"
          onClick={() => window.open(selectedForm.listUrl, '_blank', 'noopener,noreferrer')}
          className={styles.submitButton}
          style={{
            background: accentColor,
            borderColor: accentColor
          }}
        >
          Open PDF Form
        </button>

        <div style={{
          marginTop: '18px',
          fontSize: '12px',
          color: '#605e5c',
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
        border: '1px solid #e1dfdd',
        borderRadius: '4px',
        padding: '20px 24px',
        background: '#ffffff',
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
          color: '#323130',
          marginBottom: '16px'
        }}>
          Start a new request here, or open the SharePoint list view to review, search, edit, or display existing submissions.
        </p>

        <div style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          <button
            type="button"
            className={styles.submitButton}
            style={{
              background: accentColor,
              borderColor: accentColor
            }}
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
    if (!config.layout || !config.layout.rows || config.layout.rows.length === 0) return null;

    const renderedFieldNames: string[] = [];

    const rowElements = config.layout.rows.map((row, index) => {
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
    const sections = config.sections || [];

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
            borderBottom: `1px solid ${accentColor}`,
            paddingBottom: '8px',
            marginBottom: section.description ? '6px' : '16px',
            color: '#1f2937'
          }}>
            {section.title}
          </h3>

          {section.description && (
            <div style={{ fontSize: '13px', color: '#605e5c', marginBottom: '16px' }}>
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
            borderBottom: `1px solid ${accentColor}`,
            paddingBottom: '8px',
            marginBottom: '16px',
            color: '#1f2937'
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
    <div className={styles.portalLayout}>
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

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && renderableFields.length > 0 && renderSections()}

          {!isLoadingFields && selectedForm && !selectedFormIsPdf && renderableFields.length > 0 && (
            <div className={styles.submitBar}>
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

              <button
                type="button"
                onClick={() => {
                  handleSubmit().catch((error: unknown) => {
                    setErrorMessage(error instanceof Error ? error.message : 'Unexpected submit error.');
                  });
                }}
                disabled={isSubmitting || isGeneratingJson}
                className={styles.submitButton}
                style={{
                  background: accentColor,
                  borderColor: accentColor
                }}
              >
                {isSubmitting ? 'Submitting...' : 'Submit Request'}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}