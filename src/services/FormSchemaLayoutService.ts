import type { IField } from './SchemaService';

export type FormValue = string | number | boolean | string[] | undefined;
export type SectionLayout = 'oneColumn' | 'twoColumn';
export type TransformType = 'trim' | 'uppercase' | 'lowercase' | 'upper' | 'lower';

export interface ISectionConfig {
  title: string;
  description?: string;
  layout?: SectionLayout;
  fields: string[];
}

export interface IFieldUiConfig {
  placeholder?: string;
  width?: 'half' | 'full';
  readOnly?: boolean;
}

export interface IValidationRule {
  required?: boolean;
  message?: string;
  warning?: boolean;
  pattern?: string;
  minLength?: number;
  maxLength?: number;
}

export interface IFormJsonConfig {
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

const DEFAULT_HIDDEN_FIELDS: string[] = [
  'ID',
  'ContentType',
  'ContentTypeId',
  'Attachments',
  'Created',
  'Modified',
  'Author',
  'Editor',
  'Edit',
  'LinkTitle',
  'LinkTitleNoMenu',
  'DocIcon',
  'ItemChildCount',
  'FolderChildCount',
  'AppAuthor',
  'AppEditor',
  '_UIVersionString',
  '_UIVersion',
  '_ModerationStatus',
  '_ModerationComments',
  'FileRef',
  'FileDirRef',
  'FileLeafRef',
  'File_x0020_Type',
  'FSObjType',
  'HTML_x0020_File_x0020_Type',
  'ServerUrl',
  'EncodedAbsUrl',
  'BaseName',
  'MetaInfo',
  'owshiddenversion',
  'WorkflowVersion',
  'GUID',
  'ComplianceAssetId'
];

const DEFAULT_INTERNAL_WORKFLOW_KEYWORDS: string[] = [
  'workflow',
  'stage',
  'approval',
  'approver',
  'decision',
  'review',
  'internal',
  'admin',
  'system',
  'status',
  'comments',
  'history',
  'completed',
  'processed'
];

const containsAny = (text: string, keywords: string[]): boolean => {
  const normalizedText = text.toLowerCase();

  return keywords.some(keyword =>
    normalizedText.indexOf(keyword.toLowerCase()) !== -1
  );
};

const getSearchableText = (field: IField): string =>
  `${field.internalName || ''} ${field.title || ''}`.toLowerCase();

const shouldHideField = (field: IField): boolean => {
  if (DEFAULT_HIDDEN_FIELDS.indexOf(field.internalName) !== -1) return true;

  return false;
};

const hasAnyKeyword = (field: IField, keywords: string[]): boolean =>
  containsAny(getSearchableText(field), keywords);

const createLabel = (field: IField): string =>
  field.title
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/ID/g, 'ID')
    .replace(/Email/g, 'Email')
    .replace(/Prereq/g, 'Prerequisite')
    .replace(/CutScore/g, 'Cut Score')
    .trim();

const addSectionIfNotEmpty = (
  sections: ISectionConfig[],
  title: string,
  description: string,
  fields: string[],
  layout: SectionLayout = 'twoColumn'
): void => {
  if (fields.length === 0) return;

  sections.push({
    title,
    description,
    layout,
    fields
  });
};

const shouldUseFullWidth = (field: IField): boolean => {
  if (field.typeAsString === 'Note') return true;
  if (field.typeAsString === 'User') return true;
  if (field.typeAsString === 'MultiChoice') return true;

  return hasAnyKeyword(field, [
    'description',
    'justification',
    'reason',
    'notes',
    'comment',
    'explanation',
    'details',
    'address'
  ]);
};

export const generateFormLayoutFromFields = (fields: IField[]): IFormJsonConfig => {
  const hiddenFields = [...DEFAULT_HIDDEN_FIELDS];

  const visibleFields = fields.filter(field => !shouldHideField(field));

  const groupedFieldNames: string[] = [];

  const takeFields = (keywords: string[]): string[] => {
    const matches = visibleFields
      .filter(field =>
        groupedFieldNames.indexOf(field.internalName) === -1 &&
        hasAnyKeyword(field, keywords)
      )
      .map(field => field.internalName);

    matches.forEach(fieldName => groupedFieldNames.push(fieldName));

    return matches;
  };

  const requesterFields = takeFields([
    'requester',
    'requestor',
    'student',
    'employee',
    'submitter',
    'first',
    'middle',
    'last',
    'name',
    'banner',
    'sid',
    'snumber',
    'idnumber'
  ]);

  const contactFields = takeFields([
    'email',
    'phone',
    'mobile',
    'address',
    'city',
    'state',
    'zip'
  ]);

  const academicFields = takeFields([
    'course',
    'subject',
    'prefix',
    'number',
    'section',
    'crn',
    'prereq',
    'prerequisite',
    'placement',
    'cutscore',
    'cut score',
    'term',
    'semester',
    'year',
    'program',
    'degree',
    'major',
    'catalog',
    'curriculum',
    'campus'
  ]);

  const requestFields = takeFields([
    'request',
    'type',
    'category',
    'department',
    'division',
    'priority',
    'summary',
    'description',
    'reason',
    'justification',
    'purpose',
    'details'
  ]);

  const financialFields = takeFields([
    'amount',
    'cost',
    'total',
    'budget',
    'fund',
    'funding',
    'grant',
    'finance',
    'account',
    'org',
    'programcode',
    'index'
  ]);

  const acknowledgementFields = takeFields([
    'signature',
    'acknowledge',
    'acknowledgement',
    'certification',
    'confirm',
    'agreement',
    'completion',
    'date'
  ]);

  const reviewFields = visibleFields
    .filter(field =>
      groupedFieldNames.indexOf(field.internalName) === -1 &&
      (
        hasAnyKeyword(field, DEFAULT_INTERNAL_WORKFLOW_KEYWORDS) ||
        field.typeAsString === 'User'
      )
    )
    .map(field => field.internalName);

  reviewFields.forEach(fieldName => groupedFieldNames.push(fieldName));

  const otherFields = visibleFields
    .filter(field => groupedFieldNames.indexOf(field.internalName) === -1)
    .map(field => field.internalName);

  const sections: ISectionConfig[] = [];

  addSectionIfNotEmpty(
    sections,
    'Requester / Student Information',
    'Basic requester, student, or employee information.',
    requesterFields
  );

  addSectionIfNotEmpty(
    sections,
    'Contact Information',
    'Email, phone, address, and related contact details.',
    contactFields
  );

  addSectionIfNotEmpty(
    sections,
    'Academic / Course Information',
    'Course, term, program, campus, and academic details.',
    academicFields
  );

  addSectionIfNotEmpty(
    sections,
    'Request Details',
    'Main request details and supporting explanation.',
    requestFields
  );

  addSectionIfNotEmpty(
    sections,
    'Financial / Budget Information',
    'Budget, funding, cost, and finance-related fields.',
    financialFields
  );

  addSectionIfNotEmpty(
    sections,
    'Acknowledgement / Signature',
    'Acknowledgement, certification, date, and signature fields.',
    acknowledgementFields
  );

  addSectionIfNotEmpty(
    sections,
    'Review / Approval',
    'Review notes, approvals, decisions, and internal follow-up fields.',
    reviewFields
  );

  addSectionIfNotEmpty(
    sections,
    'Other Information',
    'Additional fields from the SharePoint list.',
    otherFields
  );

  const labels: { [internalName: string]: string } = {};
  const helpText: { [internalName: string]: string } = {};
  const fieldConfig: { [internalName: string]: IFieldUiConfig } = {};
  const validation: { [internalName: string]: IValidationRule } = {};
  const transform: { [internalName: string]: TransformType } = {};

  visibleFields.forEach(field => {
    const lowerName = field.internalName.toLowerCase();

    labels[field.internalName] = createLabel(field);

    fieldConfig[field.internalName] = {
      width: shouldUseFullWidth(field) ? 'full' : 'half'
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

    if (lowerName.indexOf('phone') !== -1 || lowerName.indexOf('mobile') !== -1) {
      helpText[field.internalName] = 'Enter the best contact phone number.';
    }

    if (lowerName.indexOf('name') !== -1 || lowerName.indexOf('title') !== -1) {
      transform[field.internalName] = 'trim';
    }

    if (
      lowerName.indexOf('justification') !== -1 ||
      lowerName.indexOf('description') !== -1 ||
      lowerName.indexOf('reason') !== -1 ||
      lowerName.indexOf('details') !== -1
    ) {
      helpText[field.internalName] = 'Briefly explain the reason for this request.';
      fieldConfig[field.internalName].width = 'full';
    }

    if (field.typeAsString === 'User') {
      helpText[field.internalName] = 'Enter the person’s email address.';
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

  return {
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
};