import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { SPPermission } from "@microsoft/sp-page-context";
import styles from "./FrccFormsPortal.module.scss";
import type { IFrccFormsPortalProps } from "./IFrccFormsPortalProps";
import {
  getActiveForms,
  updateFormJson,
  IFormConfig,
} from "../../../services/FormConfigService";
import {
  getListFields,
  getLookupOptions,
  IField,
  ILookupOption,
} from "../../../services/SchemaService";
import { getTheme } from "../../../services/ThemeService";

import prerequisiteOverride from "../../../config/forms/prerequisite-placement-override.form.json";
import courseSubstitution from "../../../config/forms/course-substitution.form.json";
import studentWaiver from "../../../config/forms/student-waiver.form.json";

const WORKBENCH_STYLE_ID = "frcc-workbench-fix";

type FormValue = string | number | boolean | string[] | undefined;
type SectionLayout = "oneColumn" | "twoColumn";
type TransformType = "trim" | "uppercase" | "lowercase" | "upper" | "lower";
type NavView = "all" | "favorites" | "recent" | "popular";

const CONFIG_LIST_ID = "b57ddb21-bb42-4b2f-9a70-bf9f3cc3e984";
const NAV_PAGE_SIZE = 12;
const FAVORITES_STORAGE_KEY = "frccFormsPortalV3FavoriteFormIds";
const RECENT_STORAGE_KEY = "frccFormsPortalV3RecentFormIds";
const USAGE_STORAGE_KEY = "frccFormsPortalV3UsageMap";

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
  width?: "half" | "full";
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
  studentWaiver: studentWaiver as IFormJsonConfig,
};

function applyWorkbenchFix(): void {
  if (window.location.href.indexOf("/_layouts/15/workbench.aspx") === -1)
    return;
  if (document.getElementById(WORKBENCH_STYLE_ID)) return;

  const style = document.createElement("style");
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
      console.warn(
        "Invalid SharePoint FormJson. Falling back to local JSON.",
        error,
      );
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
    lowerUrl.indexOf(".pdf") !== -1 ||
    lowerUrl.indexOf("/forms/allitems.aspx") !== -1 ||
    lowerUrl.indexOf("/forms/all%20forms.aspx") !== -1
  );
}

function isEmptyValue(value: FormValue): boolean {
  if (value === undefined || value === null || value === "") return true;
  if (Array.isArray(value) && value.length === 0) return true;
  return false;
}

function applyTransform(
  value: FormValue,
  transform?: TransformType,
): FormValue {
  if (typeof value !== "string") return value;

  if (transform === "trim") return value.trim();
  if (transform === "uppercase" || transform === "upper")
    return value.toUpperCase();
  if (transform === "lowercase" || transform === "lower")
    return value.toLowerCase();

  return value;
}

function getStringValue(source: unknown, key: string): string {
  const record = source as { [key: string]: unknown };
  const value = record ? record[key] : undefined;

  if (value === undefined || value === null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean")
    return String(value);

  return "";
}

function getArrayValue(source: unknown, key: string): string[] {
  const record = source as { [key: string]: unknown };
  const value = record ? record[key] : undefined;

  if (!value) return [];
  if (Array.isArray(value))
    return value
      .map((item) => String(item))
      .filter((item) => item.trim() !== "");

  const maybeResults = value as { results?: unknown[] };
  if (Array.isArray(maybeResults.results)) {
    return maybeResults.results
      .map((item) => String(item))
      .filter((item) => item.trim() !== "");
  }

  if (typeof value === "string") {
    return value
      .split(/[;,]/)
      .map((item) => item.trim())
      .filter((item) => item !== "");
  }

  return [];
}

function readStoredNumberArray(key: string): number[] {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return [];

    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];

    return parsed
      .map((item) => Number(item))
      .filter((item) => !Number.isNaN(item));
  } catch {
    return [];
  }
}

function readStoredUsageMap(key: string): { [id: number]: number } {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return {};

    const parsed = JSON.parse(raw) as { [key: string]: number };
    const result: { [id: number]: number } = {};

    Object.keys(parsed).forEach((formId) => {
      const numericId = Number(formId);
      const count = Number(parsed[formId]);

      if (!Number.isNaN(numericId) && !Number.isNaN(count)) {
        result[numericId] = count;
      }
    });

    return result;
  } catch {
    return {};
  }
}

function saveToLocalStorage(key: string, value: unknown): void {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Local storage is optional. The portal still works without it.
  }
}

export default function FrccFormsPortal(
  props: IFrccFormsPortalProps,
): React.ReactElement<IFrccFormsPortalProps> {
  const [forms, setForms] = React.useState<IFormConfig[]>([]);
  const [fields, setFields] = React.useState<IField[]>([]);
  const [lookupOptions, setLookupOptions] = React.useState<{
    [fieldName: string]: ILookupOption[];
  }>({});
  const [formData, setFormData] = React.useState<{
    [fieldName: string]: FormValue;
  }>({});
  const [validationErrors, setValidationErrors] = React.useState<{
    [fieldName: string]: string;
  }>({});
  const [validationWarnings, setValidationWarnings] = React.useState<{
    [fieldName: string]: string;
  }>({});
  const [searchText, setSearchText] = React.useState<string>("");
  const [selectedForm, setSelectedForm] = React.useState<
    IFormConfig | undefined
  >(undefined);
  const [isLoadingForms, setIsLoadingForms] = React.useState<boolean>(true);
  const [isLoadingFields, setIsLoadingFields] = React.useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [isGeneratingJson, setIsGeneratingJson] =
    React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string>("");
  const [successMessage, setSuccessMessage] = React.useState<string>("");
  const [divisionFilter, setDivisionFilter] = React.useState<string>("All");
  const [formTypeFilter, setFormTypeFilter] = React.useState<string>("All");
  const [audienceFilter, setAudienceFilter] = React.useState<string>("All");
  const [navView, setNavView] = React.useState<NavView>("all");
  const [currentPage, setCurrentPage] = React.useState<number>(1);
  const [favoriteFormIds, setFavoriteFormIds] = React.useState<number[]>([]);
  const [recentFormIds, setRecentFormIds] = React.useState<number[]>([]);
  const [usageMap, setUsageMap] = React.useState<{ [id: number]: number }>({});
  const [formPermissionMap, setFormPermissionMap] = React.useState<{
    [formId: number]: IFormPermissionState;
  }>({});
  const [canManageFormConfiguration, setCanManageFormConfiguration] =
    React.useState<boolean>(false);
  const accessCheckCacheRef = React.useRef<{
    [cacheKey: string]: IFormPermissionState;
  }>({});

  const theme = getTheme();

  const inputStyle: React.CSSProperties = {
    width: "100%",
    padding: theme.input.padding,
    borderRadius: theme.input.borderRadius,
    border: `1px solid ${theme.input.borderColor}`,
    fontSize: theme.input.fontSize,
    background: theme.input.background,
  };

  const labelStyle: React.CSSProperties = {
    display: "block",
    fontSize: theme.text.labelSize,
    fontWeight: 600,
    marginBottom: "6px",
    color: theme.text.primary,
  };

  const cardStyle: React.CSSProperties = {
    background: theme.card.background,
    border: `1px solid ${theme.card.borderColor}`,
    borderRadius: theme.card.borderRadius,
    padding: theme.card.padding,
    boxShadow: theme.card.shadow,
  };

  const primaryButtonStyle: React.CSSProperties = {
    background: theme.button.primaryBackground,
    borderColor: theme.button.primaryBackground,
    color: theme.button.primaryText,
    borderRadius: theme.button.borderRadius,
    padding: theme.button.padding,
    fontWeight: theme.button.fontWeight,
  };

  const sectionHeaderStyle: React.CSSProperties = {
    marginTop: 0,
    padding: "12px 16px",
    background: "linear-gradient(90deg, #005a9e 0%, #006fbf 100%)",
    color: "#ffffff",
    borderRadius: "10px",
    fontSize: "16px",
    fontWeight: 800,
    boxShadow: "0 4px 10px rgba(0, 90, 158, 0.16)",
  };

  const filterSelectStyle: React.CSSProperties = {
    width: "100%",
    minHeight: "36px",
    padding: "8px 10px",
    borderRadius: "8px",
    border: "1px solid #c8c6c4",
    background: "#ffffff",
    color: "#1f2937",
    fontSize: "13px",
    boxSizing: "border-box",
  };

  const navButtonBaseStyle: React.CSSProperties = {
    border: "1px solid #d0d7de",
    borderRadius: "999px",
    padding: "6px 10px",
    background: "#ffffff",
    color: "#374151",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: 600,
  };

  const metadataChipStyle: React.CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    border: "1px solid #dbeafe",
    borderRadius: "999px",
    background: "#eff6ff",
    color: "#1e3a8a",
    padding: "6px 10px",
    fontSize: "12px",
    fontWeight: 600,
  };

  const modernJsonHeaders: HeadersInit = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata;charset=utf-8",
    "odata-version": "",
  };

  React.useEffect(() => {
    setFavoriteFormIds(readStoredNumberArray(FAVORITES_STORAGE_KEY));
    setRecentFormIds(readStoredNumberArray(RECENT_STORAGE_KEY));
    setUsageMap(readStoredUsageMap(USAGE_STORAGE_KEY));
  }, []);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [searchText, divisionFilter, formTypeFilter, audienceFilter, navView]);

  const isSharePointUrl = (url?: string): boolean => {
    if (!url) return false;

    const lowerUrl = url.toLowerCase();

    return (
      lowerUrl.indexOf("https://cccs.sharepoint.com/") === 0 ||
      lowerUrl.indexOf(
        props.context.pageContext.web.absoluteUrl.toLowerCase(),
      ) === 0
    );
  };

  const noListPermission: IFormPermissionState = {
    canView: false,
    canAdd: false,
    canEdit: false,
  };

  const readOnlyPermission: IFormPermissionState = {
    canView: true,
    canAdd: false,
    canEdit: false,
  };

  const externalFormPermission: IFormPermissionState = {
    canView: true,
    canAdd: false,
    canEdit: false,
  };

  const getFormType = (form?: IFormConfig): string => {
    if (!form) return "";
    return getStringValue(form, "formType") || getStringValue(form, "FormType");
  };

  const getDivision = (form?: IFormConfig): string => {
    if (!form) return "";
    return getStringValue(form, "division") || getStringValue(form, "Division");
  };

  const getContactEmail = (form?: IFormConfig): string => {
    if (!form) return "";
    return (
      getStringValue(form, "contactEmail") ||
      getStringValue(form, "ContactEmail")
    );
  };

  const getOwnerTitle = (form?: IFormConfig): string => {
    if (!form) return "";
    return (
      getStringValue(form, "ownerTitle") ||
      getStringValue(form, "OwnerTitle") ||
      getStringValue(form, "Owner")
    );
  };

  const getOwnerEmail = (form?: IFormConfig): string => {
    if (!form) return "";
    return (
      getStringValue(form, "ownerEmail") || getStringValue(form, "OwnerEmail")
    );
  };

  const getAudience = (form?: IFormConfig): string[] => {
    if (!form) return [];

    const lower = getArrayValue(form, "audience");
    const upper = getArrayValue(form, "Audience");
    const combined = [...lower, ...upper];

    return Array.from(new Set(combined.filter((item) => item.trim() !== "")));
  };

  const getFormSearchText = (form: IFormConfig): string => {
    return [
      form.title,
      form.listUrl,
      getFormType(form),
      getDivision(form),
      getOwnerTitle(form),
      getOwnerEmail(form),
      getContactEmail(form),
      getAudience(form).join(" "),
    ]
      .join(" ")
      .toLowerCase();
  };

  const enrichFormsWithGovernanceMetadata = async (
    formConfigs: IFormConfig[],
  ): Promise<IFormConfig[]> => {
    if (formConfigs.length === 0) return formConfigs;

    try {
      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists(guid'${CONFIG_LIST_ID}')/items?$select=Id,FormType,Division,Audience,ContactEmail,Owner/Title,Owner/EMail&$expand=Owner&$top=5000`;

      const response = await props.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        },
      );

      if (!response.ok) {
        console.warn(
          "Could not load governance metadata from FRCC Forms Configuration. Continuing with base form data.",
        );
        return formConfigs;
      }

      const data = await response.json();
      const items = (data.value || []) as Array<{
        Id: number;
        FormType?: string;
        Division?: string;
        Audience?: { results?: string[] } | string[] | string;
        ContactEmail?: string;
        Owner?: { Title?: string; EMail?: string };
      }>;

      const itemMap: { [id: number]: (typeof items)[number] } = {};
      items.forEach((item) => {
        itemMap[item.Id] = item;
      });

      return formConfigs.map((form) => {
        const item = itemMap[form.id];
        if (!item) return form;

        const audience = Array.isArray(item.Audience)
          ? item.Audience
          : item.Audience &&
              typeof item.Audience === "object" &&
              Array.isArray(item.Audience.results)
            ? item.Audience.results
            : typeof item.Audience === "string"
              ? item.Audience.split(/[;,]/)
                  .map((value) => value.trim())
                  .filter((value) => value !== "")
              : [];

        return {
          ...form,
          formType: item.FormType || "",
          division: item.Division || "",
          audience,
          contactEmail: item.ContactEmail || "",
          ownerTitle: item.Owner?.Title || "",
          ownerEmail: item.Owner?.EMail || "",
        } as IFormConfig;
      });
    } catch (error: unknown) {
      console.warn(
        "Governance metadata enrichment failed. Continuing with base form data.",
        error,
      );
      return formConfigs;
    }
  };

  const getCurrentUserListPermissions = async (
    listId?: string,
  ): Promise<IFormPermissionState> => {
    if (!listId || listId.trim() === "") return noListPermission;

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
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        },
      );

      if (!response.ok) {
        accessCheckCacheRef.current[cacheKey] = noListPermission;
        return noListPermission;
      }

      const data = await response.json();
      const permissions = new SPPermission(data);

      const result: IFormPermissionState = {
        canView: permissions.hasPermission(SPPermission.viewListItems),
        canAdd: permissions.hasPermission(SPPermission.addListItems),
        canEdit: permissions.hasPermission(SPPermission.editListItems),
      };

      accessCheckCacheRef.current[cacheKey] = result;
      return result;
    } catch (error: unknown) {
      console.warn(
        `Permission check failed for ListId ${listId}. Hiding form from portal.`,
        error,
      );
      accessCheckCacheRef.current[cacheKey] = noListPermission;
      return noListPermission;
    }
  };

  const getCurrentUserSharePointUrlPermissions = async (
    url?: string,
  ): Promise<IFormPermissionState> => {
    if (!url || !isSharePointUrl(url)) return externalFormPermission;

    const cacheKey = `url:${url.toLowerCase()}`;

    if (accessCheckCacheRef.current[cacheKey] !== undefined) {
      return accessCheckCacheRef.current[cacheKey];
    }

    try {
      const response = await props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
      );

      const result = response.ok ? readOnlyPermission : noListPermission;
      accessCheckCacheRef.current[cacheKey] = result;
      return result;
    } catch (error: unknown) {
      console.warn(
        `Access check failed for URL ${url}. Hiding form from portal.`,
        error,
      );
      accessCheckCacheRef.current[cacheKey] = noListPermission;
      return noListPermission;
    }
  };

  const getCurrentUserFormPermissions = async (
    form: IFormConfig,
  ): Promise<IFormPermissionState> => {
    if (form.listId) {
      return getCurrentUserListPermissions(form.listId);
    }

    if (isSharePointUrl(form.listUrl)) {
      return getCurrentUserSharePointUrlPermissions(form.listUrl);
    }

    return externalFormPermission;
  };

  const filterFormsByCurrentUserPermissions = async (
    formConfigs: IFormConfig[],
  ): Promise<{
    visibleForms: IFormConfig[];
    permissionMap: { [formId: number]: IFormPermissionState };
  }> => {
    const permissionResults = await Promise.all(
      formConfigs.map(async (form) => {
        const permissions = await getCurrentUserFormPermissions(form);
        return { form, permissions };
      }),
    );

    const nextPermissionMap: { [formId: number]: IFormPermissionState } = {};
    const visibleForms: IFormConfig[] = [];

    permissionResults.forEach((result) => {
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
        const results = await getActiveForms(
          props.context.pageContext.web.absoluteUrl,
          props.context.spHttpClient,
        );

        const configListPermissions =
          await getCurrentUserListPermissions(CONFIG_LIST_ID);
        setCanManageFormConfiguration(configListPermissions.canEdit);

        const enrichedResults =
          await enrichFormsWithGovernanceMetadata(results);
        const accessFilteredResult =
          await filterFormsByCurrentUserPermissions(enrichedResults);

        setFormPermissionMap(accessFilteredResult.permissionMap);
        setForms(accessFilteredResult.visibleForms);
        setSelectedForm(accessFilteredResult.visibleForms[0]);
      } catch (error: unknown) {
        setErrorMessage(
          error instanceof Error ? error.message : "Failed to load forms.",
        );
      } finally {
        setIsLoadingForms(false);
      }
    }

    loadForms().catch((error: unknown) => {
      setErrorMessage(
        error instanceof Error ? error.message : "Failed to load forms.",
      );
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
        setErrorMessage("");
        setSuccessMessage("");
        setFormData({});
        setValidationErrors({});
        setValidationWarnings({});
        setLookupOptions({});
        setIsLoadingFields(false);
        return;
      }

      if (!selectedForm.listId) {
        setFields([]);
        setErrorMessage(
          `Missing ListId for ${selectedForm.title}. Add ListId in FRCC Forms Configuration.`,
        );
        return;
      }

      try {
        setIsLoadingFields(true);
        setErrorMessage("");
        setSuccessMessage("");
        setFormData({});
        setValidationErrors({});
        setValidationWarnings({});
        setLookupOptions({});

        const siteUrl = props.context.pageContext.web.absoluteUrl;

        const result = await getListFields(
          siteUrl,
          selectedForm.listId,
          props.context.spHttpClient,
        );

        setFields(result);

        const configForDefaults = parseFormJson(
          selectedForm.formJson,
          selectedForm.formKey,
        );

        if (configForDefaults.defaults) {
          setFormData(configForDefaults.defaults);
        }

        const lookupMap: { [fieldName: string]: ILookupOption[] } = {};

        for (const field of result) {
          if (field.typeAsString === "Lookup" && field.lookupList) {
            lookupMap[field.internalName] = await getLookupOptions(
              siteUrl,
              field.lookupList,
              field.lookupField || "Title",
              props.context.spHttpClient,
            );
          }
        }

        setLookupOptions(lookupMap);
      } catch (error: unknown) {
        setFields([]);
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "Failed to load list schema.",
        );
      } finally {
        setIsLoadingFields(false);
      }
    }

    loadFields().catch((error: unknown) => {
      setFields([]);
      setErrorMessage(
        error instanceof Error ? error.message : "Failed to load list schema.",
      );
      setIsLoadingFields(false);
    });
  }, [
    selectedForm,
    props.context.pageContext.web.absoluteUrl,
    props.context.spHttpClient,
  ]);

  const getConfig = (): IFormJsonConfig =>
    parseFormJson(selectedForm?.formJson, selectedForm?.formKey);

  const handleChange = (fieldName: string, value: FormValue): void => {
    const config = getConfig();
    const transformedValue = applyTransform(
      value,
      config.transform ? config.transform[fieldName] : undefined,
    );

    setFormData((previous) => ({
      ...previous,
      [fieldName]: transformedValue,
    }));

    setValidationErrors((previous) => ({
      ...previous,
      [fieldName]: "",
    }));

    setValidationWarnings((previous) => ({
      ...previous,
      [fieldName]: "",
    }));
  };

  const resolveUserId = async (email: string): Promise<number> => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const normalizedEmail = email.trim().toLowerCase();

    const response: SPHttpClientResponse =
      await props.context.spHttpClient.post(
        `${siteUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        {
          headers: modernJsonHeaders,
          body: JSON.stringify({
            logonName: `i:0#.f|membership|${normalizedEmail}`,
          }),
        },
      );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to resolve user "${normalizedEmail}". Status: ${response.status}. Details: ${errorText}`,
      );
    }

    const data = await response.json();
    return data.Id as number;
  };

  const getRenderableFields = (): IField[] => {
    const config = getConfig();
    const hiddenFields = config.hiddenFields || [];
    const fieldOrder = config.fieldOrder || [];

    const visibleFields = fields.filter(
      (field) => hiddenFields.indexOf(field.internalName) === -1,
    );

    if (fieldOrder.length === 0) {
      return visibleFields;
    }

    const ordered: IField[] = [];

    fieldOrder.forEach((fieldName) => {
      const match = visibleFields.filter(
        (field) => field.internalName === fieldName,
      )[0];
      if (match) ordered.push(match);
    });

    visibleFields.forEach((field) => {
      if (fieldOrder.indexOf(field.internalName) === -1) {
        ordered.push(field);
      }
    });

    return ordered;
  };

  const validateRule = (
    field: IField,
    value: FormValue,
    rule?: IValidationRule,
  ): string => {
    const required =
      rule && rule.required !== undefined ? rule.required : field.required;

    if (required && isEmptyValue(value)) {
      return rule && rule.message
        ? rule.message
        : `${field.title} is required.`;
    }

    if (typeof value === "string") {
      if (
        rule &&
        rule.minLength !== undefined &&
        value.length < rule.minLength
      ) {
        return (
          rule.message ||
          `${field.title} must be at least ${rule.minLength} characters.`
        );
      }

      if (
        rule &&
        rule.maxLength !== undefined &&
        value.length > rule.maxLength
      ) {
        return (
          rule.message ||
          `${field.title} must be ${rule.maxLength} characters or fewer.`
        );
      }

      if (rule && rule.pattern) {
        const regex = new RegExp(rule.pattern);

        if (value && !regex.test(value)) {
          return rule.message || `${field.title} format is invalid.`;
        }
      }
    }

    return "";
  };

  const validateForm = (renderableFields: IField[]): boolean => {
    const config = getConfig();
    const errors: { [fieldName: string]: string } = {};
    const warnings: { [fieldName: string]: string } = {};

    renderableFields.forEach((field) => {
      const rule = config.validation
        ? config.validation[field.internalName]
        : undefined;
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
      setErrorMessage("PDF forms do not need generated FormJson.");
      return;
    }

    try {
      setIsGeneratingJson(true);
      setErrorMessage("");
      setSuccessMessage("");

      const hiddenFields = ["Status"];

      const visibleFields = fields.filter(
        (field) => hiddenFields.indexOf(field.internalName) === -1,
      );

      const hasAnyKeyword = (field: IField, keywords: string[]): boolean => {
        const lowerName = field.internalName.toLowerCase();
        const lowerTitle = field.title.toLowerCase();
        const searchableText = `${lowerName} ${lowerTitle}`;

        return keywords.some(
          (keyword) => searchableText.indexOf(keyword) !== -1,
        );
      };

      const groupedFieldNames: string[] = [];

      const takeFields = (keywords: string[]): string[] => {
        const matches = visibleFields
          .filter(
            (field) =>
              groupedFieldNames.indexOf(field.internalName) === -1 &&
              hasAnyKeyword(field, keywords),
          )
          .map((field) => field.internalName);

        matches.forEach((fieldName) => groupedFieldNames.push(fieldName));

        return matches;
      };

      const requesterFields = takeFields([
        "requester",
        "requestor",
        "student",
        "employee",
        "first",
        "middle",
        "last",
        "name",
      ]);

      const contactFields = takeFields([
        "email",
        "phone",
        "address",
        "city",
        "state",
        "zip",
      ]);

      const requestFields = takeFields([
        "request",
        "type",
        "category",
        "campus",
        "department",
        "priority",
        "summary",
        "description",
        "reason",
        "justification",
      ]);

      const academicFields = takeFields([
        "course",
        "prereq",
        "prerequisite",
        "placement",
        "cutscore",
        "cut score",
        "term",
        "semester",
        "year",
        "program",
        "degree",
        "major",
      ]);

      const financialFields = takeFields([
        "amount",
        "cost",
        "total",
        "budget",
        "fund",
        "funding",
        "grant",
        "finance",
        "account",
      ]);

      const signatureFields = takeFields([
        "signature",
        "acknowledge",
        "acknowledgement",
        "certification",
        "completion",
        "date",
      ]);

      const reviewFields = visibleFields
        .filter(
          (field) =>
            groupedFieldNames.indexOf(field.internalName) === -1 &&
            (hasAnyKeyword(field, [
              "supervisor",
              "advisor",
              "approver",
              "approval",
              "decision",
              "workflow",
              "stage",
              "review",
              "chair",
              "dean",
              "director",
              "notes",
            ]) ||
              field.typeAsString === "User" ||
              field.typeAsString === "Note"),
        )
        .map((field) => field.internalName);

      reviewFields.forEach((fieldName) => groupedFieldNames.push(fieldName));

      const otherFields = visibleFields
        .filter((field) => groupedFieldNames.indexOf(field.internalName) === -1)
        .map((field) => field.internalName);

      const sections: ISectionConfig[] = [
        {
          title: "Requester / Student Information",
          description: "Basic requester, student, or employee identity fields.",
          layout: "twoColumn" as const,
          fields: requesterFields,
        },
        {
          title: "Contact Information",
          description: "Email, phone, address, and related contact details.",
          layout: "twoColumn" as const,
          fields: contactFields,
        },
        {
          title: "Request Details",
          description: "Main request details and supporting explanation.",
          layout: "twoColumn" as const,
          fields: requestFields,
        },
        {
          title: "Academic / Course Information",
          description: "Course, term, program, and academic details.",
          layout: "twoColumn" as const,
          fields: academicFields,
        },
        {
          title: "Financial / Budget Information",
          description: "Budget, funding, cost, and finance-related fields.",
          layout: "twoColumn" as const,
          fields: financialFields,
        },
        {
          title: "Acknowledgement / Signature",
          description:
            "Acknowledgement, certification, date, and signature fields.",
          layout: "twoColumn" as const,
          fields: signatureFields,
        },
        {
          title: "Review / Approval",
          description:
            "Review notes, approvals, decisions, and internal follow-up fields.",
          layout: "twoColumn" as const,
          fields: reviewFields,
        },
      ].filter((section) => section.fields.length > 0);

      if (otherFields.length > 0) {
        sections.push({
          title: "Other Information",
          description: "Additional fields from the SharePoint list.",
          layout: "twoColumn" as const,
          fields: otherFields,
        });
      }

      const labels = visibleFields.reduce(
        (accumulator: { [key: string]: string }, field: IField) => {
          accumulator[field.internalName] = field.title
            .replace(/([a-z])([A-Z])/g, "$1 $2")
            .replace(/ID/g, "ID")
            .replace(/Email/g, "Email")
            .replace(/Prereq/g, "Prerequisite")
            .replace(/CutScore/g, "Cut Score");

          return accumulator;
        },
        {},
      );

      const helpText: { [key: string]: string } = {};
      const fieldConfig: { [key: string]: IFieldUiConfig } = {};
      const validation: { [key: string]: IValidationRule } = {};
      const transform: { [key: string]: TransformType } = {};

      visibleFields.forEach((field) => {
        const lowerName = field.internalName.toLowerCase();

        fieldConfig[field.internalName] = {
          width:
            field.typeAsString === "Note" || field.typeAsString === "User"
              ? "full"
              : "half",
        };

        if (field.required) {
          validation[field.internalName] = {
            required: true,
            message: `${field.title} is required.`,
          };
        }

        if (lowerName.indexOf("email") !== -1) {
          helpText[field.internalName] =
            "Enter the email address that should receive updates.";
          fieldConfig[field.internalName].placeholder = "name@frontrange.edu";
          transform[field.internalName] = "lowercase";
        }

        if (
          lowerName.indexOf("name") !== -1 ||
          lowerName.indexOf("title") !== -1
        ) {
          transform[field.internalName] = "trim";
        }

        if (lowerName.indexOf("justification") !== -1) {
          helpText[field.internalName] =
            "Briefly explain why this request is needed.";
          fieldConfig[field.internalName].width = "full";
        }

        if (field.typeAsString === "User") {
          helpText[field.internalName] =
            "Enter the person’s email address. This will be upgraded to a People Picker later.";
          fieldConfig[field.internalName].placeholder = "name@frontrange.edu";
          fieldConfig[field.internalName].width = "full";
        }

        if (field.typeAsString === "Lookup") {
          helpText[field.internalName] =
            "Select the appropriate value from the list.";
        }

        if (field.typeAsString === "Note") {
          fieldConfig[field.internalName].width = "full";
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
          accentColor: "#005a9e",
        },
      };

      const formattedJson = JSON.stringify(generatedJson, null, 2);

      await updateFormJson(
        props.context.pageContext.web.absoluteUrl,
        props.context.spHttpClient,
        selectedForm.id,
        formattedJson,
      );

      setSelectedForm({
        ...selectedForm,
        formJson: formattedJson,
      });

      setForms((previousForms) =>
        previousForms.map((form) =>
          form.id === selectedForm.id
            ? { ...form, formJson: formattedJson }
            : form,
        ),
      );

      setSuccessMessage(
        "Enterprise FormJson generated and saved successfully.",
      );
    } catch (error: unknown) {
      setErrorMessage(
        error instanceof Error ? error.message : "Failed to generate FormJson.",
      );
    } finally {
      setIsGeneratingJson(false);
    }
  };

  const handleSubmit = async (): Promise<void> => {
    if (!selectedForm) return;

    setErrorMessage("");
    setSuccessMessage("");

    if (!selectedForm.listId) {
      setErrorMessage(
        `Missing ListId for ${selectedForm.title}. Add ListId in FRCC Forms Configuration.`,
      );
      return;
    }

    const selectedPermissions = formPermissionMap[selectedForm.id];

    if (!selectedPermissions || !selectedPermissions.canAdd) {
      setErrorMessage(
        "You have read-only access to this form. You can view existing requests, but you cannot submit a new request.",
      );
      return;
    }

    const renderableFields = getRenderableFields();

    if (!validateForm(renderableFields)) {
      setErrorMessage("Please complete the required fields.");
      return;
    }

    try {
      setIsSubmitting(true);

      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists(guid'${selectedForm.listId}')/items`;

      const payload: {
        [fieldName: string]: string | number | boolean | { results: string[] };
      } = {};

      for (const field of renderableFields) {
        const config = getConfig();
        const value = applyTransform(
          formData[field.internalName],
          config.transform ? config.transform[field.internalName] : undefined,
        );

        if (isEmptyValue(value)) {
          continue;
        }

        if (field.typeAsString === "User") {
          const userId = await resolveUserId(String(value));
          payload[`${field.internalName}Id`] = userId;
        } else if (field.typeAsString === "Lookup") {
          payload[`${field.internalName}Id`] = Number(value);
        } else if (
          field.typeAsString === "MultiChoice" &&
          Array.isArray(value)
        ) {
          payload[field.internalName] = { results: value };
        } else if (
          typeof value === "string" ||
          typeof value === "number" ||
          typeof value === "boolean"
        ) {
          payload[field.internalName] = value;
        }
      }

      const response: SPHttpClientResponse =
        await props.context.spHttpClient.post(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: modernJsonHeaders,
            body: JSON.stringify(payload),
          },
        );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Submit failed. Status: ${response.status}. Endpoint: ${endpoint}. Payload: ${JSON.stringify(payload)}. Details: ${errorText}`,
        );
      }

      setSuccessMessage("Form submitted successfully.");
      setFormData({});
      setValidationErrors({});
      setValidationWarnings({});
    } catch (error: unknown) {
      setErrorMessage(
        error instanceof Error ? error.message : "Unexpected submit error.",
      );
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSelectForm = (form: IFormConfig): void => {
    setSelectedForm(form);
    setSuccessMessage("");
    setErrorMessage("");

    const nextRecent = [
      form.id,
      ...recentFormIds.filter((id) => id !== form.id),
    ].slice(0, 10);
    const nextUsageMap = {
      ...usageMap,
      [form.id]: (usageMap[form.id] || 0) + 1,
    };

    setRecentFormIds(nextRecent);
    setUsageMap(nextUsageMap);
    saveToLocalStorage(RECENT_STORAGE_KEY, nextRecent);
    saveToLocalStorage(USAGE_STORAGE_KEY, nextUsageMap);
  };

  const toggleFavorite = (formId: number): void => {
    const nextFavorites =
      favoriteFormIds.indexOf(formId) === -1
        ? [...favoriteFormIds, formId]
        : favoriteFormIds.filter((id) => id !== formId);

    setFavoriteFormIds(nextFavorites);
    saveToLocalStorage(FAVORITES_STORAGE_KEY, nextFavorites);
  };

  const getUniqueValues = (values: string[]): string[] =>
    Array.from(new Set(values.filter((value) => value.trim() !== ""))).sort(
      (a, b) => a.localeCompare(b),
    );

  const divisionOptions = getUniqueValues(
    forms.map((form) => getDivision(form)),
  );
  const formTypeOptions = getUniqueValues(
    forms.map((form) => getFormType(form)),
  );
  const audienceOptions = getUniqueValues(
    forms.reduce(
      (accumulator: string[], form) => accumulator.concat(getAudience(form)),
      [],
    ),
  );

  const formsAfterSearchAndFilters = forms.filter((form) => {
    const searchableText = getFormSearchText(form);
    const matchesSearch =
      searchText.trim() === "" ||
      searchableText.indexOf(searchText.trim().toLowerCase()) !== -1;
    const matchesDivision =
      divisionFilter === "All" || getDivision(form) === divisionFilter;
    const matchesFormType =
      formTypeFilter === "All" || getFormType(form) === formTypeFilter;
    const matchesAudience =
      audienceFilter === "All" ||
      getAudience(form).indexOf(audienceFilter) !== -1;

    return (
      matchesSearch && matchesDivision && matchesFormType && matchesAudience
    );
  });

  const navScopedForms = (() => {
    if (navView === "favorites") {
      return formsAfterSearchAndFilters.filter(
        (form) => favoriteFormIds.indexOf(form.id) !== -1,
      );
    }

    if (navView === "recent") {
      const recentIndex: { [id: number]: number } = {};
      recentFormIds.forEach((id, index) => {
        recentIndex[id] = index;
      });

      return formsAfterSearchAndFilters
        .filter((form) => recentFormIds.indexOf(form.id) !== -1)
        .sort((a, b) => recentIndex[a.id] - recentIndex[b.id]);
    }

    if (navView === "popular") {
      return [...formsAfterSearchAndFilters].sort(
        (a, b) => (usageMap[b.id] || 0) - (usageMap[a.id] || 0),
      );
    }

    return [...formsAfterSearchAndFilters].sort((a, b) => {
      const divisionCompare = (getDivision(a) || "Unassigned").localeCompare(
        getDivision(b) || "Unassigned",
      );
      if (divisionCompare !== 0) return divisionCompare;
      return a.title.localeCompare(b.title);
    });
  })();

  const totalPages = Math.max(
    1,
    Math.ceil(navScopedForms.length / NAV_PAGE_SIZE),
  );
  const safeCurrentPage = Math.min(currentPage, totalPages);
  const pagedForms = navScopedForms.slice(
    (safeCurrentPage - 1) * NAV_PAGE_SIZE,
    safeCurrentPage * NAV_PAGE_SIZE,
  );

  const filteredForms = navScopedForms;

  const config = getConfig();
  const renderableFields = getRenderableFields();
  const accentColor =
    config.theme && config.theme.accentColor
      ? config.theme.accentColor
      : "#005a9e";
  const selectedFormIsPdf = isPdfForm(selectedForm?.listUrl);
  const selectedFormPermissions = selectedForm
    ? formPermissionMap[selectedForm.id]
    : undefined;
  const selectedCanSubmit = selectedFormPermissions
    ? selectedFormPermissions.canAdd
    : false;

  const renderMetadataChip = (
    label: string,
    value?: string,
  ): React.ReactElement | null => {
    if (!value || value.trim() === "") return null;

    return (
      <span style={metadataChipStyle}>
        <span style={{ color: "#64748b", fontWeight: 700 }}>{label}</span>
        <span>{value}</span>
      </span>
    );
  };

  const renderNavFormButton = (form: IFormConfig): React.ReactElement => {
    const formType =
      getFormType(form) || (isPdfForm(form.listUrl) ? "PDF" : "SharePoint");
    const division = getDivision(form);
    const isFavorite = favoriteFormIds.indexOf(form.id) !== -1;
    const permissions = formPermissionMap[form.id];

    return (
      <div key={form.id} style={{ position: "relative", marginBottom: "9px" }}>
        <button
          onClick={() => handleSelectForm(form)}
          className={
            selectedForm?.id === form.id
              ? styles.activeFormButton
              : styles.formButton
          }
          style={{
            paddingRight: "38px",
            borderRadius: "10px",
            boxShadow:
              selectedForm?.id === form.id
                ? "0 8px 18px rgba(0, 90, 158, 0.14)"
                : "0 1px 2px rgba(15, 23, 42, 0.04)",
          }}
        >
          <span
            style={{
              display: "block",
              fontWeight: selectedForm?.id === form.id ? 800 : 700,
            }}
          >
            {form.title}
          </span>

          <span
            style={{
              display: "block",
              marginTop: "5px",
              color: selectedForm?.id === form.id ? "#075985" : "#64748b",
              fontSize: "12px",
              lineHeight: 1.35,
            }}
          >
            {formType}
            {division ? ` • ${division}` : ""}
          </span>

          {permissions && !permissions.canAdd && (
            <span
              style={{
                display: "inline-flex",
                marginTop: "6px",
                borderRadius: "999px",
                background: "#fef3c7",
                color: "#92400e",
                padding: "3px 7px",
                fontSize: "11px",
                fontWeight: 800,
              }}
            >
              Read only
            </span>
          )}
        </button>

        <button
          type="button"
          onClick={() => toggleFavorite(form.id)}
          aria-label={
            isFavorite
              ? `Remove ${form.title} from favorites`
              : `Add ${form.title} to favorites`
          }
          style={{
            position: "absolute",
            top: "10px",
            right: "9px",
            border: "none",
            background: "transparent",
            color: isFavorite ? "#f97316" : "#94a3b8",
            cursor: "pointer",
            fontSize: "18px",
            lineHeight: 1,
          }}
        >
          {isFavorite ? "★" : "☆"}
        </button>
      </div>
    );
  };

  const renderPagedNavigation = (): React.ReactElement[] => {
    const elements: React.ReactElement[] = [];
    let lastDivision = "";

    pagedForms.forEach((form) => {
      const division = getDivision(form) || "Unassigned";

      if (navView === "all" && division !== lastDivision) {
        elements.push(
          <div
            key={`division-${division}`}
            style={{
              margin: elements.length === 0 ? "4px 0 8px 0" : "18px 0 8px 0",
              padding: "8px 10px",
              borderRadius: "999px",
              background: "#e2e8f0",
              color: "#334155",
              fontSize: "12px",
              fontWeight: 800,
              letterSpacing: "0.02em",
            }}
          >
            {division}
          </div>,
        );
        lastDivision = division;
      }

      elements.push(renderNavFormButton(form));
    });

    return elements;
  };

  const renderEnterpriseFormHeader = (): React.ReactElement | null => {
    if (!selectedForm) return null;

    const formType =
      getFormType(selectedForm) || (selectedFormIsPdf ? "PDF" : "SharePoint");
    const division = getDivision(selectedForm) || "Division not set";
    const ownerTitle = getOwnerTitle(selectedForm) || "Owner not set";
    const ownerEmail = getOwnerEmail(selectedForm);
    const contactEmail = getContactEmail(selectedForm);
    const audience = getAudience(selectedForm);

    return (
      <div
        style={{
          maxWidth: "1180px",
          margin: "0 auto 22px auto",
          background: "linear-gradient(135deg, #ffffff 0%, #f8fbff 100%)",
          border: "1px solid #dbeafe",
          borderRadius: "18px",
          boxShadow: "0 14px 32px rgba(15, 23, 42, 0.10)",
          overflow: "hidden",
        }}
      >
        <div
          style={{
            height: "10px",
            background:
              "linear-gradient(90deg, #003a63 0%, #005a9e 45%, #60a5fa 100%)",
          }}
        />

        <div style={{ padding: "28px 32px 24px 32px" }}>
          <div
            style={{ display: "flex", alignItems: "flex-start", gap: "18px" }}
          >
            <div
              style={{
                width: "68px",
                height: "68px",
                borderRadius: "16px",
                background: "#003a63",
                color: "#ffffff",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontWeight: 800,
                fontSize: "18px",
                letterSpacing: "0.5px",
                boxShadow: "0 8px 18px rgba(0, 58, 99, 0.28)",
                flexShrink: 0,
              }}
            >
              FRCC
            </div>

            <div style={{ flex: 1, minWidth: 0 }}>
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "center",
                  gap: "10px",
                  marginBottom: "8px",
                }}
              >
                <span
                  style={{
                    borderRadius: "999px",
                    background: selectedCanSubmit ? "#dcfce7" : "#fef3c7",
                    color: selectedCanSubmit ? "#166534" : "#92400e",
                    padding: "5px 10px",
                    fontSize: "12px",
                    fontWeight: 800,
                  }}
                >
                  {selectedCanSubmit ? "Submit Enabled" : "Read Only"}
                </span>

                <span
                  style={{
                    borderRadius: "999px",
                    background: "#e0f2fe",
                    color: "#075985",
                    padding: "5px 10px",
                    fontSize: "12px",
                    fontWeight: 800,
                  }}
                >
                  {formType}
                </span>
              </div>

              <h1
                style={{
                  margin: "0 0 8px 0",
                  color: "#0f172a",
                  fontSize: "30px",
                  lineHeight: 1.15,
                  fontWeight: 800,
                }}
              >
                {selectedForm.title}
              </h1>

              <p
                style={{
                  margin: 0,
                  color: "#475569",
                  fontSize: "14px",
                  lineHeight: 1.6,
                }}
              >
                Use this portal to start a new request, review existing
                requests, and confirm the responsible division, owner, audience,
                and support contact for this form.
              </p>
            </div>

            <button
              type="button"
              onClick={() => toggleFavorite(selectedForm.id)}
              style={{
                border: "1px solid #cbd5e1",
                borderRadius: "999px",
                background:
                  favoriteFormIds.indexOf(selectedForm.id) !== -1
                    ? "#fff7ed"
                    : "#ffffff",
                color:
                  favoriteFormIds.indexOf(selectedForm.id) !== -1
                    ? "#c2410c"
                    : "#475569",
                padding: "9px 13px",
                fontWeight: 800,
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
              aria-label="Toggle favorite form"
            >
              {favoriteFormIds.indexOf(selectedForm.id) !== -1
                ? "★ Favorite"
                : "☆ Favorite"}
            </button>
          </div>

          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "9px",
              marginTop: "22px",
            }}
          >
            {renderMetadataChip("Division", division)}
            {renderMetadataChip("Owner", ownerTitle)}
            {audience.length > 0 &&
              renderMetadataChip("Audience", audience.join(", "))}
            {contactEmail && renderMetadataChip("Contact", contactEmail)}
            {!contactEmail &&
              ownerEmail &&
              renderMetadataChip("Owner Email", ownerEmail)}
          </div>

          {canManageFormConfiguration && !selectedFormIsPdf && (
            <div
              style={{
                marginTop: "18px",
                padding: "14px 16px",
                border: "1px dashed #93c5fd",
                borderRadius: "14px",
                background: "#eff6ff",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: "14px",
                flexWrap: "wrap",
              }}
            >
              <div>
                <div
                  style={{
                    color: "#1e3a8a",
                    fontWeight: 800,
                    fontSize: "13px",
                  }}
                >
                  Admin tools
                </div>
                <div
                  style={{
                    color: "#475569",
                    fontSize: "12px",
                    marginTop: "3px",
                  }}
                >
                  Generate or refresh the FormJson configuration from the live
                  SharePoint list schema.
                </div>
              </div>

              <button
                type="button"
                onClick={() => {
                  generateFormJson().catch((error: unknown) => {
                    setErrorMessage(
                      error instanceof Error
                        ? error.message
                        : "Failed to generate FormJson.",
                    );
                  });
                }}
                disabled={isSubmitting || isLoadingFields || isGeneratingJson}
                className={styles.secondaryButton}
              >
                {isGeneratingJson ? "Generating..." : "Generate / Refresh JSON"}
              </button>
            </div>
          )}

          {selectedForm.listUrl && (
            <div
              style={{
                marginTop: "16px",
                color: "#64748b",
                fontSize: "12px",
                wordBreak: "break-all",
              }}
            >
              Source: {selectedForm.listUrl}
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderField = (field: IField): React.ReactElement => {
    const fieldUiConfig = config.fields
      ? config.fields[field.internalName]
      : undefined;

    const label =
      config.labels && config.labels[field.internalName]
        ? config.labels[field.internalName]
        : field.title;

    const help = config.helpText
      ? config.helpText[field.internalName]
      : undefined;
    const value = formData[field.internalName];
    const readOnly = fieldUiConfig && fieldUiConfig.readOnly === true;

    return (
      <div key={field.internalName} style={{ marginBottom: "12px" }}>
        <label style={labelStyle}>
          {label}
          {field.required ||
          (config.validation &&
            config.validation[field.internalName] &&
            config.validation[field.internalName].required)
            ? " *"
            : ""}
        </label>

        {field.typeAsString === "Text" && (
          <input
            type="text"
            placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined}
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {field.typeAsString === "Note" && (
          <textarea
            placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined}
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {field.typeAsString === "Choice" && (
          <select
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          >
            <option value="">Select...</option>
            {field.choices?.map((choice: string) => (
              <option key={choice} value={choice}>
                {choice}
              </option>
            ))}
          </select>
        )}

        {field.typeAsString === "MultiChoice" && (
          <select
            multiple
            disabled={readOnly}
            style={{ ...inputStyle, minHeight: "80px" }}
            value={(value as string[]) || []}
            onChange={(event) => {
              const selectedValues = Array.from(
                event.target.selectedOptions,
              ).map((option) => option.value);
              handleChange(field.internalName, selectedValues);
            }}
          >
            {field.choices?.map((choice: string) => (
              <option key={choice} value={choice}>
                {choice}
              </option>
            ))}
          </select>
        )}

        {(field.typeAsString === "Number" ||
          field.typeAsString === "Currency") && (
          <input
            type="number"
            placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined}
            disabled={readOnly}
            style={inputStyle}
            value={value === undefined ? "" : String(value)}
            onChange={(event) => {
              const nextValue = event.target.value;
              handleChange(
                field.internalName,
                nextValue === "" ? undefined : Number(nextValue),
              );
            }}
          />
        )}

        {field.typeAsString === "Boolean" && (
          <input
            type="checkbox"
            disabled={readOnly}
            checked={Boolean(value)}
            onChange={(event) =>
              handleChange(field.internalName, event.target.checked)
            }
          />
        )}

        {field.typeAsString === "DateTime" && (
          <input
            type="date"
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {field.typeAsString === "User" && (
          <input
            type="email"
            placeholder={
              fieldUiConfig && fieldUiConfig.placeholder
                ? fieldUiConfig.placeholder
                : "Enter user email"
            }
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {field.typeAsString === "Lookup" && (
          <select
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(
                field.internalName,
                event.target.value === ""
                  ? undefined
                  : Number(event.target.value),
              )
            }
          >
            <option value="">Select...</option>
            {(lookupOptions[field.internalName] || []).map(
              (option: ILookupOption) => (
                <option key={option.id} value={option.id}>
                  {option.title}
                </option>
              ),
            )}
          </select>
        )}

        {field.typeAsString === "URL" && (
          <input
            type="url"
            placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined}
            disabled={readOnly}
            style={inputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {[
          "Text",
          "Note",
          "Choice",
          "MultiChoice",
          "Number",
          "Currency",
          "Boolean",
          "DateTime",
          "User",
          "Lookup",
          "URL",
        ].indexOf(field.typeAsString) === -1 && (
          <input
            type="text"
            disabled
            placeholder={`Unsupported: ${field.typeAsString}`}
            style={inputStyle}
          />
        )}

        {help && (
          <div style={{ fontSize: "12px", color: theme.text.secondary }}>
            {help}
          </div>
        )}

        {validationWarnings[field.internalName] && (
          <div style={{ color: "#8a6d00", fontSize: "12px" }}>
            {validationWarnings[field.internalName]}
          </div>
        )}

        {validationErrors[field.internalName] && (
          <div style={{ color: "#a80000", fontSize: "12px" }}>
            {validationErrors[field.internalName]}
          </div>
        )}
      </div>
    );
  };

  const getFieldWrapperClass = (field: IField): string => {
    const fieldUiConfig = config.fields
      ? config.fields[field.internalName]
      : undefined;

    if (fieldUiConfig && fieldUiConfig.width === "full")
      return styles.sectionFull;
    if (field.typeAsString === "Note" || field.typeAsString === "User")
      return styles.sectionFull;

    return "";
  };

  const renderPdfPanel = (): React.ReactElement | null => {
    if (!selectedForm || !selectedForm.listUrl) return null;

    return (
      <div
        style={{
          maxWidth: "760px",
          ...cardStyle,
        }}
      >
        <h3
          style={{
            marginTop: 0,
            marginBottom: "12px",
            color: theme.text.primary,
          }}
        >
          PDF Form
        </h3>

        <p
          style={{
            fontSize: "14px",
            lineHeight: "1.6",
            color: theme.text.primary,
            marginBottom: "18px",
          }}
        >
          This form is provided as a PDF instead of an online SharePoint form.
          Open the PDF in a new tab to download it, fill it out, print it, or
          save a completed copy.
        </p>

        <button
          type="button"
          onClick={() =>
            window.open(selectedForm.listUrl, "_blank", "noopener,noreferrer")
          }
          className={styles.submitButton}
          style={primaryButtonStyle}
        >
          Open PDF Form
        </button>

        <div
          style={{
            marginTop: "18px",
            fontSize: "12px",
            color: theme.text.secondary,
            wordBreak: "break-all",
          }}
        >
          {selectedForm.listUrl}
        </div>
      </div>
    );
  };

  const renderFormHubPanel = (): React.ReactElement | null => {
    if (!selectedForm || selectedFormIsPdf) return null;

    return (
      <div
        style={{
          maxWidth: "900px",
          ...cardStyle,
          marginBottom: "26px",
        }}
      >
        <h3
          style={{
            marginTop: 0,
            marginBottom: "8px",
            color: "#1f2937",
          }}
        >
          Form Hub
        </h3>

        <p
          style={{
            fontSize: "14px",
            lineHeight: "1.6",
            color: theme.text.primary,
            marginBottom: "16px",
          }}
        >
          {selectedCanSubmit
            ? "Start a new request here, or open the SharePoint list view to review, search, edit, or display existing submissions."
            : "You have read-only access to this form. You can open the SharePoint list view to review existing submissions, but you cannot submit a new request from the portal."}
        </p>

        {!selectedCanSubmit && (
          <div
            style={{
              marginBottom: "16px",
              padding: "10px 12px",
              border: "1px solid #f1c232",
              borderRadius: "6px",
              background: "#fff8dc",
              color: "#5f4b00",
              fontSize: "13px",
              lineHeight: "1.5",
            }}
          >
            Read-only access: submission is disabled because your account does
            not have Add Items permission on this list.
          </div>
        )}

        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "12px",
          }}
        >
          {selectedCanSubmit && (
            <button
              type="button"
              className={styles.submitButton}
              style={primaryButtonStyle}
              onClick={() => {
                setSuccessMessage("");
                setErrorMessage("");
                setFormData({});
                setValidationErrors({});
                setValidationWarnings({});
                window.scrollTo({ top: 0, behavior: "smooth" });
              }}
            >
              Start New Request
            </button>
          )}

          {selectedForm.listUrl && (
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() =>
                window.open(
                  selectedForm.listUrl,
                  "_blank",
                  "noopener,noreferrer",
                )
              }
            >
              View Existing Requests
            </button>
          )}
        </div>
      </div>
    );
  };

  const renderRowsLayout = (): React.ReactElement | null => {
    if (
      !config.layout ||
      !config.layout.rows ||
      config.layout.rows.length === 0
    )
      return null;

    const renderedFieldNames: string[] = [];

    const rowElements = config.layout.rows.map((row, index) => {
      const rowFields = row
        .map(
          (fieldName) =>
            renderableFields.filter(
              (field) => field.internalName === fieldName,
            )[0],
        )
        .filter((field): field is IField => field !== undefined);

      rowFields.forEach((field) => renderedFieldNames.push(field.internalName));

      return (
        <div
          key={`layout-row-${index}`}
          style={{
            display: "grid",
            gridTemplateColumns:
              rowFields.length > 1
                ? `repeat(${rowFields.length}, minmax(250px, 1fr))`
                : "1fr",
            gap: "18px 28px",
            marginBottom: "18px",
          }}
        >
          {rowFields.map((field) => (
            <div key={field.internalName}>{renderField(field)}</div>
          ))}
        </div>
      );
    });

    const remainingFields = renderableFields.filter(
      (field) => renderedFieldNames.indexOf(field.internalName) === -1,
    );

    if (remainingFields.length > 0) {
      rowElements.push(
        <div key="layout-row-remaining" className={styles.sectionGrid}>
          {remainingFields.map((field) => (
            <div
              key={field.internalName}
              className={getFieldWrapperClass(field)}
            >
              {renderField(field)}
            </div>
          ))}
        </div>,
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
              {renderableFields.map((field) => (
                <div
                  key={field.internalName}
                  className={getFieldWrapperClass(field)}
                >
                  {renderField(field)}
                </div>
              ))}
            </div>
          )}
        </div>,
      ];
    }

    const renderedFieldNames: string[] = [];

    const sectionElements = sections.map((section) => {
      const sectionFields = section.fields
        .map(
          (fieldName) =>
            renderableFields.filter(
              (field) => field.internalName === fieldName,
            )[0],
        )
        .filter((field): field is IField => field !== undefined);

      sectionFields.forEach((field) =>
        renderedFieldNames.push(field.internalName),
      );

      return (
        <div key={section.title} style={{ marginBottom: "28px" }}>
          <h3
            style={{
              ...sectionHeaderStyle,
              marginBottom: section.description ? "10px" : "18px",
            }}
          >
            {section.title}
          </h3>

          {section.description && (
            <div
              style={{
                fontSize: "13px",
                color: theme.text.secondary,
                marginBottom: "16px",
              }}
            >
              {section.description}
            </div>
          )}

          <div
            className={section.layout === "oneColumn" ? "" : styles.sectionGrid}
          >
            {sectionFields.map((field) => (
              <div
                key={field.internalName}
                className={
                  section.layout === "oneColumn"
                    ? styles.sectionFull
                    : getFieldWrapperClass(field)
                }
              >
                {renderField(field)}
              </div>
            ))}
          </div>
        </div>
      );
    });

    const remainingFields = renderableFields.filter(
      (field) => renderedFieldNames.indexOf(field.internalName) === -1,
    );

    if (remainingFields.length > 0) {
      sectionElements.push(
        <div key="other-fields" style={{ marginBottom: "28px" }}>
          <h3
            style={{
              ...sectionHeaderStyle,
              marginBottom: "18px",
            }}
          >
            Other Information
          </h3>

          <div className={styles.sectionGrid}>
            {remainingFields.map((field) => (
              <div
                key={field.internalName}
                className={getFieldWrapperClass(field)}
              >
                {renderField(field)}
              </div>
            ))}
          </div>
        </div>,
      );
    }

    return sectionElements;
  };

  return (
    <div
      style={{
        background: "#eef3f8",
        fontFamily: theme.app.fontFamily,
        minHeight: "100vh",
      }}
    >
      <div
        style={{
          maxWidth: "1500px",
          margin: "0 auto",
          padding: "22px 34px 0 34px",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: "18px",
            padding: "20px 24px",
            marginBottom: "18px",
            borderRadius: "18px",
            background:
              "linear-gradient(135deg, #002e5d 0%, #005a9e 62%, #0ea5e9 100%)",
            color: "#ffffff",
            boxShadow: "0 16px 34px rgba(15, 23, 42, 0.16)",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
            <div
              style={{
                width: "58px",
                height: "58px",
                borderRadius: "16px",
                background: "#ffffff",
                color: "#003a63",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontWeight: 900,
                fontSize: "16px",
                letterSpacing: "0.5px",
              }}
            >
              FRCC
            </div>

            <div>
              <div
                style={{ fontSize: "26px", fontWeight: 900, lineHeight: 1.1 }}
              >
                FRCC Forms Portal
              </div>
              <div
                style={{ marginTop: "6px", fontSize: "14px", opacity: 0.92 }}
              >
                Centralized, permission-aware forms platform for institutional
                requests
              </div>
            </div>
          </div>

          <div
            style={{
              display: "flex",
              gap: "10px",
              flexWrap: "wrap",
              justifyContent: "flex-end",
            }}
          >
            <span
              style={{
                border: "1px solid rgba(255,255,255,0.45)",
                borderRadius: "999px",
                padding: "7px 11px",
                fontSize: "12px",
                fontWeight: 800,
              }}
            >
              Governed
            </span>
            <span
              style={{
                border: "1px solid rgba(255,255,255,0.45)",
                borderRadius: "999px",
                padding: "7px 11px",
                fontSize: "12px",
                fontWeight: 800,
              }}
            >
              Scalable 150+ Forms
            </span>
            <span
              style={{
                border: "1px solid rgba(255,255,255,0.45)",
                borderRadius: "999px",
                padding: "7px 11px",
                fontSize: "12px",
                fontWeight: 800,
              }}
            >
              SharePoint Security
            </span>
          </div>
        </div>
      </div>

      <div
        className={styles.portalLayout}
        style={{
          background: "#eef3f8",
          fontFamily: theme.app.fontFamily,
          minHeight: "calc(100vh - 140px)",
          maxWidth: "1500px",
          margin: "0 auto",
          borderRadius: "18px 18px 0 0",
          overflow: "hidden",
        }}
      >
        <div className={styles.leftNav} style={{ background: "#f8fafc" }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "10px",
              marginBottom: "18px",
            }}
          >
            <div
              style={{
                width: "42px",
                height: "42px",
                borderRadius: "12px",
                background: "#003a63",
                color: "#ffffff",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontWeight: 800,
                fontSize: "13px",
                letterSpacing: "0.4px",
              }}
            >
              FRCC
            </div>

            <div>
              <h3 style={{ margin: 0 }}>Forms Portal</h3>
              <div
                style={{ fontSize: "12px", color: "#64748b", marginTop: "2px" }}
              >
                {forms.length} accessible form{forms.length === 1 ? "" : "s"}
              </div>
            </div>
          </div>

          <input
            type="text"
            placeholder={`Search ${forms.length} forms by title, owner, division, audience...`}
            value={searchText}
            onChange={(event) => setSearchText(event.target.value)}
            className={styles.searchBox}
            aria-label="Search forms"
            style={{
              minHeight: "42px",
              borderRadius: "10px",
              border: "1px solid #94a3b8",
              boxShadow: "0 1px 2px rgba(15, 23, 42, 0.06)",
              marginBottom: "14px",
            }}
          />

          <div style={{ display: "grid", gap: "8px", marginBottom: "12px" }}>
            <select
              value={divisionFilter}
              onChange={(event) => setDivisionFilter(event.target.value)}
              style={filterSelectStyle}
              aria-label="Filter by division"
            >
              <option value="All">All divisions</option>
              {divisionOptions.map((division) => (
                <option key={division} value={division}>
                  {division}
                </option>
              ))}
            </select>

            <select
              value={formTypeFilter}
              onChange={(event) => setFormTypeFilter(event.target.value)}
              style={filterSelectStyle}
              aria-label="Filter by form type"
            >
              <option value="All">All form types</option>
              {formTypeOptions.map((formType) => (
                <option key={formType} value={formType}>
                  {formType}
                </option>
              ))}
            </select>

            <select
              value={audienceFilter}
              onChange={(event) => setAudienceFilter(event.target.value)}
              style={filterSelectStyle}
              aria-label="Filter by audience"
            >
              <option value="All">All audiences</option>
              {audienceOptions.map((audience) => (
                <option key={audience} value={audience}>
                  {audience}
                </option>
              ))}
            </select>
          </div>

          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "7px",
              marginBottom: "12px",
            }}
          >
            {(["all", "favorites", "recent", "popular"] as NavView[]).map(
              (view) => {
                const label =
                  view === "all"
                    ? "All"
                    : view === "favorites"
                      ? "Favorites"
                      : view === "recent"
                        ? "Recent"
                        : "Most Used";

                return (
                  <button
                    key={view}
                    type="button"
                    onClick={() => setNavView(view)}
                    style={{
                      ...navButtonBaseStyle,
                      background: navView === view ? "#005a9e" : "#ffffff",
                      color: navView === view ? "#ffffff" : "#374151",
                      borderColor: navView === view ? "#005a9e" : "#d0d7de",
                    }}
                  >
                    {label}
                  </button>
                );
              },
            )}
          </div>

          <div className={styles.searchMeta}>
            Showing {pagedForms.length} of {filteredForms.length} result
            {filteredForms.length === 1 ? "" : "s"}
          </div>

          {isLoadingForms && <div>Loading forms...</div>}

          {!isLoadingForms && errorMessage && forms.length === 0 && (
            <div>{errorMessage}</div>
          )}

          {!isLoadingForms && !errorMessage && filteredForms.length === 0 && (
            <div
              style={{ fontSize: "13px", color: "#64748b", lineHeight: 1.5 }}
            >
              No forms match the current filters.
            </div>
          )}

          {!isLoadingForms && renderPagedNavigation()}

          {filteredForms.length > NAV_PAGE_SIZE && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: "8px",
                marginTop: "14px",
              }}
            >
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={safeCurrentPage <= 1}
                onClick={() =>
                  setCurrentPage((previous) => Math.max(1, previous - 1))
                }
                style={{ padding: "7px 10px", fontSize: "12px" }}
              >
                Previous
              </button>

              <span
                style={{ fontSize: "12px", color: "#64748b", fontWeight: 700 }}
              >
                {safeCurrentPage} / {totalPages}
              </span>

              <button
                type="button"
                className={styles.secondaryButton}
                disabled={safeCurrentPage >= totalPages}
                onClick={() =>
                  setCurrentPage((previous) =>
                    Math.min(totalPages, previous + 1),
                  )
                }
                style={{ padding: "7px 10px", fontSize: "12px" }}
              >
                Next
              </button>
            </div>
          )}
        </div>

        <div className={styles.formArea} style={{ background: "#eef3f8" }}>
          {renderEnterpriseFormHeader()}

          <div
            className={styles.formPlaceholder}
            style={{
              borderRadius: "16px",
              border: "1px solid #dbeafe",
              boxShadow: "0 12px 28px rgba(15, 23, 42, 0.08)",
              maxWidth: "1180px",
            }}
          >
            {isLoadingFields && <div>Loading form fields...</div>}

            {!isLoadingFields &&
              errorMessage &&
              selectedForm &&
              !selectedFormIsPdf && (
                <div
                  style={{
                    color: "#991b1b",
                    background: "#fee2e2",
                    border: "1px solid #fecaca",
                    borderRadius: "10px",
                    padding: "11px 13px",
                    marginBottom: "14px",
                    fontSize: "13px",
                  }}
                >
                  {errorMessage}
                </div>
              )}

            {!isLoadingFields && successMessage && (
              <div
                style={{
                  color: "#166534",
                  background: "#dcfce7",
                  border: "1px solid #bbf7d0",
                  borderRadius: "10px",
                  padding: "11px 13px",
                  marginBottom: "14px",
                  fontSize: "13px",
                }}
              >
                {successMessage}
              </div>
            )}

            {!isLoadingFields && !selectedForm && (
              <div>Select a form from the left.</div>
            )}

            {!isLoadingFields &&
              selectedForm &&
              selectedFormIsPdf &&
              renderPdfPanel()}

            {!isLoadingFields &&
              selectedForm &&
              !selectedFormIsPdf &&
              renderableFields.length > 0 &&
              renderFormHubPanel()}

            {!isLoadingFields &&
              selectedForm &&
              !selectedFormIsPdf &&
              renderableFields.length === 0 && (
                <div>No editable fields found for this list.</div>
              )}

            {!isLoadingFields &&
              selectedForm &&
              !selectedFormIsPdf &&
              selectedCanSubmit &&
              renderableFields.length > 0 &&
              renderSections()}

            {!isLoadingFields &&
              selectedForm &&
              !selectedFormIsPdf &&
              !selectedCanSubmit &&
              renderableFields.length > 0 && (
                <div
                  style={{
                    ...cardStyle,
                    maxWidth: "900px",
                    marginBottom: "26px",
                  }}
                >
                  <h3 style={{ marginTop: 0, color: theme.text.primary }}>
                    Read-only access
                  </h3>
                  <p
                    style={{
                      fontSize: "14px",
                      lineHeight: "1.6",
                      color: theme.text.primary,
                      marginBottom: 0,
                    }}
                  >
                    This form is visible because you have Read access to the
                    list. New request fields and submission controls are hidden
                    because you do not have Add Items permission. Use View
                    Existing Requests above to review the list.
                  </p>
                </div>
              )}

            {!isLoadingFields &&
              selectedForm &&
              !selectedFormIsPdf &&
              selectedCanSubmit &&
              renderableFields.length > 0 && (
                <div className={styles.submitBar}>
                  {canManageFormConfiguration && (
                    <button
                      type="button"
                      onClick={() => {
                        generateFormJson().catch((error: unknown) => {
                          setErrorMessage(
                            error instanceof Error
                              ? error.message
                              : "Failed to generate FormJson.",
                          );
                        });
                      }}
                      disabled={
                        isSubmitting || isLoadingFields || isGeneratingJson
                      }
                      className={styles.secondaryButton}
                    >
                      {isGeneratingJson
                        ? "Generating..."
                        : "Generate / Refresh JSON"}
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() => {
                      handleSubmit().catch((error: unknown) => {
                        setErrorMessage(
                          error instanceof Error
                            ? error.message
                            : "Unexpected submit error.",
                        );
                      });
                    }}
                    disabled={isSubmitting || isGeneratingJson}
                    className={styles.submitButton}
                    style={primaryButtonStyle}
                  >
                    {isSubmitting ? "Submitting..." : "Submit Request"}
                  </button>
                </div>
              )}
          </div>
        </div>
      </div>
    </div>
  );
}
