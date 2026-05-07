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
import { generateFormLayoutFromFields } from "../../../services/FormSchemaLayoutService";

import prerequisiteOverride from "../../../config/forms/prerequisite-placement-override.form.json";
import courseSubstitution from "../../../config/forms/course-substitution.form.json";
import studentWaiver from "../../../config/forms/student-waiver.form.json";

const WORKBENCH_STYLE_ID = "frcc-workbench-fix";

type FormValue = string | number | boolean | string[] | undefined;
type SectionLayout = "oneColumn" | "twoColumn" | "single-column" | "two-column";
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
  label?: string;
  description?: string;
  helpText?: string;
  placeholder?: string;
  width?: "half" | "full" | "100%";
  readOnly?: boolean;
  hidden?: boolean;
  rows?: number;
  buttonText?: string;
}

interface IValidationRule {
  required?: boolean;
  message?: string;
  warning?: boolean;
  pattern?: string;
  minLength?: number;
  maxLength?: number;
}

interface IPeopleSuggestion {
  key: string;
  displayName: string;
  email: string;
  jobTitle?: string;
  department?: string;
}

interface IAttachmentItem {
  id: string;
  file: File;
}

interface IFormJsonConfig {
  title?: string;
  description?: string;
  layout?:
    | "single-column"
    | "two-column"
    | {
        rows?: string[][];
      };
  hideFormHub?: boolean;
  hideSectionHeaders?: boolean;
  hiddenFields?: string[];
  fieldOrder?: string[];
  labels?: { [internalName: string]: string };
  helpText?: { [internalName: string]: string };
  sections?: ISectionConfig[];
  fields?: {
    [internalName: string]: IFieldUiConfig;
  };
  fieldOverrides?: {
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
    headerColor?: string;
    pageBackground?: string;
    backgroundColor?: string;
    cardBackground?: string;
    cardBackgroundColor?: string;
    cardMaxWidth?: string;
    cardBorderRadius?: string;
    cardPadding?: string;
    fontFamily?: string;
    labelFontSize?: string;
    labelColor?: string;
    descriptionColor?: string;
    inputBackground?: string;
    inputBorder?: string;
    buttonBackground?: string;
  };
  logo?: {
    type?: "text" | "image";
    text?: string;
    url?: string;
    backgroundColor?: string;
    color?: string;
    shape?: "circle" | "rounded" | "square";
  };
  intro?: {
    showUserNotice?: boolean;
    userNoticeText?: string;
  };
  submit?: {
    text?: string;
    align?: "left" | "center" | "right";
    backgroundColor?: string;
  };
  footer?: {
    showPoweredBy?: boolean;
    poweredByText?: string;
    text?: string;
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

function normalizeLayout(layout?: SectionLayout): "oneColumn" | "twoColumn" {
  if (layout === "oneColumn" || layout === "single-column") return "oneColumn";
  return "twoColumn";
}

function hasFormSpecificPresentation(config: IFormJsonConfig): boolean {
  return Boolean(
    config.title ||
      config.description ||
      config.logo ||
      config.intro ||
      config.footer ||
      config.hideFormHub ||
      config.hideSectionHeaders ||
      config.theme?.pageBackground ||
      config.theme?.backgroundColor ||
      config.theme?.cardMaxWidth ||
      config.theme?.cardPadding,
  );
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
  const [attachmentItems, setAttachmentItems] = React.useState<IAttachmentItem[]>([]);
  const [activePeopleFieldName, setActivePeopleFieldName] = React.useState<string>("");
  const [peopleSuggestions, setPeopleSuggestions] = React.useState<{
    [fieldName: string]: IPeopleSuggestion[];
  }>({});
  const attachmentInputRef = React.useRef<HTMLInputElement | null>(null);
  const peopleSearchTimeoutRef = React.useRef<number | undefined>(undefined);
  const accessCheckCacheRef = React.useRef<{
    [cacheKey: string]: IFormPermissionState;
  }>({});

  const setAccessCheckCache = (
    cacheKey: string,
    permissions: IFormPermissionState,
  ): IFormPermissionState => {
    accessCheckCacheRef.current = {
      ...accessCheckCacheRef.current,
      [cacheKey]: permissions,
    };

    return permissions;
  };

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
        return setAccessCheckCache(cacheKey, noListPermission);
      }

      const data = await response.json();
      const permissions = new SPPermission(data);

      const result: IFormPermissionState = {
        canView: permissions.hasPermission(SPPermission.viewListItems),
        canAdd: permissions.hasPermission(SPPermission.addListItems),
        canEdit: permissions.hasPermission(SPPermission.editListItems),
      };

      return setAccessCheckCache(cacheKey, result);
    } catch (error: unknown) {
      console.warn(
        `Permission check failed for ListId ${listId}. Hiding form from portal.`,
        error,
      );
      return setAccessCheckCache(cacheKey, noListPermission);
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
      return setAccessCheckCache(cacheKey, result);
    } catch (error: unknown) {
      console.warn(
        `Access check failed for URL ${url}. Hiding form from portal.`,
        error,
      );
      return setAccessCheckCache(cacheKey, noListPermission);
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
        setAttachmentItems([]);
        setPeopleSuggestions({});
        setActivePeopleFieldName("");
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
        setAttachmentItems([]);
        setPeopleSuggestions({});
        setActivePeopleFieldName("");

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

  const getFieldUiConfig = (fieldName: string): IFieldUiConfig | undefined => {
    const currentConfig = getConfig();

    return (
      (currentConfig.fieldOverrides && currentConfig.fieldOverrides[fieldName]) ||
      (currentConfig.fields && currentConfig.fields[fieldName])
    );
  };

  const getConfigLabel = (field: IField): string => {
    const currentConfig = getConfig();
    const fieldUiConfig = getFieldUiConfig(field.internalName);

    return (
      fieldUiConfig?.label ||
      (currentConfig.labels && currentConfig.labels[field.internalName]) ||
      field.title
    );
  };

  const getConfigHelpText = (field: IField): string | undefined => {
    const currentConfig = getConfig();
    const fieldUiConfig = getFieldUiConfig(field.internalName);

    return (
      fieldUiConfig?.description ||
      fieldUiConfig?.helpText ||
      (currentConfig.helpText ? currentConfig.helpText[field.internalName] : undefined)
    );
  };

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

  const searchPeople = async (
    fieldName: string,
    query: string,
  ): Promise<void> => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const endpoint = `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

    const response = await props.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose;charset=utf-8",
          "odata-version": "",
        },
        body: JSON.stringify({
          queryParams: {
            __metadata: {
              type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters",
            },
            AllowEmailAddresses: true,
            AllowMultipleEntities: false,
            AllUrlZones: false,
            MaximumEntitySuggestions: 8,
            PrincipalSource: 15,
            PrincipalType: 1,
            QueryString: query,
            Required: false,
            SharePointGroupID: 0,
            UrlZone: 0,
            UrlZoneSpecified: false,
          },
        }),
      },
    );

    if (!response.ok) {
      setPeopleSuggestions((previous) => ({
        ...previous,
        [fieldName]: [],
      }));
      return;
    }

    const data = await response.json();
    const rawResult =
      data?.d?.ClientPeoplePickerSearchUser ||
      data?.ClientPeoplePickerSearchUser ||
      "[]";

    const parsed = JSON.parse(rawResult) as Array<{
      Key?: string;
      DisplayText?: string;
      EntityData?: {
        Email?: string;
        Title?: string;
        Department?: string;
      };
    }>;

    const suggestions = parsed
      .map((item) => {
        const email = item.EntityData?.Email || "";
        return {
          key: item.Key || email || item.DisplayText || "",
          displayName: item.DisplayText || email || item.Key || "Unknown user",
          email,
          jobTitle: item.EntityData?.Title || "",
          department: item.EntityData?.Department || "",
        };
      })
      .filter((item) => item.email !== "");

    setPeopleSuggestions((previous) => ({
      ...previous,
      [fieldName]: suggestions,
    }));
  };

  const queuePeopleSearch = (fieldName: string, query: string): void => {
    setActivePeopleFieldName(fieldName);

    if (peopleSearchTimeoutRef.current !== undefined) {
      window.clearTimeout(peopleSearchTimeoutRef.current);
    }

    const trimmedQuery = query.trim();

    if (trimmedQuery.length < 2) {
      setPeopleSuggestions((previous) => ({
        ...previous,
        [fieldName]: [],
      }));
      return;
    }

    peopleSearchTimeoutRef.current = window.setTimeout(() => {
      searchPeople(fieldName, trimmedQuery).catch((error: unknown) => {
        console.warn("People picker search failed.", error);
      });
    }, 250);
  };

  const selectPeopleSuggestion = (
    fieldName: string,
    suggestion: IPeopleSuggestion,
  ): void => {
    handleChange(fieldName, suggestion.email);
    setPeopleSuggestions((previous) => ({
      ...previous,
      [fieldName]: [],
    }));
    setActivePeopleFieldName("");
  };

  const escapeODataString = (value: string): string => value.replace(/'/g, "''");

  const uploadAttachmentsToItem = async (
    listId: string,
    itemId: number,
  ): Promise<void> => {
    if (attachmentItems.length === 0) return;

    const siteUrl = props.context.pageContext.web.absoluteUrl;

    for (const attachment of attachmentItems) {
      const fileName = escapeODataString(attachment.file.name);
      const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`;
      const fileBuffer = await attachment.file.arrayBuffer();

      const response = await props.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/octet-stream",
            "odata-version": "",
          },
          body: fileBuffer,
        },
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Attachment upload failed for ${attachment.file.name}. Status: ${response.status}. Details: ${errorText}`,
        );
      }
    }
  };

  const handleAttachmentSelection = (files: FileList | null): void => {
    if (!files || files.length === 0) return;

    const selectedFiles = Array.from(files).map((file) => ({
      id: `${file.name}-${file.size}-${file.lastModified}-${Math.random().toString(36).slice(2)}`,
      file,
    }));

    setAttachmentItems((previous) => [...previous, ...selectedFiles]);

    if (attachmentInputRef.current) {
      attachmentInputRef.current.value = "";
    }
  };

  const removeAttachment = (attachmentId: string): void => {
    setAttachmentItems((previous) =>
      previous.filter((attachment) => attachment.id !== attachmentId),
    );
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

    const visibleFields = fields.filter((field) => {
      const fieldUiConfig = getFieldUiConfig(field.internalName);

      return (
        hiddenFields.indexOf(field.internalName) === -1 &&
        fieldUiConfig?.hidden !== true
      );
    });

    const sectionFieldOrder = (config.sections || []).reduce(
      (accumulator: string[], section) => accumulator.concat(section.fields),
      [],
    );

    const effectiveOrder = fieldOrder.length > 0 ? fieldOrder : sectionFieldOrder;

    if (effectiveOrder.length === 0) {
      return visibleFields;
    }

    const ordered: IField[] = [];

    effectiveOrder.forEach((fieldName) => {
      const match = visibleFields.filter(
        (field) => field.internalName === fieldName,
      )[0];
      if (match && ordered.indexOf(match) === -1) ordered.push(match);
    });

    visibleFields.forEach((field) => {
      if (ordered.indexOf(field) === -1) {
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

      const generatedJson = generateFormLayoutFromFields(fields) as IFormJsonConfig;
      const nextHiddenFields = generatedJson.hiddenFields || [];

      generatedJson.hiddenFields =
        nextHiddenFields.indexOf("Status") === -1
          ? [...nextHiddenFields, "Status"]
          : nextHiddenFields;

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

      const createdItem = await response.json();
      const createdItemId = Number(createdItem.Id || createdItem.ID);

      if (attachmentItems.length > 0) {
        if (Number.isNaN(createdItemId)) {
          throw new Error(
            "Form submitted, but the new SharePoint item ID could not be read for attachment upload.",
          );
        }

        await uploadAttachmentsToItem(selectedForm.listId, createdItemId);
      }

      setSuccessMessage(
        attachmentItems.length > 0
          ? "Form submitted successfully with attachments."
          : "Form submitted successfully.",
      );
      setFormData({});
      setAttachmentItems([]);
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
    setAttachmentItems([]);
    setPeopleSuggestions({});
    setActivePeopleFieldName("");

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
  const formSpecificPresentation = hasFormSpecificPresentation(config);
  const formPageBackground =
    config.theme?.pageBackground || config.theme?.backgroundColor || "#eef3f8";
  const formCardBackground =
    config.theme?.cardBackground || config.theme?.cardBackgroundColor || "#ffffff";
  const formCardMaxWidth = config.theme?.cardMaxWidth || "1180px";
  const formCardPadding = config.theme?.cardPadding || "24px";
  const formCardBorderRadius = config.theme?.cardBorderRadius || "16px";
  const selectedFormIsPdf = isPdfForm(selectedForm?.listUrl);
  const selectedFormPermissions = selectedForm
    ? formPermissionMap[selectedForm.id]
    : undefined;
  const selectedCanSubmit = selectedFormPermissions
    ? selectedFormPermissions.canAdd
    : false;
  const submitAlign =
    config.submit?.align === "center"
      ? "center"
      : config.submit?.align === "right"
        ? "flex-end"
        : "flex-start";
  const submitButtonText = config.submit?.text || "Submit Request";
  const submitButtonBackground =
    config.submit?.backgroundColor ||
    config.theme?.buttonBackground ||
    theme.button.primaryBackground;

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
    if (!selectedForm || formSpecificPresentation) return null;

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

  const renderConfiguredAttachmentField = (): React.ReactElement | null => {
    const fieldUiConfig = getFieldUiConfig("Attachments");

    if (fieldUiConfig?.hidden === true) return null;

    const label = fieldUiConfig?.label || "Attachments";
    const help = fieldUiConfig?.description || fieldUiConfig?.helpText;
    const buttonText = fieldUiConfig?.buttonText || "Add attachments";

    return (
      <div
        key="Attachments"
        style={{ marginBottom: formSpecificPresentation ? "30px" : "12px" }}
      >
        <label
          style={{
            ...labelStyle,
            fontSize: config.theme?.labelFontSize || labelStyle.fontSize,
            fontWeight: formSpecificPresentation ? 500 : labelStyle.fontWeight,
            color: config.theme?.labelColor || labelStyle.color,
            marginBottom: "8px",
          }}
        >
          {label}
        </label>

        {help && (
          <div
            style={{
              fontSize: formSpecificPresentation ? "13px" : "12px",
              color: config.theme?.descriptionColor || theme.text.secondary,
              lineHeight: 1.55,
              marginBottom: "10px",
            }}
          >
            {help}
          </div>
        )}

        <input
          ref={attachmentInputRef}
          type="file"
          multiple
          style={{ display: "none" }}
          onChange={(event) => handleAttachmentSelection(event.target.files)}
        />

        <button
          type="button"
          className={styles.secondaryButton}
          onClick={() => attachmentInputRef.current?.click()}
          style={{
            borderColor: formSpecificPresentation ? "#b9b9b9" : undefined,
            borderRadius: formSpecificPresentation ? "3px" : undefined,
            background: "#ffffff",
            color: "#333333",
            padding: formSpecificPresentation ? "8px 14px" : undefined,
            fontWeight: 600,
          }}
        >
          + {buttonText}
        </button>

        {attachmentItems.length > 0 && (
          <div
            style={{
              marginTop: "12px",
              display: "grid",
              gap: "8px",
            }}
          >
            {attachmentItems.map((attachment) => (
              <div
                key={attachment.id}
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: "12px",
                  padding: "8px 10px",
                  border: "1px solid #d1d5db",
                  borderRadius: "6px",
                  background: "#ffffff",
                  fontSize: "13px",
                }}
              >
                <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
                  {attachment.file.name}
                </span>
                <button
                  type="button"
                  onClick={() => removeAttachment(attachment.id)}
                  style={{
                    border: "none",
                    background: "transparent",
                    color: "#a4262c",
                    cursor: "pointer",
                    fontWeight: 700,
                  }}
                  aria-label={`Remove ${attachment.file.name}`}
                >
                  Remove
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  };

  const renderField = (field: IField): React.ReactElement => {
    const fieldUiConfig = getFieldUiConfig(field.internalName);
    const label = getConfigLabel(field);
    const help = getConfigHelpText(field);
    const value = formData[field.internalName];
    const readOnly = fieldUiConfig && fieldUiConfig.readOnly === true;
    const required =
      field.required ||
      (config.validation &&
        config.validation[field.internalName] &&
        config.validation[field.internalName].required);
    const inputBackground = config.theme?.inputBackground || inputStyle.background;
    const inputBorder =
      config.theme?.inputBorder === "none"
        ? "none"
        : config.theme?.inputBorder || inputStyle.border;
    const fieldInputStyle: React.CSSProperties = {
      ...inputStyle,
      background: inputBackground,
      border: inputBorder,
      borderRadius: formSpecificPresentation ? "3px" : inputStyle.borderRadius,
      minHeight: formSpecificPresentation ? "38px" : undefined,
      fontSize: formSpecificPresentation ? "14px" : inputStyle.fontSize,
      boxShadow: config.theme?.inputBorder === "none" ? "none" : undefined,
    };
    const fieldLabelStyle: React.CSSProperties = {
      ...labelStyle,
      fontSize: config.theme?.labelFontSize || labelStyle.fontSize,
      fontWeight: formSpecificPresentation ? 500 : labelStyle.fontWeight,
      color: config.theme?.labelColor || labelStyle.color,
      marginBottom: help ? "7px" : "8px",
    };
    const helpStyle: React.CSSProperties = {
      fontSize: formSpecificPresentation ? "13px" : "12px",
      color: config.theme?.descriptionColor || theme.text.secondary,
      lineHeight: 1.55,
      marginBottom: "10px",
    };

    return (
      <div
        key={field.internalName}
        style={{ marginBottom: formSpecificPresentation ? "30px" : "12px" }}
      >
        <label style={fieldLabelStyle}>
          {label}
          {required && <span style={{ color: "#a4262c" }}> *</span>}
        </label>

        {help && <div style={helpStyle}>{help}</div>}

        {field.typeAsString === "Text" && (
          <input
            type="text"
            placeholder={fieldUiConfig ? fieldUiConfig.placeholder : undefined}
            disabled={readOnly}
            style={fieldInputStyle}
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
            rows={fieldUiConfig?.rows || (formSpecificPresentation ? 5 : undefined)}
            style={{
              ...fieldInputStyle,
              minHeight: formSpecificPresentation ? "96px" : undefined,
              resize: "vertical",
            }}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {field.typeAsString === "Choice" && (
          <select
            disabled={readOnly}
            style={fieldInputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          >
            <option value="">
              {fieldUiConfig?.placeholder || "Select..."}
            </option>
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
            style={{ ...fieldInputStyle, minHeight: "80px" }}
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
            style={fieldInputStyle}
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
            style={fieldInputStyle}
            value={String(value || "")}
            onChange={(event) =>
              handleChange(field.internalName, event.target.value)
            }
          />
        )}

        {field.typeAsString === "User" && (
          <div style={{ position: "relative" }}>
            <input
              type="text"
              placeholder={
                fieldUiConfig && fieldUiConfig.placeholder
                  ? fieldUiConfig.placeholder
                  : "Enter a name or email address"
              }
              disabled={readOnly}
              style={fieldInputStyle}
              value={String(value || "")}
              onFocus={() => setActivePeopleFieldName(field.internalName)}
              onBlur={() => {
                window.setTimeout(() => {
                  setActivePeopleFieldName((current) =>
                    current === field.internalName ? "" : current,
                  );
                }, 180);
              }}
              onChange={(event) => {
                const nextValue = event.target.value;
                handleChange(field.internalName, nextValue);
                queuePeopleSearch(field.internalName, nextValue);
              }}
            />

            {activePeopleFieldName === field.internalName &&
              (peopleSuggestions[field.internalName] || []).length > 0 && (
                <div
                  style={{
                    position: "absolute",
                    zIndex: 1000,
                    top: "calc(100% + 4px)",
                    left: 0,
                    right: 0,
                    maxHeight: "240px",
                    overflowY: "auto",
                    border: "1px solid #c8c6c4",
                    borderRadius: "4px",
                    background: "#ffffff",
                    boxShadow: "0 8px 18px rgba(0,0,0,0.18)",
                  }}
                >
                  {(peopleSuggestions[field.internalName] || []).map(
                    (suggestion) => (
                      <button
                        key={suggestion.key}
                        type="button"
                        onMouseDown={(event) => event.preventDefault()}
                        onClick={() =>
                          selectPeopleSuggestion(field.internalName, suggestion)
                        }
                        style={{
                          width: "100%",
                          display: "flex",
                          alignItems: "center",
                          gap: "10px",
                          padding: "9px 10px",
                          border: "none",
                          borderBottom: "1px solid #f3f4f6",
                          background: "#ffffff",
                          cursor: "pointer",
                          textAlign: "left",
                        }}
                      >
                        <span
                          style={{
                            width: "30px",
                            height: "30px",
                            borderRadius: "50%",
                            background: "#e5e7eb",
                            color: "#64748b",
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontWeight: 800,
                            flexShrink: 0,
                          }}
                        >
                          {(suggestion.displayName || suggestion.email)
                            .slice(0, 1)
                            .toUpperCase()}
                        </span>
                        <span style={{ minWidth: 0 }}>
                          <span
                            style={{
                              display: "block",
                              color: "#111827",
                              fontSize: "13px",
                              fontWeight: 600,
                            }}
                          >
                            {suggestion.displayName}
                          </span>
                          <span
                            style={{
                              display: "block",
                              color: "#6b7280",
                              fontSize: "12px",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                            }}
                          >
                            {suggestion.email}
                          </span>
                        </span>
                      </button>
                    ),
                  )}
                </div>
              )}
          </div>
        )}

        {field.typeAsString === "Lookup" && (
          <select
            disabled={readOnly}
            style={fieldInputStyle}
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
            <option value="">
              {fieldUiConfig?.placeholder || "Select..."}
            </option>
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
            style={fieldInputStyle}
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
            style={fieldInputStyle}
          />
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
    const fieldUiConfig = getFieldUiConfig(field.internalName);

    if (
      fieldUiConfig &&
      (fieldUiConfig.width === "full" || fieldUiConfig.width === "100%")
    )
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
          maxWidth: config.theme?.cardMaxWidth || "900px",
          ...cardStyle,
          marginBottom: "26px",
          borderRadius: formSpecificPresentation
            ? config.theme?.cardBorderRadius || "10px"
            : cardStyle.borderRadius,
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
                setAttachmentItems([]);
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
      typeof config.layout === "string" ||
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

  const renderFormJsonHeader = (): React.ReactElement | null => {
    if (!formSpecificPresentation) return null;

    const logo = config.logo;
    const title = config.title || selectedForm?.title || "Form";
    const description = config.description;
    const userNoticeText = config.intro?.showUserNotice
      ? config.intro.userNoticeText
      : undefined;
    const logoShape = logo?.shape || "circle";

    return (
      <div style={{ marginBottom: "30px" }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "18px",
            marginBottom: description ? "22px" : "12px",
          }}
        >
          {logo && (
            <div
              style={{
                width: "54px",
                height: "54px",
                borderRadius:
                  logoShape === "circle"
                    ? "50%"
                    : logoShape === "square"
                      ? "0"
                      : "12px",
                background: logo.backgroundColor || "#002f6c",
                color: logo.color || "#ffffff",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontWeight: 800,
                fontSize: "13px",
                letterSpacing: "0.2px",
                flexShrink: 0,
                overflow: "hidden",
              }}
            >
              {logo.type === "image" && logo.url ? (
                <img
                  src={logo.url}
                  alt="Form logo"
                  style={{ width: "100%", height: "100%", objectFit: "cover" }}
                />
              ) : (
                logo.text || "FRCC"
              )}
            </div>
          )}

          <h1
            style={{
              margin: 0,
              color: config.theme?.headerColor || "#242424",
              fontSize: "28px",
              lineHeight: 1.15,
              fontWeight: 700,
            }}
          >
            {title}
          </h1>
        </div>

        {description && (
          <p
            style={{
              margin: "0 0 20px 0",
              color: config.theme?.descriptionColor || "#333333",
              fontSize: "14px",
              lineHeight: 1.55,
            }}
          >
            {description}
          </p>
        )}

        {userNoticeText && (
          <p
            style={{
              margin: "0 0 26px 0",
              color: "#666666",
              fontSize: "14px",
              lineHeight: 1.55,
            }}
          >
            {userNoticeText}
          </p>
        )}
      </div>
    );
  };

  const renderConfiguredFieldByName = (
    fieldName: string,
  ): React.ReactElement | null => {
    if (fieldName === "Attachments") return renderConfiguredAttachmentField();

    const field = renderableFields.filter(
      (candidate) => candidate.internalName === fieldName,
    )[0];

    if (!field) return null;

    return renderField(field);
  };

  const renderFormJsonFooter = (): React.ReactElement | null => {
    if (!config.footer) return null;

    return (
      <div
        style={{
          marginTop: "26px",
          paddingTop: "18px",
          borderTop: "1px solid #e5e7eb",
          color: "#666666",
          fontSize: "11px",
          lineHeight: 1.55,
        }}
      >
        {config.footer.showPoweredBy && (
          <div style={{ marginBottom: "10px", fontWeight: 600 }}>
            🧱 {config.footer.poweredByText || "Powered by Microsoft Lists"}
          </div>
        )}

        {config.footer.text && <div>{config.footer.text}</div>}
      </div>
    );
  };

  const renderSections = (): React.ReactElement[] => {
    const sections = config.sections || [];
    const defaultLayout =
      typeof config.layout === "string"
        ? normalizeLayout(config.layout)
        : "twoColumn";

    if (sections.length === 0) {
      return [
        <div key="default-section">
          {renderRowsLayout() || (
            <div
              className={
                defaultLayout === "oneColumn" ? "" : styles.sectionGrid
              }
            >
              {renderableFields.map((field) => (
                <div
                  key={field.internalName}
                  className={
                    defaultLayout === "oneColumn"
                      ? styles.sectionFull
                      : getFieldWrapperClass(field)
                  }
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

    const sectionElements = sections.map((section, sectionIndex) => {
      const sectionLayout = normalizeLayout(section.layout || defaultLayout);
      const sectionFields = section.fields
        .map((fieldName) => {
          if (fieldName === "Attachments") return undefined;

          return renderableFields.filter(
            (field) => field.internalName === fieldName,
          )[0];
        })
        .filter((field): field is IField => field !== undefined);

      section.fields.forEach((fieldName) => renderedFieldNames.push(fieldName));

      const configuredElements = section.fields
        .map((fieldName) => renderConfiguredFieldByName(fieldName))
        .filter((element): element is React.ReactElement => element !== null);

      const showSectionHeader =
        !config.hideSectionHeaders && section.title && section.title.trim() !== "";

      return (
        <div
          key={section.title || `section-${sectionIndex}`}
          style={{ marginBottom: formSpecificPresentation ? "8px" : "28px" }}
        >
          {showSectionHeader && (
            <h3
              style={{
                ...sectionHeaderStyle,
                background: `linear-gradient(90deg, ${accentColor} 0%, #006fbf 100%)`,
                marginBottom: section.description ? "10px" : "18px",
              }}
            >
              {section.title}
            </h3>
          )}

          {section.description && !config.hideSectionHeaders && (
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
            className={sectionLayout === "oneColumn" ? "" : styles.sectionGrid}
            style={
              sectionLayout === "oneColumn"
                ? undefined
                : formSpecificPresentation
                  ? { gap: "0 32px" }
                  : undefined
            }
          >
            {configuredElements.map((element, index) => {
              const field = sectionFields[index];

              return (
                <div
                  key={element.key || `configured-field-${index}`}
                  className={
                    sectionLayout === "oneColumn"
                      ? styles.sectionFull
                      : field
                        ? getFieldWrapperClass(field)
                        : styles.sectionFull
                  }
                >
                  {element}
                </div>
              );
            })}
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
          {!config.hideSectionHeaders && (
            <h3
              style={{
                ...sectionHeaderStyle,
                background: `linear-gradient(90deg, ${accentColor} 0%, #006fbf 100%)`,
                marginBottom: "18px",
              }}
            >
              Other Information
            </h3>
          )}

          <div
            className={defaultLayout === "oneColumn" ? "" : styles.sectionGrid}
          >
            {remainingFields.map((field) => (
              <div
                key={field.internalName}
                className={
                  defaultLayout === "oneColumn"
                    ? styles.sectionFull
                    : getFieldWrapperClass(field)
                }
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

        <div
          className={styles.formArea}
          style={{
            background: formSpecificPresentation ? formPageBackground : "#eef3f8",
            fontFamily: config.theme?.fontFamily || theme.app.fontFamily,
          }}
        >
          {renderEnterpriseFormHeader()}

          <div
            className={styles.formPlaceholder}
            style={{
              borderRadius: formSpecificPresentation ? formCardBorderRadius : "16px",
              border: formSpecificPresentation ? "1px solid #d7d7d7" : "1px solid #dbeafe",
              boxShadow: formSpecificPresentation
                ? "0 2px 10px rgba(0, 0, 0, 0.10)"
                : "0 12px 28px rgba(15, 23, 42, 0.08)",
              maxWidth: formSpecificPresentation ? formCardMaxWidth : "1180px",
              background: formSpecificPresentation ? formCardBackground : undefined,
              padding: formSpecificPresentation ? formCardPadding : undefined,
            }}
          >
            {renderFormJsonHeader()}
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
                <div
                  className={styles.submitBar}
                  style={{
                    justifyContent: submitAlign,
                    borderTop: formSpecificPresentation ? "none" : undefined,
                    marginTop: formSpecificPresentation ? "8px" : undefined,
                  }}
                >
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
                    style={{
                      ...primaryButtonStyle,
                      background: submitButtonBackground,
                      borderColor: submitButtonBackground,
                      borderRadius: formSpecificPresentation ? "3px" : primaryButtonStyle.borderRadius,
                    }}
                  >
                    {isSubmitting ? "Submitting..." : submitButtonText}
                  </button>
                </div>
              )}

            {!isLoadingFields &&
              selectedForm &&
              !selectedFormIsPdf &&
              selectedCanSubmit &&
              formSpecificPresentation &&
              renderFormJsonFooter()}
          </div>
        </div>
      </div>
    </div>
  );
}
