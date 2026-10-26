import { reactive } from 'vue';

export type AiPermissionCategory =
  'sheetData' | 'sheetFormat' | 'sheetStructure' | 'docEdits';

export type AiPermissionCategoryMeta = {
  id: AiPermissionCategory;
  label: string;
};

const STORAGE_KEY = 'office_ai_tool_permissions';

const AI_PERMISSION_CATEGORIES: AiPermissionCategoryMeta[] = [
  { id: 'sheetData', label: 'Данные в таблице' },
  { id: 'sheetFormat', label: 'Форматирование таблицы' },
  { id: 'sheetStructure', label: 'Структура листов' },
  { id: 'docEdits', label: 'Правки в документах' },
];

const TOOL_TO_CATEGORY: Record<string, AiPermissionCategory> = {
  set_cells: 'sheetData',
  set_formulas: 'sheetData',
  copy_formulas: 'sheetData',
  clear_range: 'sheetData',
  replace_in_sheet: 'sheetData',
  sort_range: 'sheetData',
  create_filter: 'sheetData',
  remove_filter: 'sheetData',
  set_column_filter: 'sheetData',
  set_list_validation: 'sheetData',
  set_number_validation: 'sheetData',
  clear_data_validation: 'sheetData',
  format_range: 'sheetFormat',
  set_number_format: 'sheetFormat',
  set_borders: 'sheetFormat',
  merge_range: 'sheetFormat',
  set_column_widths: 'sheetFormat',
  set_row_heights: 'sheetFormat',
  insert_image: 'sheetFormat',
  add_conditional_highlight: 'sheetFormat',
  clear_conditional_formatting: 'sheetFormat',
  insert_sheet: 'sheetStructure',
  rename_sheet: 'sheetStructure',
  delete_sheet: 'sheetStructure',
  set_freeze: 'sheetStructure',
  replace_text: 'docEdits',
  replace_selection: 'docEdits',
  insert_text: 'docEdits',
  insert_paragraph: 'docEdits',
  set_heading: 'docEdits',
  set_list: 'docEdits',
  format_text: 'docEdits',
};

type PermissionsState = Record<AiPermissionCategory, boolean>;

function defaultPermissions(): PermissionsState {
  return {
    sheetData: false,
    sheetFormat: false,
    sheetStructure: false,
    docEdits: false,
  };
}

function readStored(): PermissionsState {
  const next = defaultPermissions();
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return next;
    const parsed = JSON.parse(raw) as Partial<
      Record<AiPermissionCategory | 'deleteSheet', unknown>
    >;
    AI_PERMISSION_CATEGORIES.forEach(({ id }) => {
      if (typeof parsed[id] === 'boolean') next[id] = parsed[id];
    });
    // migrate legacy deleteSheet → sheetStructure
    if (
      typeof parsed.sheetStructure !== 'boolean' &&
      parsed.deleteSheet === true
    ) {
      next.sheetStructure = true;
    }
  } catch {
    /* ignore */
  }
  return next;
}

function persist(state: PermissionsState) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    /* ignore */
  }
}

const permissions = reactive<PermissionsState>(readStored());

export function isToolAutoAllowed(name: string): boolean {
  const category = TOOL_TO_CATEGORY[name];
  if (!category) return false;
  return permissions[category];
}

function setCategoryAllowed(category: AiPermissionCategory, enabled: boolean) {
  permissions[category] = enabled;
  persist(permissions);
}

export function useAiToolPermissions() {
  return {
    permissions,
    categories: AI_PERMISSION_CATEGORIES,
    isToolAutoAllowed,
    setCategoryAllowed,
  };
}
