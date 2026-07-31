export interface PlandayApiCredentials {
  clientId: string;
  refreshToken: string;
}

export interface Department {
  id: number;
  name: string;
  number?: string | null;
}

export interface EmployeeGroup {
  id: number;
  name: string;
}

export type StructureItemType = 'department' | 'employee_group';
export type ImportAction = 'create' | 'skip_duplicate' | 'delete';

export interface StructureReviewItem {
  id?: number; // Existing Planday ID if deleting or existing
  type: StructureItemType;
  name: string;
  number?: string | null;
  action: ImportAction;
  status?: 'pending' | 'processing' | 'success' | 'error' | 'skipped';
  resultMessage?: string;
  createdId?: number;
}

export type OverwriteMode = 'keep' | 'overwrite';

export interface CurrentStructure {
  departments: Department[];
  employeeGroups: EmployeeGroup[];
}

export interface ExcelParsedSheet {
  headers: string[];
  rows: Record<string, any>[];
}
