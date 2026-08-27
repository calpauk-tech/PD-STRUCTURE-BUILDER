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

export interface Skill {
  id: number;
  name: string;
  description?: string;
  isTimeLimited?: boolean;
}

export interface RevenueUnit {
  id: number;
  name: string;
  departmentId: number;
}

export interface Position {
  id: number;
  name: string;
}

export type StructureItemType = 'department' | 'employee_group' | 'skill' | 'position';
export type ImportAction = 'create' | 'skip_duplicate' | 'delete' | 'update';

export interface StructureReviewItem {
  id?: number; // Existing Planday ID if deleting, updating or existing
  type: StructureItemType;
  name: string;
  number?: string | null;
  action: ImportAction;
  status?: 'pending' | 'processing' | 'success' | 'error' | 'skipped';
  resultMessage?: string;
  createdId?: number;
  updatePayload?: any; // For default rates or other update data
}

export type OverwriteMode = 'keep' | 'overwrite';

export interface CurrentStructure {
  departments: Department[];
  employeeGroups: EmployeeGroup[];
  skills: Skill[];
  revenueUnits: RevenueUnit[];
}

export interface ExcelParsedSheet {
  headers: string[];
  rows: Record<string, any>[];
}
