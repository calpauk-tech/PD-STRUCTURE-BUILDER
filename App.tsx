import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx-js-style';
import type { 
    PlandayApiCredentials, 
    Department, 
    EmployeeGroup, 
    StructureReviewItem, 
    OverwriteMode, 
    CurrentStructure 
} from './types';
import { 
    initializeService, 
    fetchDepartments, 
    fetchEmployeeGroups, 
    createDepartment, 
    deleteDepartment, 
    createEmployeeGroup, 
    deleteEmployeeGroup, 
    resetService, 
    fetchPortalInfo 
} from './services/plandayService';

// --- SVG Icons ---
const CheckIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" /></svg>);
const InfoIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>);
const CopyIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" /></svg>);
const DownloadIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>);
const AlertIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" /></svg>);
const CloudUploadIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" /></svg>);
const UserGroupIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z" /></svg>);
const ArrowRightIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14 5l7 7m0 0l-7 7m7-7H3" /></svg>);
const BuildingOfficeIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5m0 0v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" /></svg>);
const TrashIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>);
const PlusIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" /></svg>);
const RefreshIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" /></svg>);

// --- Component Modals ---

const ConfirmModal: React.FC<{ 
    isOpen: boolean; 
    onClose: () => void; 
    onConfirm: () => void; 
    title: string; 
    message: string | React.ReactNode; 
    confirmText: string; 
    cancelText: string;
    isDangerous?: boolean;
}> = ({ isOpen, onClose, onConfirm, title, message, confirmText, cancelText, isDangerous = false }) => {
    if (!isOpen) return null;
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-gray-900/60 backdrop-blur-sm p-4 transition-opacity">
             <div className="bg-white rounded-xl shadow-2xl max-w-lg w-full overflow-hidden transform transition-all scale-100 opacity-100">
                <div className="p-6">
                    <h3 className={`text-xl font-bold mb-3 ${isDangerous ? 'text-red-700' : 'text-gray-900'}`}>{title}</h3>
                    <div className="text-gray-600 mb-8 text-base leading-relaxed">{message}</div>
                    <div className="flex justify-end gap-3">
                        <button 
                            onClick={onClose}
                            className="px-5 py-2.5 rounded-lg text-gray-700 bg-gray-100 hover:bg-gray-200 font-medium transition-colors focus:ring-2 focus:ring-gray-300 focus:outline-none"
                        >
                            {cancelText}
                        </button>
                        <button 
                            onClick={() => { onConfirm(); onClose(); }}
                            className={`px-5 py-2.5 rounded-lg text-white font-bold transition-colors focus:ring-2 focus:outline-none ${isDangerous ? 'bg-red-600 hover:bg-red-700 focus:ring-red-300' : 'bg-blue-600 hover:bg-blue-700 focus:ring-blue-300'}`}
                        >
                            {confirmText}
                        </button>
                    </div>
                </div>
             </div>
        </div>
    );
};

const TokenGuideModal: React.FC<{ isOpen: boolean; onClose: () => void }> = ({ isOpen, onClose }) => {
    if (!isOpen) return null;
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-gray-900/60 backdrop-blur-sm p-4 transition-opacity">
            <div className="bg-white rounded-xl shadow-2xl max-w-2xl w-full max-h-[90vh] flex flex-col overflow-hidden">
                <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50">
                    <h3 className="text-xl font-bold text-gray-900 flex items-center gap-2">
                        <span className="text-blue-600">🔑</span> How to Get Your Planday Refresh Token
                    </h3>
                    <button onClick={onClose} className="text-gray-400 hover:text-gray-600 focus:outline-none rounded-lg p-1">
                        <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" /></svg>
                    </button>
                </div>
                <div className="p-6 overflow-y-auto space-y-4 text-gray-700 text-sm">
                    <ol className="list-decimal list-inside space-y-3 leading-relaxed">
                        <li>Log in to your Planday portal as an Administrator.</li>
                        <li>Navigate to <strong>Settings &rarr; API Access</strong> in the top-right menu.</li>
                        <li>
                            Click <span className="bg-blue-100 text-blue-800 px-1.5 py-0.5 rounded font-medium">"Connect APP"</span> and connect using the App ID below:
                            <div className="mt-2 flex items-center justify-between bg-gray-100 p-3 rounded-lg border border-gray-200 font-mono text-xs text-gray-800">
                                <span className="break-all font-semibold">e05c91de-25a0-46a2-95be-f83020dc761c</span>
                                <button 
                                    className="text-gray-500 hover:text-blue-600 ml-2 p-1 transition-colors flex items-center gap-1 text-xs font-sans font-medium" 
                                    onClick={() => navigator.clipboard.writeText('e05c91de-25a0-46a2-95be-f83020dc761c')} 
                                    title="Copy App ID"
                                >
                                    <CopyIcon className="w-4 h-4"/> Copy
                                </button>
                            </div>
                        </li>
                        <li>Authorize permissions for Employee Read, Department Create/Read/Delete, Employee Group Create/Read/Delete.</li>
                        <li>Copy the generated <strong>"Token"</strong> (Refresh Token) and paste it into Step 1.</li>
                    </ol>
                </div>
                <div className="p-4 border-t border-gray-100 bg-gray-50 flex justify-end">
                    <button onClick={onClose} className="px-6 py-2 rounded-lg text-white bg-blue-600 hover:bg-blue-700 font-bold transition-colors">
                        Got It
                    </button>
                </div>
            </div>
        </div>
    );
};

// --- Page Header & Components ---

const PageHeader: React.FC = () => (
    <div className="text-center">
        <h1 className="text-4xl font-extrabold text-gray-900 flex items-center justify-center gap-3">
            Planday Structure Builder
            <span className="bg-blue-500 text-white text-xs font-bold px-2.5 py-1 rounded-full uppercase tracking-wider">BETA</span>
        </h1>
        <p className="mt-2 text-lg text-gray-600">Import your Planday Structure in bulk via Excel file</p>
    </div>
);

const Stepper: React.FC<{ 
    current: number; 
    steps: { title: string; subtitle: string }[]; 
    onStepClick: (index: number) => void; 
}> = ({ current, steps, onStepClick }) => (
    <nav aria-label="Progress">
        <ol role="list" className="flex items-center">
            {steps.map((step, index) => {
                const isLastStep = index === steps.length - 1;
                const isCompleted = index < current || (index === current && isLastStep);
                const isCurrent = index === current;
                const canClick = index < current; 
                
                return (
                    <li key={step.title} className={`relative ${!isLastStep ? 'flex-1' : ''}`}>
                        <div 
                            className={`flex items-center text-sm font-medium ${canClick ? 'cursor-pointer group' : ''}`}
                            onClick={() => canClick && onStepClick(index)}
                        >
                            <span className={`flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-full transition-colors ${isCompleted ? 'bg-green-600 text-white' : isCurrent ? 'bg-blue-600 text-white shadow-md' : 'bg-gray-200 text-gray-600'}`}>
                                {isCompleted ? <CheckIcon className="h-6 w-6" /> : <span>{index + 1}</span>}
                            </span>
                            <div className="ml-4 hidden md:block">
                                <span className={`block text-sm font-semibold ${isCompleted ? 'text-green-600' : isCurrent ? 'text-blue-600' : 'text-gray-500'}`}>{step.title}</span>
                                <span className="block text-xs text-gray-500">{step.subtitle}</span>
                            </div>
                        </div>
                        {!isLastStep && <div className={`absolute top-5 left-10 -ml-px mt-px h-0.5 w-full ${isCompleted ? 'bg-green-600' : 'bg-gray-200'}`} />}
                    </li>
                );
            })}
        </ol>
    </nav>
);

const ProgressBar: React.FC<{ percentage: number; current?: number; total?: number; label?: string }> = ({ percentage, current, total, label = "Processing Structure Import..." }) => (
    <div className="w-full">
        <div className="flex justify-between mb-2 items-end">
            <span className="text-sm font-bold text-blue-900">{label}</span>
            {typeof current === 'number' && typeof total === 'number' && total > 0 && (
                <span className="text-xs font-semibold text-blue-700 bg-blue-100 px-3 py-1 rounded-full">
                    {current} / {total} items
                </span>
            )}
        </div>
        <div className="w-full bg-gray-200 rounded-full h-3.5 overflow-hidden shadow-inner">
            <div 
                className="bg-blue-600 h-3.5 rounded-full transition-all duration-300 ease-out" 
                style={{ width: `${Math.min(100, Math.max(0, percentage))}%` }}
            />
        </div>
    </div>
);

// --- Main App Component ---

export default function App() {
    // --- State ---
    const FIXED_CLIENT_ID = 'e05c91de-25a0-46a2-95be-f83020dc761c';

    const [step, setStep] = useState<number>(0);
    const [credentials, setCredentials] = useState<PlandayApiCredentials>({
        clientId: FIXED_CLIENT_ID,
        refreshToken: ''
    });
    const [isConnecting, setIsConnecting] = useState(false);
    const [authError, setAuthError] = useState<string | null>(null);
    const [showTokenGuide, setShowTokenGuide] = useState(false);
    const [portalInfo, setPortalInfo] = useState<any>(null);
    const [copiedAppId, setCopiedAppId] = useState(false);

    // Current Planday Structure fetched from API
    const [currentStructure, setCurrentStructure] = useState<CurrentStructure>({
        departments: [],
        employeeGroups: []
    });

    // Template options
    const [includeDepartments, setIncludeDepartments] = useState(true);
    const [includeEmployeeGroups, setIncludeEmployeeGroups] = useState(true);

    // File Import state
    const [overwriteMode, setOverwriteMode] = useState<OverwriteMode>('keep');
    const [showOverwriteConfirm, setShowOverwriteConfirm] = useState(false);
    const [uploadedFileName, setUploadedFileName] = useState<string | null>(null);
    const [parsedHeaders, setParsedHeaders] = useState<string[]>([]);
    const [parsedRows, setParsedRows] = useState<Record<string, any>[]>([]);

    // Mapping state
    const [deptColumn, setDeptColumn] = useState<string>('');
    const [groupColumn, setGroupColumn] = useState<string>('');

    // Review state
    const [reviewItems, setReviewItems] = useState<StructureReviewItem[]>([]);
    const [reviewTab, setReviewTab] = useState<'all' | 'create' | 'skip' | 'delete'>('all');
    const [reviewSearch, setReviewSearch] = useState('');
    const [newItemName, setNewItemName] = useState('');
    const [newItemType, setNewItemType] = useState<'department' | 'employee_group'>('department');

    // Execution state
    const [isExecuting, setIsExecuting] = useState(false);
    const [executionProgress, setExecutionProgress] = useState({ current: 0, total: 0, label: '' });
    const [executionCompleted, setExecutionCompleted] = useState(false);
    const [showExecuteConfirm, setShowExecuteConfirm] = useState(false);

    // Initial credentials load
    useEffect(() => {
        const saved = sessionStorage.getItem('plandayCredentials');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                if (parsed.refreshToken) {
                    setCredentials({
                        clientId: FIXED_CLIENT_ID,
                        refreshToken: parsed.refreshToken
                    });
                }
            } catch { }
        }
    }, []);

    // Helper: auto-detect column matching
    const autoDetectColumns = (headers: string[]) => {
        let detectedDept = '';
        let detectedGroup = '';

        for (const h of headers) {
            const clean = h.trim().toLowerCase();
            if (!detectedDept && (clean.includes('department') || clean.includes('afdeling') || clean.includes('avdelning') || clean.includes('abteilung') || clean === 'dept')) {
                detectedDept = h;
            }
            if (!detectedGroup && (clean.includes('group') || clean.includes('rolle') || clean.includes('stilling') || clean.includes('position') || clean.includes('employee group') || clean === 'role' || clean === 'groups')) {
                detectedGroup = h;
            }
        }

        // Default fallbacks if standard 2-column template
        if (!detectedDept && headers.length > 0 && includeDepartments) {
            detectedDept = headers[0];
        }
        if (!detectedGroup && headers.length > 1 && includeEmployeeGroups) {
            detectedGroup = headers[1];
        } else if (!detectedGroup && headers.length === 1 && !includeDepartments && includeEmployeeGroups) {
            detectedGroup = headers[0];
        }

        setDeptColumn(detectedDept);
        setGroupColumn(detectedGroup);
    };

    // --- Actions ---

    // Step 1: Connect
    const handleConnect = async () => {
        if (!credentials.refreshToken.trim()) {
            setAuthError("Please enter your Planday Refresh Token.");
            return;
        }

        setIsConnecting(true);
        setAuthError(null);

        const credsToUse = {
            clientId: FIXED_CLIENT_ID,
            refreshToken: credentials.refreshToken.trim()
        };

        try {
            initializeService(credsToUse);
            const [depts, groups, info] = await Promise.all([
                fetchDepartments(),
                fetchEmployeeGroups(),
                fetchPortalInfo().catch(() => null)
            ]);

            setCurrentStructure({
                departments: depts,
                employeeGroups: groups
            });
            setPortalInfo(info);

            sessionStorage.setItem('plandayCredentials', JSON.stringify(credsToUse));
            setStep(1); // Move to Step 2
        } catch (err: any) {
            setAuthError(err.message || "Failed to connect to Planday API.");
        } finally {
            setIsConnecting(false);
        }
    };

    // Step 2: Download Current Structure Excel
    const handleDownloadCurrentStructure = () => {
        const wb = XLSX.utils.book_new();
        
        // Prepare rows based on chosen options
        const maxLen = Math.max(
            includeDepartments ? currentStructure.departments.length : 0,
            includeEmployeeGroups ? currentStructure.employeeGroups.length : 0
        );

        const data: any[] = [];
        
        // Build Header
        const headerRow: any = {};
        if (includeDepartments && includeEmployeeGroups) {
            headerRow.A = "Departments";
            headerRow.B = "Employee Groups";
        } else if (includeDepartments) {
            headerRow.A = "Departments";
        } else if (includeEmployeeGroups) {
            headerRow.A = "Employee Groups";
        }

        for (let i = 0; i < Math.max(1, maxLen); i++) {
            const row: any = {};
            if (includeDepartments && includeEmployeeGroups) {
                row["Departments"] = currentStructure.departments[i]?.name || "";
                row["Employee Groups"] = currentStructure.employeeGroups[i]?.name || "";
            } else if (includeDepartments) {
                row["Departments"] = currentStructure.departments[i]?.name || "";
            } else if (includeEmployeeGroups) {
                row["Employee Groups"] = currentStructure.employeeGroups[i]?.name || "";
            }
            data.push(row);
        }

        const ws = XLSX.utils.json_to_sheet(data);

        // Styling
        const headerStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "1E40AF" } }, // Deep Blue
            alignment: { horizontal: "center" }
        };

        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
            if (ws[cellAddress]) {
                ws[cellAddress].s = headerStyle;
            }
        }

        ws['!cols'] = [
            { wch: 35 },
            { wch: 35 }
        ];

        XLSX.utils.book_append_sheet(wb, ws, "Planday Structure");
        XLSX.writeFile(wb, `Planday_Structure_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    // Step 2: Download Blank Excel Template
    const handleDownloadBlankTemplate = () => {
        const wb = XLSX.utils.book_new();
        const headers: string[] = [];

        if (includeDepartments && includeEmployeeGroups) {
            headers.push("Departments", "Employee Groups");
        } else if (includeDepartments) {
            headers.push("Departments");
        } else if (includeEmployeeGroups) {
            headers.push("Employee Groups");
        }

        const sampleRows = [
            headers.reduce((acc, h) => {
                acc[h] = h === "Departments" ? "e.g. Sales" : "e.g. Sales Associate";
                return acc;
            }, {} as Record<string, string>)
        ];

        const ws = XLSX.utils.json_to_sheet(sampleRows);

        const headerStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "2563EB" } },
            alignment: { horizontal: "center" }
        };

        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
            if (ws[cellAddress]) {
                ws[cellAddress].s = headerStyle;
            }
        }

        ws['!cols'] = [{ wch: 35 }, { wch: 35 }];

        XLSX.utils.book_append_sheet(wb, ws, "Structure Template");
        XLSX.writeFile(wb, `Planday_Structure_Template.xlsx`);
    };

    // Step 2: File Upload Parsing
    const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        setUploadedFileName(file.name);
        const reader = new FileReader();

        reader.onload = (evt) => {
            try {
                const bstr = evt.target?.result;
                const wb = XLSX.read(bstr, { type: 'binary' });
                const wsname = wb.SheetNames[0];
                const ws = wb.Sheets[wsname];

                const rawJson: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                if (!rawJson || rawJson.length === 0) {
                    alert("The uploaded file appears to be empty.");
                    return;
                }

                // Header is first row
                const headers: string[] = (rawJson[0] || []).map((h: any, i: number) => h ? String(h).trim() : `Column ${i + 1}`);
                const rowsData: Record<string, any>[] = [];

                for (let r = 1; r < rawJson.length; r++) {
                    const row = rawJson[r];
                    if (!row || row.length === 0) continue;
                    const rowObj: Record<string, any> = {};
                    let hasValue = false;
                    headers.forEach((h, colIdx) => {
                        const val = row[colIdx] !== undefined && row[colIdx] !== null ? String(row[colIdx]).trim() : '';
                        rowObj[h] = val;
                        if (val) hasValue = true;
                    });
                    if (hasValue) rowsData.push(rowObj);
                }

                setParsedHeaders(headers);
                setParsedRows(rowsData);
                autoDetectColumns(headers);

                // Move to Mapping step
                setStep(2);
            } catch (err) {
                alert("Failed to parse Excel file. Please ensure it is a valid .xlsx, .xls, or .csv file.");
            }
        };

        reader.readAsBinaryString(file);
    };

    // Step 3 -> 4: Process File Rows and Generate Review Items
    const generateReviewList = () => {
        const existingDeptNames = new Set(currentStructure.departments.map(d => d.name.trim().toLowerCase()));
        const existingGroupNames = new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));

        const newDeptsSet = new Set<string>();
        const newGroupsSet = new Set<string>();

        parsedRows.forEach(row => {
            if (deptColumn && row[deptColumn]) {
                const val = String(row[deptColumn]).trim();
                if (val) newDeptsSet.add(val);
            }
            if (groupColumn && row[groupColumn]) {
                const val = String(row[groupColumn]).trim();
                if (val) newGroupsSet.add(val);
            }
        });

        const items: StructureReviewItem[] = [];

        // If Overwrite mode: mark all existing Planday structure for deletion
        if (overwriteMode === 'overwrite') {
            currentStructure.departments.forEach(d => {
                items.push({
                    id: d.id,
                    type: 'department',
                    name: d.name,
                    action: 'delete',
                    status: 'pending'
                });
            });
            currentStructure.employeeGroups.forEach(g => {
                items.push({
                    id: g.id,
                    type: 'employee_group',
                    name: g.name,
                    action: 'delete',
                    status: 'pending'
                });
            });
        }

        // Departments from file
        newDeptsSet.forEach(name => {
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            if (isDuplicate && overwriteMode === 'keep') {
                items.push({
                    type: 'department',
                    name,
                    action: 'skip_duplicate',
                    status: 'skipped',
                    resultMessage: 'Already exists in Planday'
                });
            } else {
                items.push({
                    type: 'department',
                    name,
                    action: 'create',
                    status: 'pending'
                });
            }
        });

        // Employee Groups from file
        newGroupsSet.forEach(name => {
            const isDuplicate = existingGroupNames.has(name.toLowerCase());
            if (isDuplicate && overwriteMode === 'keep') {
                items.push({
                    type: 'employee_group',
                    name,
                    action: 'skip_duplicate',
                    status: 'skipped',
                    resultMessage: 'Already exists in Planday'
                });
            } else {
                items.push({
                    type: 'employee_group',
                    name,
                    action: 'create',
                    status: 'pending'
                });
            }
        });

        setReviewItems(items);
        setStep(3); // Review step
    };

    // Add custom item in Review
    const handleAddCustomReviewItem = () => {
        if (!newItemName.trim()) return;
        const cleanName = newItemName.trim();

        const existingSet = newItemType === 'department'
            ? new Set(currentStructure.departments.map(d => d.name.trim().toLowerCase()))
            : new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));

        const isDup = existingSet.has(cleanName.toLowerCase());

        setReviewItems(prev => [
            ...prev,
            {
                type: newItemType,
                name: cleanName,
                action: isDup && overwriteMode === 'keep' ? 'skip_duplicate' : 'create',
                status: isDup && overwriteMode === 'keep' ? 'skipped' : 'pending',
                resultMessage: isDup && overwriteMode === 'keep' ? 'Already exists in Planday' : undefined
            }
        ]);

        setNewItemName('');
    };

    // Remove item from Review
    const handleRemoveReviewItem = (index: number) => {
        setReviewItems(prev => prev.filter((_, i) => i !== index));
    };

    // Step 4 -> 5: Execute Import API Calls
    const handleExecuteImport = async () => {
        setIsExecuting(true);
        setExecutionCompleted(false);

        const activeItems = [...reviewItems];
        const toDelete = activeItems.filter(i => i.action === 'delete');
        const toCreate = activeItems.filter(i => i.action === 'create');

        const totalOps = toDelete.length + toCreate.length;
        let processedCount = 0;

        // Clone review items array to record progress
        const updatedItems = [...activeItems];

        // 1. Process Deletions (if Overwrite mode)
        for (let i = 0; i < updatedItems.length; i++) {
            const item = updatedItems[i];
            if (item.action === 'delete' && item.id) {
                processedCount++;
                item.status = 'processing';
                setReviewItems([...updatedItems]);
                setExecutionProgress({
                    current: processedCount,
                    total: totalOps,
                    label: `Deleting ${item.type === 'department' ? 'Department' : 'Employee Group'} "${item.name}"...`
                });

                try {
                    if (item.type === 'department') {
                        await deleteDepartment(item.id);
                    } else {
                        await deleteEmployeeGroup(item.id);
                    }
                    item.status = 'success';
                    item.resultMessage = 'Deleted successfully';
                } catch (err: any) {
                    item.status = 'error';
                    item.resultMessage = err.message || 'Failed to delete';
                }
                setReviewItems([...updatedItems]);
            }
        }

        // 2. Process Creations
        for (let i = 0; i < updatedItems.length; i++) {
            const item = updatedItems[i];
            if (item.action === 'create') {
                processedCount++;
                item.status = 'processing';
                setReviewItems([...updatedItems]);
                setExecutionProgress({
                    current: processedCount,
                    total: totalOps,
                    label: `Creating ${item.type === 'department' ? 'Department' : 'Employee Group'} "${item.name}"...`
                });

                try {
                    if (item.type === 'department') {
                        const created = await createDepartment(item.name);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = `Created (ID: ${created.id})`;
                    } else {
                        const created = await createEmployeeGroup(item.name);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = `Created (ID: ${created.id})`;
                    }
                } catch (err: any) {
                    item.status = 'error';
                    item.resultMessage = err.message || 'Failed to create';
                }
                setReviewItems([...updatedItems]);
            }
        }

        setIsExecuting(false);
        setExecutionCompleted(true);

        // Refresh current Planday structure in memory
        try {
            const [depts, groups] = await Promise.all([
                fetchDepartments(),
                fetchEmployeeGroups()
            ]);
            setCurrentStructure({ departments: depts, employeeGroups: groups });
        } catch { }
    };

    // Export execution report as Excel
    const handleExportExecutionReport = () => {
        const data = reviewItems.map(item => ({
            "Type": item.type === 'department' ? 'Department' : 'Employee Group',
            "Name": item.name,
            "Action": item.action === 'create' ? 'Create' : item.action === 'delete' ? 'Delete' : 'Skip Duplicate',
            "Status": item.status === 'success' ? 'Success' : item.status === 'error' ? 'Error' : 'Skipped',
            "Planday ID": item.createdId || item.id || '-',
            "Details / Result": item.resultMessage || ''
        }));

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Import Results");
        XLSX.writeFile(wb, `Planday_Structure_Import_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    // Filter review items for table
    const filteredReviewItems = useMemo(() => {
        return reviewItems.filter(item => {
            if (reviewTab === 'create' && item.action !== 'create') return false;
            if (reviewTab === 'skip' && item.action !== 'skip_duplicate') return false;
            if (reviewTab === 'delete' && item.action !== 'delete') return false;

            if (reviewSearch) {
                const lower = reviewSearch.toLowerCase();
                return item.name.toLowerCase().includes(lower) || item.type.toLowerCase().includes(lower);
            }
            return true;
        });
    }, [reviewItems, reviewTab, reviewSearch]);

    // Review counts
    const reviewStats = useMemo(() => {
        const toCreateDept = reviewItems.filter(i => i.type === 'department' && i.action === 'create').length;
        const toCreateGroup = reviewItems.filter(i => i.type === 'employee_group' && i.action === 'create').length;
        const skipped = reviewItems.filter(i => i.action === 'skip_duplicate').length;
        const toDelete = reviewItems.filter(i => i.action === 'delete').length;
        return { toCreateDept, toCreateGroup, skipped, toDelete };
    }, [reviewItems]);

    // Steps configuration
    const stepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Template & Upload", subtitle: "Export or Import Excel" },
        { title: "Field Mapping", subtitle: "Select Columns" },
        { title: "Review & Check", subtitle: "Deduplicate & Confirm" },
        { title: "Import Structure", subtitle: "Execute API Creation" },
    ];

    return (
        <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
            {/* Top Navigation / Header */}
            <header className="bg-white border-b border-gray-200 sticky top-0 z-30 shadow-sm">
                <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex justify-between items-center">
                    <div className="flex items-center gap-3">
                        <div className="bg-blue-600 text-white p-2 rounded-xl shadow-md">
                            <BuildingOfficeIcon className="w-6 h-6" />
                        </div>
                        <div>
                            <span className="font-extrabold text-xl text-gray-900 tracking-tight">Planday Structure Builder</span>
                            <span className="ml-2 bg-blue-100 text-blue-800 text-xs font-bold px-2 py-0.5 rounded-full uppercase">BETA</span>
                        </div>
                    </div>

                    {step > 0 && (
                        <div className="flex items-center gap-4 text-sm text-gray-600">
                            {portalInfo?.portalName && (
                                <span className="bg-gray-100 border border-gray-200 px-3 py-1 rounded-full text-xs font-semibold text-gray-700">
                                    Connected: <strong>{portalInfo.portalName}</strong>
                                </span>
                            )}
                            <button 
                                onClick={() => {
                                    resetService();
                                    sessionStorage.removeItem('plandayCredentials');
                                    setStep(0);
                                }}
                                className="text-gray-500 hover:text-red-600 text-xs font-medium underline transition-colors"
                            >
                                Disconnect
                            </button>
                        </div>
                    )}
                </div>
            </header>

            {/* Main Content */}
            <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 flex flex-col gap-8">
                <PageHeader />

                {/* Progress Stepper Bar */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200">
                    <Stepper current={step} steps={stepsList} onStepClick={(s) => setStep(s)} />
                </div>

                {/* --- STEP 1: AUTHENTICATION --- */}
                {step === 0 && (
                    <div className="grid md:grid-cols-2 gap-8 items-start">
                        {/* Credentials Card */}
                        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 flex flex-col">
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">
                                Connect to Planday
                            </h2>
                            <p className="text-gray-500 text-sm mb-6">
                                Enter your Planday refresh token to connect with the App.
                            </p>

                            {authError && (
                                <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm flex items-start gap-3">
                                    <AlertIcon className="w-5 h-5 flex-shrink-0 mt-0.5" />
                                    <span>{authError}</span>
                                </div>
                            )}

                            <div className="space-y-4">
                                <div>
                                    <label className="block text-sm font-semibold text-gray-700 mb-1.5">
                                        Refresh Token
                                    </label>
                                    <input 
                                        type="password" 
                                        value={credentials.refreshToken}
                                        onChange={(e) => setCredentials(prev => ({ ...prev, refreshToken: e.target.value }))}
                                        placeholder="Enter Refresh Token"
                                        className="w-full px-4 py-3 rounded-xl border border-gray-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm bg-white shadow-sm"
                                    />
                                </div>

                                <button 
                                    onClick={handleConnect}
                                    disabled={isConnecting}
                                    className="w-full mt-2 py-3 px-6 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow transition-all flex items-center justify-center gap-2 disabled:opacity-50"
                                >
                                    {isConnecting ? (
                                        <>
                                            <svg className="animate-spin h-5 w-5 text-white" fill="none" viewBox="0 0 24 24">
                                                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                                                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                                            </svg>
                                            <span>Connecting to Planday...</span>
                                        </>
                                    ) : (
                                        <span>Connect to Planday</span>
                                    )}
                                </button>
                            </div>
                        </div>

                        {/* Instructions Card */}
                        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 flex flex-col">
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">
                                How to get your refresh token
                            </h2>
                            <p className="text-gray-500 text-sm mb-6">
                                Follow these steps to generate the necessary credentials from your Planday portal.
                            </p>

                            <ol className="space-y-4 text-sm text-gray-700">
                                <li className="flex gap-2">
                                    <span className="font-medium text-gray-900">1.</span>
                                    <span>Log in to your Planday portal</span>
                                </li>
                                <li className="flex gap-2">
                                    <span className="font-medium text-gray-900">2.</span>
                                    <span>Go to Settings &rarr; API Access</span>
                                </li>
                                <li className="space-y-2">
                                    <div className="flex gap-2">
                                        <span className="font-medium text-gray-900">3.</span>
                                        <span>Click &quot;Connect APP&quot; and connect to app:</span>
                                    </div>
                                    <div className="flex items-center justify-between bg-gray-100/80 px-4 py-3 rounded-xl border border-gray-200/60 font-mono text-sm text-gray-800">
                                        <span className="select-all font-medium text-gray-800 break-all">e05c91de-25a0-46a2-95be-f83020dc761c</span>
                                        <button 
                                            onClick={() => {
                                                navigator.clipboard.writeText('e05c91de-25a0-46a2-95be-f83020dc761c');
                                                setCopiedAppId(true);
                                                setTimeout(() => setCopiedAppId(false), 2000);
                                            }}
                                            className="text-gray-500 hover:text-blue-600 p-1 rounded-lg transition-colors ml-2 flex-shrink-0"
                                            title="Copy App ID"
                                        >
                                            {copiedAppId ? <CheckIcon className="w-4 h-4 text-green-600" /> : <CopyIcon className="w-4 h-4" />}
                                        </button>
                                    </div>
                                </li>
                                <li className="flex gap-2">
                                    <span className="font-medium text-gray-900">4.</span>
                                    <span>Authorize the app when prompted</span>
                                </li>
                                <li className="flex gap-2">
                                    <span className="font-medium text-gray-900">5.</span>
                                    <span>Copy the &quot;Token&quot; value (this is your Refresh Token)</span>
                                </li>
                            </ol>
                        </div>
                    </div>
                )}

                {/* --- STEP 2: CURRENT STRUCTURE & EXPORT / IMPORT --- */}
                {step === 1 && (
                    <div className="space-y-8">
                        {/* Current Structure Overview Banner */}
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 grid sm:grid-cols-2 lg:grid-cols-3 gap-6">
                            <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 flex items-center gap-4">
                                <div className="p-3 bg-blue-600 text-white rounded-lg">
                                    <BuildingOfficeIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-blue-900">{currentStructure.departments.length}</span>
                                    <span className="text-xs font-semibold text-blue-700 uppercase tracking-wider">Current Departments</span>
                                </div>
                            </div>

                            <div className="bg-indigo-50 p-4 rounded-xl border border-indigo-100 flex items-center gap-4">
                                <div className="p-3 bg-indigo-600 text-white rounded-lg">
                                    <UserGroupIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-indigo-900">{currentStructure.employeeGroups.length}</span>
                                    <span className="text-xs font-semibold text-indigo-700 uppercase tracking-wider">Current Employee Groups</span>
                                </div>
                            </div>

                            <div className="bg-gray-50 p-4 rounded-xl border border-gray-200 flex flex-col justify-center sm:col-span-2 lg:col-span-1">
                                <span className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-1">Planday Portal</span>
                                <span className="text-sm font-bold text-gray-800 truncate">{portalInfo?.portalName || 'Connected Planday Account'}</span>
                            </div>
                        </div>

                        {/* Split Action Panels */}
                        <div className="grid md:grid-cols-2 gap-8">
                            {/* Panel A: Template Generator / Structure Export */}
                            <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 flex flex-col justify-between">
                                <div>
                                    <div className="flex items-center gap-3 mb-4">
                                        <div className="p-2 bg-green-100 text-green-700 rounded-lg">
                                            <DownloadIcon className="w-5 h-5" />
                                        </div>
                                        <h3 className="text-xl font-bold text-gray-900">1. Download Template or Structure</h3>
                                    </div>
                                    <p className="text-sm text-gray-600 mb-6">
                                        Select which structure elements to include in your Excel file. Download either your existing Planday structure or a clean blank template.
                                    </p>

                                    {/* Selection Controls */}
                                    <div className="space-y-3 bg-gray-50 p-4 rounded-xl border border-gray-200 mb-6">
                                        <span className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-2">Include in Excel file:</span>
                                        <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                            <input 
                                                type="checkbox" 
                                                checked={includeDepartments} 
                                                onChange={(e) => setIncludeDepartments(e.target.checked)}
                                                className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                                            />
                                            <span>Departments (Column A)</span>
                                        </label>
                                        <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                            <input 
                                                type="checkbox" 
                                                checked={includeEmployeeGroups} 
                                                onChange={(e) => setIncludeEmployeeGroups(e.target.checked)}
                                                className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                                            />
                                            <span>Employee Groups ({includeDepartments ? 'Column B' : 'Column A'})</span>
                                        </label>
                                    </div>
                                </div>

                                <div className="space-y-3">
                                    <button 
                                        onClick={handleDownloadCurrentStructure}
                                        disabled={!includeDepartments && !includeEmployeeGroups}
                                        className="w-full py-3 px-4 bg-green-600 hover:bg-green-700 text-white font-bold rounded-xl shadow-sm hover:shadow transition-all flex items-center justify-center gap-2 text-sm disabled:opacity-50"
                                    >
                                        <DownloadIcon className="w-4 h-4" />
                                        <span>Download Current Structure Excel</span>
                                    </button>

                                    <button 
                                        onClick={handleDownloadBlankTemplate}
                                        disabled={!includeDepartments && !includeEmployeeGroups}
                                        className="w-full py-2.5 px-4 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 font-bold rounded-xl transition-all flex items-center justify-center gap-2 text-sm disabled:opacity-50"
                                    >
                                        <span>Download Blank Template</span>
                                    </button>
                                </div>
                            </div>

                            {/* Panel B: Upload File for Import */}
                            <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 flex flex-col justify-between">
                                <div>
                                    <div className="flex items-center gap-3 mb-4">
                                        <div className="p-2 bg-blue-100 text-blue-700 rounded-lg">
                                            <CloudUploadIcon className="w-5 h-5" />
                                        </div>
                                        <h3 className="text-xl font-bold text-gray-900">2. Upload File for Import</h3>
                                    </div>
                                    <p className="text-sm text-gray-600 mb-6">
                                        Upload your updated Excel file (.xlsx, .xls) or CSV. You will be able to map columns in the next step.
                                    </p>

                                    {/* Existing Structure Prompt Choice */}
                                    <div className="mb-6 bg-amber-50/60 border border-amber-200 p-4 rounded-xl">
                                        <label className="block text-xs font-bold text-amber-900 uppercase tracking-wider mb-2">
                                            Existing Structure Action:
                                        </label>
                                        <div className="space-y-2 text-sm">
                                            <label className="flex items-center gap-3 cursor-pointer text-gray-800 font-medium">
                                                <input 
                                                    type="radio" 
                                                    name="overwriteMode" 
                                                    value="keep"
                                                    checked={overwriteMode === 'keep'} 
                                                    onChange={() => setOverwriteMode('keep')}
                                                    className="w-4 h-4 text-blue-600 focus:ring-blue-500"
                                                />
                                                <span>Keep existing structure (Add new items only)</span>
                                            </label>
                                            <label className="flex items-center gap-3 cursor-pointer text-red-800 font-medium">
                                                <input 
                                                    type="radio" 
                                                    name="overwriteMode" 
                                                    value="overwrite"
                                                    checked={overwriteMode === 'overwrite'} 
                                                    onChange={() => {
                                                        setOverwriteMode('overwrite');
                                                        setShowOverwriteConfirm(true);
                                                    }}
                                                    className="w-4 h-4 text-red-600 focus:ring-red-500"
                                                />
                                                <span className="text-red-700 font-semibold">Delete / Overwrite existing structure</span>
                                            </label>
                                        </div>
                                    </div>
                                </div>

                                {/* Drag and Drop / File Input Dropzone */}
                                <div className="relative border-2 border-dashed border-gray-300 hover:border-blue-500 rounded-2xl p-6 text-center transition-colors bg-gray-50/50 hover:bg-blue-50/30 group">
                                    <input 
                                        type="file" 
                                        accept=".xlsx,.xls,.csv" 
                                        onChange={handleFileUpload}
                                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                                    />
                                    <CloudUploadIcon className="w-10 h-10 text-gray-400 group-hover:text-blue-600 mx-auto mb-2 transition-colors" />
                                    <span className="block text-sm font-bold text-gray-800">
                                        Click or Drag & Drop Excel File
                                    </span>
                                    <span className="block text-xs text-gray-500 mt-1">
                                        Supports .xlsx, .xls, .csv
                                    </span>
                                </div>
                            </div>
                        </div>

                        {/* Collapsible View of Current Structure */}
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200">
                            <h4 className="text-md font-bold text-gray-900 mb-4 flex items-center gap-2">
                                <span>📋</span> Current Structure Preview
                            </h4>
                            <div className="grid md:grid-cols-2 gap-6 text-sm">
                                {/* Departments List */}
                                <div className="border border-gray-200 rounded-xl overflow-hidden">
                                    <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200 flex justify-between items-center">
                                        <span>Departments</span>
                                        <span className="text-xs bg-blue-100 text-blue-800 px-2 py-0.5 rounded-full">{currentStructure.departments.length}</span>
                                    </div>
                                    <div className="max-h-48 overflow-y-auto p-3 space-y-1 bg-white">
                                        {currentStructure.departments.length > 0 ? (
                                            currentStructure.departments.map(d => (
                                                <div key={d.id} className="py-1 px-2 hover:bg-gray-50 rounded text-gray-700 text-xs flex justify-between">
                                                    <span>{d.name}</span>
                                                    <span className="text-gray-400">ID: {d.id}</span>
                                                </div>
                                            ))
                                        ) : (
                                            <span className="text-xs text-gray-400 italic">No departments currently in Planday</span>
                                        )}
                                    </div>
                                </div>

                                {/* Employee Groups List */}
                                <div className="border border-gray-200 rounded-xl overflow-hidden">
                                    <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200 flex justify-between items-center">
                                        <span>Employee Groups</span>
                                        <span className="text-xs bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-full">{currentStructure.employeeGroups.length}</span>
                                    </div>
                                    <div className="max-h-48 overflow-y-auto p-3 space-y-1 bg-white">
                                        {currentStructure.employeeGroups.length > 0 ? (
                                            currentStructure.employeeGroups.map(g => (
                                                <div key={g.id} className="py-1 px-2 hover:bg-gray-50 rounded text-gray-700 text-xs flex justify-between">
                                                    <span>{g.name}</span>
                                                    <span className="text-gray-400">ID: {g.id}</span>
                                                </div>
                                            ))
                                        ) : (
                                            <span className="text-xs text-gray-400 italic">No employee groups currently in Planday</span>
                                        )}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                )}

                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <div className="flex justify-between items-center mb-2">
                                <h2 className="text-2xl font-bold text-gray-900">Map File Columns</h2>
                                <span className="text-xs font-semibold text-gray-500 bg-gray-100 px-3 py-1 rounded-full">
                                    File: <strong>{uploadedFileName}</strong> ({parsedRows.length} rows)
                                </span>
                            </div>
                            <p className="text-sm text-gray-600">
                                Select which columns from your uploaded file represent Departments and Employee Groups.
                            </p>
                        </div>

                        {/* Selectors */}
                        <div className="grid md:grid-cols-2 gap-6 bg-gray-50 p-6 rounded-2xl border border-gray-200">
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">
                                    Departments Column:
                                </label>
                                <select 
                                    value={deptColumn} 
                                    onChange={(e) => setDeptColumn(e.target.value)}
                                    className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent font-medium"
                                >
                                    <option value="">-- None / Do Not Import --</option>
                                    {parsedHeaders.map(h => (
                                        <option key={h} value={h}>{h}</option>
                                    ))}
                                </select>
                            </div>

                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">
                                    Employee Groups Column:
                                </label>
                                <select 
                                    value={groupColumn} 
                                    onChange={(e) => setGroupColumn(e.target.value)}
                                    className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent font-medium"
                                >
                                    <option value="">-- None / Do Not Import --</option>
                                    {parsedHeaders.map(h => (
                                        <option key={h} value={h}>{h}</option>
                                    ))}
                                </select>
                            </div>
                        </div>

                        {/* Mapped Data Live Preview Table */}
                        <div>
                            <h3 className="text-md font-bold text-gray-900 mb-3">Preview Mapped Data (First 10 Rows)</h3>
                            <div className="border border-gray-200 rounded-xl overflow-x-auto shadow-sm">
                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-3 w-12 text-center">#</th>
                                            <th className="p-3">Department (Mapped: {deptColumn || 'None'})</th>
                                            <th className="p-3">Employee Group (Mapped: {groupColumn || 'None'})</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {parsedRows.slice(0, 10).map((row, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50/80">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-gray-800">
                                                    {deptColumn && row[deptColumn] ? row[deptColumn] : <span className="text-gray-300 italic">-</span>}
                                                </td>
                                                <td className="p-3 font-medium text-gray-800">
                                                    {groupColumn && row[groupColumn] ? row[groupColumn] : <span className="text-gray-300 italic">-</span>}
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* Actions */}
                        <div className="flex justify-between items-center pt-4 border-t border-gray-100">
                            <button 
                                onClick={() => setStep(1)}
                                className="px-5 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl transition-colors text-sm"
                            >
                                Back
                            </button>

                            <button 
                                onClick={generateReviewList}
                                disabled={!deptColumn && !groupColumn}
                                className="px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow-md transition-all text-sm flex items-center gap-2 disabled:opacity-50"
                            >
                                <span>Continue to Review</span>
                                <ArrowRightIcon className="w-4 h-4" />
                            </button>
                        </div>
                    </div>
                )}

                {/* --- STEP 4: REVIEW & DUPLICATE CHECK --- */}
                {step === 3 && (
                    <div className="space-y-6">
                        {/* Overview Stats */}
                        <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
                            <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">New Departments</span>
                                <span className="text-2xl font-extrabold text-blue-600 mt-1 block">{reviewStats.toCreateDept}</span>
                            </div>

                            <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">New Employee Groups</span>
                                <span className="text-2xl font-extrabold text-indigo-600 mt-1 block">{reviewStats.toCreateGroup}</span>
                            </div>

                            <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">Duplicates Skipped</span>
                                <span className="text-2xl font-extrabold text-amber-600 mt-1 block">{reviewStats.skipped}</span>
                            </div>

                            <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">Items To Delete</span>
                                <span className="text-2xl font-extrabold text-red-600 mt-1 block">{reviewStats.toDelete}</span>
                            </div>
                        </div>

                        {/* Add Custom Quick Item */}
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 flex flex-wrap items-center gap-3">
                            <span className="text-sm font-bold text-gray-800 shrink-0">Quick Add Item:</span>
                            <input 
                                type="text" 
                                value={newItemName} 
                                onChange={(e) => setNewItemName(e.target.value)}
                                placeholder="Enter name..."
                                className="flex-1 min-w-[200px] px-3.5 py-2 border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white"
                            />
                            <select 
                                value={newItemType} 
                                onChange={(e) => setNewItemType(e.target.value as any)}
                                className="px-3.5 py-2 border border-gray-300 rounded-xl text-sm bg-white font-medium"
                            >
                                <option value="department">Department</option>
                                <option value="employee_group">Employee Group</option>
                            </select>
                            <button 
                                onClick={handleAddCustomReviewItem}
                                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl text-sm flex items-center gap-1 shadow-sm"
                            >
                                <PlusIcon className="w-4 h-4" /> Add
                            </button>
                        </div>

                        {/* Review Table Container */}
                        <div className="bg-white rounded-2xl shadow-sm border border-gray-200 overflow-hidden">
                            {/* Table Controls Header */}
                            <div className="p-6 border-b border-gray-200 flex flex-wrap justify-between items-center gap-4 bg-gray-50/50">
                                <div className="flex gap-2">
                                    <button 
                                        onClick={() => setReviewTab('all')}
                                        className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-colors ${reviewTab === 'all' ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                    >
                                        All ({reviewItems.length})
                                    </button>
                                    <button 
                                        onClick={() => setReviewTab('create')}
                                        className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-colors ${reviewTab === 'create' ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                    >
                                        To Create ({reviewStats.toCreateDept + reviewStats.toCreateGroup})
                                    </button>
                                    <button 
                                        onClick={() => setReviewTab('skip')}
                                        className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-colors ${reviewTab === 'skip' ? 'bg-amber-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                    >
                                        Skipped ({reviewStats.skipped})
                                    </button>
                                    {reviewStats.toDelete > 0 && (
                                        <button 
                                            onClick={() => setReviewTab('delete')}
                                            className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-colors ${reviewTab === 'delete' ? 'bg-red-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                        >
                                            To Delete ({reviewStats.toDelete})
                                        </button>
                                    )}
                                </div>

                                <input 
                                    type="text" 
                                    value={reviewSearch} 
                                    onChange={(e) => setReviewSearch(e.target.value)}
                                    placeholder="Search structure..."
                                    className="px-3.5 py-1.5 border border-gray-300 rounded-xl text-xs bg-white w-60"
                                />
                            </div>

                            {/* Table */}
                            <div className="overflow-x-auto">
                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-4 w-12 text-center">#</th>
                                            <th className="p-4">Type</th>
                                            <th className="p-4">Name</th>
                                            <th className="p-4">Action</th>
                                            <th className="p-4">Status / Details</th>
                                            <th className="p-4 text-center w-16">Remove</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {filteredReviewItems.length > 0 ? (
                                            filteredReviewItems.map((item, idx) => (
                                                <tr key={idx} className="hover:bg-gray-50/80 transition-colors">
                                                    <td className="p-4 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                    <td className="p-4">
                                                        <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${item.type === 'department' ? 'bg-blue-100 text-blue-800' : 'bg-indigo-100 text-indigo-800'}`}>
                                                            {item.type === 'department' ? 'Department' : 'Employee Group'}
                                                        </span>
                                                    </td>
                                                    <td className="p-4 font-bold text-gray-900">{item.name}</td>
                                                    <td className="p-4">
                                                        {item.action === 'create' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-green-100 text-green-800">
                                                                Create
                                                            </span>
                                                        )}
                                                        {item.action === 'skip_duplicate' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">
                                                                Skip Duplicate
                                                            </span>
                                                        )}
                                                        {item.action === 'delete' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-red-100 text-red-800">
                                                                Delete
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td className="p-4 text-xs text-gray-600">
                                                        {item.resultMessage || (item.action === 'create' ? 'Ready to create' : item.action === 'delete' ? 'Scheduled for deletion' : 'Existing item')}
                                                    </td>
                                                    <td className="p-4 text-center">
                                                        <button 
                                                            onClick={() => handleRemoveReviewItem(idx)}
                                                            className="text-gray-400 hover:text-red-600 transition-colors p-1"
                                                            title="Remove from queue"
                                                        >
                                                            <TrashIcon className="w-4 h-4" />
                                                        </button>
                                                    </td>
                                                </tr>
                                            ))
                                        ) : (
                                            <tr>
                                                <td colSpan={6} className="p-8 text-center text-gray-400 text-sm italic">
                                                    No structure items match the selected filter.
                                                </td>
                                            </tr>
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* Actions */}
                        <div className="flex justify-between items-center pt-4">
                            <button 
                                onClick={() => setStep(2)}
                                className="px-5 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl transition-colors text-sm"
                            >
                                Back to Mapping
                            </button>

                            <button 
                                onClick={() => {
                                    if (reviewItems.filter(i => i.action === 'create' || i.action === 'delete').length === 0) {
                                        alert("There are no structure items to create or delete.");
                                        return;
                                    }
                                    setShowExecuteConfirm(true);
                                }}
                                className="px-8 py-3.5 bg-green-600 hover:bg-green-700 text-white font-bold rounded-xl shadow-lg hover:shadow-xl transition-all text-sm flex items-center gap-2"
                            >
                                <span>Confirm & Import Structure</span>
                                <ArrowRightIcon className="w-5 h-5" />
                            </button>
                        </div>
                    </div>
                )}

                {/* --- STEP 5: IMPORT EXECUTION & RESULTS --- */}
                {step === 4 && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Structure Import Execution</h2>
                            <p className="text-sm text-gray-600">
                                Creating new Departments and Employee Groups in your Planday portal...
                            </p>
                        </div>

                        {/* Progress */}
                        {isExecuting && (
                            <div className="bg-blue-50/50 p-6 rounded-2xl border border-blue-100">
                                <ProgressBar 
                                    percentage={(executionProgress.current / Math.max(1, executionProgress.total)) * 100}
                                    current={executionProgress.current}
                                    total={executionProgress.total}
                                    label={executionProgress.label}
                                />
                            </div>
                        )}

                        {executionCompleted && (
                            <div className="p-4 bg-green-50 border border-green-200 rounded-xl text-green-800 text-sm flex items-center justify-between">
                                <div className="flex items-center gap-2 font-bold">
                                    <CheckIcon className="w-5 h-5 text-green-600" />
                                    <span>Import execution completed!</span>
                                </div>
                                <button 
                                    onClick={handleExportExecutionReport}
                                    className="px-4 py-2 bg-green-600 hover:bg-green-700 text-white font-bold rounded-lg text-xs transition-colors flex items-center gap-1.5"
                                >
                                    <DownloadIcon className="w-4 h-4" /> Download Report
                                </button>
                            </div>
                        )}

                        {/* Execution Results Table */}
                        <div className="border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                            <div className="bg-gray-100 px-4 py-3 font-bold text-gray-800 text-xs uppercase tracking-wider border-b border-gray-200">
                                Results Log
                            </div>
                            <div className="max-h-96 overflow-y-auto">
                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-50 text-gray-600 text-xs font-bold border-b border-gray-200 sticky top-0">
                                        <tr>
                                            <th className="p-3 w-12 text-center">#</th>
                                            <th className="p-3">Type</th>
                                            <th className="p-3">Name</th>
                                            <th className="p-3">Action</th>
                                            <th className="p-3">Status</th>
                                            <th className="p-3">Planday ID / Result</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {reviewItems.map((item, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-xs">
                                                    {item.type === 'department' ? 'Department' : 'Employee Group'}
                                                </td>
                                                <td className="p-3 font-bold text-gray-900">{item.name}</td>
                                                <td className="p-3 text-xs">
                                                    {item.action === 'create' ? 'Create' : item.action === 'delete' ? 'Delete' : 'Skip'}
                                                </td>
                                                <td className="p-3">
                                                    {item.status === 'success' && <span className="text-xs font-bold text-green-600">✓ Success</span>}
                                                    {item.status === 'error' && <span className="text-xs font-bold text-red-600">✗ Error</span>}
                                                    {item.status === 'skipped' && <span className="text-xs font-bold text-amber-600">Skipped</span>}
                                                    {item.status === 'processing' && <span className="text-xs font-bold text-blue-600 animate-pulse">Processing...</span>}
                                                    {item.status === 'pending' && <span className="text-xs text-gray-400">Pending</span>}
                                                </td>
                                                <td className="p-3 text-xs text-gray-600">
                                                    {item.resultMessage || '-'}
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* Reset / Done Button */}
                        {executionCompleted && (
                            <div className="flex justify-end pt-4 border-t border-gray-100">
                                <button 
                                    onClick={() => setStep(1)}
                                    className="px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow-md transition-all text-sm flex items-center gap-2"
                                >
                                    <RefreshIcon className="w-4 h-4" /> Start New Import
                                </button>
                            </div>
                        )}
                    </div>
                )}
            </main>

            {/* Footer */}
            <footer className="py-8 text-center text-gray-400 text-xs border-t border-gray-200 mt-auto bg-white">
                <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row justify-between items-center gap-2">
                    <div className="flex items-center gap-1">
                        <span className="font-semibold text-gray-600">Planday Structure Builder (Beta)</span>
                        <span>• Client-Side Processing</span>
                    </div>
                    <div className="text-gray-400">
                        Made with ❤️ for Planday Community
                    </div>
                </div>
            </footer>

            {/* Token Guide Modal */}
            <TokenGuideModal 
                isOpen={showTokenGuide} 
                onClose={() => setShowTokenGuide(false)} 
            />

            {/* Overwrite Confirmation Modal */}
            <ConfirmModal 
                isOpen={showOverwriteConfirm}
                onClose={() => {
                    setShowOverwriteConfirm(false);
                    setOverwriteMode('keep');
                }}
                onConfirm={() => setShowOverwriteConfirm(false)}
                title="⚠️ Danger: Overwrite Existing Structure?"
                message="Selecting Overwrite will delete ALL existing Departments and Employee Groups in your Planday account before creating the new ones from your file. Are you sure you want to proceed?"
                confirmText="Yes, Overwrite Existing"
                cancelText="Cancel, Keep Existing"
                isDangerous={true}
            />

            {/* Execute Import Confirmation Modal */}
            <ConfirmModal 
                isOpen={showExecuteConfirm}
                onClose={() => setShowExecuteConfirm(false)}
                onConfirm={() => {
                    setShowExecuteConfirm(false);
                    setStep(4); // Execution step
                    handleExecuteImport();
                }}
                title="Confirm Planday Structure Import"
                message={
                    <div className="space-y-2 text-sm">
                        <p>You are about to execute the following structure changes in Planday:</p>
                        <ul className="list-disc pl-5 font-semibold text-gray-800 space-y-1">
                            {reviewStats.toDelete > 0 && <li className="text-red-600">Delete {reviewStats.toDelete} existing structure items</li>}
                            {reviewStats.toCreateDept > 0 && <li>Create {reviewStats.toCreateDept} new Departments</li>}
                            {reviewStats.toCreateGroup > 0 && <li>Create {reviewStats.toCreateGroup} new Employee Groups</li>}
                            {reviewStats.skipped > 0 && <li className="text-amber-600">Skip {reviewStats.skipped} existing duplicate items</li>}
                        </ul>
                    </div>
                }
                confirmText="Start Import Now"
                cancelText="Go Back"
            />
        </div>
    );
}
