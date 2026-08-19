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
    fetchPortalInfo,
    updateEmployeeGroupDefaultRate,
    updateDepartment,
    fetchEmployeeGroupDefaultRates
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

    const [appMode, setAppMode] = useState<'import' | 'update' | 'delete'>('import');

    // Bulk Delete / Update state
    const [deleteInputMethod, setDeleteInputMethod] = useState<'text' | 'file' | 'selection'>('selection');
    const [deleteDeptText, setDeleteDeptText] = useState('');
    const [deleteGroupText, setDeleteGroupText] = useState('');
    const [deleteSelectedDepts, setDeleteSelectedDepts] = useState<Set<number>>(new Set());
    const [deleteSelectedGroups, setDeleteSelectedGroups] = useState<Set<number>>(new Set());
    const [deleteSearchTerm, setDeleteSearchTerm] = useState('');

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
    const [includeWageColumns, setIncludeWageColumns] = useState(false);
    const [parsedHeaders, setParsedHeaders] = useState<string[]>([]);
    const [parsedRows, setParsedRows] = useState<Record<string, any>[]>([]);

    // Mapping state
    const [deptColumn, setDeptColumn] = useState<string>('');
    const [groupColumn, setGroupColumn] = useState<string>('');
    const [deptNumberColumn, setDeptNumberColumn] = useState<string>('');
    const [wageTypeColumn, setWageTypeColumn] = useState<string>('');
    const [rateColumn, setRateColumn] = useState<string>('');
    const [salaryCodeColumn, setSalaryCodeColumn] = useState<string>('');
    const [salaryCode2Column, setSalaryCode2Column] = useState<string>('');

    // Review state
    const [reviewItems, setReviewItems] = useState<StructureReviewItem[]>([]);
    const [reviewTab, setReviewTab] = useState<'all' | 'create' | 'skip' | 'delete' | 'update'>('all');
    const [reviewSearch, setReviewSearch] = useState('');
    const [newItemName, setNewItemName] = useState('');
    const [newItemType, setNewItemType] = useState<'department' | 'employee_group'>('department');

    // Execution state
    const [isExecuting, setIsExecuting] = useState(false);
    const [executionProgress, setExecutionProgress] = useState({ current: 0, total: 0, label: '' });
    const [executionCompleted, setExecutionCompleted] = useState(false);
    const [showExecuteConfirm, setShowExecuteConfirm] = useState(false);

    // Employee Group Defaults
    const [applyGroupDefaults, setApplyGroupDefaults] = useState(false);
    const [globalDefaultWageType, setGlobalDefaultWageType] = useState<'HourlyRate' | 'ShiftRate' | ''>('');
    const [globalHourlyRate, setGlobalHourlyRate] = useState<number | ''>('');
    const [globalShiftRate, setGlobalShiftRate] = useState<number | ''>('');
    const [globalSalaryCode, setGlobalSalaryCode] = useState<string>('');
    const [globalSalaryCode2, setGlobalSalaryCode2] = useState<string>('');

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
    }, []);    // Helper: auto-detect column matching
    const autoDetectColumns = (headers: string[]) => {
        let detectedDept = '';
        let detectedGroup = '';
        let detectedDeptNum = '';
        let detectedWageType = '';
        let detectedRate = '';
        let detectedSalCode = '';
        let detectedSalCode2 = '';

        for (const h of headers) {
            const clean = h.trim().toLowerCase();
            
            if (!detectedDept && (clean === 'departments' || clean === 'department' || clean === 'dept' || clean === 'afdeling')) {
                detectedDept = h;
            }
            if (!detectedDeptNum && (clean.includes('department number') || clean.includes('dept number'))) {
                detectedDeptNum = h;
            }
            if (!detectedGroup && (clean.includes('employee group') || clean === 'employee groups' || clean === 'role' || clean === 'stilling')) {
                detectedGroup = h;
            }
            if (!detectedWageType && (clean.includes('wage type') || clean.includes('default wage type'))) {
                detectedWageType = h;
            }
            if (!detectedRate && (clean === 'rate' || clean === 'hourly rate' || clean === 'shift rate')) {
                detectedRate = h;
            }
            if (!detectedSalCode && (clean === 'salary code' || clean === 'salarycode')) {
                detectedSalCode = h;
            }
            if (!detectedSalCode2 && (clean.includes('salary code 2'))) {
                detectedSalCode2 = h;
            }
        }

        setDeptColumn(detectedDept);
        setGroupColumn(detectedGroup);
        setDeptNumberColumn(detectedDeptNum);
        setWageTypeColumn(detectedWageType);
        setRateColumn(detectedRate);
        setSalaryCodeColumn(detectedSalCode);
        setSalaryCode2Column(detectedSalCode2);
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
    };    // Step 2: Download Current Structure Excel
    const handleDownloadCurrentStructure = async () => {
        const wb = XLSX.utils.book_new();
        
        // Fetch wage rates if requested
        let wageRatesMap = new Map<number, any>();
        if (includeEmployeeGroups && includeWageColumns) {
            try {
                const rates = await fetchEmployeeGroupDefaultRates();
                rates.forEach(rate => {
                    wageRatesMap.set(rate.employeeGroupId, rate);
                });
            } catch (err) {
                console.error("Failed to fetch default wage rates:", err);
            }
        }

        // Prepare rows based on chosen options
        const maxLen = Math.max(
            includeDepartments ? currentStructure.departments.length : 0,
            includeEmployeeGroups ? currentStructure.employeeGroups.length : 0
        );
        const data: any[] = [];
        
        for (let i = 0; i < Math.max(1, maxLen); i++) {
            const row: any = {};
            if (includeDepartments) {
                row["Departments"] = currentStructure.departments[i]?.name || "";
                row["Department Number (optional)"] = currentStructure.departments[i]?.number || "";
            }
            if (includeEmployeeGroups) {
                const group = currentStructure.employeeGroups[i];
                row["Employee Groups"] = group?.name || "";
                if (includeWageColumns) {
                    const rateData = group ? wageRatesMap.get(group.id) : null;
                    row["Default Wage Type (HourlyRate/ShiftRate)"] = rateData?.defaultWageType || "";
                    row["Rate"] = rateData?.defaultWageType === 'HourlyRate' ? (rateData?.hourlyRate ?? "") : (rateData?.defaultWageType === 'ShiftRate' ? (rateData?.shiftRate ?? "") : "");
                    row["Salary Code"] = rateData?.salaryCode || "";
                    row["Salary Code 2 (Rare)"] = rateData?.salaryCode2 || "";
                }
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
    };    // Step 2: Download Blank Excel Template
    const handleDownloadBlankTemplate = () => {
        const wb = XLSX.utils.book_new();
        const headers: string[] = [];
        if (includeDepartments) {
            headers.push("Departments", "Department Number (optional)");
        }
        if (includeEmployeeGroups) {
            headers.push("Employee Groups");
            if (includeWageColumns) {
                headers.push("Default Wage Type (HourlyRate/ShiftRate)", "Rate", "Salary Code", "Salary Code 2 (Rare)");
            }
        }

        const sampleRows = [
            headers.reduce((acc, h) => {
                if (h === "Departments") acc[h] = "e.g. Sales";
                else if (h === "Department Number (optional)") acc[h] = "e.g. 101";
                else if (h === "Employee Groups") acc[h] = "e.g. Sales Associate";
                else if (h === "Default Wage Type (HourlyRate/ShiftRate)") acc[h] = "HourlyRate";
                else if (h === "Rate") acc[h] = "15.50";
                else if (h === "Salary Code") acc[h] = "SC1";
                else if (h === "Salary Code 2 (Rare)") acc[h] = "";
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

    // Step 3 -> 4: Process File Rows and Generate Review Items (Import)
    const generateReviewList = () => {
        const existingDeptNames = new Set(currentStructure.departments.map(d => d.name.trim().toLowerCase()));
        const existingGroupNames = new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));        const newDeptsMap = new Map<string, string | null>();
        const newGroupsMap = new Map<string, any>();
        
        parsedRows.forEach(row => {
            if (deptColumn && row[deptColumn]) {
                const val = String(row[deptColumn]).trim();
                let deptNum = null;
                if (deptNumberColumn && row[deptNumberColumn]) {
                    deptNum = String(row[deptNumberColumn]).trim();
                }
                if (val) {
                    if (!newDeptsMap.has(val) || deptNum) {
                        newDeptsMap.set(val, deptNum);
                    }
                }
            }
            if (groupColumn && row[groupColumn]) {
                const val = String(row[groupColumn]).trim();
                if (val) {
                    const groupData: any = newGroupsMap.get(val) || {};
                    if (wageTypeColumn && row[wageTypeColumn]) groupData.wageType = String(row[wageTypeColumn]).trim();
                    if (rateColumn && row[rateColumn]) groupData.rate = String(row[rateColumn]).trim();
                    if (salaryCodeColumn && row[salaryCodeColumn]) groupData.salaryCode = String(row[salaryCodeColumn]).trim();
                    if (salaryCode2Column && row[salaryCode2Column]) groupData.salaryCode2 = String(row[salaryCode2Column]).trim();
                    newGroupsMap.set(val, groupData);
                }
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
                    action: appMode === 'update' ? 'update' : 'delete',
                    status: 'pending'
                });
            });
            currentStructure.employeeGroups.forEach(g => {
                items.push({
                    id: g.id,
                    type: 'employee_group',
                    name: g.name,
                    action: appMode === 'update' ? 'update' : 'delete',
                    status: 'pending'
                });
            });
        }        // Departments from file
        newDeptsMap.forEach((deptNum, name) => {
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            const existingDept = currentStructure.departments.find(d => d.name.trim().toLowerCase() === name.toLowerCase());
            
            if (appMode === 'update') {
                if (existingDept) {
                    items.push({
                        id: existingDept.id,
                        type: 'department',
                        name,
                        number: deptNum,
                        action: 'update',
                        status: 'pending',
                        updatePayload: { name, number: deptNum },
                        resultMessage: 'Update scheduled'
                    });
                } else {
                    items.push({
                        type: 'department',
                        name,
                        action: 'skip_duplicate',
                        status: 'error',
                        resultMessage: 'Not found in Planday'
                    });
                }
            } else {
                if (isDuplicate && overwriteMode === 'keep') {
                    if (deptNum && existingDept && existingDept.number !== deptNum) {
                        items.push({
                            id: existingDept.id,
                            type: 'department',
                            name,
                            number: deptNum,
                            action: 'update',
                            status: 'pending',
                            resultMessage: 'Updating number from file'
                        });
                    } else {
                        items.push({
                            type: 'department',
                            name,
                            number: deptNum,
                            action: 'skip_duplicate',
                            status: 'skipped',
                            resultMessage: 'Already exists in Planday'
                        });
                    }
                } else {
                    items.push({
                        type: 'department',
                        name,
                        number: deptNum,
                        action: 'create',
                        status: 'pending'
                    });
                }
            }
        });

        // Employee Groups from file
        newGroupsMap.forEach((groupData, name) => {
            const isDuplicate = existingGroupNames.has(name.toLowerCase());
            const existingGroup = currentStructure.employeeGroups.find(g => g.name.trim().toLowerCase() === name.toLowerCase());
            
            let updatePayload: any = undefined;
            if (Object.keys(groupData).length > 0 || appMode === 'update') {
                updatePayload = {};
                if (groupData.wageType) updatePayload.defaultWageType = groupData.wageType;
                if (groupData.rate && groupData.wageType === 'HourlyRate') updatePayload.hourlyRate = groupData.rate;
                if (groupData.rate && groupData.wageType === 'ShiftRate') updatePayload.shiftRate = groupData.rate;
                if (groupData.salaryCode) updatePayload.salaryCode = groupData.salaryCode;
                if (groupData.salaryCode2) updatePayload.salaryCode2 = groupData.salaryCode2;
            }

            if (appMode === 'update') {
                if (existingGroup) {
                    items.push({
                        id: existingGroup.id,
                        type: 'employee_group',
                        name,
                        action: 'update',
                        status: 'pending',
                        updatePayload,
                        resultMessage: 'Update scheduled'
                    });
                } else {
                    items.push({
                        type: 'employee_group',
                        name,
                        action: 'skip_duplicate',
                        status: 'error',
                        resultMessage: 'Not found in Planday'
                    });
                }
            } else {
                if (isDuplicate && overwriteMode === 'keep') {
                    if (updatePayload && existingGroup && Object.keys(updatePayload).length > 0) {
                        items.push({
                            id: existingGroup.id,
                            type: 'employee_group',
                            name,
                            action: 'update',
                            status: 'pending',
                            updatePayload,
                            resultMessage: 'Updating wage data from file'
                        });
                    } else {
                        items.push({
                            type: 'employee_group',
                            name,
                            action: 'skip_duplicate',
                            status: 'skipped',
                            resultMessage: 'Already exists in Planday'
                        });
                    }
                } else {
                    items.push({
                        type: 'employee_group',
                        name,
                        action: 'create',
                        status: 'pending',
                        updatePayload
                    });
                }
            }
        });        setReviewItems(items);
        setStep(3); // Always go to Review from Mapping
    };

    // Process File/Text/Selection and Generate Review Items (Delete)
    const generateActionReviewList = () => {
        const items: StructureReviewItem[] = [];
        
        if (deleteInputMethod === 'selection') {
            currentStructure.departments.forEach(d => {
                if (deleteSelectedDepts.has(d.id)) {
                    items.push({ id: d.id, type: 'department', name: d.name, action: appMode === 'update' ? 'update' : 'delete', status: 'pending' });
                }
            });
            currentStructure.employeeGroups.forEach(g => {
                if (deleteSelectedGroups.has(g.id)) {
                    items.push({ id: g.id, type: 'employee_group', name: g.name, action: appMode === 'update' ? 'update' : 'delete', status: 'pending' });
                }
            });
        } else if (deleteInputMethod === 'text') {
            const dNames = deleteDeptText.split(/[,;\n]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
            const gNames = deleteGroupText.split(/[,;\n]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
            
            currentStructure.departments.forEach(d => {
                if (dNames.includes(d.name.toLowerCase())) {
                    items.push({ id: d.id, type: 'department', name: d.name, action: appMode === 'update' ? 'update' : 'delete', status: 'pending' });
                }
            });
            currentStructure.employeeGroups.forEach(g => {
                if (gNames.includes(g.name.toLowerCase())) {
                    items.push({ id: g.id, type: 'employee_group', name: g.name, action: appMode === 'update' ? 'update' : 'delete', status: 'pending' });
                }
            });
        } else if (deleteInputMethod === 'file') {
            const dNames = new Set<string>();
            const gNames = new Set<string>();
            parsedRows.forEach(row => {
                if (deptColumn && row[deptColumn]) dNames.add(String(row[deptColumn]).trim().toLowerCase());
                if (groupColumn && row[groupColumn]) gNames.add(String(row[groupColumn]).trim().toLowerCase());
            });
            
            currentStructure.departments.forEach(d => {
                if (dNames.has(d.name.toLowerCase())) {
                    items.push({ id: d.id, type: 'department', name: d.name, action: appMode === 'update' ? 'update' : 'delete', status: 'pending' });
                }
            });
            currentStructure.employeeGroups.forEach(g => {
                if (gNames.has(g.name.toLowerCase())) {
                    items.push({ id: g.id, type: 'employee_group', name: g.name, action: appMode === 'update' ? 'update' : 'delete', status: 'pending' });
                }
            });
        }        setReviewItems(items);
        setStep(appMode === 'update' && deleteInputMethod !== 'file' ? 2 : 3); // Config or Review
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

    const handleUpdatePayloadChange = (id: number | undefined, field: string, value: any) => {
        setReviewItems(prev => prev.map(item => {
            if (item.id === id && item.type === 'employee_group' && item.action === 'update') {
                const currentPayload = item.updatePayload || {};
                const newPayload = { ...currentPayload, [field]: value };
                if (field === 'defaultWageType') {
                    if (value === 'HourlyRate') newPayload.shiftRate = undefined;
                    if (value === 'ShiftRate') newPayload.hourlyRate = undefined;
                }
                return { ...item, updatePayload: newPayload };
            }
            return item;
        }));
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
        const toUpdate = activeItems.filter(i => i.action === 'update');

        const totalOps = toDelete.length + toCreate.length + toUpdate.length;
        let processedCount = 0;

        // Clone review items array to record progress
        const updatedItems = [...activeItems];

        // 1. Process Deletions
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

                try {                    if (item.type === 'department') {
                        const created = await createDepartment(item.name, item.number || null);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = `Created (ID: ${created.id})`;
                    } else {                        const created = await createEmployeeGroup(item.name);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = `Created (ID: ${created.id})`;
                        
                        // Apply rates from file upload or global defaults
                        let payloadToApply = item.updatePayload;
                        
                        // If no file payload but global is enabled, use global
                        if (!payloadToApply && applyGroupDefaults) {
                            payloadToApply = {
                                defaultWageType: globalDefaultWageType,
                                hourlyRate: globalHourlyRate !== '' ? Number(globalHourlyRate) : undefined,
                                shiftRate: globalShiftRate !== '' ? Number(globalShiftRate) : undefined,
                                salaryCode: globalSalaryCode,
                                salaryCode2: globalSalaryCode2
                            };
                        }

                        if (payloadToApply && Object.keys(payloadToApply).length > 0) {
                            try {
                                await updateEmployeeGroupDefaultRate(created.id, payloadToApply);
                                item.resultMessage += ' + Rates applied';
                            } catch (rateErr: any) {
                                item.resultMessage += ` (Warning: Rates failed - ${rateErr.message})`;
                            }
                        }
                    }
                } catch (err: any) {
                    item.status = 'error';
                    item.resultMessage = err.message || 'Failed to create';
                }
                setReviewItems([...updatedItems]);
            }
        }

        // 3. Process Updates
        for (let i = 0; i < updatedItems.length; i++) {
            const item = updatedItems[i];
            if (item.action === 'update' && item.id) {
                processedCount++;
                item.status = 'processing';
                setReviewItems([...updatedItems]);
                setExecutionProgress({
                    current: processedCount,
                    total: totalOps,
                    label: `Updating ${item.type === 'department' ? 'Department' : 'Employee Group'} "${item.name}"...`
                });

                try {                    if (item.type === 'department') {
                        const newName = item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name;
                        const newNum = item.updatePayload?.number !== undefined ? item.updatePayload.number : (item.number || null);
                        if (newName !== item.name || newNum !== (item.number || null) || appMode === 'import') {
                            await updateDepartment(item.id, newName, newNum);
                            item.status = 'success';
                            item.resultMessage = 'Updated successfully';
                        } else {
                            item.status = 'skipped';
                            item.resultMessage = 'No updates configured';
                        }
                    } else {
                        const payload = item.updatePayload;
                        if (payload && (payload.defaultWageType || payload.salaryCode || payload.salaryCode2)) {
                            await updateEmployeeGroupDefaultRate(item.id, {
                                defaultWageType: payload.defaultWageType,
                                hourlyRate: payload.hourlyRate !== '' && payload.hourlyRate !== undefined ? Number(payload.hourlyRate) : undefined,
                                shiftRate: payload.shiftRate !== '' && payload.shiftRate !== undefined ? Number(payload.shiftRate) : undefined,
                                salaryCode: payload.salaryCode,
                                salaryCode2: payload.salaryCode2
                            });
                            item.status = 'success';
                            item.resultMessage = 'Updated successfully';
                        } else {
                            item.status = 'skipped';
                            item.resultMessage = 'No updates configured';
                        }
                    }
                } catch (err: any) {
                    item.status = 'error';
                    item.resultMessage = err.message || 'Failed to update';
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
    };    // Export execution report as Excel
    const handleExportExecutionReport = () => {
        const data = reviewItems.map(item => {
            let actionLabel = 'Skip Duplicate';
            if (item.action === 'create') actionLabel = 'Create';
            else if (item.action === 'delete') actionLabel = 'Delete';
            else if (item.action === 'update') actionLabel = 'Update';

            const row: any = {
                "Type": item.type === 'department' ? 'Department' : 'Employee Group',
                "Name": item.name,
                "Action": actionLabel,
                "Status": item.status === 'success' ? 'Success' : item.status === 'error' ? 'Error' : 'Skipped',
                "Planday ID": item.createdId || item.id || '-',
                "Details / Result": item.resultMessage || ''
            };

            if (item.type === 'department') {
                row["Number"] = item.updatePayload?.number !== undefined ? item.updatePayload.number : (item.number || "");
            }

            if (item.type === 'employee_group') {
                row["Wage Type"] = item.updatePayload?.defaultWageType || "";
                row["Rate"] = item.updatePayload?.defaultWageType === 'HourlyRate' ? (item.updatePayload?.hourlyRate ?? "") : (item.updatePayload?.defaultWageType === 'ShiftRate' ? (item.updatePayload?.shiftRate ?? "") : "");
                row["Salary Code"] = item.updatePayload?.salaryCode || "";
                row["Salary Code 2"] = item.updatePayload?.salaryCode2 || "";
            }

            return row;
        });
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Import Results");
        XLSX.writeFile(wb, `Planday_Structure_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    // Filter review items for table
    const filteredReviewItems = useMemo(() => {
        return reviewItems.filter(item => {
            if (reviewTab === 'create' && item.action !== 'create') return false;
            if (reviewTab === 'skip' && item.action !== 'skip_duplicate') return false;
            if (reviewTab === 'delete' && item.action !== 'delete') return false;
            if (reviewTab === 'update' && item.action !== 'update') return false;

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
        const toUpdate = reviewItems.filter(i => i.action === 'update').length;
        return { toCreateDept, toCreateGroup, skipped, toDelete, toUpdate };
    }, [reviewItems]);

    // Steps configuration
    const importStepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Template & Upload", subtitle: "Export or Import Excel" },
        { title: "Field Mapping", subtitle: "Select Columns" },
        { title: "Review & Check", subtitle: "Deduplicate & Confirm" },
        { title: "Import Structure", subtitle: "Execute API Creation" },
    ];

    const deleteStepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Select Items", subtitle: "Choose what to delete" },
        { title: "Mapping / Review", subtitle: "Confirm items" },
        { title: "Review & Check", subtitle: "Final review" },
        { title: "Execute", subtitle: "API Deletion" },
    ];

    const updateStepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Select Items", subtitle: "Choose what to update" },
        { title: "Configure", subtitle: "Set new rates" },
        { title: "Review", subtitle: "Confirm updates" },
        { title: "Execute", subtitle: "API Update" },
    ];

    const stepsList = appMode === 'import' ? importStepsList : appMode === 'update' ? updateStepsList : deleteStepsList;

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

                {step > 0 && (
                    <div className="flex justify-center mb-0">
                        <div className="bg-white p-1 rounded-xl shadow-sm border border-gray-200 inline-flex">
                            <button
                                onClick={() => { setAppMode('import'); setStep(1); }}
                                className={`px-6 py-2 rounded-lg text-sm font-bold transition-colors ${appMode === 'import' ? 'bg-blue-600 text-white' : 'text-gray-600 hover:bg-gray-100'}`}
                            >
                                Import
                            </button>
                            <button
                                onClick={() => { setAppMode('update'); setStep(1); }}
                                className={`px-6 py-2 rounded-lg text-sm font-bold transition-colors ${appMode === 'update' ? 'bg-amber-600 text-white' : 'text-gray-600 hover:bg-gray-100'}`}
                            >
                                Update
                            </button>
                            <button
                                onClick={() => { setAppMode('delete'); setStep(1); }}
                                className={`px-6 py-2 rounded-lg text-sm font-bold transition-colors ${appMode === 'delete' ? 'bg-red-600 text-white' : 'text-gray-600 hover:bg-gray-100'}`}
                            >
                                Delete
                            </button>
                        </div>
                    </div>
                )}

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
                {step === 1 && appMode === 'import' && (
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
                                        {includeEmployeeGroups && (                                            <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800 ml-7">
                                                <input 
                                                    type="checkbox" 
                                                    checked={includeWageColumns} 
                                                    onChange={(e) => setIncludeWageColumns(e.target.checked)}
                                                    className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                                                />
                                                <span>Include Default Wage Data columns</span>
                                            </label>
                                        )}
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

                {/* --- STEP 1 (DELETE MODE): SELECT ITEMS --- */}
                {step === 1 && (appMode === 'delete' || appMode === 'update') && (
                    <div className="space-y-8">
                        {/* Current Structure Overview Banner (reused) */}
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

                        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200">
                            <h2 className="text-2xl font-bold text-gray-900 mb-6">{appMode === 'delete' ? 'Select Items to Delete' : 'Select Items to Update'}</h2>

                            {/* Input Method Tabs */}
                            <div className="flex gap-4 mb-6">
                                {['selection', 'text', 'file'].map(method => (
                                    <button 
                                        key={method}
                                        onClick={() => setDeleteInputMethod(method as any)}
                                        className={`px-4 py-2 rounded-lg font-bold text-sm transition-colors ${deleteInputMethod === method ? 'bg-red-100 text-red-800' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                    >
                                        {method === 'selection' ? 'Select from List' : method === 'text' ? 'Paste Names' : 'Upload File'}
                                    </button>
                                ))}
                            </div>

                            {/* Method: Selection */}
                            {deleteInputMethod === 'selection' && (
                                <div className="space-y-6">
                                    <input 
                                        type="text" 
                                        placeholder="Search by name..."
                                        value={deleteSearchTerm}
                                        onChange={(e) => setDeleteSearchTerm(e.target.value)}
                                        className="w-full px-4 py-2 border border-gray-200 rounded-xl"
                                    />
                                    <div className="grid md:grid-cols-2 gap-6">
                                        {/* Depts */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden">
                                            <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200">Departments</div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1">
                                                {currentStructure.departments.filter(d => d.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).map(d => (
                                                    <label key={d.id} className="flex items-center gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                        <input 
                                                            type="checkbox" 
                                                            checked={deleteSelectedDepts.has(d.id)}
                                                            onChange={(e) => {
                                                                const newSet = new Set(deleteSelectedDepts);
                                                                if (e.target.checked) newSet.add(d.id);
                                                                else newSet.delete(d.id);
                                                                setDeleteSelectedDepts(newSet);
                                                            }}
                                                            className="w-4 h-4 text-red-600 rounded focus:ring-red-500"
                                                        />
                                                        <span className="text-sm text-gray-700">{d.name}</span>
                                                    </label>
                                                ))}
                                            </div>
                                        </div>
                                        {/* Groups */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden">
                                            <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200">Employee Groups</div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1">
                                                {currentStructure.employeeGroups.filter(g => g.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).map(g => (
                                                    <label key={g.id} className="flex items-center gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                        <input 
                                                            type="checkbox" 
                                                            checked={deleteSelectedGroups.has(g.id)}
                                                            onChange={(e) => {
                                                                const newSet = new Set(deleteSelectedGroups);
                                                                if (e.target.checked) newSet.add(g.id);
                                                                else newSet.delete(g.id);
                                                                setDeleteSelectedGroups(newSet);
                                                            }}
                                                            className="w-4 h-4 text-red-600 rounded focus:ring-red-500"
                                                        />
                                                        <span className="text-sm text-gray-700">{g.name}</span>
                                                    </label>
                                                ))}
                                            </div>
                                        </div>
                                    </div>
                                    <button 
                                        onClick={generateActionReviewList}
                                        disabled={deleteSelectedDepts.size === 0 && deleteSelectedGroups.size === 0}
                                        className="w-full py-3 px-4 bg-red-600 hover:bg-red-700 text-white font-bold rounded-xl shadow-sm transition-all disabled:opacity-50"
                                    >
                                        Review Items ({deleteSelectedDepts.size + deleteSelectedGroups.size} items)
                                    </button>
                                </div>
                            )}

                            {/* Method: Text */}
                            {deleteInputMethod === 'text' && (
                                <div className="space-y-6">
                                    <div className="grid md:grid-cols-2 gap-6">
                                        <div>
                                            <label className="block text-sm font-bold text-gray-700 mb-2">Departments (comma, semicolon, or newline separated)</label>
                                            <textarea 
                                                value={deleteDeptText}
                                                onChange={e => setDeleteDeptText(e.target.value)}
                                                className="w-full h-32 px-4 py-2 border border-gray-200 rounded-xl"
                                                placeholder="Sales, Marketing; IT&#10;Finance"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-sm font-bold text-gray-700 mb-2">Employee Groups (comma, semicolon, or newline separated)</label>
                                            <textarea 
                                                value={deleteGroupText}
                                                onChange={e => setDeleteGroupText(e.target.value)}
                                                className="w-full h-32 px-4 py-2 border border-gray-200 rounded-xl"
                                                placeholder="Manager, Associate; Admin&#10;Intern"
                                            />
                                        </div>
                                    </div>
                                    <button 
                                        onClick={generateActionReviewList}
                                        disabled={!deleteDeptText.trim() && !deleteGroupText.trim()}
                                        className="w-full py-3 px-4 bg-red-600 hover:bg-red-700 text-white font-bold rounded-xl shadow-sm transition-all disabled:opacity-50"
                                    >
                                        Review Items
                                    </button>
                                </div>
                            )}

                            {/* Method: File */}
                            {deleteInputMethod === 'file' && (
                                <div className="space-y-6">
                                    <div className="relative border-2 border-dashed border-gray-300 hover:border-red-500 rounded-2xl p-6 text-center transition-colors bg-gray-50/50 hover:bg-red-50/30 group">
                                        <input 
                                            type="file" 
                                            accept=".xlsx,.xls,.csv" 
                                            onChange={handleFileUpload}
                                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                                        />
                                        <CloudUploadIcon className="w-10 h-10 text-gray-400 group-hover:text-red-600 mx-auto mb-2 transition-colors" />
                                        <span className="block text-sm font-bold text-gray-800">
                                            Click or Drag & Drop Excel File
                                        </span>
                                        <span className="block text-xs text-gray-500 mt-1">
                                            Names in file will be matched against Planday structure.
                                        </span>
                                    </div>
                                </div>
                            )}

                        </div>
                    </div>                )}                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && (appMode === 'import' || deleteInputMethod === 'file') && (
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
                        </div>                        {/* Selectors */}
                        <div className="grid md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4 bg-gray-50 p-6 rounded-2xl border border-gray-200">
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Departments:</label>
                                <select value={deptColumn} onChange={(e) => setDeptColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Dept Number (Opt):</label>
                                <select value={deptNumberColumn} onChange={(e) => setDeptNumberColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Employee Groups:</label>
                                <select value={groupColumn} onChange={(e) => setGroupColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Default Wage Type:</label>
                                <select value={wageTypeColumn} onChange={(e) => setWageTypeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Rate:</label>
                                <select value={rateColumn} onChange={(e) => setRateColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Salary Code:</label>
                                <select value={salaryCodeColumn} onChange={(e) => setSalaryCodeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Salary Code 2 (Rare):</label>
                                <select value={salaryCode2Column} onChange={(e) => setSalaryCode2Column(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                        </div>                        {/* Mapped Data Live Preview Table */}
                        <div>
                            <h3 className="text-md font-bold text-gray-900 mb-3">Preview Mapped Data (First 10 Rows)</h3>
                            <div className="border border-gray-200 rounded-xl overflow-x-auto shadow-sm">
                                <table className="w-full text-left text-sm whitespace-nowrap">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-3 w-12 text-center">#</th>
                                            <th className="p-3">Department (Mapped: {deptColumn || 'None'})</th>
                                            <th className="p-3">Dept # (Mapped: {deptNumberColumn || 'None'})</th>
                                            <th className="p-3">Group (Mapped: {groupColumn || 'None'})</th>
                                            <th className="p-3">Wage Type (Mapped: {wageTypeColumn || 'None'})</th>
                                            <th className="p-3">Rate (Mapped: {rateColumn || 'None'})</th>
                                            <th className="p-3">Sal. Code (Mapped: {salaryCodeColumn || 'None'})</th>
                                            <th className="p-3">Sal. Code 2 (Mapped: {salaryCode2Column || 'None'})</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {parsedRows.slice(0, 10).map((row, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50/80">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-gray-800">{deptColumn && row[deptColumn] ? row[deptColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{deptNumberColumn && row[deptNumberColumn] ? row[deptNumberColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{groupColumn && row[groupColumn] ? row[groupColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{wageTypeColumn && row[wageTypeColumn] ? row[wageTypeColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{rateColumn && row[rateColumn] ? row[rateColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{salaryCodeColumn && row[salaryCodeColumn] ? row[salaryCodeColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{salaryCode2Column && row[salaryCode2Column] ? row[salaryCode2Column] : <span className="text-gray-300 italic">-</span>}</td>
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
                            </button>                            <button 
                                onClick={appMode === 'import' || (appMode === 'update' && deleteInputMethod === 'file') ? generateReviewList : generateActionReviewList}
                                disabled={!deptColumn && !groupColumn}
                                className="px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow-md transition-all text-sm flex items-center gap-2 disabled:opacity-50"
                            >
                                <span>Continue to Review</span>
                                <ArrowRightIcon className="w-4 h-4" />
                            </button>
                        </div>
                    </div>
                )}                {/* --- STEP 3 (UPDATE MODE): CONFIGURE DEFAULTS --- */}
                {step === 2 && appMode === 'update' && deleteInputMethod !== 'file' && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <div className="flex justify-between items-center mb-2">
                                <h2 className="text-2xl font-bold text-gray-900">Configure Update Details</h2>
                            </div>
                            <p className="text-sm text-gray-600">
                                Apply the following default rates/codes to the selected employee groups.
                            </p>
                        </div>
                        
                        {reviewItems.filter(i => i.action === 'update').length === 0 ? (
                            <div className="p-4 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-sm font-medium">
                                No items selected for update.
                            </div>
                        ) : (
                            <div className="space-y-4 max-h-[60vh] overflow-y-auto pr-2">
                                {reviewItems.filter(i => i.type === 'department' && i.action === 'update').map(item => (
                                    <div key={item.id} className="border border-green-100 bg-green-50/20 rounded-xl p-4 shadow-sm">
                                        <h3 className="font-bold text-green-900 mb-3">{item.name}</h3>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Name</label>
                                                <input 
                                                    type="text"
                                                    value={item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'name', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-500"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Number (Optional)</label>
                                                <input 
                                                    type="text"
                                                    value={item.updatePayload?.number !== undefined ? item.updatePayload.number : (item.number || '')}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'number', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-500"
                                                />
                                            </div>
                                        </div>
                                    </div>
                                ))}
                                {reviewItems.filter(i => i.type === 'employee_group' && i.action === 'update').map(item => (
                                    <div key={item.id} className="border border-indigo-100 bg-indigo-50/20 rounded-xl p-4 shadow-sm">
                                        <h3 className="font-bold text-indigo-900 mb-3">{item.name}</h3>
                                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">Wage Type</label>
                                                <select 
                                                    value={item.updatePayload?.defaultWageType || ''} 
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'defaultWageType', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-indigo-500"
                                                >
                                                    <option value="">-- No Update --</option>
                                                    <option value="HourlyRate">Hourly Rate</option>
                                                    <option value="ShiftRate">Shift Rate</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">
                                                    Rate Value {item.updatePayload?.defaultWageType === 'HourlyRate' ? '(Hourly)' : item.updatePayload?.defaultWageType === 'ShiftRate' ? '(Shift)' : ''}
                                                </label>
                                                <input 
                                                    type="number" 
                                                    step="0.01"
                                                    value={item.updatePayload?.defaultWageType === 'HourlyRate' ? (item.updatePayload?.hourlyRate || '') : item.updatePayload?.defaultWageType === 'ShiftRate' ? (item.updatePayload?.shiftRate || '') : ''}
                                                    onChange={(e) => {
                                                        if (item.updatePayload?.defaultWageType === 'HourlyRate') handleUpdatePayloadChange(item.id, 'hourlyRate', e.target.value);
                                                        if (item.updatePayload?.defaultWageType === 'ShiftRate') handleUpdatePayloadChange(item.id, 'shiftRate', e.target.value);
                                                    }}
                                                    disabled={!item.updatePayload?.defaultWageType}
                                                    placeholder={!item.updatePayload?.defaultWageType ? "Select Wage Type first" : "0.00"}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm disabled:bg-gray-100 disabled:text-gray-400 bg-white focus:ring-2 focus:ring-indigo-500"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">Salary Code (Optional)</label>
                                                <input 
                                                    type="text" 
                                                    value={item.updatePayload?.salaryCode || ''}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'salaryCode', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-indigo-500"
                                                    placeholder="e.g. 1000"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">Salary Code 2 <span className="font-normal text-gray-500">(If active)</span></label>
                                                <input 
                                                    type="text" 
                                                    value={item.updatePayload?.salaryCode2 || ''}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'salaryCode2', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-indigo-500"
                                                    placeholder="e.g. 2000"
                                                />
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        )}

                        <div className="flex justify-between items-center pt-4 border-t border-gray-100">
                            <button 
                                onClick={() => setStep(1)}
                                className="px-5 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl transition-colors text-sm"
                            >
                                Back
                            </button>
                            <button 
                                onClick={() => setStep(3)}
                                className="px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow-md transition-all text-sm flex items-center gap-2"
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

                            {appMode !== 'update' ? (
                                <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                    <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">Items To Delete</span>
                                    <span className="text-2xl font-extrabold text-red-600 mt-1 block">{reviewStats.toDelete}</span>
                                </div>
                            ) : (
                                <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                    <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">Items To Update</span>
                                    <span className="text-2xl font-extrabold text-amber-600 mt-1 block">{reviewStats.toUpdate}</span>
                                </div>
                            )}
                        </div>

                        {/* Add Custom Quick Item */}
                        {appMode === 'import' && (
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
                        )}

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
                                    {reviewStats.toUpdate > 0 && (
                                        <button 
                                            onClick={() => setReviewTab('update')}
                                            className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-colors ${reviewTab === 'update' ? 'bg-amber-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                        >
                                            To Update ({reviewStats.toUpdate})
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
                            <div className="overflow-x-auto">                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-4 w-12 text-center">#</th>
                                            <th className="p-4">Type</th>
                                            <th className="p-4">Name</th>
                                            <th className="p-4">Action</th>
                                            <th className="p-4">Payload Details</th>
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
                                                        {item.action === 'update' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">
                                                                Update
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td className="p-4 text-xs text-gray-600">
                                                        {item.type === 'department' && (item.updatePayload?.number || item.number) && (
                                                            <div><span className="font-semibold text-gray-800">Number:</span> {item.updatePayload?.number || item.number}</div>
                                                        )}
                                                        {item.type === 'employee_group' && item.updatePayload ? (
                                                            <div className="space-y-1">
                                                                {item.updatePayload.defaultWageType && <div><span className="font-semibold text-gray-800">Wage Type:</span> {item.updatePayload.defaultWageType}</div>}
                                                                {item.updatePayload.hourlyRate !== undefined && <div><span className="font-semibold text-gray-800">Hourly Rate:</span> {item.updatePayload.hourlyRate}</div>}
                                                                {item.updatePayload.shiftRate !== undefined && <div><span className="font-semibold text-gray-800">Shift Rate:</span> {item.updatePayload.shiftRate}</div>}
                                                                {item.updatePayload.salaryCode && <div><span className="font-semibold text-gray-800">Sal Code:</span> {item.updatePayload.salaryCode}</div>}
                                                                {item.updatePayload.salaryCode2 && <div><span className="font-semibold text-gray-800">Sal Code 2:</span> {item.updatePayload.salaryCode2}</div>}
                                                            </div>
                                                        ) : null}
                                                        {(item.type === 'department' && !(item.updatePayload?.number || item.number)) || (item.type === 'employee_group' && !item.updatePayload) ? (
                                                            <span className="italic text-gray-400">None</span>
                                                        ) : null}
                                                    </td>
                                                    <td className="p-4 text-xs text-gray-600">
                                                        {item.resultMessage || (item.action === 'create' ? 'Ready to create' : item.action === 'delete' ? 'Scheduled for deletion' : item.action === 'update' ? 'Scheduled for update' : 'Existing item')}
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
                                                <td colSpan={7} className="p-8 text-center text-gray-400 text-sm italic">
                                                    No structure items match the selected filter.
                                                </td>
                                            </tr>
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* Global Employee Group Defaults */}
                        {appMode === 'import' && reviewStats.toCreateGroup > 0 && (
                            <div className="border border-indigo-200 bg-indigo-50/50 rounded-xl p-5 mb-6">
                                <div className="flex items-center gap-2 mb-4">
                                    <input 
                                        type="checkbox" 
                                        id="applyDefaultsToggle" 
                                        checked={applyGroupDefaults} 
                                        onChange={(e) => setApplyGroupDefaults(e.target.checked)} 
                                        className="w-4 h-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500"
                                    />
                                    <label htmlFor="applyDefaultsToggle" className="font-bold text-gray-900">
                                        Apply Default Rate Settings to all newly created Employee Groups
                                    </label>
                                </div>
                                
                                {applyGroupDefaults && (
                                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mt-4 p-4 bg-white rounded-lg border border-indigo-100">
                                        <div>
                                            <label className="block text-xs font-bold text-gray-700 mb-1">Wage Type</label>
                                            <select 
                                                value={globalDefaultWageType} 
                                                onChange={(e) => {
                                                    setGlobalDefaultWageType(e.target.value as any);
                                                    if (e.target.value === 'HourlyRate') setGlobalShiftRate('');
                                                    if (e.target.value === 'ShiftRate') setGlobalHourlyRate('');
                                                }}
                                                className="w-full p-2 border border-gray-300 rounded-lg text-sm"
                                            >
                                                <option value="">-- Select Type --</option>
                                                <option value="HourlyRate">Hourly Rate</option>
                                                <option value="ShiftRate">Shift Rate</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-xs font-bold text-gray-700 mb-1">
                                                Rate Value {globalDefaultWageType === 'HourlyRate' ? '(Hourly)' : globalDefaultWageType === 'ShiftRate' ? '(Shift)' : ''}
                                            </label>
                                            <input 
                                                type="number" 
                                                step="0.01"
                                                value={globalDefaultWageType === 'HourlyRate' ? globalHourlyRate : globalDefaultWageType === 'ShiftRate' ? globalShiftRate : ''}
                                                onChange={(e) => {
                                                    if (globalDefaultWageType === 'HourlyRate') setGlobalHourlyRate(e.target.value ? Number(e.target.value) : '');
                                                    if (globalDefaultWageType === 'ShiftRate') setGlobalShiftRate(e.target.value ? Number(e.target.value) : '');
                                                }}
                                                disabled={!globalDefaultWageType}
                                                placeholder={!globalDefaultWageType ? "Select Wage Type first" : "0.00"}
                                                className="w-full p-2 border border-gray-300 rounded-lg text-sm disabled:bg-gray-100 disabled:text-gray-400"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs font-bold text-gray-700 mb-1">Salary Code (Optional)</label>
                                            <input 
                                                type="text" 
                                                value={globalSalaryCode}
                                                onChange={(e) => setGlobalSalaryCode(e.target.value)}
                                                className="w-full p-2 border border-gray-300 rounded-lg text-sm"
                                                placeholder="e.g. 1000"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs font-bold text-gray-700 mb-1">Salary Code 2 <span className="font-normal text-gray-500">(If active)</span></label>
                                            <input 
                                                type="text" 
                                                value={globalSalaryCode2}
                                                onChange={(e) => setGlobalSalaryCode2(e.target.value)}
                                                className="w-full p-2 border border-gray-300 rounded-lg text-sm"
                                                placeholder="e.g. 2000"
                                                title="Only use if the portal setting is activated in the portal"
                                            />
                                        </div>
                                    </div>
                                )}
                            </div>
                        )}

                        {/* Actions */}
                        <div className="flex justify-between items-center pt-4">
                            <button 
                                onClick={() => setStep(appMode === 'delete' && deleteInputMethod !== 'file' ? 1 : 2)}
                                className="px-5 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl transition-colors text-sm"
                            >
                                {appMode === 'delete' && deleteInputMethod !== 'file' ? 'Back to Selection' : 'Back to Mapping'}
                            </button>

                            <button 
                                onClick={() => {
                                    if (reviewItems.filter(i => i.action === 'create' || i.action === 'delete' || i.action === 'update').length === 0) {
                                        alert("There are no structure items to process.");
                                        return;
                                    }
                                    setShowExecuteConfirm(true);
                                }}
                                className={`px-8 py-3.5 ${appMode === 'delete' ? 'bg-red-600 hover:bg-red-700' : 'bg-green-600 hover:bg-green-700'} text-white font-bold rounded-xl shadow-lg hover:shadow-xl transition-all text-sm flex items-center gap-2`}
                            >
                                <span>{appMode === 'delete' ? 'Confirm Deletions' : 'Confirm & Import Structure'}</span>
                                <ArrowRightIcon className="w-5 h-5" />
                            </button>
                        </div>
                    </div>
                )}

                {/* --- STEP 5: IMPORT EXECUTION & RESULTS --- */}
                {step === 4 && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">{appMode === 'delete' ? 'Bulk Deletion Execution' : appMode === 'update' ? 'Bulk Update Execution' : 'Structure Import Execution'}</h2>
                            <p className="text-sm text-gray-600">
                                {appMode === 'delete' ? 'Deleting selected Departments and Employee Groups in your Planday portal...' : appMode === 'update' ? 'Updating selected items in your Planday portal...' : 'Creating new Departments and Employee Groups in your Planday portal...'}
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
                                    <span>{appMode === 'delete' ? 'Deletion execution completed!' : appMode === 'update' ? 'Update execution completed!' : 'Import execution completed!'}</span>
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
                                                    {item.action === 'create' ? 'Create' : item.action === 'delete' ? 'Delete' : item.action === 'update' ? 'Update' : 'Skip'}
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
                                    <RefreshIcon className="w-4 h-4" /> {appMode === 'delete' ? 'Start New Bulk Delete' : appMode === 'update' ? 'Start New Bulk Update' : 'Start New Import'}
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
                title={appMode === 'delete' ? "⚠️ Confirm Bulk Deletion" : appMode === 'update' ? "Confirm Bulk Update" : "Confirm Planday Structure Import"}
                message={
                    <div className="space-y-2 text-sm">
                        <p>{appMode === 'delete' ? "You are about to execute the following DELETIONS in Planday. This is irreversible:" : appMode === 'update' ? "You are about to UPDATE the following items in Planday:" : "You are about to execute the following structure changes in Planday:"}</p>
                        <ul className="list-disc pl-5 font-semibold text-gray-800 space-y-1">
                            {reviewStats.toDelete > 0 && <li className="text-red-600">Delete {reviewStats.toDelete} existing structure items</li>}
                            {reviewStats.toCreateDept > 0 && <li>Create {reviewStats.toCreateDept} new Departments</li>}
                            {reviewStats.toCreateGroup > 0 && <li>Create {reviewStats.toCreateGroup} new Employee Groups</li>}
                            {reviewStats.toUpdate > 0 && <li className="text-amber-600">Update {reviewStats.toUpdate} existing items</li>}
                            {reviewStats.skipped > 0 && <li className="text-gray-500">Skip {reviewStats.skipped} items</li>}
                        </ul>
                        {appMode === 'delete' && <p className="mt-4 text-red-600 font-bold">Are you sure you want to permanently delete these items?</p>}
                        {appMode === 'update' && <p className="mt-4 text-amber-600 font-bold">This will apply the configured updates to the selected items.</p>}
                    </div>
                }
                confirmText={appMode === 'delete' ? "Yes, Execute Deletion" : appMode === 'update' ? "Execute Updates" : "Start Import Now"}
                cancelText="Go Back"
                isDangerous={appMode === 'delete'}
            />
        </div>
    );
}
