import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx-js-style';
import type { 
    PlandayApiCredentials, 
    Department, 
    EmployeeGroup, 
    StructureReviewItem, 
    OverwriteMode, 
    CurrentStructure,
    Section,
    Position,
    Skill,
    RevenueUnit
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
    revokeToken, 
    fetchPortalInfo,
    updateEmployeeGroupDefaultRate,
    updateDepartment,
    fetchEmployeeGroupDefaultRates,
    fetchSkills,
    createSkill,
    updateSkill,
    deleteSkill,
    fetchPositions,
    fetchPositionById,
    fetchRevenueUnits,
    createPosition,
    updatePosition,
    deletePosition,
    fetchSections,
    fetchAllSections,
    createSection,
    updateSection,
    deleteSection
} from './services/plandayService';
import { checkForUSDateFormat, parseDateWithFormat } from './utils/dateParsing';

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
const StarIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11.049 2.927c.3-.921 1.603-.921 1.902 0l1.519 4.674a1 1 0 00.95.69h4.915c.969 0 1.371 1.24.588 1.81l-3.976 2.888a1 1 0 00-.363 1.118l1.518 4.674c.3.922-.755 1.688-1.538 1.118l-3.976-2.888a1 1 0 00-1.176 0l-3.976 2.888c-.783.57-1.838-.197-1.538-1.118l1.518-4.674a1 1 0 00-.363-1.118l-3.976-2.888c-.784-.57-.38-1.81.588-1.81h4.914a1 1 0 00.951-.69l1.519-4.674z" /></svg>);
const BriefcaseIcon: React.FC<{ className?: string }> = ({ className }) => (<svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 13.255A23.931 23.931 0 0112 15c-3.183 0-6.22-.62-9-1.745M16 6V4a2 2 0 00-2-2h-4a2 2 0 00-2 2v2m4 6h.01M5 20h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" /></svg>);
const SectionIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM14 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zM14 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z" />
    </svg>
);

const TYPE_ORDER: Record<string, number> = {
    'department': 1,
    'employee_group': 2,
    'section': 3,
    'position': 4,
    'skill': 5
};

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

    // Bulk Import state (File / Template upload)
    
    // Bulk Delete / Update state
    const [deleteInputMethod, setDeleteInputMethod] = useState<'file' | 'selection'>('file');
    const [deleteSelectedDepts, setDeleteSelectedDepts] = useState<Set<number>>(new Set());
    const [deleteSelectedGroups, setDeleteSelectedGroups] = useState<Set<number>>(new Set());
    const [deleteSelectedSections, setDeleteSelectedSections] = useState<Set<number>>(new Set());
    const [deleteSelectedPositions, setDeleteSelectedPositions] = useState<Set<number>>(new Set());
    const [deleteSelectedSkills, setDeleteSelectedSkills] = useState<Set<number>>(new Set());
    const [deleteSectionDeptId, setDeleteSectionDeptId] = useState<string>('');
    const [deletePositionDeptId, setDeletePositionDeptId] = useState<string>('');
    const [deleteSearchTerm, setDeleteSearchTerm] = useState('');

    const [step, setStep] = useState<number>(0);
    const [credentials, setCredentials] = useState<PlandayApiCredentials>({
        clientId: FIXED_CLIENT_ID,
        refreshToken: ''
    });
    const [isConnecting, setIsConnecting] = useState(false);
    const [isDisconnecting, setIsDisconnecting] = useState(false);
    const [authError, setAuthError] = useState<string | null>(null);
    const [showTokenGuide, setShowTokenGuide] = useState(false);
    const [portalInfo, setPortalInfo] = useState<any>(null);
    const [portalName, setPortalName] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [copiedAppId, setCopiedAppId] = useState(false);

    // Current Planday Structure fetched from API
    const [currentStructure, setCurrentStructure] = useState<CurrentStructure>({
        departments: [],
        employeeGroups: [],
        skills: [],
        positions: [],
        sections: []
    });

    // Template options
    const [includeDepartments, setIncludeDepartments] = useState(true);
    const [includeEmployeeGroups, setIncludeEmployeeGroups] = useState(true);
    const [includeSkills, setIncludeSkills] = useState(false);
    const [includeSections, setIncludeSections] = useState(false);
    const [includePositions, setIncludePositions] = useState(false);

    // Current Structure Preview department filters
    const [previewSectionDeptId, setPreviewSectionDeptId] = useState<string>('');
    const [previewPositionDeptId, setPreviewPositionDeptId] = useState<string>('');

    // Select from List department filters
    const [selectListSectionDeptId, setSelectListSectionDeptId] = useState<string>('all');
    const [selectListPositionDeptId, setSelectListPositionDeptId] = useState<string>('all');

    // File Import state
    const [uploadedFileName, setUploadedFileName] = useState<string | null>(null);
    const [includeWageColumns, setIncludeWageColumns] = useState(false);
    const [populateWithCurrent, setPopulateWithCurrent] = useState(false);
    const [parsedHeaders, setParsedHeaders] = useState<string[]>([]);
    const [parsedRows, setParsedRows] = useState<Record<string, any>[]>([]);

    // Mapping state
    const [deptColumn, setDeptColumn] = useState<string>('');
    const [deptUpdateNameColumn, setDeptUpdateNameColumn] = useState<string>('');
    const [groupColumn, setGroupColumn] = useState<string>('');
    const [deptNumberColumn, setDeptNumberColumn] = useState<string>('');
    const [wageTypeColumn, setWageTypeColumn] = useState<string>('');
    const [rateColumn, setRateColumn] = useState<string>('');
    const [salaryCodeColumn, setSalaryCodeColumn] = useState<string>('');
    const [salaryCode2Column, setSalaryCode2Column] = useState<string>('');
    
    // Skills Mapping State
    const [skillNameColumn, setSkillNameColumn] = useState<string>('');
    const [skillDescColumn, setSkillDescColumn] = useState<string>('');

    // Sections Mapping State
    const [sectionNameColumn, setSectionNameColumn] = useState<string>('');
    const [sectionDeptColumn, setSectionDeptColumn] = useState<string>('');
    const [sectionColorColumn, setSectionColorColumn] = useState<string>('');
    const [sectionUpdateNameColumn, setSectionUpdateNameColumn] = useState<string>('');

    // Positions Mapping State
    const [posNameColumn, setPosNameColumn] = useState<string>('');
    const [posUpdateNameColumn, setPosUpdateNameColumn] = useState<string>('');
    const [posDeptColumn, setPosDeptColumn] = useState<string>('');
    const [posSectionColumn, setPosSectionColumn] = useState<string>('');
    const [posGroupColumn, setPosGroupColumn] = useState<string>('');
    const [posColorColumn, setPosColorColumn] = useState<string>('');
    const [posValidFromColumn, setPosValidFromColumn] = useState<string>('');
    const [posValidToColumn, setPosValidToColumn] = useState<string>('');
    const [posSkillsColumn, setPosSkillsColumn] = useState<string>('');
    const [posPayrollColumn, setPosPayrollColumn] = useState<string>('');
    const [posRevenueUnitColumn, setPosRevenueUnitColumn] = useState<string>('');

    // Review state
    const [reviewItems, setReviewItems] = useState<StructureReviewItem[]>([]);
    const [validationWarnings, setValidationWarnings] = useState<string[]>([]);
    const [reviewTab, setReviewTab] = useState<'all' | 'create' | 'skip' | 'delete' | 'update'>('all');
    const [reviewSearch, setReviewSearch] = useState('');
    const [newItemName, setNewItemName] = useState('');
    const [newItemType, setNewItemType] = useState<'department' | 'employee_group' | 'skill' | 'position' | 'section'>('department');
    const [newItemDeptId, setNewItemDeptId] = useState<number | ''>('');
    const [newItemColor, setNewItemColor] = useState<string>('');

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
        let detectedDeptUpdateName = '';
        let detectedGroup = '';
        let detectedDeptNum = '';
        let detectedWageType = '';
        let detectedRate = '';
        let detectedSalCode = '';
        let detectedSalCode2 = '';

        let detSkillName = '';
        let detSkillDesc = '';

        let detSectionName = '';
        let detSectionDept = '';
        let detSectionColor = '';
        let detSectionUpdateName = '';

        let detPosName = '';
        let detPosUpdateName = '';
        let detPosDept = '';
        let detPosSection = '';
        let detPosGroup = '';
        let detPosColor = '';
        let detPosValidFrom = '';
        let detPosValidTo = '';
        let detPosSkills = '';
        let detPosPayroll = '';
        let detPosRevenueUnit = '';

        for (const h of headers) {
            const clean = h.trim().toLowerCase();
            
            // UPDATE - Department Name (renaming target)
            if (!detectedDeptUpdateName && (
                clean.includes('update - department name') ||
                clean.includes('update - dept name') ||
                clean.includes('update department name') ||
                clean.includes('update dept name') ||
                clean.includes('new department name') ||
                clean.includes('new dept name') ||
                clean === 'update - department' ||
                clean === 'new department'
            )) {
                detectedDeptUpdateName = h;
            }

            // UPDATE - Department Number / Department Number
            if (!detectedDeptNum && (
                clean.includes('update - department number') ||
                clean.includes('update - dept number') ||
                clean.includes('update department number') ||
                clean.includes('update dept number') ||
                clean.includes('department number') || 
                clean.includes('dept number')
            )) {
                detectedDeptNum = h;
            }

            // Current / Identifier Department Name
            if (!detectedDept && 
                !clean.includes('update') && 
                !clean.includes('new') && 
                !clean.includes('section') && 
                !clean.includes('position') && 
                (
                    clean === 'departments' || 
                    clean === 'department' || 
                    clean === 'dept' || 
                    clean === 'department name' || 
                    clean === 'dept name' || 
                    clean === 'current department' ||
                    clean === 'current department name' ||
                    clean === 'afdeling'
                )
            ) {
                detectedDept = h;
            }

            if (!detectedGroup && (clean.includes('employee group') || clean === 'employee groups' || clean === 'role' || clean === 'stilling')) {
                detectedGroup = h;
            }
            if (!detectedWageType && (clean.includes('wage type') || clean.includes('default wage type'))) {
                detectedWageType = h;
            }
            if (!detectedRate && !clean.includes('wage type') && (
                clean === 'rate' || 
                clean === 'hourly rate' || 
                clean === 'shift rate' || 
                clean === 'group: default rate' || 
                clean === 'default rate' || 
                clean.includes('default rate')
            )) {
                detectedRate = h;
            }
            if (!detectedSalCode && (
                clean === 'salary code' || 
                clean === 'salarycode' || 
                clean === 'group: salary code' || 
                (clean.includes('salary code') && !clean.includes('2'))
            )) {
                detectedSalCode = h;
            }
            if (!detectedSalCode2 && (
                clean.includes('salary code 2')
            )) {
                detectedSalCode2 = h;
            }

            if (!detSkillName && (clean.includes('skill: name') || clean === 'skills' || clean === 'skill' || clean === 'skill name')) detSkillName = h;
            if (!detSkillDesc && clean.includes('skill: description')) detSkillDesc = h;

            // Section Column detection
            if (!detSectionUpdateName && (
                clean.includes('update - section name') ||
                clean.includes('update section name') ||
                clean.includes('new section name') ||
                clean === 'update - section'
            )) {
                detSectionUpdateName = h;
            }

            if (!detSectionName && (
                clean.includes('section: name') ||
                clean === 'section name' ||
                clean === 'section' ||
                clean === 'sections'
            ) && !clean.includes('position') && !clean.includes('update')) {
                detSectionName = h;
            }

            if (!detSectionDept && (
                clean.includes('section: department') ||
                clean.includes('section: dept') ||
                clean === 'section department' ||
                clean === 'section dept'
            )) {
                detSectionDept = h;
            }

            if (!detSectionColor && (
                clean.includes('section: color') ||
                clean === 'section color' ||
                clean.includes('section color hex') ||
                clean === 'section color hex # (optional)'
            )) {
                detSectionColor = h;
            }

            // Position Column detection
            if (!detPosUpdateName && (
                clean.includes('update - position name') ||
                clean.includes('update - pos name') ||
                clean.includes('update position name') ||
                clean.includes('update pos name') ||
                clean.includes('new position name') ||
                clean.includes('new pos name') ||
                clean === 'update - position' ||
                clean === 'new position'
            )) {
                detPosUpdateName = h;
            }

            if (!detPosName && !clean.includes('update') && !clean.includes('new') && (
                clean.includes('position: name') ||
                clean === 'positions' ||
                clean === 'position' ||
                clean === 'position name' ||
                clean === 'pos name'
            )) {
                detPosName = h;
            }

            if (!detPosDept && (
                clean.includes('position: department') ||
                clean.includes('position: dept') ||
                clean === 'position department' ||
                clean === 'position dept'
            )) {
                detPosDept = h;
            }

            if (!detPosSection && (
                clean.includes('position: section') ||
                clean === 'position section'
            )) {
                detPosSection = h;
            }

            if (!detPosGroup && (
                clean.includes('position: employee group') ||
                clean === 'position employee group' ||
                clean === 'position group'
            )) {
                detPosGroup = h;
            }

            if (!detPosColor && (
                clean.includes('position: color') ||
                clean === 'position color' ||
                clean.includes('position color hex') ||
                clean.includes('pos. color') ||
                clean.includes('pos color')
            )) {
                detPosColor = h;
            }

            if (!detPosValidFrom && (
                clean.includes('position: valid from') ||
                clean.includes('position valid from') ||
                clean.includes('pos. valid from') ||
                clean.includes('pos valid from')
            )) {
                detPosValidFrom = h;
            }

            if (!detPosValidTo && (
                clean.includes('position: valid to') ||
                clean.includes('position valid to') ||
                clean.includes('pos. valid to') ||
                clean.includes('pos valid to')
            )) {
                detPosValidTo = h;
            }

            if (!detPosSkills && (
                clean.includes('position: skills') ||
                clean.includes('position skills') ||
                clean.includes('pos. skills') ||
                clean.includes('pos skills')
            )) {
                detPosSkills = h;
            }

            if (!detPosPayroll && (
                clean.includes('position: include in payroll') ||
                clean.includes('include in payroll') ||
                clean.includes('affect revenue') ||
                clean.includes('affectrevenue') ||
                clean.includes('pos. in payroll') ||
                clean.includes('pos payroll')
            )) {
                detPosPayroll = h;
            }

            if (!detPosRevenueUnit && (
                clean.includes('position: revenue unit') ||
                clean.includes('position revenue unit') ||
                clean.includes('pos. revenue unit') ||
                clean.includes('pos. rev unit') ||
                clean.includes('pos revenue unit')
            )) {
                detPosRevenueUnit = h;
            }
        }

        setDeptColumn(detectedDept);
        setDeptUpdateNameColumn(detectedDeptUpdateName);
        setGroupColumn(detectedGroup);
        setDeptNumberColumn(detectedDeptNum);
        setWageTypeColumn(detectedWageType);
        setRateColumn(detectedRate);
        setSalaryCodeColumn(detectedSalCode);
        setSalaryCode2Column(detectedSalCode2);

        setSkillNameColumn(detSkillName);
        setSkillDescColumn(detSkillDesc);

        setSectionNameColumn(detSectionName);
        setSectionDeptColumn(detSectionDept);
        setSectionColorColumn(detSectionColor);
        setSectionUpdateNameColumn(detSectionUpdateName);

        setPosNameColumn(detPosName);
        setPosUpdateNameColumn(detPosUpdateName);
        setPosDeptColumn(detPosDept);
        setPosSectionColumn(detPosSection);
        setPosGroupColumn(detPosGroup);
        setPosColorColumn(detPosColor);
        setPosValidFromColumn(detPosValidFrom);
        setPosValidToColumn(detPosValidTo);
        setPosSkillsColumn(detPosSkills);
        setPosPayrollColumn(detPosPayroll);
        setPosRevenueUnitColumn(detPosRevenueUnit);
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
            const [depts, groups, info, skills, revUnits] = await Promise.all([
                fetchDepartments(),
                fetchEmployeeGroups(),
                fetchPortalInfo().catch(() => null),
                fetchSkills().catch(() => []),
                fetchRevenueUnits().catch(() => [])
            ]);
            const positions = await fetchPositions(skills).catch(() => []);
            const sections = await fetchAllSections(depts).catch(() => []);

            setCurrentStructure({
                departments: depts,
                employeeGroups: groups,
                skills: (skills || []).filter((s: any) => s.isActive !== false && s.IsActive !== false && s.isActive !== 'false' && s.IsActive !== 'false'),
                positions: (positions || []).filter((p: any) => p.isActive !== false && p.IsActive !== false && p.isActive !== 'false' && p.IsActive !== 'false'),
                sections: (sections || []).filter((s: any) => s.isActive !== false && s.IsActive !== false && s.isActive !== 'false' && s.IsActive !== 'false'),
                revenueUnits: revUnits
            });
            setPortalInfo(info);
            if (info?.name) setPortalName(info.name);

            sessionStorage.setItem('plandayCredentials', JSON.stringify(credsToUse));
            setStep(1); // Move to Step 2
        } catch (err: any) {
            setAuthError(err.message || "Failed to connect to Planday API.");
        } finally {
            setIsConnecting(false);
        }
    };

    const handleChangeCredentials = async () => {
        setIsDisconnecting(true);
        try {
            await revokeToken();
        } catch (error) {
            console.error("Failed to revoke token:", error);
        } finally {
            sessionStorage.removeItem('plandayCredentials');
            setPortalName(null);
            setPortalInfo(null);
            resetService();
            setCredentials({
                clientId: FIXED_CLIENT_ID,
                refreshToken: ''
            });
            setCurrentStructure({
                departments: [],
                employeeGroups: [],
                skills: [],
                positions: [],
                sections: []
            });
            setReviewItems([]);
            setParsedHeaders([]);
            setParsedRows([]);
            setUploadedFileName(null);
            setStep(0);
            setIsDisconnecting(false);
        }
    };

    const handleManualRefresh = async () => {
        setIsLoading(true);
        try {
            const [depts, groups, info, skills, revUnits] = await Promise.all([
                fetchDepartments(),
                fetchEmployeeGroups(),
                fetchPortalInfo().catch(() => null),
                fetchSkills().catch(() => []),
                fetchRevenueUnits().catch(() => [])
            ]);
            const positions = await fetchPositions(skills).catch(() => []);
            const sections = await fetchAllSections(depts).catch(() => []);
            setCurrentStructure({
                departments: depts,
                employeeGroups: groups,
                skills: (skills || []).filter((s: any) => s.isActive !== false && s.IsActive !== false && s.isActive !== 'false' && s.IsActive !== 'false'),
                positions: (positions || []).filter((p: any) => p.isActive !== false && p.IsActive !== false && p.isActive !== 'false' && p.IsActive !== 'false'),
                sections: (sections || []).filter((s: any) => s.isActive !== false && s.IsActive !== false && s.isActive !== 'false' && s.IsActive !== 'false'),
                revenueUnits: revUnits
            });
            setPortalInfo(info);
            if (info?.name) setPortalName(info.name);
        } catch (error) {
            console.error("Failed to refresh", error);
        } finally {
            setIsLoading(false);
        }
    };

    // Helper: Build "Reference Data" sheet listing all API IDs next to descriptive names
    // Reflects the selections made: only types currently selected will be visible
    const buildReferenceDataSheet = (opts?: {
        includeDepartments?: boolean;
        includeSections?: boolean;
        includeEmployeeGroups?: boolean;
        includeSkills?: boolean;
        includePositions?: boolean;
    }) => {
        const rows: any[][] = [];
        const cellStyles: { r: number; c: number; s: any }[] = [];

        const deptHeaderStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
            fill: { fgColor: { rgb: "1E3A8A" } }, // Dark Navy Blue (Departments)
            alignment: { horizontal: "left" }
        };

        const groupHeaderStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
            fill: { fgColor: { rgb: "6B21A8" } }, // Purple (Employee Groups)
            alignment: { horizontal: "left" }
        };

        const sectionHeaderStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
            fill: { fgColor: { rgb: "374151" } }, // Dark Grey (Sections)
            alignment: { horizontal: "left" }
        };

        const positionHeaderStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
            fill: { fgColor: { rgb: "0F766E" } }, // Teal (Positions)
            alignment: { horizontal: "left" }
        };

        const skillHeaderStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
            fill: { fgColor: { rgb: "854D0E" } }, // Dark Yellow (Skills)
            alignment: { horizontal: "left" }
        };

        const colHeaderStyle = {
            font: { bold: true, color: { rgb: "1E293B" }, sz: 10 },
            fill: { fgColor: { rgb: "E2E8F0" } }, // Slate 200
            alignment: { horizontal: "left" },
            border: {
                bottom: { style: "thin", color: { rgb: "94A3B8" } }
            }
        };

        const dataTextStyle = {
            font: { color: { rgb: "1E293B" }, sz: 10 },
            alignment: { horizontal: "left" }
        };

        const dataIdStyle = {
            font: { bold: true, color: { rgb: "0F172A" }, sz: 10 },
            alignment: { horizontal: "left" }
        };

        const emptyStyle = {
            font: { italic: true, color: { rgb: "64748B" }, sz: 10 },
            alignment: { horizontal: "left" }
        };

        let currentRow = 0;

        const showDepts = opts?.includeDepartments !== undefined ? opts.includeDepartments : includeDepartments;
        const showGroups = opts?.includeEmployeeGroups !== undefined ? opts.includeEmployeeGroups : includeEmployeeGroups;
        const showSections = opts?.includeSections !== undefined ? opts.includeSections : includeSections;
        const showPositions = opts?.includePositions !== undefined ? opts.includePositions : includePositions;
        const showSkills = opts?.includeSkills !== undefined ? opts.includeSkills : includeSkills;

        let hasAnySection = false;

        // 1. Section: Departments (only if selected)
        if (showDepts) {
            hasAnySection = true;
            rows.push(["DEPARTMENTS", "", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: deptHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: deptHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: deptHeaderStyle });
            currentRow++;

            rows.push(["Department Name", "API ID", "Department Number"]);
            cellStyles.push({ r: currentRow, c: 0, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: colHeaderStyle });
            currentRow++;

            if (currentStructure.departments && currentStructure.departments.length > 0) {
                currentStructure.departments.forEach(dept => {
                    rows.push([dept.name, dept.id, dept.number || ""]);
                    cellStyles.push({ r: currentRow, c: 0, s: dataTextStyle });
                    cellStyles.push({ r: currentRow, c: 1, s: dataIdStyle });
                    cellStyles.push({ r: currentRow, c: 2, s: dataTextStyle });
                    currentRow++;
                });
            } else {
                rows.push(["(No departments found in Planday)", "", ""]);
                cellStyles.push({ r: currentRow, c: 0, s: emptyStyle });
                currentRow++;
            }

            // Spacing row
            rows.push(["", "", ""]);
            currentRow++;
        }

        // 2. Section: Employee Groups (only if selected)
        if (showGroups) {
            hasAnySection = true;
            rows.push(["EMPLOYEE GROUPS", "", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: groupHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: groupHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: groupHeaderStyle });
            currentRow++;

            rows.push(["Employee Group Name", "API ID", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: colHeaderStyle });
            currentRow++;

            if (currentStructure.employeeGroups && currentStructure.employeeGroups.length > 0) {
                currentStructure.employeeGroups.forEach(group => {
                    rows.push([group.name, group.id, ""]);
                    cellStyles.push({ r: currentRow, c: 0, s: dataTextStyle });
                    cellStyles.push({ r: currentRow, c: 1, s: dataIdStyle });
                    currentRow++;
                });
            } else {
                rows.push(["(No employee groups found in Planday)", "", ""]);
                cellStyles.push({ r: currentRow, c: 0, s: emptyStyle });
                currentRow++;
            }

            // Spacing row
            rows.push(["", "", ""]);
            currentRow++;
        }

        // 3. Section: Sections (only if selected)
        if (showSections) {
            hasAnySection = true;
            rows.push(["SECTIONS", "", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: sectionHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: sectionHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: sectionHeaderStyle });
            currentRow++;

            rows.push(["Section Name", "API ID", "Department / Color"]);
            cellStyles.push({ r: currentRow, c: 0, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: colHeaderStyle });
            currentRow++;

            const activeSections = (currentStructure.sections || []).filter(s => s.isActive !== false && (s as any).IsActive !== false && (s as any).isActive !== 'false' && (s as any).IsActive !== 'false');
            if (activeSections.length > 0) {
                const deptMap = new Map<number, string>(currentStructure.departments.map(d => [d.id, d.name]));
                activeSections.forEach(sec => {
                    const deptName = deptMap.get(sec.departmentId) || `Dept ID: ${sec.departmentId}`;
                    const details = [deptName, sec.color ? `Color: ${sec.color}` : ''].filter(Boolean).join(' | ');
                    rows.push([sec.name, sec.id, details]);
                    cellStyles.push({ r: currentRow, c: 0, s: dataTextStyle });
                    cellStyles.push({ r: currentRow, c: 1, s: dataIdStyle });
                    cellStyles.push({ r: currentRow, c: 2, s: dataTextStyle });
                    currentRow++;
                });
            } else {
                rows.push(["(No sections found in Planday)", "", ""]);
                cellStyles.push({ r: currentRow, c: 0, s: emptyStyle });
                currentRow++;
            }

            // Spacing row
            rows.push(["", "", ""]);
            currentRow++;
        }

        // 4. Section: Positions (only if positions is selected)
        if (showPositions) {
            hasAnySection = true;
            rows.push(["POSITIONS", "", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: positionHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: positionHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: positionHeaderStyle });
            currentRow++;

            rows.push(["Position Name", "API ID", "Associated Department / Group"]);
            cellStyles.push({ r: currentRow, c: 0, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: colHeaderStyle });
            currentRow++;

            const activePositions = (currentStructure.positions || []).filter(p => p.isActive !== false && (p as any).IsActive !== false && (p as any).isActive !== 'false' && (p as any).IsActive !== 'false');
            if (activePositions.length > 0) {
                const deptMap = new Map<number, string>(currentStructure.departments.map(d => [d.id, d.name]));
                const groupMap = new Map<number, string>(currentStructure.employeeGroups.map(g => [g.id, g.name]));
                activePositions.forEach(pos => {
                    const deptName = pos.departmentId ? deptMap.get(pos.departmentId) : undefined;
                    const groupName = pos.employeeGroupId ? groupMap.get(pos.employeeGroupId) : undefined;
                    const details: string[] = [];
                    if (deptName) details.push(`Dept: ${deptName}`);
                    else if (pos.departmentId) details.push(`Dept ID: ${pos.departmentId}`);
                    if (groupName) details.push(`Group: ${groupName}`);
                    else if (pos.employeeGroupId) details.push(`Group ID: ${pos.employeeGroupId}`);
                    if (pos.sectionId && currentStructure.sections) {
                        const sec = currentStructure.sections.find(s => s.id === pos.sectionId);
                        if (sec) details.push(`Section: ${sec.name}`);
                    }
                    if (pos.revenueUnitId && currentStructure.revenueUnits) {
                        const ru = currentStructure.revenueUnits.find(r => r.id === pos.revenueUnitId);
                        if (ru) details.push(`Revenue Unit: ${ru.name}`);
                        else details.push(`RevUnit ID: ${pos.revenueUnitId}`);
                    }
                    if (pos.affectRevenue !== undefined) {
                        details.push(`Include in Payroll: ${pos.affectRevenue}`);
                    }
                    if (pos.skillIds && pos.skillIds.length > 0 && currentStructure.skills) {
                        const sNames = currentStructure.skills.filter(s => pos.skillIds?.includes(s.id)).map(s => s.name).join(', ');
                        if (sNames) details.push(`Skills: ${sNames}`);
                    }

                    rows.push([pos.name, pos.id, details.join(' | ')]);
                    cellStyles.push({ r: currentRow, c: 0, s: dataTextStyle });
                    cellStyles.push({ r: currentRow, c: 1, s: dataIdStyle });
                    cellStyles.push({ r: currentRow, c: 2, s: dataTextStyle });
                    currentRow++;
                });
            } else {
                rows.push(["(No positions found in Planday)", "", ""]);
                cellStyles.push({ r: currentRow, c: 0, s: emptyStyle });
                currentRow++;
            }

            // Spacing row
            rows.push(["", "", ""]);
            currentRow++;
        }

        // 5. Section: Skills (only if selected)
        if (showSkills) {
            hasAnySection = true;
            rows.push(["SKILLS", "", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: skillHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: skillHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: skillHeaderStyle });
            currentRow++;

            rows.push(["Skill Name", "API ID", "Description"]);
            cellStyles.push({ r: currentRow, c: 0, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 1, s: colHeaderStyle });
            cellStyles.push({ r: currentRow, c: 2, s: colHeaderStyle });
            currentRow++;

            if (currentStructure.skills && currentStructure.skills.length > 0) {
                currentStructure.skills.forEach(skill => {
                    rows.push([skill.name, skill.id, skill.description || ""]);
                    cellStyles.push({ r: currentRow, c: 0, s: dataTextStyle });
                    cellStyles.push({ r: currentRow, c: 1, s: dataIdStyle });
                    cellStyles.push({ r: currentRow, c: 2, s: dataTextStyle });
                    currentRow++;
                });
            } else {
                rows.push(["(No skills found in Planday)", "", ""]);
                cellStyles.push({ r: currentRow, c: 0, s: emptyStyle });
                currentRow++;
            }

            // Spacing row
            rows.push(["", "", ""]);
            currentRow++;
        }

        if (!hasAnySection) {
            rows.push(["(No structure types selected for reference data)", "", ""]);
            cellStyles.push({ r: currentRow, c: 0, s: emptyStyle });
            currentRow++;
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);

        // Apply styles to sheet
        cellStyles.forEach(({ r, c, s }) => {
            const addr = XLSX.utils.encode_cell({ r, c });
            if (!ws[addr]) {
                ws[addr] = { t: 's', v: '' };
            }
            ws[addr].s = s;
        });

        ws['!cols'] = [
            { wch: 38 }, // Descriptive Name
            { wch: 18 }, // API ID
            { wch: 38 }  // Additional info
        ];

        return ws;
    };

    // Helper: Determine distinct dark header background color by entity type for Excel template export files
    // Departments: Dark Navy Blue (#1E3A8A)
    // Employee Groups: Purple (#6B21A8)
    // Sections: Dark Grey (#374151)
    // Positions: Teal (#0F766E)
    // Skills: Dark Yellow (#854D0E)
    const getHeaderBgColor = (headerName: string): string => {
        const h = (headerName || '').trim().toLowerCase();

        // 1. Positions (check first: contains position-specific prefix or title)
        if (h.startsWith('position') || h.includes('position')) {
            return "0F766E"; // Teal (Positions)
        }

        // 2. Sections (check before department because 'Section: Department' exists)
        if (h.startsWith('section') || h.includes('section')) {
            return "374151"; // Dark Grey (Sections)
        }

        // 3. Skills
        if (h.startsWith('skill') || h.includes('skill')) {
            return "854D0E"; // Dark Yellow (Skills)
        }

        // 4. Employee Groups (includes group-level wage and salary fields)
        if (h.startsWith('group:') || h.includes('employee group') || h.includes('wage') || h.includes('salary code')) {
            return "6B21A8"; // Purple (Employee Groups)
        }

        // 5. Departments
        if (h.includes('department')) {
            return "1E3A8A"; // Dark Navy Blue (Departments)
        }

        return "1E293B"; // Dark Slate fallback
    };

    // Step 2: Download Current Structure Excel (Populated)
    const handleDownloadCurrentStructure = async () => {
        const wb = XLSX.utils.book_new();
        
        // Fetch wage rates if requested
        let wageRatesMap = new Map<number, any>();
        if (includeEmployeeGroups && includeWageColumns && appMode !== 'delete') {
            try {
                const rates = await fetchEmployeeGroupDefaultRates();
                rates.forEach(rate => {
                    wageRatesMap.set(rate.employeeGroupId, rate);
                });
            } catch (err) {
                console.error("Failed to fetch default wage rates:", err);
            }
        }

        const activeSections = (currentStructure.sections || []).filter(s => s.isActive !== false && (s as any).IsActive !== false && (s as any).isActive !== 'false' && (s as any).IsActive !== 'false');
        const activePositions = (currentStructure.positions || []).filter(p => p.isActive !== false && (p as any).IsActive !== false && (p as any).isActive !== 'false' && (p as any).IsActive !== 'false');

        // Prepare rows based on chosen options
        const maxLen = Math.max(
            includeDepartments ? currentStructure.departments.length : 0,
            includeEmployeeGroups ? currentStructure.employeeGroups.length : 0,
            includeSections ? activeSections.length : 0,
            includePositions ? activePositions.length : 0,
            includeSkills ? (currentStructure.skills?.length || 0) : 0
        );
        const data: any[] = [];
        
        for (let i = 0; i < Math.max(1, maxLen); i++) {
            const row: any = {};
            if (includeDepartments) {
                if (appMode === 'delete') {
                    row["Departments"] = currentStructure.departments[i]?.name || "";
                } else if (appMode === 'update') {
                    row["Departments"] = currentStructure.departments[i]?.name || "";
                    row["UPDATE - Department Name"] = "";
                    row["Department Number (optional)"] = currentStructure.departments[i]?.number || "";
                } else {
                    row["Departments"] = currentStructure.departments[i]?.name || "";
                    row["Department Number (optional)"] = currentStructure.departments[i]?.number || "";
                }
            }
            if (includeEmployeeGroups) {
                const group = currentStructure.employeeGroups[i];
                if (appMode === 'delete') {
                    row["Employee Groups"] = group?.name || "";
                } else if (appMode === 'update') {
                    row["Employee Groups"] = group?.name || "";
                    if (includeWageColumns) {
                        const rateData = group ? wageRatesMap.get(group.id) : null;
                        row["Group: Default Wage Type (HourlyRate/ShiftRate)"] = rateData?.defaultWageType || "";
                        row["Group: Default Rate"] = rateData?.defaultWageType === 'HourlyRate' ? (rateData?.hourlyRate ?? "") : (rateData?.defaultWageType === 'ShiftRate' ? (rateData?.shiftRate ?? "") : "");
                        row["Group: Salary Code"] = rateData?.salaryCode || "";
                        row["Group: Salary Code 2 (needs activation)"] = rateData?.salaryCode2 || "";
                    }
                } else {
                    row["Employee Groups"] = group?.name || "";
                    if (includeWageColumns) {
                        const rateData = group ? wageRatesMap.get(group.id) : null;
                        row["Group: Default Wage Type (HourlyRate/ShiftRate)"] = rateData?.defaultWageType || "";
                        row["Group: Default Rate"] = rateData?.defaultWageType === 'HourlyRate' ? (rateData?.hourlyRate ?? "") : (rateData?.defaultWageType === 'ShiftRate' ? (rateData?.shiftRate ?? "") : "");
                        row["Group: Salary Code"] = rateData?.salaryCode || "";
                        row["Group: Salary Code 2 (needs activation)"] = rateData?.salaryCode2 || "";
                    }
                }
            }
            if (includeSections) {
                const sec = activeSections[i];
                const deptName = sec?.departmentId ? (currentStructure.departments.find(d => d.id === sec.departmentId)?.name || "") : "";
                if (appMode === 'delete') {
                    row["Section: Name"] = sec?.name || "";
                    row["Section: Department"] = deptName;
                } else if (appMode === 'update') {
                    row["Section: Name"] = sec?.name || "";
                    row["Section: Department"] = deptName;
                    row["UPDATE - Section Name"] = "";
                    row["Section: Color Hex # (optional)"] = sec?.color || "";
                } else {
                    row["Section: Name"] = sec?.name || "";
                    row["Section: Department"] = deptName;
                    row["Section: Color Hex # (optional)"] = sec?.color || "";
                }
            }
            if (includePositions) {
                if (appMode === 'delete') {
                    row["Positions"] = activePositions[i]?.name || "";
                } else if (appMode === 'update') {
                    const pos = activePositions[i];
                    const deptName = pos?.departmentId ? (currentStructure.departments.find(d => d.id === pos.departmentId)?.name || "") : "";
                    const secName = pos?.sectionId ? (activeSections.find(s => s.id === pos.sectionId)?.name || "") : "";
                    const revUnitName = pos?.revenueUnitId ? (currentStructure.revenueUnits?.find(r => r.id === pos.revenueUnitId)?.name || "") : "";
                    let skillNames = "";
                    if (pos?.skillIds && currentStructure.skills) {
                        skillNames = currentStructure.skills.filter(s => pos.skillIds?.includes(s.id)).map(s => s.name).join(', ');
                    }
                    row["Positions"] = pos?.name || "";
                    row["UPDATE - Position Name"] = "";
                    row["Position: Department (optional)"] = deptName;
                    row["Position: Section (optional)"] = secName;
                    row["Position: Color Hex # (optional)"] = pos?.color || "";
                    row["Position: Valid from (optional)"] = pos?.validFrom || "";
                    row["Position: Valid to (optional)"] = pos?.validTo || "";
                    row["Position: Skills (optional)"] = skillNames;
                    row["Position: Include in Payroll % - true/false"] = pos?.affectRevenue !== undefined ? String(pos.affectRevenue) : "true";
                    row["Position: Revenue Unit"] = revUnitName;
                } else {
                    const pos = activePositions[i];
                    const deptName = pos?.departmentId ? (currentStructure.departments.find(d => d.id === pos.departmentId)?.name || "") : "";
                    const groupName = pos?.employeeGroupId ? (currentStructure.employeeGroups.find(g => g.id === pos.employeeGroupId)?.name || "") : "";
                    const secName = pos?.sectionId ? (activeSections.find(s => s.id === pos.sectionId)?.name || "") : "";
                    const revUnitName = pos?.revenueUnitId ? (currentStructure.revenueUnits?.find(r => r.id === pos.revenueUnitId)?.name || "") : "";
                    let skillNames = "";
                    if (pos?.skillIds && currentStructure.skills) {
                        skillNames = currentStructure.skills.filter(s => pos.skillIds?.includes(s.id)).map(s => s.name).join(', ');
                    }
                    row["Position: Department"] = deptName;
                    row["Position: Section (optional)"] = secName;
                    row["Position: Employee Group"] = groupName;
                    row["Position: Name"] = pos?.name || "";
                    row["Position: Color Hex # (optional)"] = pos?.color || "";
                    row["Position: Valid from (optional)"] = pos?.validFrom || "";
                    row["Position: Valid to (optional)"] = pos?.validTo || "";
                    row["Position: Skills (optional)"] = skillNames;
                    row["Position: Include in Payroll % - true/false"] = pos?.affectRevenue !== undefined ? String(pos.affectRevenue) : "";
                    row["Position: Revenue Unit"] = revUnitName;
                }
            }
            if (includeSkills) {
                if (appMode === 'delete') {
                    row["Skills"] = currentStructure.skills[i]?.name || "";
                } else {
                    row["Skill: Name"] = currentStructure.skills[i]?.name || "";
                    row["Skill: Description"] = currentStructure.skills[i]?.description || "";
                }
            }
            data.push(row);
        }

        const ws = XLSX.utils.json_to_sheet(data);

        // Styling: Each structure type has a distinct dark background color for clear visual recognition
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
            if (ws[cellAddress]) {
                const headerText = String(ws[cellAddress].v || '');
                const bg = getHeaderBgColor(headerText);
                ws[cellAddress].s = {
                    font: { bold: true, color: { rgb: "FFFFFF" } },
                    fill: { fgColor: { rgb: bg } },
                    alignment: { horizontal: "center" }
                };
            }
        }

        const colCount = Math.max(Object.keys(data[0] || {}).length, 2);
        ws['!cols'] = Array(colCount).fill({ wch: 32 });

        const mainSheetName = appMode === 'delete' ? "Items to Delete" : appMode === 'update' ? "Planday Update" : "Planday Structure";
        XLSX.utils.book_append_sheet(wb, ws, mainSheetName);

        // Always include second sheet for "Reference Data" listing all API IDs next to descriptive names reflecting selection
        const refWs = buildReferenceDataSheet();
        XLSX.utils.book_append_sheet(wb, refWs, "Reference Data");

        const dateStr = new Date().toISOString().slice(0, 10);
        const fileName = appMode === 'delete'
            ? `Planday_Delete_Structure_${dateStr}.xlsx`
            : appMode === 'update'
                ? `Planday_Update_Structure_${dateStr}.xlsx`
                : `Planday_Structure_${dateStr}.xlsx`;
        XLSX.writeFile(wb, fileName);
    };

    // Step 2: Download Blank Excel Template
    const handleDownloadBlankTemplate = () => {
        const wb = XLSX.utils.book_new();
        const headers: string[] = [];
        if (includeDepartments) {
            if (appMode === 'delete') {
                headers.push("Departments");
            } else if (appMode === 'update') {
                headers.push("Departments", "UPDATE - Department Name", "Department Number (optional)");
            } else {
                headers.push("Departments", "Department Number (optional)");
            }
        }
        if (includeEmployeeGroups) {
            if (appMode === 'delete') {
                headers.push("Employee Groups");
            } else if (appMode === 'update') {
                headers.push("Employee Groups");
                if (includeWageColumns) {
                    headers.push("Group: Default Wage Type (HourlyRate/ShiftRate)", "Group: Default Rate", "Group: Salary Code", "Group: Salary Code 2 (needs activation)");
                }
            } else {
                headers.push("Employee Groups");
                if (includeWageColumns) {
                    headers.push("Group: Default Wage Type (HourlyRate/ShiftRate)", "Group: Default Rate", "Group: Salary Code", "Group: Salary Code 2 (needs activation)");
                }
            }
        }
        if (includeSections) {
            if (appMode === 'delete') {
                headers.push("Section: Name", "Section: Department");
            } else if (appMode === 'update') {
                headers.push("Section: Name", "Section: Department", "UPDATE - Section Name", "Section: Color Hex # (optional)");
            } else {
                headers.push("Section: Name", "Section: Department", "Section: Color Hex # (optional)");
            }
        }
        if (includePositions) {
            if (appMode === 'delete') {
                headers.push("Positions");
            } else if (appMode === 'update') {
                headers.push(
                    "Positions", 
                    "UPDATE - Position Name",
                    "Position: Department (optional)",
                    "Position: Section (optional)",
                    "Position: Color Hex # (optional)",
                    "Position: Valid from (optional)",
                    "Position: Valid to (optional)",
                    "Position: Skills (optional)",
                    "Position: Include in Payroll % - true/false",
                    "Position: Revenue Unit"
                );
            } else {
                headers.push(
                    "Position: Department",
                    "Position: Section (optional)",
                    "Position: Employee Group",
                    "Position: Name",
                    "Position: Color Hex # (optional)",
                    "Position: Valid from (optional)",
                    "Position: Valid to (optional)",
                    "Position: Skills (optional)",
                    "Position: Include in Payroll % - true/false",
                    "Position: Revenue Unit"
                );
            }
        }
        if (includeSkills) {
            if (appMode === 'delete') {
                headers.push("Skills");
            } else {
                headers.push("Skill: Name", "Skill: Description");
            }
        }

        const sampleRows = [
            headers.reduce((acc, h) => {
                acc[h] = "";
                if (h === "Departments") acc[h] = "e.g. Care Home Burnley";
                else if (h === "UPDATE - Department Name") acc[h] = "e.g. Care Home Burnley North";
                else if (h === "Department Number (optional)") acc[h] = "e.g. 101";
                else if (h === "Section: Name") acc[h] = "e.g. Ground Floor";
                else if (h === "Section: Department") acc[h] = "e.g. Care Home Burnley";
                else if (h === "UPDATE - Section Name") acc[h] = "e.g. Ground Floor North";
                else if (h === "Section: Color Hex # (optional)") acc[h] = "#2CD08E";
                else if (h === "Employee Groups") acc[h] = "e.g. General Nurses";
                else if (h === "Skills") acc[h] = "e.g. Drivers License";
                else if (h === "Positions") acc[h] = "e.g. Senior Nurse";
                else if (h === "UPDATE - Position Name") acc[h] = "e.g. Lead Nurse";
                else if (h === "Group: Default Wage Type (HourlyRate/ShiftRate)" || h === "Default Wage Type (HourlyRate/ShiftRate)") acc[h] = "HourlyRate";
                else if (h === "Group: Default Rate" || h === "Rate") acc[h] = "15.50";
                else if (h === "Group: Salary Code" || h === "Salary Code") acc[h] = "SC1";
                else if (h === "Group: Salary Code 2 (needs activation)" || h === "Salary Code 2 (Rare)") acc[h] = "";
                else if (h === "Skill: Name") acc[h] = "e.g. Drivers License";
                else if (h === "Skill: Description") acc[h] = "Valid category B license";
                else if (h === "Position: Department") acc[h] = "e.g. Care Home Burnley";
                else if (h === "Position: Employee Group") acc[h] = "e.g. General Nurses";
                else if (h === "Position: Name") acc[h] = "e.g. General Nurse";
                return acc;
            }, {} as Record<string, string>)
        ];

        const ws = XLSX.utils.json_to_sheet(sampleRows);

        // Styling: Each structure type has a distinct dark background color for clear visual recognition
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
            if (ws[cellAddress]) {
                const headerText = String(ws[cellAddress].v || '');
                const bg = getHeaderBgColor(headerText);
                ws[cellAddress].s = {
                    font: { bold: true, color: { rgb: "FFFFFF" } },
                    fill: { fgColor: { rgb: bg } },
                    alignment: { horizontal: "center" }
                };
            }
        }

        const colCount = Math.max(headers.length, 2);
        ws['!cols'] = Array(colCount).fill({ wch: 32 });

        const mainSheetName = appMode === 'delete' ? "Items to Delete" : appMode === 'update' ? "Planday Update" : "Structure Template";
        XLSX.utils.book_append_sheet(wb, ws, mainSheetName);

        // Include "Reference Data" sheet reflecting current selections
        const refWs = buildReferenceDataSheet();
        XLSX.utils.book_append_sheet(wb, refWs, "Reference Data");

        const fileName = appMode === 'delete'
            ? `Planday_Delete_Template.xlsx`
            : appMode === 'update'
                ? `Planday_Update_Template.xlsx`
                : `Planday_Structure_Template.xlsx`;
        XLSX.writeFile(wb, fileName);
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
                // Prefer data sheet if workbook contains "Reference Data"
                const wsname = wb.SheetNames.find(n => n.toLowerCase() !== 'reference data') || wb.SheetNames[0];
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
        const existingGroupNames = new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));
        const existingSkillNames = new Set(currentStructure.skills.map(s => s.name.trim().toLowerCase()));

        interface DeptImportData {
            lookupName: string;
            updateName?: string | null;
            number?: string | null;
            hasNumber: boolean;
        }
        interface SectionImportData {
            name: string;
            departmentName: string;
            updateName?: string | null;
            color?: string | null;
        }

        const newDeptsMap = new Map<string, DeptImportData>();
        const newSectionsMap = new Map<string, SectionImportData>();
        const newGroupsMap = new Map<string, any>();
        const newSkillsMap = new Map<string, string>();
        const newPositionsList: any[] = [];
        const warnings: string[] = [];
        
        const dateKeys = [posValidFromColumn, posValidToColumn].filter(Boolean);
        const isUSFormat = checkForUSDateFormat(parsedRows, dateKeys);

        parsedRows.forEach((row, index) => {
                const rowNum = index + 2; // +1 for 0-index, +1 for header row
                
                // Department
                const hasDeptName = !!(deptColumn && row[deptColumn]);
                const hasDeptUpdateName = !!(deptUpdateNameColumn && row[deptUpdateNameColumn]);
                const hasDeptNum = !!(deptNumberColumn && row[deptNumberColumn] !== undefined && row[deptNumberColumn] !== null && String(row[deptNumberColumn]).trim() !== '');
                
                if (hasDeptNum && !hasDeptName && !hasDeptUpdateName) {
                    warnings.push(`Row ${rowNum}: Excluded Department row. Missing Department Name identifier. Entered: [Dept Number: '${row[deptNumberColumn]}']`);
                }

                const rawVal = hasDeptName ? String(row[deptColumn]).trim() : '';
                const updateNameVal = hasDeptUpdateName ? String(row[deptUpdateNameColumn]).trim() : null;
                let deptNum: string | null = null;
                if (hasDeptNum) {
                    deptNum = String(row[deptNumberColumn]).trim();
                }

                if (rawVal) {
                    const key = rawVal.toLowerCase();
                    const existingEntry = newDeptsMap.get(key) || {
                        lookupName: rawVal,
                        updateName: null,
                        number: null,
                        hasNumber: false
                    };
                    if (updateNameVal) {
                        existingEntry.updateName = updateNameVal;
                    }
                    if (hasDeptNum) {
                        existingEntry.number = deptNum;
                        existingEntry.hasNumber = true;
                    }
                    newDeptsMap.set(key, existingEntry);
                }

                // Section
                const hasSecName = !!(sectionNameColumn && row[sectionNameColumn]);
                const hasSecDept = !!(sectionDeptColumn && row[sectionDeptColumn]);
                const hasSecColor = !!(sectionColorColumn && row[sectionColorColumn]);
                const hasSecUpdateName = !!(sectionUpdateNameColumn && row[sectionUpdateNameColumn]);

                if (appMode !== 'update' && (hasSecColor || hasSecUpdateName) && (!hasSecName || !hasSecDept)) {
                    const entered = [];
                    if (hasSecName) entered.push(`Section Name: '${row[sectionNameColumn]}'`);
                    if (hasSecDept) entered.push(`Department: '${row[sectionDeptColumn]}'`);
                    if (hasSecColor) entered.push(`Color: '${row[sectionColorColumn]}'`);
                    if (hasSecUpdateName) entered.push(`Update Name: '${row[sectionUpdateNameColumn]}'`);
                    warnings.push(`Row ${rowNum}: Excluded Section row. Missing required: ${!hasSecName ? "'Section Name' " : ""}${!hasSecDept ? "'Section Department'" : ""}. Entered: [${entered.join(', ')}]`);
                }

                if (hasSecName && (hasSecDept || appMode === 'update')) {
                    const sName = String(row[sectionNameColumn]).trim();
                    const dName = hasSecDept ? String(row[sectionDeptColumn]).trim() : '';
                    const cVal = hasSecColor ? String(row[sectionColorColumn]).trim() : null;
                    const uName = hasSecUpdateName ? String(row[sectionUpdateNameColumn]).trim() : null;
                    if (sName) {
                        const key = dName ? `${dName.toLowerCase()}:::${sName.toLowerCase()}` : `:::${sName.toLowerCase()}`;
                        newSectionsMap.set(key, {
                            name: sName,
                            departmentName: dName,
                            color: cVal,
                            updateName: uName
                        });
                    }
                }
                
                // Employee Group
                const hasGroupName = !!(groupColumn && row[groupColumn]);
                const hasWageType = !!(wageTypeColumn && row[wageTypeColumn]);
                const hasRate = !!(rateColumn && row[rateColumn]);
                const hasSalCode = !!(salaryCodeColumn && row[salaryCodeColumn]);
                const hasSalCode2 = !!(salaryCode2Column && row[salaryCode2Column]);

                if ((hasWageType || hasRate || hasSalCode || hasSalCode2) && !hasGroupName) {
                    const entered = [];
                    if (hasWageType) entered.push(`Wage Type: '${row[wageTypeColumn]}'`);
                    if (hasRate) entered.push(`Rate: '${row[rateColumn]}'`);
                    if (hasSalCode) entered.push(`Salary Code: '${row[salaryCodeColumn]}'`);
                    if (hasSalCode2) entered.push(`Salary Code 2: '${row[salaryCode2Column]}'`);
                    warnings.push(`Row ${rowNum}: Excluded Employee Group creation. Missing required: 'Employee Group Name'. Entered: [${entered.join(', ')}]`);
                }

                if (hasGroupName) {
                    const val = String(row[groupColumn]).trim();
                    if (val) {
                        const groupData: any = newGroupsMap.get(val) || {};
                        if (hasWageType) groupData.wageType = String(row[wageTypeColumn]).trim();
                        if (hasRate) groupData.rate = String(row[rateColumn]).trim();
                        if (hasSalCode) groupData.salaryCode = String(row[salaryCodeColumn]).trim();
                        if (hasSalCode2) groupData.salaryCode2 = String(row[salaryCode2Column]).trim();
                        newGroupsMap.set(val, groupData);
                    }
                }
                
                // Skills
                const hasSkillDesc = !!(skillDescColumn && row[skillDescColumn]);
                const hasSkillName = !!(skillNameColumn && row[skillNameColumn]);
                
                if (hasSkillDesc && !hasSkillName) {
                    warnings.push(`Row ${rowNum}: Excluded Skill creation. Missing required: 'Skill Name'. Entered: [Skill Description: '${row[skillDescColumn]}']`);
                }

                if (hasSkillName) {
                    const val = String(row[skillNameColumn]).trim();
                    if (val) {
                        const desc = hasSkillDesc ? String(row[skillDescColumn]).trim() : '';
                        if (!newSkillsMap.has(val) || desc) {
                            newSkillsMap.set(val, desc);
                        }
                    }
                }
                
                // Positions
                const hasPosName = !!(posNameColumn && row[posNameColumn]);
                const hasPosDept = !!(posDeptColumn && row[posDeptColumn]);
                const hasPosGroup = !!(posGroupColumn && row[posGroupColumn]);
                const hasPosUpdateName = !!(posUpdateNameColumn && row[posUpdateNameColumn]);
                
                const hasAnyPosData = hasPosName || hasPosDept || hasPosGroup || hasPosUpdateName ||
                                      !!(posSectionColumn && row[posSectionColumn]) ||
                                      !!(posColorColumn && row[posColorColumn]) ||
                                      !!(posValidFromColumn && row[posValidFromColumn]) ||
                                      !!(posValidToColumn && row[posValidToColumn]) ||
                                      !!(posSkillsColumn && row[posSkillsColumn]) ||
                                      !!(posPayrollColumn && row[posPayrollColumn]) ||
                                      !!(posRevenueUnitColumn && row[posRevenueUnitColumn]);

                if (hasAnyPosData) {
                    if (appMode === 'update') {
                        if (hasPosName) {
                            newPositionsList.push({
                                name: String(row[posNameColumn]).trim(),
                                updateName: hasPosUpdateName ? String(row[posUpdateNameColumn]).trim() : undefined,
                                departmentName: hasPosDept ? String(row[posDeptColumn]).trim() : '',
                                sectionName: posSectionColumn && row[posSectionColumn] ? String(row[posSectionColumn]).trim() : '',
                                employeeGroupName: hasPosGroup ? String(row[posGroupColumn]).trim() : '',
                                color: posColorColumn ? (row[posColorColumn] ? String(row[posColorColumn]).trim() : null) : undefined,
                                validFrom: posValidFromColumn ? (row[posValidFromColumn] ? parseDateWithFormat(row[posValidFromColumn], isUSFormat) : null) : undefined,
                                validTo: posValidToColumn ? (row[posValidToColumn] ? parseDateWithFormat(row[posValidToColumn], isUSFormat) : null) : undefined,
                                skills: posSkillsColumn ? (row[posSkillsColumn] ? String(row[posSkillsColumn]).trim() : '') : undefined,
                                affectRevenue: posPayrollColumn ? (row[posPayrollColumn] !== undefined && row[posPayrollColumn] !== '' ? (String(row[posPayrollColumn]).trim().toLowerCase() !== 'false' && String(row[posPayrollColumn]).trim() !== '0') : true) : undefined,
                                revenueUnitName: posRevenueUnitColumn ? (row[posRevenueUnitColumn] ? String(row[posRevenueUnitColumn]).trim() : '') : undefined
                            });
                        } else {
                            warnings.push(`Row ${rowNum}: Excluded Position update. Missing required: 'Position Name' to match.`);
                        }
                    } else {
                        if (hasPosName && hasPosDept && hasPosGroup) {
                            newPositionsList.push({
                                name: String(row[posNameColumn]).trim(),
                                updateName: hasPosUpdateName ? String(row[posUpdateNameColumn]).trim() : undefined,
                                departmentName: String(row[posDeptColumn]).trim(),
                                sectionName: posSectionColumn && row[posSectionColumn] ? String(row[posSectionColumn]).trim() : '',
                                employeeGroupName: String(row[posGroupColumn]).trim(),
                                color: posColorColumn && row[posColorColumn] ? String(row[posColorColumn]).trim() : undefined,
                                validFrom: posValidFromColumn && row[posValidFromColumn] ? parseDateWithFormat(row[posValidFromColumn], isUSFormat) : undefined,
                                validTo: posValidToColumn && row[posValidToColumn] ? parseDateWithFormat(row[posValidToColumn], isUSFormat) : undefined,
                                skills: posSkillsColumn && row[posSkillsColumn] ? String(row[posSkillsColumn]).trim() : '',
                                affectRevenue: posPayrollColumn && row[posPayrollColumn] ? (String(row[posPayrollColumn]).trim().toLowerCase() !== 'false') : true,
                                revenueUnitName: posRevenueUnitColumn && row[posRevenueUnitColumn] ? String(row[posRevenueUnitColumn]).trim() : ''
                            });
                        } else {
                            const missing = [];
                            if (!hasPosName) missing.push('Name');
                            if (!hasPosDept) missing.push('Department');
                            if (!hasPosGroup) missing.push('Employee Group');
                            
                            const entered = [];
                            if (hasPosName) entered.push(`Name: '${row[posNameColumn]}'`);
                            if (hasPosDept) entered.push(`Department: '${row[posDeptColumn]}'`);
                            if (hasPosGroup) entered.push(`Employee Group: '${row[posGroupColumn]}'`);
                            if (posSectionColumn && row[posSectionColumn]) entered.push(`Section: '${row[posSectionColumn]}'`);
                            if (posColorColumn && row[posColorColumn]) entered.push(`Color: '${row[posColorColumn]}'`);
                            if (posValidFromColumn && row[posValidFromColumn]) entered.push(`Valid From: '${row[posValidFromColumn]}'`);
                            if (posValidToColumn && row[posValidToColumn]) entered.push(`Valid To: '${row[posValidToColumn]}'`);
                            if (posSkillsColumn && row[posSkillsColumn]) entered.push(`Skills: '${row[posSkillsColumn]}'`);
                            if (posPayrollColumn && row[posPayrollColumn]) entered.push(`Include in Payroll: '${row[posPayrollColumn]}'`);
                            if (posRevenueUnitColumn && row[posRevenueUnitColumn]) entered.push(`Revenue Unit: '${row[posRevenueUnitColumn]}'`);
                            
                            warnings.push(`Row ${rowNum}: Excluded Position creation. Missing required: ${missing.join(', ')}. Entered: [${entered.join(', ')}]`);
                        }
                    }
                }
            });

        const items: StructureReviewItem[] = [];

        // Departments from file
        newDeptsMap.forEach((deptData) => {
            const name = deptData.lookupName;
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            const existingDept = currentStructure.departments.find(d => d.name.trim().toLowerCase() === name.toLowerCase());
            
            if (appMode === 'update') {
                if (existingDept) {
                    const targetName = (deptData.updateName && deptData.updateName.trim()) 
                        ? deptData.updateName.trim() 
                        : existingDept.name;
                    const targetNum = deptData.hasNumber 
                        ? (deptData.number ? String(deptData.number).trim() : null)
                        : (existingDept.number || null);

                    const hasNameChange = targetName !== existingDept.name;
                    const hasNumChange = targetNum !== (existingDept.number || null);

                    let msg = 'Update scheduled';
                    if (hasNameChange && hasNumChange) {
                        msg = `Rename to '${targetName}' & update number to '${targetNum || 'None'}'`;
                    } else if (hasNameChange) {
                        msg = `Rename to '${targetName}'`;
                    } else if (hasNumChange) {
                        msg = `Update number to '${targetNum || 'None'}'`;
                    } else {
                        msg = 'Existing values match file';
                    }

                    items.push({
                        id: existingDept.id,
                        type: 'department',
                        name: existingDept.name,
                        number: existingDept.number || null,
                        action: 'update',
                        status: 'pending',
                        updatePayload: { 
                            name: targetName, 
                            number: targetNum 
                        },
                        resultMessage: msg
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
                const targetName = deptData.updateName || name;
                const targetNum = deptData.hasNumber ? deptData.number : null;

                if (isDuplicate) {
                    if (existingDept && (
                        (targetNum !== null && existingDept.number !== targetNum) ||
                        (deptData.updateName && deptData.updateName !== existingDept.name)
                    )) {
                        items.push({
                            id: existingDept.id,
                            type: 'department',
                            name: existingDept.name,
                            number: existingDept.number || null,
                            action: 'update',
                            status: 'pending',
                            updatePayload: {
                                name: targetName,
                                number: targetNum !== null ? targetNum : (existingDept.number || null)
                            },
                            resultMessage: deptData.updateName 
                                ? `Updating name to '${targetName}' from file` 
                                : 'Updating number from file'
                        });
                    } else {
                        items.push({
                            type: 'department',
                            name,
                            number: targetNum,
                            action: 'skip_duplicate',
                            status: 'skipped',
                            resultMessage: 'Already exists in Planday'
                        });
                    }
                } else {
                    items.push({
                        type: 'department',
                        name: targetName,
                        number: targetNum,
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
                if (isDuplicate) {
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
        });

        // Sections from file
        const isSectionColorMapped = !!sectionColorColumn;
        const isSectionUpdateNameMapped = !!sectionUpdateNameColumn;

        newSectionsMap.forEach((secData) => {
            let matchedDept = secData.departmentName ? currentStructure.departments.find(d => 
                d.name.trim().toLowerCase() === secData.departmentName.toLowerCase() ||
                String(d.id) === secData.departmentName
            ) : undefined;
            const deptBeingCreated = secData.departmentName ? newDeptsMap.has(secData.departmentName.toLowerCase()) : false;

            let existingSec = (currentStructure.sections || []).find(s => {
                const matchesName = s.name.trim().toLowerCase() === secData.name.toLowerCase();
                if (!matchesName) return false;
                if (matchedDept && s.departmentId === matchedDept.id) return true;
                return false;
            });

            // If not found yet and department wasn't specified (or didn't match), check if there is a matching section by name
            if (!existingSec) {
                const candidateSecs = (currentStructure.sections || []).filter(s => 
                    s.name.trim().toLowerCase() === secData.name.toLowerCase()
                );
                if (candidateSecs.length === 1) {
                    existingSec = candidateSecs[0];
                    if (!matchedDept) {
                        matchedDept = currentStructure.departments.find(d => d.id === existingSec!.departmentId);
                    }
                }
            }

            const isDuplicate = !!existingSec;

            if (appMode === 'update') {
                if (existingSec && matchedDept) {
                    const targetName = (isSectionUpdateNameMapped && secData.updateName && secData.updateName.trim()) 
                        ? secData.updateName.trim() 
                        : existingSec.name;
                    
                    const nameChanged = isSectionUpdateNameMapped && targetName !== existingSec.name;
                    
                    let targetColor: string | null | undefined = undefined;
                    let colorChanged = false;
                    if (isSectionColorMapped) {
                        targetColor = secData.color ? secData.color.trim() : null;
                        colorChanged = targetColor !== (existingSec.color || null);
                    }

                    let msg = 'Update scheduled';
                    if (nameChanged && colorChanged) {
                        msg = `Rename to '${targetName}' & update color to '${targetColor || 'None'}'`;
                    } else if (nameChanged) {
                        msg = `Rename to '${targetName}'`;
                    } else if (colorChanged) {
                        msg = `Update color to '${targetColor || 'None'}'`;
                    } else if (!nameChanged && !colorChanged && (isSectionUpdateNameMapped || isSectionColorMapped)) {
                        msg = 'Existing values match file';
                    }

                    const updatePayload: any = {
                        departmentId: matchedDept.id,
                        departmentName: matchedDept.name,
                        name: targetName
                    };
                    // ONLY include color in updatePayload if color column was mapped!
                    if (isSectionColorMapped) {
                        updatePayload.color = targetColor;
                    }

                    items.push({
                        id: existingSec.id,
                        type: 'section',
                        name: existingSec.name,
                        action: 'update',
                        status: 'pending',
                        updatePayload,
                        resultMessage: msg
                    });
                } else {
                    items.push({
                        type: 'section',
                        name: secData.name,
                        action: 'skip_duplicate',
                        status: 'error',
                        resultMessage: secData.departmentName && !matchedDept 
                            ? `Department '${secData.departmentName}' not found in Planday` 
                            : `Section '${secData.name}' not found in Planday`
                    });
                }
            } else {
                // Import mode
                if (isDuplicate && existingSec && matchedDept) {
                    const targetColor = isSectionColorMapped ? (secData.color ? secData.color.trim() : null) : (existingSec.color || null);
                    const targetName = (isSectionUpdateNameMapped && secData.updateName) ? secData.updateName.trim() : existingSec.name;
                    if ((isSectionColorMapped && targetColor !== (existingSec.color || null)) || (isSectionUpdateNameMapped && targetName !== existingSec.name)) {
                        const updatePayload: any = {
                            departmentId: matchedDept.id,
                            departmentName: matchedDept.name,
                            name: targetName
                        };
                        if (isSectionColorMapped) {
                            updatePayload.color = targetColor;
                        }
                        items.push({
                            id: existingSec.id,
                            type: 'section',
                            name: existingSec.name,
                            action: 'update',
                            status: 'pending',
                            updatePayload,
                            resultMessage: (isSectionUpdateNameMapped && targetName !== existingSec.name) ? `Updating name to '${targetName}' from file` : 'Updating color from file'
                        });
                    } else {
                        items.push({
                            id: existingSec.id,
                            type: 'section',
                            name: secData.name,
                            action: 'skip_duplicate',
                            status: 'skipped',
                            resultMessage: `Already exists in Department '${matchedDept.name}'`
                        });
                    }
                } else {
                    items.push({
                        type: 'section',
                        name: secData.updateName || secData.name,
                        action: 'create',
                        status: 'pending',
                        updatePayload: {
                            departmentName: matchedDept ? matchedDept.name : secData.departmentName,
                            departmentId: matchedDept?.id,
                            color: secData.color || null
                        },
                        resultMessage: !matchedDept && !deptBeingCreated 
                            ? `Warning: Department '${secData.departmentName}' not found in Planday` 
                            : undefined
                    });
                }
            }
        });

        // Positions from file
        const isPosUpdateNameMapped = !!posUpdateNameColumn;
        const isPosColorMapped = !!posColorColumn;
        const isPosSectionMapped = !!posSectionColumn;
        const isPosValidFromMapped = !!posValidFromColumn;
        const isPosValidToMapped = !!posValidToColumn;
        const isPosSkillsMapped = !!posSkillsColumn;
        const isPosPayrollMapped = !!posPayrollColumn;
        const isPosRevenueUnitMapped = !!posRevenueUnitColumn;

        newPositionsList.forEach(pos => {
            if (appMode === 'update') {
                let matchedPos: Position | undefined = undefined;
                let candidatePositions = currentStructure.positions || [];

                if (pos.departmentName) {
                    const matchedDept = currentStructure.departments.find(d => 
                        d.name.trim().toLowerCase() === pos.departmentName.toLowerCase() ||
                        String(d.id) === pos.departmentName
                    );
                    if (matchedDept) {
                        candidatePositions = candidatePositions.filter(p => p.departmentId === matchedDept.id);
                    }
                }

                if (pos.employeeGroupName && candidatePositions.length > 0) {
                    const matchedGroup = currentStructure.employeeGroups.find(g => 
                        g.name.trim().toLowerCase() === pos.employeeGroupName.toLowerCase() ||
                        String(g.id) === pos.employeeGroupName
                    );
                    if (matchedGroup) {
                        const grpFiltered = candidatePositions.filter(p => p.employeeGroupId === matchedGroup.id);
                        if (grpFiltered.length > 0) {
                            candidatePositions = grpFiltered;
                        }
                    }
                }

                matchedPos = candidatePositions.find(p => p.name.trim().toLowerCase() === pos.name.toLowerCase());

                if (!matchedPos) {
                    matchedPos = (currentStructure.positions || []).find(p => p.name.trim().toLowerCase() === pos.name.toLowerCase());
                }

                if (!matchedPos) {
                    items.push({
                        type: 'position',
                        name: pos.name,
                        action: 'skip_duplicate',
                        status: 'error',
                        resultMessage: pos.departmentName 
                            ? `Position '${pos.name}' not found in Department '${pos.departmentName}'` 
                            : `Position '${pos.name}' not found in Planday`
                    });
                    return;
                }

                const targetName = (isPosUpdateNameMapped && pos.updateName && pos.updateName.trim()) 
                    ? pos.updateName.trim() 
                    : matchedPos.name;

                // affectRevenue is always required for updating positions. Set to true if not specified.
                const targetAffectRevenue = (isPosPayrollMapped && pos.affectRevenue !== undefined && pos.affectRevenue !== null)
                    ? Boolean(pos.affectRevenue)
                    : true;

                const payload: any = {
                    name: targetName,
                    affectRevenue: targetAffectRevenue
                };

                let targetSectionName: string | undefined = undefined;
                if (isPosSectionMapped) {
                    if (pos.sectionName) {
                        targetSectionName = pos.sectionName;
                        const matchedSec = (currentStructure.sections || []).find(s => 
                            s.name.trim().toLowerCase() === pos.sectionName.toLowerCase() &&
                            (!matchedPos?.departmentId || s.departmentId === matchedPos.departmentId)
                        ) || (currentStructure.sections || []).find(s => s.name.trim().toLowerCase() === pos.sectionName.toLowerCase());
                        payload.sectionId = matchedSec ? matchedSec.id : null;
                    } else {
                        payload.sectionId = null;
                    }
                }

                if (isPosColorMapped) {
                    payload.color = pos.color ? pos.color.trim() : null;
                }

                if (isPosValidFromMapped) {
                    payload.validFrom = pos.validFrom || null;
                }

                if (isPosValidToMapped) {
                    payload.validTo = pos.validTo || null;
                }

                if (isPosSkillsMapped) {
                    if (pos.skills && String(pos.skills).trim()) {
                        const skillTokens = String(pos.skills).split(',').map((s: string) => s.trim().toLowerCase());
                        const skillIds: number[] = [];
                        (currentStructure.skills || []).forEach(sk => {
                            if (skillTokens.includes(sk.name.toLowerCase()) || skillTokens.includes(String(sk.id))) {
                                skillIds.push(sk.id);
                            }
                        });
                        payload.skillIds = skillIds;
                    } else {
                        payload.skillIds = [];
                    }
                }

                let targetRevUnitName: string | undefined = undefined;
                if (isPosRevenueUnitMapped) {
                    if (pos.revenueUnitName) {
                        targetRevUnitName = pos.revenueUnitName;
                        const matchedRev = (currentStructure.revenueUnits || []).find(r => 
                            r.name.trim().toLowerCase() === pos.revenueUnitName.toLowerCase() &&
                            (!matchedPos?.departmentId || r.departmentId === matchedPos.departmentId)
                        ) || (currentStructure.revenueUnits || []).find(r => r.name.trim().toLowerCase() === pos.revenueUnitName.toLowerCase());
                        payload.revenueUnitId = matchedRev ? matchedRev.id : null;
                    } else {
                        payload.revenueUnitId = null;
                    }
                }

                const changes: string[] = [];
                if (isPosUpdateNameMapped && targetName !== matchedPos.name) {
                    changes.push(`Rename to '${targetName}'`);
                }
                if (isPosColorMapped) {
                    changes.push(`Color: '${payload.color || 'None'}'`);
                }
                if (isPosSectionMapped) {
                    changes.push(`Section: '${pos.sectionName || 'None'}'`);
                }
                if (isPosPayrollMapped) {
                    changes.push(`Payroll: ${payload.affectRevenue}`);
                }
                if (isPosValidFromMapped) {
                    changes.push(`ValidFrom: '${payload.validFrom || 'None'}'`);
                }
                if (isPosValidToMapped) {
                    changes.push(`ValidTo: '${payload.validTo || 'None'}'`);
                }
                if (isPosSkillsMapped) {
                    changes.push(`Skills: '${pos.skills || 'None'}'`);
                }
                if (isPosRevenueUnitMapped) {
                    changes.push(`RevUnit: '${pos.revenueUnitName || 'None'}'`);
                }

                const matchedDeptObj = currentStructure.departments.find(d => d.id === matchedPos!.departmentId);
                const matchedGroupObj = currentStructure.employeeGroups.find(g => g.id === matchedPos!.employeeGroupId);

                items.push({
                    id: matchedPos.id,
                    type: 'position',
                    name: matchedPos.name,
                    action: 'update',
                    status: 'pending',
                    updatePayload: {
                        ...payload,
                        departmentName: matchedDeptObj?.name || pos.departmentName || '',
                        employeeGroupName: matchedGroupObj?.name || pos.employeeGroupName || '',
                        sectionName: targetSectionName,
                        skills: pos.skills,
                        revenueUnitName: targetRevUnitName
                    },
                    resultMessage: changes.length > 0 ? changes.join(', ') : 'Update scheduled'
                });
            } else {
                // Import mode: check if position already exists in this Department & Employee Group
                let matchedPos: Position | undefined = undefined;
                let candidatePositions = currentStructure.positions || [];

                if (pos.departmentName) {
                    const matchedDept = currentStructure.departments.find(d => 
                        d.name.trim().toLowerCase() === pos.departmentName.toLowerCase() ||
                        String(d.id) === pos.departmentName
                    );
                    if (matchedDept) {
                        candidatePositions = candidatePositions.filter(p => p.departmentId === matchedDept.id);
                    }
                }

                if (pos.employeeGroupName && candidatePositions.length > 0) {
                    const matchedGroup = currentStructure.employeeGroups.find(g => 
                        g.name.trim().toLowerCase() === pos.employeeGroupName.toLowerCase() ||
                        String(g.id) === pos.employeeGroupName
                    );
                    if (matchedGroup) {
                        const grpFiltered = candidatePositions.filter(p => p.employeeGroupId === matchedGroup.id);
                        if (grpFiltered.length > 0) {
                            candidatePositions = grpFiltered;
                        }
                    }
                }

                matchedPos = candidatePositions.find(p => p.name.trim().toLowerCase() === pos.name.toLowerCase());

                if (matchedPos) {
                    const targetName = (isPosUpdateNameMapped && pos.updateName && pos.updateName.trim()) 
                        ? pos.updateName.trim() 
                        : matchedPos.name;

                    const targetAffectRevenue = (isPosPayrollMapped && pos.affectRevenue !== undefined && pos.affectRevenue !== null)
                        ? Boolean(pos.affectRevenue)
                        : true;

                    const payload: any = {
                        name: targetName,
                        affectRevenue: targetAffectRevenue
                    };

                    let targetSectionName: string | undefined = undefined;
                    if (isPosSectionMapped) {
                        if (pos.sectionName) {
                            targetSectionName = pos.sectionName;
                            const matchedSec = (currentStructure.sections || []).find(s => 
                                s.name.trim().toLowerCase() === pos.sectionName.toLowerCase() &&
                                (!matchedPos?.departmentId || s.departmentId === matchedPos.departmentId)
                            ) || (currentStructure.sections || []).find(s => s.name.trim().toLowerCase() === pos.sectionName.toLowerCase());
                            payload.sectionId = matchedSec ? matchedSec.id : null;
                        } else {
                            payload.sectionId = null;
                        }
                    }

                    if (isPosColorMapped) {
                        payload.color = pos.color ? pos.color.trim() : null;
                    }

                    if (isPosValidFromMapped) {
                        payload.validFrom = pos.validFrom || null;
                    }

                    if (isPosValidToMapped) {
                        payload.validTo = pos.validTo || null;
                    }

                    if (isPosSkillsMapped) {
                        if (pos.skills && String(pos.skills).trim()) {
                            const skillTokens = String(pos.skills).split(',').map((s: string) => s.trim().toLowerCase());
                            const skillIds: number[] = [];
                            (currentStructure.skills || []).forEach(sk => {
                                if (skillTokens.includes(sk.name.toLowerCase()) || skillTokens.includes(String(sk.id))) {
                                    skillIds.push(sk.id);
                                }
                            });
                            payload.skillIds = skillIds;
                        } else {
                            payload.skillIds = [];
                        }
                    }

                    let targetRevUnitName: string | undefined = undefined;
                    if (isPosRevenueUnitMapped) {
                        if (pos.revenueUnitName) {
                            targetRevUnitName = pos.revenueUnitName;
                            const matchedRev = (currentStructure.revenueUnits || []).find(r => 
                                r.name.trim().toLowerCase() === pos.revenueUnitName.toLowerCase() &&
                                (!matchedPos?.departmentId || r.departmentId === matchedPos.departmentId)
                            ) || (currentStructure.revenueUnits || []).find(r => r.name.trim().toLowerCase() === pos.revenueUnitName.toLowerCase());
                            payload.revenueUnitId = matchedRev ? matchedRev.id : null;
                        } else {
                            payload.revenueUnitId = null;
                        }
                    }

                    const changes: string[] = [];
                    if (isPosUpdateNameMapped && targetName !== matchedPos.name) changes.push(`Rename to '${targetName}'`);
                    if (isPosColorMapped && payload.color !== (matchedPos.color || null)) changes.push(`Color: '${payload.color || 'None'}'`);
                    if (isPosSectionMapped && payload.sectionId !== (matchedPos.sectionId || null)) changes.push(`Section: '${pos.sectionName || 'None'}'`);
                    if (isPosPayrollMapped && targetAffectRevenue !== (matchedPos.affectRevenue !== undefined ? matchedPos.affectRevenue : true)) changes.push(`Payroll: ${targetAffectRevenue}`);
                    if (isPosValidFromMapped && payload.validFrom !== (matchedPos.validFrom || null)) changes.push(`ValidFrom: '${payload.validFrom || 'None'}'`);
                    if (isPosValidToMapped && payload.validTo !== (matchedPos.validTo || null)) changes.push(`ValidTo: '${payload.validTo || 'None'}'`);
                    if (isPosSkillsMapped) changes.push(`Skills: '${pos.skills || 'None'}'`);
                    if (isPosRevenueUnitMapped && payload.revenueUnitId !== (matchedPos.revenueUnitId || null)) changes.push(`RevUnit: '${pos.revenueUnitName || 'None'}'`);

                    const matchedDeptObj = currentStructure.departments.find(d => d.id === matchedPos!.departmentId);
                    const matchedGroupObj = currentStructure.employeeGroups.find(g => g.id === matchedPos!.employeeGroupId);

                    if (changes.length > 0) {
                        items.push({
                            id: matchedPos.id,
                            type: 'position',
                            name: matchedPos.name,
                            action: 'update',
                            status: 'pending',
                            updatePayload: {
                                ...payload,
                                departmentName: matchedDeptObj?.name || pos.departmentName || '',
                                employeeGroupName: matchedGroupObj?.name || pos.employeeGroupName || '',
                                sectionName: targetSectionName,
                                skills: pos.skills,
                                revenueUnitName: targetRevUnitName
                            },
                            resultMessage: changes.join(', ')
                        });
                    } else {
                        items.push({
                            id: matchedPos.id,
                            type: 'position',
                            name: matchedPos.name,
                            action: 'skip_duplicate',
                            status: 'skipped',
                            resultMessage: `Already exists in Department '${matchedDeptObj?.name || pos.departmentName}'`
                        });
                    }
                } else {
                    items.push({
                        type: 'position',
                        name: pos.name,
                        action: 'create',
                        status: 'pending',
                        updatePayload: pos
                    });
                }
            }
        });

        // Skills from file
        newSkillsMap.forEach((desc, name) => {
            const isDuplicate = existingSkillNames.has(name.toLowerCase());
            if (appMode === 'update') {
                items.push({
                    type: 'skill',
                    name,
                    action: 'skip_duplicate',
                    status: 'error',
                    resultMessage: 'Update for skills not supported'
                });
            } else {
                if (isDuplicate) {
                    items.push({
                        type: 'skill',
                        name,
                        action: 'skip_duplicate',
                        status: 'skipped',
                        resultMessage: 'Already exists in Planday'
                    });
                } else {
                    items.push({
                        type: 'skill',
                        name,
                        action: 'create',
                        status: 'pending',
                        updatePayload: { description: desc }
                    });
                }
            }
        });

        items.sort((a, b) => (TYPE_ORDER[a.type] || 99) - (TYPE_ORDER[b.type] || 99));

        setReviewItems(items);
        setValidationWarnings(warnings);
        setStep(3); // Always go to Review from Mapping
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };

    // Process File/Text/Selection and Generate Review Items (Delete / Update)
    const generateActionReviewList = () => {
        const items: StructureReviewItem[] = [];
        
        if (deleteInputMethod === 'selection') {
            currentStructure.departments.forEach(d => {
                if (deleteSelectedDepts.has(d.id)) {
                    items.push({ 
                        id: d.id, 
                        type: 'department', 
                        name: d.name, 
                        number: d.number || null,
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: appMode === 'update' ? { name: d.name, number: d.number || '' } : undefined,
                        resultMessage: appMode === 'update' ? 'Update scheduled' : 'Scheduled for deletion'
                    });
                }
            });
            currentStructure.employeeGroups.forEach(g => {
                if (deleteSelectedGroups.has(g.id)) {
                    items.push({ 
                        id: g.id, 
                        type: 'employee_group', 
                        name: g.name, 
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: appMode === 'update' ? {} : undefined,
                        resultMessage: appMode === 'update' ? 'Update scheduled' : 'Scheduled for deletion'
                    });
                }
            });
            (currentStructure.sections || []).forEach(s => {
                if (deleteSelectedSections.has(s.id)) {
                    const dept = currentStructure.departments.find(d => d.id === s.departmentId);
                    items.push({ 
                        id: s.id, 
                        type: 'section', 
                        name: s.name, 
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: { 
                            departmentId: s.departmentId, 
                            departmentName: dept?.name || `Dept #${s.departmentId}`,
                            name: s.name, 
                            color: s.color || '' 
                        },
                        resultMessage: appMode === 'update' ? 'Update scheduled' : `Scheduled for deletion (Dept: ${dept?.name || s.departmentId})`
                    });
                }
            });
            (currentStructure.positions || []).forEach(p => {
                if (deleteSelectedPositions.has(p.id)) {
                    const dept = currentStructure.departments.find(d => d.id === p.departmentId);
                    const group = currentStructure.employeeGroups.find(g => g.id === p.employeeGroupId);
                    items.push({ 
                        id: p.id, 
                        type: 'position', 
                        name: p.name, 
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: { 
                            departmentId: p.departmentId, 
                            departmentName: dept?.name || `Dept #${p.departmentId}`,
                            employeeGroupId: p.employeeGroupId,
                            employeeGroupName: group?.name || `Group #${p.employeeGroupId}`,
                            name: p.name, 
                            color: p.color || '',
                            validFrom: p.validFrom || '',
                            validTo: p.validTo || '',
                            revenueUnitId: p.revenueUnitId,
                            affectRevenue: p.affectRevenue !== undefined ? p.affectRevenue : true
                        },
                        resultMessage: appMode === 'update' ? 'Update scheduled' : `Scheduled for deletion (Dept: ${dept?.name || p.departmentId})`
                    });
                }
            });
            (currentStructure.skills || []).forEach(sk => {
                if (deleteSelectedSkills.has(sk.id)) {
                    items.push({ 
                        id: sk.id, 
                        type: 'skill', 
                        name: sk.name, 
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: { 
                            name: sk.name, 
                            description: sk.description || '' 
                        },
                        resultMessage: appMode === 'update' ? 'Update scheduled' : 'Scheduled for deletion'
                    });
                }
            });
        } else if (deleteInputMethod === 'file') {
            const dNames = new Set<string>();
            const gNames = new Set<string>();
            const sEntries: { name: string; dept?: string }[] = [];

            parsedRows.forEach(row => {
                if (deptColumn && row[deptColumn]) dNames.add(String(row[deptColumn]).trim().toLowerCase());
                if (groupColumn && row[groupColumn]) gNames.add(String(row[groupColumn]).trim().toLowerCase());
                if (sectionNameColumn && row[sectionNameColumn]) {
                    const sName = String(row[sectionNameColumn]).trim();
                    const dName = sectionDeptColumn && row[sectionDeptColumn] ? String(row[sectionDeptColumn]).trim() : '';
                    if (sName) sEntries.push({ name: sName, dept: dName });
                }
            });
            
            const matchedDepts = new Set<string>();
            currentStructure.departments.forEach(d => {
                if (dNames.has(d.name.toLowerCase())) {
                    matchedDepts.add(d.name.toLowerCase());
                    items.push({ 
                        id: d.id, 
                        type: 'department', 
                        name: d.name, 
                        number: d.number || null,
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: appMode === 'update' ? { name: d.name, number: d.number || '' } : undefined,
                        resultMessage: appMode === 'update' ? 'Update scheduled' : 'Scheduled for deletion'
                    });
                }
            });
            dNames.forEach(name => {
                if (!matchedDepts.has(name)) {
                    items.push({
                        id: 0,
                        type: 'department',
                        name: name,
                        action: 'skip_duplicate',
                        status: 'skipped',
                        resultMessage: 'Not found in Planday structure'
                    });
                }
            });

            const matchedGroups = new Set<string>();
            currentStructure.employeeGroups.forEach(g => {
                if (gNames.has(g.name.toLowerCase())) {
                    matchedGroups.add(g.name.toLowerCase());
                    items.push({ 
                        id: g.id, 
                        type: 'employee_group', 
                        name: g.name, 
                        action: appMode === 'update' ? 'update' : 'delete', 
                        status: 'pending',
                        updatePayload: appMode === 'update' ? {} : undefined,
                        resultMessage: appMode === 'update' ? 'Update scheduled' : 'Scheduled for deletion'
                    });
                }
            });
            gNames.forEach(name => {
                if (!matchedGroups.has(name)) {
                    items.push({
                        id: 0,
                        type: 'employee_group',
                        name: name,
                        action: 'skip_duplicate',
                        status: 'skipped',
                        resultMessage: 'Not found in Planday structure'
                    });
                }
            });

            // Sections from file
            const matchedSecKeys = new Set<string>();
            sEntries.forEach(entry => {
                const sNameLower = entry.name.toLowerCase();
                const dNameLower = (entry.dept || '').toLowerCase();
                const matched = (currentStructure.sections || []).filter(s => {
                    if (s.name.trim().toLowerCase() !== sNameLower) return false;
                    if (dNameLower) {
                        const d = currentStructure.departments.find(dept => dept.id === s.departmentId);
                        return d && (d.name.trim().toLowerCase() === dNameLower || String(d.id) === dNameLower);
                    }
                    return true;
                });

                if (matched.length > 0) {
                    matched.forEach(s => {
                        const key = `${s.departmentId}:::${s.id}`;
                        if (!matchedSecKeys.has(key)) {
                            matchedSecKeys.add(key);
                            const dept = currentStructure.departments.find(d => d.id === s.departmentId);
                            items.push({
                                id: s.id,
                                type: 'section',
                                name: s.name,
                                action: appMode === 'update' ? 'update' : 'delete',
                                status: 'pending',
                                updatePayload: {
                                    departmentId: s.departmentId,
                                    departmentName: dept?.name || `Dept #${s.departmentId}`,
                                    name: s.name,
                                    color: s.color || ''
                                },
                                resultMessage: appMode === 'update' ? 'Update scheduled' : `Scheduled for deletion (${dept?.name || s.departmentId})`
                            });
                        }
                    });
                } else {
                    items.push({
                        id: 0,
                        type: 'section',
                        name: entry.dept ? `${entry.dept}: ${entry.name}` : entry.name,
                        action: 'skip_duplicate',
                        status: 'skipped',
                        resultMessage: 'Not found in Planday structure'
                    });
                }
            });

            // Positions from file
            const pEntries: { name: string; dept?: string; group?: string }[] = [];
            parsedRows.forEach(row => {
                if (posNameColumn && row[posNameColumn]) {
                    const pName = String(row[posNameColumn]).trim();
                    const dName = posDeptColumn && row[posDeptColumn] ? String(row[posDeptColumn]).trim() : '';
                    const gName = posGroupColumn && row[posGroupColumn] ? String(row[posGroupColumn]).trim() : '';
                    if (pName) pEntries.push({ name: pName, dept: dName, group: gName });
                }
            });

            const matchedPosKeys = new Set<number>();
            pEntries.forEach(entry => {
                const pNameLower = entry.name.toLowerCase();
                const dNameLower = (entry.dept || '').toLowerCase();
                const gNameLower = (entry.group || '').toLowerCase();

                const matched = (currentStructure.positions || []).filter(p => {
                    if (p.name.trim().toLowerCase() !== pNameLower) return false;
                    if (dNameLower) {
                        const d = currentStructure.departments.find(dept => dept.id === p.departmentId);
                        if (!d || (d.name.trim().toLowerCase() !== dNameLower && String(d.id) !== dNameLower)) return false;
                    }
                    if (gNameLower) {
                        const g = currentStructure.employeeGroups.find(grp => grp.id === p.employeeGroupId);
                        if (!g || (g.name.trim().toLowerCase() !== gNameLower && String(g.id) !== gNameLower)) return false;
                    }
                    return true;
                });

                if (matched.length > 0) {
                    matched.forEach(p => {
                        if (!matchedPosKeys.has(p.id)) {
                            matchedPosKeys.add(p.id);
                            const dept = currentStructure.departments.find(d => d.id === p.departmentId);
                            const group = currentStructure.employeeGroups.find(g => g.id === p.employeeGroupId);
                            items.push({
                                id: p.id,
                                type: 'position',
                                name: p.name,
                                action: appMode === 'update' ? 'update' : 'delete',
                                status: 'pending',
                                updatePayload: {
                                    departmentId: p.departmentId,
                                    departmentName: dept?.name || `Dept #${p.departmentId}`,
                                    employeeGroupId: p.employeeGroupId,
                                    employeeGroupName: group?.name || `Group #${p.employeeGroupId}`,
                                    name: p.name,
                                    color: p.color || '',
                                    validFrom: p.validFrom || '',
                                    validTo: p.validTo || '',
                                    revenueUnitId: p.revenueUnitId,
                                    affectRevenue: p.affectRevenue !== undefined ? p.affectRevenue : true
                                },
                                resultMessage: appMode === 'update' ? 'Update scheduled' : `Scheduled for deletion (${dept?.name || p.departmentId})`
                            });
                        }
                    });
                } else {
                    items.push({
                        id: 0,
                        type: 'position',
                        name: entry.dept ? `${entry.dept}: ${entry.name}` : entry.name,
                        action: 'skip_duplicate',
                        status: 'skipped',
                        resultMessage: 'Not found in Planday structure'
                    });
                }
            });
        }
        items.sort((a, b) => (TYPE_ORDER[a.type] || 99) - (TYPE_ORDER[b.type] || 99));
        setReviewItems(items);
        setStep(appMode === 'update' && deleteInputMethod !== 'file' ? 2 : 3); // Config or Review
        if (appMode !== 'update' || deleteInputMethod === 'file') {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
    };

    // Add custom item in Review
    const handleAddCustomReviewItem = () => {
        if (!newItemName.trim()) return;
        const cleanName = newItemName.trim();

        if (newItemType === 'section') {
            const deptId = typeof newItemDeptId === 'number' ? newItemDeptId : (currentStructure.departments[0]?.id || 0);
            const dept = currentStructure.departments.find(d => d.id === deptId);
            const existingSec = (currentStructure.sections || []).find(s => 
                s.name.trim().toLowerCase() === cleanName.toLowerCase() && s.departmentId === deptId
            );
            const isDup = !!existingSec;

            setReviewItems(prev => [
                ...prev,
                {
                    type: 'section',
                    name: cleanName,
                    action: isDup ? 'skip_duplicate' : 'create',
                    status: isDup ? 'skipped' : 'pending',
                    updatePayload: {
                        departmentId: deptId,
                        departmentName: dept?.name || `Dept #${deptId}`,
                        color: newItemColor.trim() || null
                    },
                    resultMessage: isDup ? `Already exists in Department '${dept?.name || deptId}'` : undefined
                }
            ]);

            setNewItemName('');
            setNewItemColor('');
            return;
        }

        const existingSet = newItemType === 'department'
            ? new Set(currentStructure.departments.map(d => d.name.trim().toLowerCase()))
            : new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));

        const isDup = existingSet.has(cleanName.toLowerCase());

        setReviewItems(prev => [
            ...prev,
            {
                type: newItemType,
                name: cleanName,
                action: isDup ? 'skip_duplicate' : 'create',
                status: isDup ? 'skipped' : 'pending',
                resultMessage: isDup ? 'Already exists in Planday' : undefined
            }
        ]);

        setNewItemName('');
    };

    const handleUpdatePayloadChange = (id: number | undefined, field: string, value: any) => {
        setReviewItems(prev => prev.map(item => {
            if (item.id === id && item.action === 'update') {
                if (item.type === 'section') {
                    const currentPayload = item.updatePayload || { name: item.name, color: '' };
                    return { ...item, updatePayload: { ...currentPayload, [field]: value } };
                }
                if (item.type === 'employee_group') {
                    const currentPayload = item.updatePayload || {};
                    const newPayload = { ...currentPayload, [field]: value };
                    if (field === 'defaultWageType') {
                        if (value === 'HourlyRate') newPayload.shiftRate = undefined;
                        if (value === 'ShiftRate') newPayload.hourlyRate = undefined;
                    }
                    return { ...item, updatePayload: newPayload };
                }
                if (item.type === 'department') {
                    const currentPayload = item.updatePayload || { name: item.name, number: item.number || '' };
                    const newPayload = { ...currentPayload, [field]: value };
                    return { ...item, updatePayload: newPayload };
                }
                if (item.type === 'position') {
                    const currentPayload = item.updatePayload || { name: item.name };
                    const newPayload = { ...currentPayload, [field]: value };
                    return { ...item, updatePayload: newPayload };
                }
                if (item.type === 'skill') {
                    const currentPayload = item.updatePayload || { name: item.name, description: '' };
                    const newPayload = { ...currentPayload, [field]: value };
                    return { ...item, updatePayload: newPayload };
                }
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
        const successfullyDeletedDeptIds = new Set<number>();
        const successfullyDeletedGroupIds = new Set<number>();
        const successfullyDeletedSectionIds = new Set<number>();
        const successfullyDeletedPositionIds = new Set<number>();
        const successfullyDeletedSkillIds = new Set<number>();

        for (let i = 0; i < updatedItems.length; i++) {
            const item = updatedItems[i];
            if (item.action === 'delete' && item.id) {
                processedCount++;
                item.status = 'processing';
                setReviewItems([...updatedItems]);
                setExecutionProgress({
                    current: processedCount,
                    total: totalOps,
                    label: `Deleting ${item.type === 'department' ? 'Department' : item.type === 'section' ? 'Section' : 'Employee Group'} "${item.name}"...`
                });

                try {
                    if (item.type === 'section') {
                        let deptId = item.updatePayload?.departmentId;
                        if (!deptId) {
                            const sec = (currentStructure.sections || []).find(s => s.id === item.id);
                            deptId = sec?.departmentId;
                        }
                        if (!deptId && item.updatePayload?.departmentName) {
                            const d = currentStructure.departments.find(dept => dept.name.toLowerCase() === item.updatePayload.departmentName.toLowerCase());
                            deptId = d?.id;
                        }
                        if (!deptId) throw new Error('Missing department ID for section deletion');
                        await deleteSection(deptId, item.id);
                        successfullyDeletedSectionIds.add(item.id);
                    } else if (item.type === 'department') {
                        await deleteDepartment(item.id);
                        successfullyDeletedDeptIds.add(item.id);
                    } else if (item.type === 'employee_group') {
                        await deleteEmployeeGroup(item.id);
                        successfullyDeletedGroupIds.add(item.id);
                    } else if (item.type === 'position') {
                        await deletePosition(item.id);
                        successfullyDeletedPositionIds.add(item.id);
                    } else if (item.type === 'skill') {
                        await deleteSkill(item.id);
                        successfullyDeletedSkillIds.add(item.id);
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

        if (successfullyDeletedDeptIds.size > 0 || 
            successfullyDeletedGroupIds.size > 0 || 
            successfullyDeletedSectionIds.size > 0 || 
            successfullyDeletedPositionIds.size > 0 || 
            successfullyDeletedSkillIds.size > 0) {
            setCurrentStructure(prev => ({
                ...prev,
                departments: prev.departments.filter(d => !successfullyDeletedDeptIds.has(d.id)),
                employeeGroups: prev.employeeGroups.filter(g => !successfullyDeletedGroupIds.has(g.id)),
                sections: (prev.sections || []).filter(s => !successfullyDeletedSectionIds.has(s.id)),
                skills: prev.skills.filter(s => !successfullyDeletedSkillIds.has(s.id)),
                positions: prev.positions
                    .filter(p => !successfullyDeletedPositionIds.has(p.id))
                    .map(p => ({
                        ...p,
                        skillIds: (p.skillIds || []).filter(id => !successfullyDeletedSkillIds.has(id))
                    }))
            }));
            setDeleteSelectedDepts(prev => { const s = new Set(prev); successfullyDeletedDeptIds.forEach(id => s.delete(id)); return s; });
            setDeleteSelectedGroups(prev => { const s = new Set(prev); successfullyDeletedGroupIds.forEach(id => s.delete(id)); return s; });
            setDeleteSelectedSections(prev => { const s = new Set(prev); successfullyDeletedSectionIds.forEach(id => s.delete(id)); return s; });
            setDeleteSelectedPositions(prev => { const s = new Set(prev); successfullyDeletedPositionIds.forEach(id => s.delete(id)); return s; });
            setDeleteSelectedSkills(prev => { const s = new Set(prev); successfullyDeletedSkillIds.forEach(id => s.delete(id)); return s; });
        }

        // 2. Process Creations
        // Order: Departments -> Employee Groups -> Sections -> Positions -> Skills
        const creationOrder = ['department', 'employee_group', 'section', 'position', 'skill'];
        
        for (const type of creationOrder) {
            for (let i = 0; i < updatedItems.length; i++) {
                const item = updatedItems[i];
                if (item.action === 'create' && item.type === type) {
                    processedCount++;
                    item.status = 'processing';
                    setReviewItems([...updatedItems]);
                    setExecutionProgress({
                        current: processedCount,
                        total: totalOps,
                        label: `Creating ${item.type === 'section' ? 'Section' : item.type} "${item.name}"...`
                    });

                    try {                    
                        if (item.type === 'department') {
                            const created = await createDepartment(item.name, item.number || null);
                            item.status = 'success';
                            item.createdId = created.id;
                            item.resultMessage = `Created (ID: ${created.id})`;
                        } else if (item.type === 'section') {
                            let deptId = item.updatePayload?.departmentId;
                            if (!deptId && item.updatePayload?.departmentName) {
                                const dName = item.updatePayload.departmentName.trim().toLowerCase();
                                const existingDept = currentStructure.departments.find(d => d.name.trim().toLowerCase() === dName);
                                if (existingDept) {
                                    deptId = existingDept.id;
                                } else {
                                    const newlyCreated = updatedItems.find(it => it.type === 'department' && it.name.trim().toLowerCase() === dName && it.createdId);
                                    if (newlyCreated && newlyCreated.createdId) {
                                        deptId = newlyCreated.createdId;
                                    }
                                }
                            }
                            if (!deptId) {
                                throw new Error(`Cannot create section without a valid Department (Department '${item.updatePayload?.departmentName || 'Unknown'}' not found)`);
                            }
                            const color = item.updatePayload?.color || null;
                            const created = await createSection(deptId, item.name, color);
                            item.status = 'success';
                            item.createdId = created.id;
                            item.resultMessage = `Created (ID: ${created.id}, Dept ID: ${deptId})`;
                        } else if (item.type === 'skill') {
                            const created = await createSkill(item.name, item.updatePayload?.description);
                            item.status = 'success';
                            item.createdId = created.id;
                            item.resultMessage = `Created (ID: ${created.id})`;
                        } else if (item.type === 'position') {
                            const posPayload = { ...item.updatePayload };
                            
                            // Try to fetch mapping IDs
                            try {
                                const [currentDepts, currentGroups, currentSkills, currentRevUnits] = await Promise.all([
                                    fetchDepartments(),
                                    fetchEmployeeGroups(),
                                    fetchSkills(),
                                    fetchRevenueUnits()
                                ]);
                                
                                const deptMatch = currentDepts.find(d => d.name.toLowerCase() === posPayload.departmentName?.toLowerCase());
                                const groupMatch = currentGroups.find(g => g.name.toLowerCase() === posPayload.employeeGroupName?.toLowerCase());
                                
                                if (!deptMatch) throw new Error(`Department '${posPayload.departmentName}' not found`);
                                if (!groupMatch) throw new Error(`Employee Group '${posPayload.employeeGroupName}' not found`);
                                
                                const apiPayload: any = {
                                    name: posPayload.name,
                                    departmentId: deptMatch.id,
                                    employeeGroupId: groupMatch.id,
                                    color: posPayload.color,
                                    validFrom: posPayload.validFrom,
                                    validTo: posPayload.validTo,
                                    affectRevenue: posPayload.affectRevenue
                                };

                                // Map sectionId if provided
                                if (posPayload.sectionName) {
                                    const sNameLower = posPayload.sectionName.trim().toLowerCase();
                                    const existingSec = (currentStructure.sections || []).find(s => s.name.trim().toLowerCase() === sNameLower && s.departmentId === deptMatch.id);
                                    if (existingSec) {
                                        apiPayload.sectionId = existingSec.id;
                                    } else {
                                        const newlyCreatedSec = updatedItems.find(it => it.type === 'section' && it.name.trim().toLowerCase() === sNameLower && it.createdId);
                                        if (newlyCreatedSec && newlyCreatedSec.createdId) {
                                            apiPayload.sectionId = newlyCreatedSec.createdId;
                                        }
                                    }
                                }

                                if (posPayload.skills) {
                                    const skillNames = posPayload.skills.split(',').map((s: string) => s.trim().toLowerCase());
                                    const skillIds = currentSkills.filter(s => skillNames.includes(s.name.toLowerCase())).map(s => s.id);
                                    if (skillIds.length > 0) apiPayload.skillIds = skillIds;
                                }

                                if (posPayload.revenueUnitName) {
                                    const revMatch = currentRevUnits.find(r => r.name.toLowerCase() === posPayload.revenueUnitName?.toLowerCase() && r.departmentId === deptMatch.id);
                                    if (revMatch) apiPayload.revenueUnitId = revMatch.id;
                                } else {
                                    // Auto assign if only 1 revenue unit in that dept
                                    const deptRevUnits = currentRevUnits.filter(r => r.departmentId === deptMatch.id);
                                    if (deptRevUnits.length === 1) apiPayload.revenueUnitId = deptRevUnits[0].id;
                                }

                                // Remove undefined/null properties to avoid payload issues
                                Object.keys(apiPayload).forEach(key => {
                                    if (apiPayload[key] === undefined || apiPayload[key] === null || apiPayload[key] === '') {
                                        delete apiPayload[key];
                                    }
                                });

                                const created = await createPosition(apiPayload);
                                item.status = 'success';
                                item.createdId = created.id;
                                item.resultMessage = `Created (ID: ${created.id})`;

                            } catch (mappingErr: any) {
                                throw new Error(`Mapping failed: ${mappingErr.message}`);
                            }

                        } else if (item.type === 'employee_group') {                        
                            const created = await createEmployeeGroup(item.name);
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
                    label: `Updating ${item.type === 'department' ? 'Department' : item.type === 'section' ? 'Section' : item.type === 'position' ? 'Position' : item.type === 'skill' ? 'Skill' : 'Employee Group'} "${item.name}"...`
                });

                try {
                    if (item.type === 'section') {
                        const existingSec = (currentStructure.sections || []).find(s => s.id === item.id);
                        const originalName = existingSec?.name || item.name;
                        const originalColor = existingSec?.color || null;
                        let deptId = item.updatePayload?.departmentId || existingSec?.departmentId;
                        if (!deptId && item.updatePayload?.departmentName) {
                            const d = currentStructure.departments.find(dept => dept.name.toLowerCase() === item.updatePayload.departmentName.toLowerCase());
                            if (d) deptId = d.id;
                        }
                        if (!deptId) throw new Error('Missing department ID for section update');

                        const targetName = (item.updatePayload?.name !== undefined && item.updatePayload?.name !== null && String(item.updatePayload.name).trim() !== '') 
                            ? String(item.updatePayload.name).trim() 
                            : originalName;
                        
                        const sectionPayload: { name: string; color?: string | null } = {
                            name: targetName
                        };
                        if (item.updatePayload?.color !== undefined) {
                            sectionPayload.color = item.updatePayload.color ? String(item.updatePayload.color).trim() : null;
                        }

                        const nameChanged = targetName !== originalName;
                        const colorChanged = item.updatePayload?.color !== undefined && sectionPayload.color !== originalColor;

                        await updateSection(deptId, item.id, sectionPayload);
                        item.status = 'success';
                        const changes: string[] = [];
                        if (nameChanged) changes.push(`Name: '${originalName}' -> '${targetName}'`);
                        if (colorChanged) changes.push(`Color: '${originalColor || 'None'}' -> '${sectionPayload.color || 'None'}'`);
                        item.resultMessage = changes.length > 0 ? `Updated successfully (${changes.join(', ')})` : 'Updated successfully';
                        if (existingSec) {
                            existingSec.name = targetName;
                            if (sectionPayload.color !== undefined) {
                                existingSec.color = sectionPayload.color;
                            }
                        }
                    } else if (item.type === 'department') {
                        const existingDept = currentStructure.departments.find(d => d.id === item.id || d.name.trim().toLowerCase() === item.name.trim().toLowerCase());
                        const originalName = existingDept?.name || item.name;
                        const originalNum = existingDept?.number || item.number || null;

                        const targetName = (item.updatePayload?.name !== undefined && item.updatePayload?.name !== null && String(item.updatePayload.name).trim() !== '') 
                            ? String(item.updatePayload.name).trim() 
                            : originalName;
                        const targetNum = item.updatePayload?.number !== undefined 
                            ? (item.updatePayload.number ? String(item.updatePayload.number).trim() : null) 
                            : originalNum;

                        const nameChanged = targetName !== originalName;
                        const numChanged = targetNum !== originalNum;

                        if (nameChanged || numChanged || (appMode === 'update' && (item.updatePayload?.name || item.updatePayload?.number !== undefined)) || appMode === 'import') {
                            await updateDepartment(item.id, targetName, targetNum);
                            item.status = 'success';
                            const changes: string[] = [];
                            if (nameChanged) changes.push(`Name: '${originalName}' -> '${targetName}'`);
                            if (numChanged) changes.push(`Number: '${originalNum || 'None'}' -> '${targetNum || 'None'}'`);
                            item.resultMessage = changes.length > 0 ? `Updated successfully (${changes.join(', ')})` : 'Updated successfully';
                            if (existingDept) {
                                existingDept.name = targetName;
                                existingDept.number = targetNum;
                            }
                        } else {
                            item.status = 'skipped';
                            item.resultMessage = 'No changes detected (Name and Number already match Planday)';
                        }
                    } else if (item.type === 'position') {
                        const existingPos = (currentStructure.positions || []).find(p => p.id === item.id);
                        const originalName = existingPos?.name || item.name;
                        const targetName = (item.updatePayload?.name !== undefined && String(item.updatePayload.name).trim() !== '') 
                            ? String(item.updatePayload.name).trim() 
                            : originalName;
                        
                        // PUT /scheduling/v1.0/positions/{positionId}
                        // Schema: sectionId, name (required), color, validFrom, validTo, skillIds, affectRevenue (required), revenueUnitId
                        // Only include mapped payload properties!
                        // affectRevenue is always required. Set to true if not specified.
                        const updateBody: any = {
                            name: targetName,
                            affectRevenue: item.updatePayload?.affectRevenue !== undefined 
                                ? Boolean(item.updatePayload.affectRevenue) 
                                : true
                        };

                        if (item.updatePayload?.sectionId !== undefined) {
                            updateBody.sectionId = item.updatePayload.sectionId;
                        }
                        if (item.updatePayload?.color !== undefined) {
                            updateBody.color = item.updatePayload.color ? String(item.updatePayload.color).trim() : null;
                        }
                        if (item.updatePayload?.validFrom !== undefined) {
                            updateBody.validFrom = item.updatePayload.validFrom || null;
                        }
                        if (item.updatePayload?.validTo !== undefined) {
                            updateBody.validTo = item.updatePayload.validTo || null;
                        }
                        if (item.updatePayload?.skillIds !== undefined) {
                            updateBody.skillIds = Array.isArray(item.updatePayload.skillIds) ? item.updatePayload.skillIds : [];
                        }
                        if (item.updatePayload?.revenueUnitId !== undefined) {
                            updateBody.revenueUnitId = item.updatePayload.revenueUnitId;
                        }

                        await updatePosition(item.id, updateBody);
                        item.status = 'success';
                        const nameChanged = targetName !== originalName;
                        item.resultMessage = nameChanged ? `Updated successfully (Name: '${originalName}' -> '${targetName}')` : 'Updated successfully';
                        if (existingPos) {
                            existingPos.name = targetName;
                            if (updateBody.color !== undefined) existingPos.color = updateBody.color;
                            if (updateBody.sectionId !== undefined) existingPos.sectionId = updateBody.sectionId;
                            if (updateBody.affectRevenue !== undefined) existingPos.affectRevenue = updateBody.affectRevenue;
                            if (updateBody.validFrom !== undefined) existingPos.validFrom = updateBody.validFrom;
                            if (updateBody.validTo !== undefined) existingPos.validTo = updateBody.validTo;
                            if (updateBody.skillIds !== undefined) existingPos.skillIds = updateBody.skillIds;
                            if (updateBody.revenueUnitId !== undefined) existingPos.revenueUnitId = updateBody.revenueUnitId;
                        }
                    } else if (item.type === 'skill') {
                        const existingSkill = (currentStructure.skills || []).find(s => s.id === item.id);
                        const originalName = existingSkill?.name || item.name;
                        const targetName = (item.updatePayload?.name !== undefined && String(item.updatePayload.name).trim() !== '')
                            ? String(item.updatePayload.name).trim()
                            : originalName;
                        const targetDesc = item.updatePayload?.description !== undefined ? item.updatePayload.description : (existingSkill?.description || '');
                        await updateSkill(item.id, targetName, targetDesc);
                        item.status = 'success';
                        const nameChanged = targetName !== originalName;
                        item.resultMessage = nameChanged ? `Updated successfully (Name: '${originalName}' -> '${targetName}')` : 'Updated successfully';
                        if (existingSkill) {
                            existingSkill.name = targetName;
                            existingSkill.description = targetDesc;
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
            const [depts, groups, skills, revUnits] = await Promise.all([
                fetchDepartments(),
                fetchEmployeeGroups(),
                fetchSkills().catch(() => []),
                fetchRevenueUnits().catch(() => [])
            ]);
            const positions = await fetchPositions(skills).catch(() => []);
            const sections = await fetchAllSections(depts).catch(() => []);
            setCurrentStructure({ 
                departments: depts, 
                employeeGroups: groups, 
                skills: (skills || []).filter((s: any) => s.isActive !== false && s.IsActive !== false && s.isActive !== 'false' && s.IsActive !== 'false'),
                positions: (positions || []).filter((p: any) => p.isActive !== false && p.IsActive !== false && p.isActive !== 'false' && p.IsActive !== 'false'),
                sections: (sections || []).filter((s: any) => s.isActive !== false && s.IsActive !== false && s.isActive !== 'false' && s.IsActive !== 'false'),
                revenueUnits: revUnits 
            });
        } catch { }
    };

    // Export execution report as Excel
    const handleExportExecutionReport = () => {
        const sortedItems = [...reviewItems].sort((a, b) => {
            const orderA = TYPE_ORDER[a.type] || 99;
            const orderB = TYPE_ORDER[b.type] || 99;
            if (orderA !== orderB) return orderA - orderB;
            return a.name.localeCompare(b.name);
        });

        const data = sortedItems.map(item => {
            let actionLabel = 'Skip Duplicate';
            if (item.action === 'create') actionLabel = 'Create';
            else if (item.action === 'delete') actionLabel = 'Delete';
            else if (item.action === 'update') actionLabel = 'Update';

            const typeLabel = item.type === 'department' ? 'Department' :
                              item.type === 'employee_group' ? 'Employee Group' :
                              item.type === 'section' ? 'Section' :
                              item.type === 'position' ? 'Position' :
                              item.type === 'skill' ? 'Skill' : 'Unknown';

            const row: any = {
                "Type": typeLabel,
                "Name": item.name,
                "Action": actionLabel,
                "Status": item.status === 'success' ? 'Success' : item.status === 'error' ? 'Error' : 'Skipped',
                "Planday ID": item.createdId || item.id || '-',
                "Details / Result": item.resultMessage || ''
            };

            if (item.type === 'department') {
                if (item.updatePayload?.name && item.updatePayload.name !== item.name) {
                    row["Updated Name"] = item.updatePayload.name;
                }
                row["Number"] = item.updatePayload?.number !== undefined ? item.updatePayload.number : (item.number || "");
            }

            if (item.type === 'employee_group') {
                row["Wage Type"] = item.updatePayload?.defaultWageType || "";
                row["Rate"] = item.updatePayload?.defaultWageType === 'HourlyRate' ? (item.updatePayload?.hourlyRate ?? "") : (item.updatePayload?.defaultWageType === 'ShiftRate' ? (item.updatePayload?.shiftRate ?? "") : "");
                row["Salary Code"] = item.updatePayload?.salaryCode || "";
                row["Salary Code 2"] = item.updatePayload?.salaryCode2 || "";
            }

            if (item.type === 'section') {
                row["Department"] = item.updatePayload?.departmentName || "";
                if (item.updatePayload?.color !== undefined) {
                    row["Color"] = item.updatePayload?.color || "";
                }
                if (item.updatePayload?.name && item.updatePayload.name !== item.name) {
                    row["Updated Name"] = item.updatePayload.name;
                }
            }

            if (item.type === 'position') {
                if (item.updatePayload?.name && item.updatePayload.name !== item.name) {
                    row["Updated Name"] = item.updatePayload.name;
                }
                row["Department Name"] = item.updatePayload?.departmentName || "";
                row["Employee Group Name"] = item.updatePayload?.employeeGroupName || "";
                if (item.updatePayload?.sectionName) {
                    row["Section Name"] = item.updatePayload.sectionName;
                }
                if (item.updatePayload?.color !== undefined) {
                    row["Color"] = item.updatePayload.color || "";
                }
                if (item.updatePayload?.affectRevenue !== undefined) {
                    row["Include in Payroll %"] = String(item.updatePayload.affectRevenue);
                }
                if (item.updatePayload?.validFrom !== undefined) {
                    row["Valid From"] = item.updatePayload.validFrom || "";
                }
                if (item.updatePayload?.validTo !== undefined) {
                    row["Valid To"] = item.updatePayload.validTo || "";
                }
                if (item.updatePayload?.skills) {
                    row["Skills"] = item.updatePayload.skills;
                }
                if (item.updatePayload?.revenueUnitName) {
                    row["Revenue Unit"] = item.updatePayload.revenueUnitName;
                }
            }

            if (item.type === 'skill') {
                row["Description"] = item.updatePayload?.description || "";
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
        return reviewItems
            .filter(item => {
                if (reviewTab === 'create' && item.action !== 'create') return false;
                if (reviewTab === 'skip' && item.action !== 'skip_duplicate') return false;
                if (reviewTab === 'delete' && item.action !== 'delete') return false;
                if (reviewTab === 'update' && item.action !== 'update') return false;

                if (reviewSearch) {
                    const lower = reviewSearch.toLowerCase();
                    return item.name.toLowerCase().includes(lower) || item.type.toLowerCase().includes(lower);
                }
                return true;
            })
            .sort((a, b) => {
                const orderA = TYPE_ORDER[a.type] || 99;
                const orderB = TYPE_ORDER[b.type] || 99;
                if (orderA !== orderB) return orderA - orderB;
                return a.name.localeCompare(b.name);
            });
    }, [reviewItems, reviewTab, reviewSearch]);

    // Review counts
    const reviewStats = useMemo(() => {
        const toCreateDept = reviewItems.filter(i => i.type === 'department' && i.action === 'create').length;
        const toCreateGroup = reviewItems.filter(i => i.type === 'employee_group' && i.action === 'create').length;
        const toCreateSection = reviewItems.filter(i => i.type === 'section' && i.action === 'create').length;
        const toCreatePosition = reviewItems.filter(i => i.type === 'position' && i.action === 'create').length;
        const toCreateSkill = reviewItems.filter(i => i.type === 'skill' && i.action === 'create').length;

        const toUpdateDept = reviewItems.filter(i => i.type === 'department' && i.action === 'update').length;
        const toUpdateGroup = reviewItems.filter(i => i.type === 'employee_group' && i.action === 'update').length;
        const toUpdateSection = reviewItems.filter(i => i.type === 'section' && i.action === 'update').length;
        const toUpdatePosition = reviewItems.filter(i => i.type === 'position' && i.action === 'update').length;
        const toUpdateSkill = reviewItems.filter(i => i.type === 'skill' && i.action === 'update').length;

        const toDeleteDept = reviewItems.filter(i => i.type === 'department' && i.action === 'delete').length;
        const toDeleteGroup = reviewItems.filter(i => i.type === 'employee_group' && i.action === 'delete').length;
        const toDeleteSection = reviewItems.filter(i => i.type === 'section' && i.action === 'delete').length;
        const toDeletePosition = reviewItems.filter(i => i.type === 'position' && i.action === 'delete').length;
        const toDeleteSkill = reviewItems.filter(i => i.type === 'skill' && i.action === 'delete').length;

        const skipped = reviewItems.filter(i => i.action === 'skip_duplicate').length;
        const toDelete = reviewItems.filter(i => i.action === 'delete').length;
        const toUpdate = reviewItems.filter(i => i.action === 'update').length;
        const toCreate = toCreateDept + toCreateGroup + toCreateSection + toCreatePosition + toCreateSkill;

        return { 
            toCreateDept, toCreateGroup, toCreateSection, toCreatePosition, toCreateSkill, toCreate,
            toUpdateDept, toUpdateGroup, toUpdateSection, toUpdatePosition, toUpdateSkill, toUpdate,
            toDeleteDept, toDeleteGroup, toDeleteSection, toDeletePosition, toDeleteSkill, toDelete,
            skipped 
        };
    }, [reviewItems]);

    // Overview pending action tiles based on active flow (New for import, Update for updates, Delete for deletions)
    const overviewTiles = useMemo(() => {
        if (appMode === 'delete') {
            return [
                { label: 'Delete Departments', count: reviewStats.toDeleteDept, color: 'text-blue-600' },
                { label: 'Delete Employee Groups', count: reviewStats.toDeleteGroup, color: 'text-purple-600' },
                { label: 'Delete Sections', count: reviewStats.toDeleteSection, color: 'text-gray-700' },
                { label: 'Delete Positions', count: reviewStats.toDeletePosition, color: 'text-teal-600' },
                { label: 'Delete Skills', count: reviewStats.toDeleteSkill, color: 'text-yellow-700' },
                { label: 'Skipped', count: reviewStats.skipped, color: 'text-amber-600' },
            ];
        }
        if (appMode === 'update') {
            return [
                { label: 'Update Departments', count: reviewStats.toUpdateDept, color: 'text-blue-600' },
                { label: 'Update Employee Groups', count: reviewStats.toUpdateGroup, color: 'text-purple-600' },
                { label: 'Update Sections', count: reviewStats.toUpdateSection, color: 'text-gray-700' },
                { label: 'Update Positions', count: reviewStats.toUpdatePosition, color: 'text-teal-600' },
                { label: 'Update Skills', count: reviewStats.toUpdateSkill, color: 'text-yellow-700' },
                { label: 'Skipped', count: reviewStats.skipped, color: 'text-amber-600' },
            ];
        }
        // appMode === 'import'
        const tiles = [
            { label: 'New Departments', count: reviewStats.toCreateDept, color: 'text-blue-600' },
            { label: 'New Employee Groups', count: reviewStats.toCreateGroup, color: 'text-purple-600' },
            { label: 'New Sections', count: reviewStats.toCreateSection, color: 'text-gray-700' },
            { label: 'New Positions', count: reviewStats.toCreatePosition, color: 'text-teal-600' },
            { label: 'New Skills', count: reviewStats.toCreateSkill, color: 'text-yellow-700' },
            { label: 'Duplicates Skipped', count: reviewStats.skipped, color: 'text-amber-600' },
        ];
        if (reviewStats.toUpdate > 0) {
            tiles.push({ label: 'Items To Update', count: reviewStats.toUpdate, color: 'text-amber-600' });
        }
        if (reviewStats.toDelete > 0) {
            tiles.push({ label: 'Items To Delete', count: reviewStats.toDelete, color: 'text-red-600' });
        }
        return tiles;
    }, [appMode, reviewStats]);

    // Current Structure Preview filtered items by department
    const filteredPreviewSections = useMemo(() => {
        if (!previewSectionDeptId) return [];
        const activeSections = (currentStructure.sections || []).filter(s => s.isActive !== false && (s as any).IsActive !== false && (s as any).isActive !== 'false' && (s as any).IsActive !== 'false');
        if (previewSectionDeptId === 'all') return activeSections;
        const targetDeptId = Number(previewSectionDeptId);
        return activeSections.filter(s => Number(s.departmentId) === targetDeptId);
    }, [currentStructure.sections, previewSectionDeptId]);

    const filteredPreviewPositions = useMemo(() => {
        if (!previewPositionDeptId) return [];
        const activePositions = (currentStructure.positions || []).filter(p => p.isActive !== false && (p as any).IsActive !== false && (p as any).isActive !== 'false' && (p as any).IsActive !== 'false');
        if (previewPositionDeptId === 'all') return activePositions;
        const targetDeptId = Number(previewPositionDeptId);
        return activePositions.filter(p => {
            if (p.departmentId !== undefined && p.departmentId !== null) {
                return Number(p.departmentId) === targetDeptId;
            }
            if (Array.isArray((p as any).departmentIds) && (p as any).departmentIds.length > 0) {
                return (p as any).departmentIds.some((id: any) => Number(id) === targetDeptId);
            }
            if (p.sectionId && currentStructure.sections) {
                const sec = currentStructure.sections.find(s => s.id === p.sectionId);
                if (sec && Number(sec.departmentId) === targetDeptId) {
                    return true;
                }
            }
            return false;
        });
    }, [currentStructure.positions, currentStructure.sections, previewPositionDeptId]);

    // Steps configuration
    const importStepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Make Selections", subtitle: "Choose what to Import" },
        { title: "Configure", subtitle: "Map/Input data" },
        { title: "Review", subtitle: "Confirm Import" },
        { title: "Execute", subtitle: "Bulk Import" },
    ];

    const deleteStepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Make Selections", subtitle: "Choose what to Delete" },
        { title: "Configure", subtitle: "Map/Input data" },
        { title: "Review", subtitle: "Confirm Deletions" },
        { title: "Execute", subtitle: "Bulk Delete" },
    ];

    const updateStepsList = [
        { title: "Authenticate", subtitle: "Connect Planday API" },
        { title: "Make Selections", subtitle: "Choose what to Update" },
        { title: "Configure", subtitle: "Map/Input data" },
        { title: "Review", subtitle: "Confirm Updates" },
        { title: "Execute", subtitle: "Bulk Update" },
    ];

    const stepsList = appMode === 'import' ? importStepsList : appMode === 'update' ? updateStepsList : deleteStepsList;

    return (
        <div className="min-h-screen bg-gray-50 flex flex-col font-sans text-gray-900">
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
                                onClick={() => { setAppMode('update'); setStep(1); setDeleteInputMethod('file'); }}
                                className={`px-6 py-2 rounded-lg text-sm font-bold transition-colors ${appMode === 'update' ? 'bg-amber-600 text-white' : 'text-gray-600 hover:bg-gray-100'}`}
                            >
                                Update
                            </button>
                            <button
                                onClick={() => { setAppMode('delete'); setStep(1); setDeleteInputMethod('file'); }}
                                className={`px-6 py-2 rounded-lg text-sm font-bold transition-colors ${appMode === 'delete' ? 'bg-red-600 text-white' : 'text-gray-600 hover:bg-gray-100'}`}
                            >
                                Delete
                            </button>
                        </div>
                    </div>
                )}

                {/* Progress Stepper Bar */}
                <div className="w-full py-2">
                    <Stepper current={step} steps={stepsList} onStepClick={(s) => setStep(s)} />
                </div>

                {step !== 0 && (
                    <div className="flex justify-center mt-6 text-sm items-center">
                        {/* Refresh Button */}
                        <button
                            onClick={handleManualRefresh}
                            className="mr-2 text-gray-500 hover:text-blue-600 transition-colors"
                            title="Refresh portal data"
                            disabled={isLoading}
                        >
                            <svg 
                                className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} 
                                fill="none" 
                                viewBox="0 0 24 24" 
                                stroke="currentColor"
                            >
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                            </svg>
                        </button>

                        {/* Portal/User Name */}
                        <span className="text-gray-600">
                            Logged in: <strong>{portalName || 'Loading...'}</strong>
                        </span>
                        
                        <span className="text-gray-400 mx-2">|</span>
                        
                        {/* Logout Button */}
                        <button 
                            onClick={handleChangeCredentials} 
                            disabled={isDisconnecting}
                            className="text-gray-600 hover:text-red-700 underline transition-colors disabled:opacity-50 inline-flex items-center gap-1.5"
                        >
                            {isDisconnecting && (
                                <svg className="animate-spin h-3.5 w-3.5 text-red-600" fill="none" viewBox="0 0 24 24">
                                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                                </svg>
                            )}
                            {isDisconnecting ? 'Disconnecting...' : 'Change credentials (log out)'}
                        </button>
                    </div>
                )}

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
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 grid sm:grid-cols-2 lg:grid-cols-5 gap-4">
                            <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 flex items-center gap-4">
                                <div className="p-3 bg-blue-600 text-white rounded-lg">
                                    <BuildingOfficeIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-blue-900">{currentStructure.departments.length}</span>
                                    <span className="text-xs font-semibold text-blue-700 uppercase tracking-wider">Departments</span>
                                </div>
                            </div>

                            <div className="bg-purple-50 p-4 rounded-xl border border-purple-200 flex items-center gap-4">
                                <div className="p-3 bg-purple-600 text-white rounded-lg">
                                    <UserGroupIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-purple-900">{currentStructure.employeeGroups.length}</span>
                                    <span className="text-xs font-semibold text-purple-700 uppercase tracking-wider">Empl. Groups</span>
                                </div>
                            </div>

                            <div className="bg-gray-100 p-4 rounded-xl border border-gray-300 flex items-center gap-4">
                                <div className="p-3 bg-gray-700 text-white rounded-lg">
                                    <SectionIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-gray-900">{currentStructure.sections?.length || 0}</span>
                                    <span className="text-xs font-semibold text-gray-700 uppercase tracking-wider">Sections</span>
                                </div>
                            </div>

                            <div className="bg-teal-50 p-4 rounded-xl border border-teal-200 flex items-center gap-4">
                                <div className="p-3 bg-teal-600 text-white rounded-lg">
                                    <BriefcaseIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-teal-900">{currentStructure.positions?.length || 0}</span>
                                    <span className="text-xs font-semibold text-teal-700 uppercase tracking-wider">Positions</span>
                                </div>
                            </div>

                            <div className="bg-yellow-50 p-4 rounded-xl border border-yellow-200 flex items-center gap-4">
                                <div className="p-3 bg-yellow-600 text-white rounded-lg">
                                    <StarIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-yellow-900">{currentStructure.skills?.length || 0}</span>
                                    <span className="text-xs font-semibold text-yellow-800 uppercase tracking-wider">Skills</span>
                                </div>
                            </div>
                        </div>

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
                                            <span>Departments</span>
                                        </label>
                                        <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                            <input 
                                                type="checkbox" 
                                                checked={includeEmployeeGroups} 
                                                onChange={(e) => setIncludeEmployeeGroups(e.target.checked)}
                                                className="w-4 h-4 text-purple-600 rounded focus:ring-purple-500"
                                            />
                                            <span>Employee Groups</span>
                                        </label>
                                        {includeEmployeeGroups && (
                                            <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800 ml-7">
                                                <input 
                                                    type="checkbox" 
                                                    checked={includeWageColumns} 
                                                    onChange={(e) => setIncludeWageColumns(e.target.checked)}
                                                    className="w-4 h-4 text-purple-600 rounded focus:ring-purple-500"
                                                />
                                                <span>Include Default Wage Data columns</span>
                                            </label>
                                        )}
                                        <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                            <input 
                                                type="checkbox" 
                                                checked={includeSections} 
                                                onChange={(e) => setIncludeSections(e.target.checked)}
                                                className="w-4 h-4 text-gray-700 rounded focus:ring-gray-500"
                                            />
                                            <span>Sections</span>
                                        </label>
                                        <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                            <input 
                                                type="checkbox" 
                                                checked={includePositions} 
                                                onChange={(e) => setIncludePositions(e.target.checked)}
                                                className="w-4 h-4 text-teal-600 rounded focus:ring-teal-500"
                                            />
                                            <span>Positions</span>
                                        </label>
                                        <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                            <input 
                                                type="checkbox" 
                                                checked={includeSkills} 
                                                onChange={(e) => setIncludeSkills(e.target.checked)}
                                                className="w-4 h-4 text-yellow-600 rounded focus:ring-yellow-500"
                                            />
                                            <span>Skills</span>
                                        </label>
                                    </div>
                                </div>

                                <div className="space-y-3 mt-4">
                                    <button 
                                        onClick={populateWithCurrent ? handleDownloadCurrentStructure : handleDownloadBlankTemplate}
                                        disabled={!includeDepartments && !includeSections && !includeEmployeeGroups && !includeSkills && !includePositions}
                                        className="w-full py-3 px-4 bg-green-600 hover:bg-green-700 text-white font-bold rounded-xl shadow-sm hover:shadow transition-all flex items-center justify-center gap-2 text-sm disabled:opacity-50"
                                    >
                                        <DownloadIcon className="w-4 h-4" />
                                        <span>Download Template</span>
                                    </button>
                                    
                                    <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-700 justify-center">
                                        <input 
                                            type="checkbox" 
                                            checked={populateWithCurrent} 
                                            onChange={(e) => setPopulateWithCurrent(e.target.checked)}
                                            className="w-4 h-4 text-green-600 rounded focus:ring-green-500"
                                        />
                                        <span>Populate with current portal structure?</span>
                                    </label>
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


                                </div>

                                {/* Drag and Drop / File Input Dropzone */}
                                <div className="relative border-2 border-dashed border-gray-300 hover:border-blue-500 rounded-2xl p-6 text-center transition-colors bg-gray-50/50 hover:bg-blue-50/30 group flex-1 flex flex-col items-center justify-center min-h-[160px]">
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
                            <div className="space-y-6 text-sm">
                                {/* Row 1: Departments and Employee Groups */}
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
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
                                            <span className="text-xs bg-purple-100 text-purple-800 px-2 py-0.5 rounded-full">{currentStructure.employeeGroups.length}</span>
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

                                {/* Row 2: Sections and Positions */}
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                    {/* Sections List */}
                                    <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                        <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200 flex justify-between items-center">
                                            <span>Sections</span>
                                            <span className="text-xs bg-gray-200 text-gray-800 px-2 py-0.5 rounded-full font-semibold">
                                                {previewSectionDeptId ? (previewSectionDeptId === 'all' ? (currentStructure.sections?.length || 0) : filteredPreviewSections.length) : 0}
                                            </span>
                                        </div>
                                        <div className="p-2.5 bg-slate-50 border-b border-gray-200">
                                            <label className="block text-[11px] font-semibold text-gray-600 mb-1">
                                                Select Department:
                                            </label>
                                            <select
                                                value={previewSectionDeptId}
                                                onChange={(e) => setPreviewSectionDeptId(e.target.value)}
                                                className="w-full text-xs bg-white border border-gray-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-slate-500 focus:border-slate-500 text-gray-800 font-medium"
                                            >
                                                <option value="">-- Select Department --</option>
                                                {currentStructure.departments.map(d => (
                                                    <option key={d.id} value={d.id}>{d.name}</option>
                                                ))}
                                                <option value="all">All Departments</option>
                                            </select>
                                        </div>
                                        <div className="max-h-48 overflow-y-auto p-3 space-y-1 bg-white flex-1">
                                            {!previewSectionDeptId ? (
                                                <div className="py-6 px-2 text-center text-xs text-gray-400 italic">
                                                    Select a department above to review sections
                                                </div>
                                            ) : filteredPreviewSections.length > 0 ? (
                                                filteredPreviewSections.map(s => {
                                                    const dept = currentStructure.departments.find(d => d.id === s.departmentId);
                                                    return (
                                                        <div key={s.id} className="py-1 px-2 hover:bg-gray-50 rounded text-gray-700 text-xs flex justify-between items-center">
                                                            <div className="flex items-center gap-1.5 truncate">
                                                                {s.color && (
                                                                    <span className="inline-block w-2.5 h-2.5 rounded-full shrink-0 border border-gray-200" style={{ backgroundColor: s.color }} />
                                                                )}
                                                                <span className="font-medium truncate">{s.name}</span>
                                                                <span className="text-gray-400 text-[10px]">({dept?.name || `Dept #${s.departmentId}`})</span>
                                                            </div>
                                                            <span className="text-gray-400 text-[10px] shrink-0">ID: {s.id}</span>
                                                        </div>
                                                    );
                                                })
                                            ) : (
                                                <div className="py-6 px-2 text-center text-xs text-gray-400 italic">
                                                    No sections available in this department
                                                </div>
                                            )}
                                        </div>
                                    </div>

                                    {/* Positions List */}
                                    <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                        <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200 flex justify-between items-center">
                                            <span>Positions</span>
                                            <span className="text-xs bg-teal-100 text-teal-800 px-2 py-0.5 rounded-full font-semibold">
                                                {previewPositionDeptId ? (previewPositionDeptId === 'all' ? (currentStructure.positions?.length || 0) : filteredPreviewPositions.length) : 0}
                                            </span>
                                        </div>
                                        <div className="p-2.5 bg-teal-50/40 border-b border-gray-200">
                                            <label className="block text-[11px] font-semibold text-gray-600 mb-1">
                                                Select Department:
                                            </label>
                                            <select
                                                value={previewPositionDeptId}
                                                onChange={(e) => setPreviewPositionDeptId(e.target.value)}
                                                className="w-full text-xs bg-white border border-gray-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 text-gray-800 font-medium"
                                            >
                                                <option value="">-- Select Department --</option>
                                                {currentStructure.departments.map(d => (
                                                    <option key={d.id} value={d.id}>{d.name}</option>
                                                ))}
                                                <option value="all">All Departments</option>
                                            </select>
                                        </div>
                                        <div className="max-h-48 overflow-y-auto p-3 space-y-1 bg-white flex-1">
                                            {!previewPositionDeptId ? (
                                                <div className="py-6 px-2 text-center text-xs text-gray-400 italic">
                                                    Select a department above to review positions
                                                </div>
                                            ) : filteredPreviewPositions.length > 0 ? (
                                                filteredPreviewPositions.map(p => {
                                                    const group = currentStructure.employeeGroups.find(g => g.id === p.employeeGroupId);
                                                    const assignedSkillNames = (p.skillIds && p.skillIds.length > 0 && currentStructure.skills)
                                                        ? currentStructure.skills.filter(s => p.skillIds?.includes(s.id)).map(s => s.name)
                                                        : [];
                                                    const revUnit = p.revenueUnitId ? currentStructure.revenueUnits?.find(r => r.id === p.revenueUnitId) : undefined;
                                                    return (
                                                        <div key={p.id} className="py-1 px-2 hover:bg-gray-50 rounded text-gray-700 text-xs flex justify-between items-center">
                                                            <div className="truncate flex items-center gap-1.5 flex-1 min-w-0 mr-2">
                                                                <span className="font-medium truncate">{p.name}</span>
                                                                {group && (
                                                                    <span className="text-gray-400 text-[10px] shrink-0">({group.name})</span>
                                                                )}
                                                                {revUnit && (
                                                                    <span 
                                                                        className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-emerald-50 text-emerald-800 border border-emerald-200 shrink-0" 
                                                                        title={`Revenue Unit: ${revUnit.name}`}
                                                                    >
                                                                        {revUnit.name}
                                                                    </span>
                                                                )}
                                                                {p.affectRevenue === false && (
                                                                    <span 
                                                                        className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-rose-50 text-rose-700 border border-rose-200 shrink-0" 
                                                                        title="Exclude from Payroll % (affectRevenue: false)"
                                                                    >
                                                                        Payroll: No
                                                                    </span>
                                                                )}
                                                                {assignedSkillNames.length > 0 && (
                                                                    <span 
                                                                        className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-amber-50 text-amber-800 border border-amber-200 shrink-0" 
                                                                        title={`Skills: ${assignedSkillNames.join(', ')}`}
                                                                    >
                                                                        {assignedSkillNames.length} {assignedSkillNames.length === 1 ? 'skill' : 'skills'}
                                                                    </span>
                                                                )}
                                                            </div>
                                                            <span className="text-gray-400 text-[10px] shrink-0">ID: {p.id}</span>
                                                        </div>
                                                    );
                                                })
                                            ) : (
                                                <div className="py-6 px-2 text-center text-xs text-gray-400 italic">
                                                    No positions available in this department
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                </div>

                                {/* Row 3: Skills */}
                                <div className="grid grid-cols-1 gap-6">
                                    {/* Skills List */}
                                    <div className="border border-gray-200 rounded-xl overflow-hidden bg-white">
                                        <div className="bg-gray-50 px-4 py-2.5 font-bold text-gray-800 border-b border-gray-200 flex justify-between items-center">
                                            <span>Skills</span>
                                            <span className="text-xs bg-yellow-100 text-yellow-900 border border-yellow-300 px-2 py-0.5 rounded-full font-semibold">{currentStructure.skills?.length || 0}</span>
                                        </div>
                                        <div className="max-h-48 overflow-y-auto p-3 space-y-1 bg-white">
                                            {(currentStructure.skills && currentStructure.skills.length > 0) ? (
                                                currentStructure.skills.map(s => (
                                                    <div key={s.id} className="py-1 px-2 hover:bg-gray-50 rounded text-gray-700 text-xs flex justify-between">
                                                        <span>{s.name}</span>
                                                        <span className="text-gray-400">ID: {s.id}</span>
                                                    </div>
                                                ))
                                            ) : (
                                                <span className="text-xs text-gray-400 italic">No skills currently in Planday</span>
                                            )}
                                        </div>
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
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 grid sm:grid-cols-2 lg:grid-cols-5 gap-4">
                            <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 flex items-center gap-4">
                                <div className="p-3 bg-blue-600 text-white rounded-lg">
                                    <BuildingOfficeIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-blue-900">{currentStructure.departments.length}</span>
                                    <span className="text-xs font-semibold text-blue-700 uppercase tracking-wider">Departments</span>
                                </div>
                            </div>

                            <div className="bg-purple-50 p-4 rounded-xl border border-purple-200 flex items-center gap-4">
                                <div className="p-3 bg-purple-600 text-white rounded-lg">
                                    <UserGroupIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-purple-900">{currentStructure.employeeGroups.length}</span>
                                    <span className="text-xs font-semibold text-purple-700 uppercase tracking-wider">Empl. Groups</span>
                                </div>
                            </div>

                            <div className="bg-gray-100 p-4 rounded-xl border border-gray-300 flex items-center gap-4">
                                <div className="p-3 bg-gray-700 text-white rounded-lg">
                                    <SectionIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-gray-900">{currentStructure.sections?.length || 0}</span>
                                    <span className="text-xs font-semibold text-gray-700 uppercase tracking-wider">Sections</span>
                                </div>
                            </div>

                            <div className="bg-teal-50 p-4 rounded-xl border border-teal-200 flex items-center gap-4">
                                <div className="p-3 bg-teal-600 text-white rounded-lg">
                                    <BriefcaseIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-teal-900">{currentStructure.positions?.length || 0}</span>
                                    <span className="text-xs font-semibold text-teal-700 uppercase tracking-wider">Positions</span>
                                </div>
                            </div>

                            <div className="bg-yellow-50 p-4 rounded-xl border border-yellow-200 flex items-center gap-4">
                                <div className="p-3 bg-yellow-600 text-white rounded-lg">
                                    <StarIcon className="w-6 h-6" />
                                </div>
                                <div>
                                    <span className="block text-2xl font-extrabold text-yellow-900">{currentStructure.skills?.length || 0}</span>
                                    <span className="text-xs font-semibold text-yellow-800 uppercase tracking-wider">Skills</span>
                                </div>
                            </div>
                        </div>

                        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200">
                            <h2 className="text-2xl font-bold text-gray-900 mb-6">{appMode === 'delete' ? 'Select Items to Delete' : 'Select Items to Update'}</h2>

                            {/* Input Method Tabs */}
                            <div className="flex flex-wrap items-center gap-2 mb-6">
                                {(['file', 'selection'] as const).map(method => (
                                    <button 
                                        key={method}
                                        onClick={() => setDeleteInputMethod(method)}
                                        className={`px-4 py-2 rounded-lg font-bold text-sm transition-colors ${
                                            deleteInputMethod === method 
                                                ? (appMode === 'update' ? 'bg-amber-100 text-amber-800 ring-1 ring-amber-300' : 'bg-red-100 text-red-800 ring-1 ring-red-300') 
                                                : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                                        }`}
                                    >
                                        {method === 'file' ? 'Upload File / Template' : 'Select from List'}
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
                                    <div className="space-y-6 text-sm">
                                        {/* Row 1: Departments and Employee Groups */}
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                            {/* Departments */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                            <div className="bg-gray-50 px-4 py-2.5 border-b border-gray-200 flex justify-between items-center">
                                                <div className="flex items-center gap-2">
                                                    <span className="font-bold text-gray-800">Departments</span>
                                                    <span className="text-xs bg-blue-100 text-blue-800 px-2 py-0.5 rounded-full font-semibold">
                                                        {deleteSelectedDepts.size}/{currentStructure.departments.length}
                                                    </span>
                                                </div>
                                                <button
                                                    type="button"
                                                    onClick={() => {
                                                        const filtered = currentStructure.departments.filter(d => d.name.toLowerCase().includes(deleteSearchTerm.toLowerCase()));
                                                        if (filtered.length > 0 && filtered.every(d => deleteSelectedDepts.has(d.id))) {
                                                            const newSet = new Set(deleteSelectedDepts);
                                                            filtered.forEach(d => newSet.delete(d.id));
                                                            setDeleteSelectedDepts(newSet);
                                                        } else {
                                                            const newSet = new Set(deleteSelectedDepts);
                                                            filtered.forEach(d => newSet.add(d.id));
                                                            setDeleteSelectedDepts(newSet);
                                                        }
                                                    }}
                                                    className="text-xs font-bold text-blue-600 hover:text-blue-800"
                                                >
                                                    {currentStructure.departments.filter(d => d.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).length > 0 && currentStructure.departments.filter(d => d.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).every(d => deleteSelectedDepts.has(d.id)) ? 'Deselect All' : 'Select All'}
                                                </button>
                                            </div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1 flex-1">
                                                {currentStructure.departments.filter(d => d.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).map(d => (
                                                    <label key={d.id} className="flex items-center justify-between gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                        <div className="flex items-center gap-2.5 truncate">
                                                            <input 
                                                                type="checkbox" 
                                                                checked={deleteSelectedDepts.has(d.id)}
                                                                onChange={(e) => {
                                                                    const newSet = new Set(deleteSelectedDepts);
                                                                    if (e.target.checked) newSet.add(d.id);
                                                                    else newSet.delete(d.id);
                                                                    setDeleteSelectedDepts(newSet);
                                                                }}
                                                                className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                                                            />
                                                            <span className="text-xs text-gray-800 font-medium truncate">{d.name}</span>
                                                        </div>
                                                        <span className="text-[10px] text-gray-400 shrink-0">ID: {d.id}</span>
                                                    </label>
                                                ))}
                                                {currentStructure.departments.length === 0 && (
                                                    <span className="text-xs text-gray-400 italic block p-2">No departments found</span>
                                                )}
                                            </div>
                                        </div>

                                        {/* Employee Groups */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                            <div className="bg-gray-50 px-4 py-2.5 border-b border-gray-200 flex justify-between items-center">
                                                <div className="flex items-center gap-2">
                                                    <span className="font-bold text-gray-800">Employee Groups</span>
                                                    <span className="text-xs bg-purple-100 text-purple-800 px-2 py-0.5 rounded-full font-semibold">
                                                        {deleteSelectedGroups.size}/{currentStructure.employeeGroups.length}
                                                    </span>
                                                </div>
                                                <button
                                                    type="button"
                                                    onClick={() => {
                                                        const filtered = currentStructure.employeeGroups.filter(g => g.name.toLowerCase().includes(deleteSearchTerm.toLowerCase()));
                                                        if (filtered.length > 0 && filtered.every(g => deleteSelectedGroups.has(g.id))) {
                                                            const newSet = new Set(deleteSelectedGroups);
                                                            filtered.forEach(g => newSet.delete(g.id));
                                                            setDeleteSelectedGroups(newSet);
                                                        } else {
                                                            const newSet = new Set(deleteSelectedGroups);
                                                            filtered.forEach(g => newSet.add(g.id));
                                                            setDeleteSelectedGroups(newSet);
                                                        }
                                                    }}
                                                    className="text-xs font-bold text-purple-600 hover:text-purple-800"
                                                >
                                                    {currentStructure.employeeGroups.filter(g => g.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).length > 0 && currentStructure.employeeGroups.filter(g => g.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).every(g => deleteSelectedGroups.has(g.id)) ? 'Deselect All' : 'Select All'}
                                                </button>
                                            </div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1 flex-1">
                                                {currentStructure.employeeGroups.filter(g => g.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).map(g => (
                                                    <label key={g.id} className="flex items-center justify-between gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                        <div className="flex items-center gap-2.5 truncate">
                                                            <input 
                                                                type="checkbox" 
                                                                checked={deleteSelectedGroups.has(g.id)}
                                                                onChange={(e) => {
                                                                    const newSet = new Set(deleteSelectedGroups);
                                                                    if (e.target.checked) newSet.add(g.id);
                                                                    else newSet.delete(g.id);
                                                                    setDeleteSelectedGroups(newSet);
                                                                }}
                                                                className="w-4 h-4 text-purple-600 rounded focus:ring-purple-500"
                                                            />
                                                            <span className="text-xs text-gray-800 font-medium truncate">{g.name}</span>
                                                        </div>
                                                        <span className="text-[10px] text-gray-400 shrink-0">ID: {g.id}</span>
                                                    </label>
                                                ))}
                                                {currentStructure.employeeGroups.length === 0 && (
                                                    <span className="text-xs text-gray-400 italic block p-2">No employee groups found</span>
                                                )}
                                            </div>
                                        </div>
                                    </div>

                                    {/* Row 2: Sections and Positions */}
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                        {/* Sections */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                            <div className="bg-gray-50 px-4 py-2.5 border-b border-gray-200 flex justify-between items-center">
                                                <div className="flex items-center gap-2">
                                                    <span className="font-bold text-gray-800">Sections</span>
                                                    <span className="text-xs bg-gray-200 text-gray-800 px-2 py-0.5 rounded-full font-semibold">
                                                        {deleteSelectedSections.size}/{(currentStructure.sections || []).length}
                                                    </span>
                                                </div>
                                                <button
                                                    type="button"
                                                    onClick={() => {
                                                        const filtered = (currentStructure.sections || []).filter(s => {
                                                            const matchesSearch = s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase());
                                                            const matchesDept = selectListSectionDeptId === 'all' || !selectListSectionDeptId || String(s.departmentId) === String(selectListSectionDeptId);
                                                            return matchesSearch && matchesDept;
                                                        });
                                                        if (filtered.length > 0 && filtered.every(s => deleteSelectedSections.has(s.id))) {
                                                            const newSet = new Set(deleteSelectedSections);
                                                            filtered.forEach(s => newSet.delete(s.id));
                                                            setDeleteSelectedSections(newSet);
                                                        } else {
                                                            const newSet = new Set(deleteSelectedSections);
                                                            filtered.forEach(s => newSet.add(s.id));
                                                            setDeleteSelectedSections(newSet);
                                                        }
                                                    }}
                                                    className="text-xs font-bold text-gray-700 hover:text-gray-900"
                                                >
                                                    {(() => {
                                                        const filtered = (currentStructure.sections || []).filter(s => {
                                                            const matchesSearch = s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase());
                                                            const matchesDept = selectListSectionDeptId === 'all' || !selectListSectionDeptId || String(s.departmentId) === String(selectListSectionDeptId);
                                                            return matchesSearch && matchesDept;
                                                        });
                                                        return filtered.length > 0 && filtered.every(s => deleteSelectedSections.has(s.id)) ? 'Deselect All' : 'Select All';
                                                    })()}
                                                </button>
                                            </div>
                                            <div className="p-2.5 bg-slate-50 border-b border-gray-200">
                                                <label className="block text-[11px] font-semibold text-gray-600 mb-1">
                                                    Select Department:
                                                </label>
                                                <select
                                                    value={selectListSectionDeptId}
                                                    onChange={(e) => setSelectListSectionDeptId(e.target.value)}
                                                    className="w-full text-xs bg-white border border-gray-300 rounded-lg px-2 py-1.5 focus:ring-2 focus:ring-slate-500 text-gray-800 font-medium"
                                                >
                                                    <option value="all">All Departments</option>
                                                    {currentStructure.departments.map(d => (
                                                        <option key={d.id} value={d.id}>{d.name}</option>
                                                    ))}
                                                </select>
                                            </div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1 flex-1">
                                                {(currentStructure.sections || [])
                                                    .filter(s => {
                                                        const matchesSearch = s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase());
                                                        const matchesDept = selectListSectionDeptId === 'all' || !selectListSectionDeptId || String(s.departmentId) === String(selectListSectionDeptId);
                                                        return matchesSearch && matchesDept;
                                                    })
                                                    .map(s => {
                                                        const dept = currentStructure.departments.find(d => d.id === s.departmentId);
                                                        return (
                                                            <label key={s.id} className="flex items-center justify-between gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                                <div className="flex items-center gap-2 truncate">
                                                                    <input 
                                                                        type="checkbox" 
                                                                        checked={deleteSelectedSections.has(s.id)}
                                                                        onChange={(e) => {
                                                                            const newSet = new Set(deleteSelectedSections);
                                                                            if (e.target.checked) newSet.add(s.id);
                                                                            else newSet.delete(s.id);
                                                                            setDeleteSelectedSections(newSet);
                                                                        }}
                                                                        className="w-4 h-4 text-gray-700 rounded focus:ring-gray-600"
                                                                    />
                                                                    {s.color && (
                                                                        <span className="inline-block w-2.5 h-2.5 rounded-full border border-gray-200 shrink-0" style={{ backgroundColor: s.color }} />
                                                                    )}
                                                                    <div className="truncate">
                                                                        <span className="text-xs text-gray-800 font-medium truncate block">{s.name}</span>
                                                                        <span className="text-[10px] text-gray-400 block truncate">({dept?.name || `Dept #${s.departmentId}`})</span>
                                                                    </div>
                                                                </div>
                                                                <span className="text-[10px] text-gray-400 shrink-0">ID: {s.id}</span>
                                                            </label>
                                                        );
                                                    })}
                                                {(currentStructure.sections || []).length === 0 && (
                                                    <span className="text-xs text-gray-400 italic block p-2">No sections found</span>
                                                )}
                                            </div>
                                        </div>

                                        {/* Positions */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                            <div className="bg-gray-50 px-4 py-2.5 border-b border-gray-200 flex justify-between items-center">
                                                <div className="flex items-center gap-2">
                                                    <span className="font-bold text-gray-800">Positions</span>
                                                    <span className="text-xs bg-teal-100 text-teal-800 px-2 py-0.5 rounded-full font-semibold">
                                                        {deleteSelectedPositions.size}/{(currentStructure.positions || []).length}
                                                    </span>
                                                </div>
                                                <button
                                                    type="button"
                                                    onClick={() => {
                                                        const filtered = (currentStructure.positions || []).filter(p => {
                                                            const matchesSearch = p.name.toLowerCase().includes(deleteSearchTerm.toLowerCase());
                                                            const matchesDept = selectListPositionDeptId === 'all' || !selectListPositionDeptId || String(p.departmentId) === String(selectListPositionDeptId);
                                                            return matchesSearch && matchesDept;
                                                        });
                                                        if (filtered.length > 0 && filtered.every(p => deleteSelectedPositions.has(p.id))) {
                                                            const newSet = new Set(deleteSelectedPositions);
                                                            filtered.forEach(p => newSet.delete(p.id));
                                                            setDeleteSelectedPositions(newSet);
                                                        } else {
                                                            const newSet = new Set(deleteSelectedPositions);
                                                            filtered.forEach(p => newSet.add(p.id));
                                                            setDeleteSelectedPositions(newSet);
                                                        }
                                                    }}
                                                    className="text-xs font-bold text-teal-600 hover:text-teal-800"
                                                >
                                                    {(() => {
                                                        const filtered = (currentStructure.positions || []).filter(p => {
                                                            const matchesSearch = p.name.toLowerCase().includes(deleteSearchTerm.toLowerCase());
                                                            const matchesDept = selectListPositionDeptId === 'all' || !selectListPositionDeptId || String(p.departmentId) === String(selectListPositionDeptId);
                                                            return matchesSearch && matchesDept;
                                                        });
                                                        return filtered.length > 0 && filtered.every(p => deleteSelectedPositions.has(p.id)) ? 'Deselect All' : 'Select All';
                                                    })()}
                                                </button>
                                            </div>
                                            <div className="p-2.5 bg-teal-50/40 border-b border-gray-200">
                                                <label className="block text-[11px] font-semibold text-gray-600 mb-1">
                                                    Select Department:
                                                </label>
                                                <select
                                                    value={selectListPositionDeptId}
                                                    onChange={(e) => setSelectListPositionDeptId(e.target.value)}
                                                    className="w-full text-xs bg-white border border-gray-300 rounded-lg px-2 py-1.5 focus:ring-2 focus:ring-teal-500 text-gray-800 font-medium"
                                                >
                                                    <option value="all">All Departments</option>
                                                    {currentStructure.departments.map(d => (
                                                        <option key={d.id} value={d.id}>{d.name}</option>
                                                    ))}
                                                </select>
                                            </div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1 flex-1">
                                                {(currentStructure.positions || [])
                                                    .filter(p => {
                                                        const matchesSearch = p.name.toLowerCase().includes(deleteSearchTerm.toLowerCase());
                                                        const matchesDept = selectListPositionDeptId === 'all' || !selectListPositionDeptId || String(p.departmentId) === String(selectListPositionDeptId);
                                                        return matchesSearch && matchesDept;
                                                    })
                                                    .map(p => {
                                                        const group = currentStructure.employeeGroups.find(g => g.id === p.employeeGroupId);
                                                        return (
                                                            <label key={p.id} className="flex items-center justify-between gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                                <div className="flex items-center gap-2 truncate">
                                                                    <input 
                                                                        type="checkbox" 
                                                                        checked={deleteSelectedPositions.has(p.id)}
                                                                        onChange={(e) => {
                                                                            const newSet = new Set(deleteSelectedPositions);
                                                                            if (e.target.checked) newSet.add(p.id);
                                                                            else newSet.delete(p.id);
                                                                            setDeleteSelectedPositions(newSet);
                                                                        }}
                                                                        className="w-4 h-4 text-teal-600 rounded focus:ring-teal-500"
                                                                    />
                                                                    <div className="truncate flex-1 min-w-0">
                                                                        <div className="flex items-center gap-1.5">
                                                                            <span className="text-xs text-gray-800 font-medium truncate block">{p.name}</span>
                                                                            {p.skillIds && p.skillIds.length > 0 && currentStructure.skills && (
                                                                                <span 
                                                                                    className="inline-flex items-center px-1 py-0.2 rounded text-[9px] font-medium bg-amber-50 text-amber-700 border border-amber-200 shrink-0"
                                                                                    title={`Skills: ${currentStructure.skills.filter(s => p.skillIds?.includes(s.id)).map(s => s.name).join(', ')}`}
                                                                                >
                                                                                    {p.skillIds.length} {p.skillIds.length === 1 ? 'skill' : 'skills'}
                                                                                </span>
                                                                            )}
                                                                        </div>
                                                                        {group && (
                                                                            <span className="text-[10px] text-gray-400 block truncate">({group.name})</span>
                                                                        )}
                                                                    </div>
                                                                </div>
                                                                <span className="text-[10px] text-gray-400 shrink-0">ID: {p.id}</span>
                                                            </label>
                                                        );
                                                    })}
                                                {(currentStructure.positions || []).length === 0 && (
                                                    <span className="text-xs text-gray-400 italic block p-2">No positions found</span>
                                                )}
                                            </div>
                                        </div>
                                    </div>

                                    {/* Row 3: Skills */}
                                    <div className="grid grid-cols-1 gap-6">
                                        {/* Skills */}
                                        <div className="border border-gray-200 rounded-xl overflow-hidden flex flex-col bg-white">
                                            <div className="bg-gray-50 px-4 py-2.5 border-b border-gray-200 flex justify-between items-center">
                                                <div className="flex items-center gap-2">
                                                    <span className="font-bold text-gray-800">Skills</span>
                                                    <span className="text-xs bg-yellow-100 text-yellow-900 border border-yellow-300 px-2 py-0.5 rounded-full font-semibold">
                                                        {deleteSelectedSkills.size}/{(currentStructure.skills || []).length}
                                                    </span>
                                                </div>
                                                <button
                                                    type="button"
                                                    onClick={() => {
                                                        const filtered = (currentStructure.skills || []).filter(s => s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase()));
                                                        if (filtered.length > 0 && filtered.every(s => deleteSelectedSkills.has(s.id))) {
                                                            const newSet = new Set(deleteSelectedSkills);
                                                            filtered.forEach(s => newSet.delete(s.id));
                                                            setDeleteSelectedSkills(newSet);
                                                        } else {
                                                            const newSet = new Set(deleteSelectedSkills);
                                                            filtered.forEach(s => newSet.add(s.id));
                                                            setDeleteSelectedSkills(newSet);
                                                        }
                                                    }}
                                                    className="text-xs font-bold text-yellow-700 hover:text-yellow-900"
                                                >
                                                    {(currentStructure.skills || []).filter(s => s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).length > 0 && (currentStructure.skills || []).filter(s => s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase())).every(s => deleteSelectedSkills.has(s.id)) ? 'Deselect All' : 'Select All'}
                                                </button>
                                            </div>
                                            <div className="max-h-64 overflow-y-auto p-3 space-y-1 flex-1">
                                                {(currentStructure.skills || [])
                                                    .filter(s => s.name.toLowerCase().includes(deleteSearchTerm.toLowerCase()))
                                                    .map(s => (
                                                        <label key={s.id} className="flex items-center justify-between gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer">
                                                            <div className="flex items-center gap-2.5 truncate">
                                                                <input 
                                                                    type="checkbox" 
                                                                    checked={deleteSelectedSkills.has(s.id)}
                                                                    onChange={(e) => {
                                                                        const newSet = new Set(deleteSelectedSkills);
                                                                        if (e.target.checked) newSet.add(s.id);
                                                                        else newSet.delete(s.id);
                                                                        setDeleteSelectedSkills(newSet);
                                                                    }}
                                                                    className="w-4 h-4 text-yellow-600 rounded focus:ring-yellow-500"
                                                                />
                                                                <span className="text-xs text-gray-800 font-medium truncate">{s.name}</span>
                                                            </div>
                                                            <span className="text-[10px] text-gray-400 shrink-0">ID: {s.id}</span>
                                                        </label>
                                                    ))}
                                                {(currentStructure.skills || []).length === 0 && (
                                                    <span className="text-xs text-gray-400 italic block p-2">No skills found</span>
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                {/* Action Footer */}
                                    <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pt-4 border-t border-gray-200">
                                        <div className="text-sm text-gray-600">
                                            <span className="font-bold text-gray-900">
                                                {deleteSelectedDepts.size + deleteSelectedGroups.size + deleteSelectedSections.size + deleteSelectedPositions.size + deleteSelectedSkills.size}
                                            </span> items selected across all types
                                        </div>
                                        <div className="flex items-center gap-3 w-full sm:w-auto">
                                            {(deleteSelectedDepts.size + deleteSelectedGroups.size + deleteSelectedSections.size + deleteSelectedPositions.size + deleteSelectedSkills.size) > 0 && (
                                                <button
                                                    type="button"
                                                    onClick={() => {
                                                        setDeleteSelectedDepts(new Set());
                                                        setDeleteSelectedGroups(new Set());
                                                        setDeleteSelectedSections(new Set());
                                                        setDeleteSelectedPositions(new Set());
                                                        setDeleteSelectedSkills(new Set());
                                                    }}
                                                    className="px-4 py-2.5 text-sm font-semibold text-gray-600 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors"
                                                >
                                                    Clear All
                                                </button>
                                            )}
                                            <button 
                                                onClick={generateActionReviewList}
                                                disabled={deleteSelectedDepts.size === 0 && deleteSelectedSections.size === 0 && deleteSelectedGroups.size === 0 && deleteSelectedPositions.size === 0 && deleteSelectedSkills.size === 0}
                                                className={`w-full sm:w-auto px-6 py-3 font-bold rounded-xl shadow-sm transition-all disabled:opacity-50 text-white ${
                                                    appMode === 'update' 
                                                        ? 'bg-amber-600 hover:bg-amber-700' 
                                                        : 'bg-red-600 hover:bg-red-700'
                                                }`}
                                            >
                                                {appMode === 'update' ? 'Configure & Review' : 'Review Selected Items'} ({deleteSelectedDepts.size + deleteSelectedSections.size + deleteSelectedGroups.size + deleteSelectedPositions.size + deleteSelectedSkills.size} items)
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            )}

                            {/* Method: File */}
                            {deleteInputMethod === 'file' && (
                                <div className="grid md:grid-cols-2 gap-8">
                                    {/* Panel A: Template Generator / Structure Export */}
                                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 flex flex-col justify-between">
                                        <div>
                                            <div className="flex items-center gap-3 mb-3">
                                                <div className={`p-2 rounded-lg ${appMode === 'update' ? 'bg-amber-100 text-amber-700' : 'bg-red-100 text-red-700'}`}>
                                                    <DownloadIcon className="w-5 h-5" />
                                                </div>
                                                <div>
                                                    <h3 className="text-lg font-bold text-gray-900">
                                                        1. Download {appMode === 'update' ? 'Update' : 'Delete'} Template
                                                    </h3>
                                                    <span className="text-xs font-semibold text-gray-500">
                                                        {appMode === 'update' ? 'Includes update columns & Reference Data sheet' : 'Simple sheet with Names only & Reference Data'}
                                                    </span>
                                                </div>
                                            </div>
                                            <p className="text-xs text-gray-600 mb-5">
                                                {appMode === 'update'
                                                    ? "Select which structure elements to include. The template provides columns for your updates, and the Reference Data sheet lists all existing API IDs."
                                                    : "Select which elements to include. The delete sheet only includes the item names for matching, while the Reference Data sheet lists your existing API IDs."}
                                            </p>

                                            {/* Selection Controls */}
                                            <div className="space-y-2.5 bg-gray-50 p-4 rounded-xl border border-gray-200 mb-5">
                                                <span className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-2">Include in Excel file:</span>
                                                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                                    <input 
                                                        type="checkbox" 
                                                        checked={includeDepartments} 
                                                        onChange={(e) => setIncludeDepartments(e.target.checked)}
                                                        className={`w-4 h-4 rounded ${appMode === 'update' ? 'text-amber-600 focus:ring-amber-500' : 'text-red-600 focus:ring-red-500'}`}
                                                    />
                                                    <span>Departments {appMode === 'delete' ? '(Names only)' : ''}</span>
                                                </label>
                                                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                                    <input 
                                                        type="checkbox" 
                                                        checked={includeEmployeeGroups} 
                                                        onChange={(e) => setIncludeEmployeeGroups(e.target.checked)}
                                                        className={`w-4 h-4 rounded ${appMode === 'update' ? 'text-amber-600 focus:ring-amber-500' : 'text-red-600 focus:ring-red-500'}`}
                                                    />
                                                    <span>Employee Groups {appMode === 'delete' ? '(Names only)' : ''}</span>
                                                </label>
                                                {appMode === 'update' && includeEmployeeGroups && (
                                                    <label className="flex items-center gap-3 cursor-pointer text-xs font-medium text-gray-700 ml-7">
                                                        <input 
                                                            type="checkbox" 
                                                            checked={includeWageColumns} 
                                                            onChange={(e) => setIncludeWageColumns(e.target.checked)}
                                                            className="w-3.5 h-3.5 text-amber-600 rounded focus:ring-amber-500"
                                                        />
                                                        <span>Include Default Wage Data columns</span>
                                                    </label>
                                                )}
                                                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                                    <input 
                                                        type="checkbox" 
                                                        checked={includeSections} 
                                                        onChange={(e) => setIncludeSections(e.target.checked)}
                                                        className={`w-4 h-4 rounded ${appMode === 'update' ? 'text-amber-600 focus:ring-amber-500' : 'text-red-600 focus:ring-red-500'}`}
                                                    />
                                                    <span>Sections {appMode === 'delete' ? '(Names only)' : ''}</span>
                                                </label>
                                                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                                    <input 
                                                        type="checkbox" 
                                                        checked={includePositions} 
                                                        onChange={(e) => setIncludePositions(e.target.checked)}
                                                        className={`w-4 h-4 rounded ${appMode === 'update' ? 'text-amber-600 focus:ring-amber-500' : 'text-red-600 focus:ring-red-500'}`}
                                                    />
                                                    <span>Positions {appMode === 'delete' ? '(Names only)' : ''}</span>
                                                </label>
                                                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800">
                                                    <input 
                                                        type="checkbox" 
                                                        checked={includeSkills} 
                                                        onChange={(e) => setIncludeSkills(e.target.checked)}
                                                        className={`w-4 h-4 rounded ${appMode === 'update' ? 'text-amber-600 focus:ring-amber-500' : 'text-red-600 focus:ring-red-500'}`}
                                                    />
                                                    <span>Skills {appMode === 'delete' ? '(Names only)' : ''}</span>
                                                </label>
                                            </div>
                                        </div>

                                        <div className="space-y-2.5 mt-2">
                                            <button 
                                                type="button"
                                                onClick={populateWithCurrent ? handleDownloadCurrentStructure : handleDownloadBlankTemplate}
                                                disabled={!includeDepartments && !includeSections && !includeEmployeeGroups && !includeSkills && !includePositions}
                                                className={`w-full py-2.5 px-4 text-white font-bold rounded-xl shadow-sm hover:shadow transition-all flex items-center justify-center gap-2 text-sm disabled:opacity-50 ${
                                                    appMode === 'update' ? 'bg-amber-600 hover:bg-amber-700' : 'bg-red-600 hover:bg-red-700'
                                                }`}
                                            >
                                                <DownloadIcon className="w-4 h-4" />
                                                <span>Download Template</span>
                                            </button>
                                            
                                            <label className="flex items-center gap-2 cursor-pointer text-xs font-medium text-gray-700 justify-center">
                                                <input 
                                                    type="checkbox" 
                                                    checked={populateWithCurrent} 
                                                    onChange={(e) => setPopulateWithCurrent(e.target.checked)}
                                                    className={`w-4 h-4 rounded ${appMode === 'update' ? 'text-amber-600 focus:ring-amber-500' : 'text-red-600 focus:ring-red-500'}`}
                                                />
                                                <span>Populate with current portal structure?</span>
                                            </label>
                                        </div>
                                    </div>

                                    {/* Panel B: Upload File */}
                                    <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 flex flex-col justify-between">
                                        <div>
                                            <div className="flex items-center gap-3 mb-3">
                                                <div className={`p-2 rounded-lg ${appMode === 'update' ? 'bg-amber-100 text-amber-700' : 'bg-red-100 text-red-700'}`}>
                                                    <CloudUploadIcon className="w-5 h-5" />
                                                </div>
                                                <h3 className="text-lg font-bold text-gray-900">
                                                    2. Upload File for {appMode === 'update' ? 'Update' : 'Deletion'}
                                                </h3>
                                            </div>
                                            <p className="text-xs text-gray-600 mb-5">
                                                {appMode === 'update'
                                                    ? 'Upload your Excel file (.xlsx, .xls) or CSV. Columns will be matched in the next step.'
                                                    : 'Upload your Excel file (.xlsx, .xls) or CSV. Department and Employee Group names will be matched for deletion.'}
                                            </p>
                                        </div>

                                        <div className={`relative border-2 border-dashed ${
                                            appMode === 'update' 
                                                ? 'border-amber-300 hover:border-amber-500 bg-amber-50/30 hover:bg-amber-50/60' 
                                                : 'border-red-300 hover:border-red-500 bg-red-50/30 hover:bg-red-50/60'
                                        } rounded-2xl p-6 text-center transition-colors group flex-1 flex flex-col items-center justify-center min-h-[180px]`}>
                                            <input 
                                                type="file" 
                                                accept=".xlsx,.xls,.csv" 
                                                onChange={handleFileUpload}
                                                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                                            />
                                            <CloudUploadIcon className={`w-10 h-10 ${appMode === 'update' ? 'text-amber-500 group-hover:text-amber-600' : 'text-red-500 group-hover:text-red-600'} mx-auto mb-2 transition-colors`} />
                                            <span className="block text-sm font-bold text-gray-800">
                                                Click or Drag & Drop Excel File
                                            </span>
                                            <span className="block text-xs text-gray-500 mt-1">
                                                Supports .xlsx, .xls, .csv
                                            </span>
                                        </div>
                                    </div>
                                </div>
                            )}

                        </div>
                    </div>                )}                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && (appMode === 'import' || deleteInputMethod === 'file') && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <div className="flex justify-between items-center mb-2">
                                <h2 className="text-2xl font-bold text-gray-900">
                                    {appMode === 'delete' ? 'Map File Columns for Deletion' : appMode === 'update' ? 'Map File Columns for Update' : 'Map File Columns'}
                                </h2>
                                <span className="text-xs font-semibold text-gray-500 bg-gray-100 px-3 py-1 rounded-full">
                                    File: <strong>{uploadedFileName}</strong> ({parsedRows.length} rows)
                                </span>
                            </div>
                            <p className="text-sm text-gray-600">
                                {appMode === 'delete' 
                                    ? 'Select which columns from your uploaded file contain the Names of Departments and Employee Groups to delete.'
                                    : 'Select which columns from your uploaded file represent Departments and Employee Groups.'}
                            </p>
                        </div>

                        {/* Mapping Cards */}
                        <div className="space-y-6">
                            {/* Departments Card */}
                            <div className="bg-gray-50 p-6 rounded-2xl border border-gray-200">
                                <div className="flex justify-between items-center mb-4 border-b border-gray-200 pb-2">
                                    <h3 className="text-lg font-bold text-gray-900">Departments</h3>
                                    {appMode === 'update' && (
                                        <span className="text-xs bg-amber-100 text-amber-800 font-semibold px-2.5 py-0.5 rounded-full">
                                            Update Mode
                                        </span>
                                    )}
                                    {appMode === 'delete' && (
                                        <span className="text-xs bg-red-100 text-red-800 font-semibold px-2.5 py-0.5 rounded-full">
                                            Delete Mode
                                        </span>
                                    )}
                                </div>
                                <div className={`grid ${appMode === 'delete' ? 'grid-cols-1 md:grid-cols-2' : 'md:grid-cols-3'} gap-4`}>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">
                                            {appMode === 'delete' ? 'Department Name (to Delete):' : appMode === 'update' ? 'Department (Match by):' : 'Departments:'}
                                        </label>
                                        <span className="block text-xs text-gray-500 mb-2">
                                            {appMode === 'delete' ? 'Column with department names to delete' : appMode === 'update' ? 'Existing department name in Planday' : 'Department name'}
                                        </span>
                                        <select value={deptColumn} onChange={(e) => setDeptColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    {appMode !== 'delete' && (
                                        <div>
                                            <label className="block text-sm font-bold text-gray-800 mb-1 flex items-center gap-1.5">
                                                <span>UPDATE - Department Name:</span>
                                                <span className="text-xs font-normal text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">New Name</span>
                                            </label>
                                            <span className="block text-xs text-gray-500 mb-2">
                                                New name to rename department to (Optional)
                                            </span>
                                            <select value={deptUpdateNameColumn} onChange={(e) => setDeptUpdateNameColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-amber-500 font-medium">
                                                <option value="">-- None --</option>
                                                {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                    )}
                                    {appMode !== 'delete' && (
                                        <div>
                                            <label className="block text-sm font-bold text-gray-800 mb-1">
                                                {appMode === 'update' ? 'Department Number (Optional):' : 'Dept Number (Opt):'}
                                            </label>
                                            <span className="block text-xs text-gray-500 mb-2">
                                                {appMode === 'update' ? 'New or updated department number' : 'Optional department number'}
                                            </span>
                                            <select value={deptNumberColumn} onChange={(e) => setDeptNumberColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                                <option value="">-- None --</option>
                                                {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                    )}
                                </div>
                            </div>

                            {/* Employee Groups Card */}
                            <div className="bg-gray-50 p-6 rounded-2xl border border-gray-200">
                                <div className="flex justify-between items-center mb-4 border-b border-gray-200 pb-2">
                                    <h3 className="text-lg font-bold text-gray-900">Employee Groups</h3>
                                    {appMode === 'delete' && (
                                        <span className="text-xs bg-red-100 text-red-800 font-semibold px-2.5 py-0.5 rounded-full">
                                            Delete Mode
                                        </span>
                                    )}
                                </div>
                                <div className={`grid ${appMode === 'delete' ? 'grid-cols-1 md:grid-cols-2' : 'md:grid-cols-2 lg:grid-cols-3'} gap-4`}>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-2">
                                            {appMode === 'delete' ? 'Employee Group Name (to Delete):' : 'Employee Groups:'}
                                        </label>
                                        <select value={groupColumn} onChange={(e) => setGroupColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    {appMode !== 'delete' && (
                                        <>
                                            <div>
                                                <label className="block text-sm font-bold text-gray-800 mb-2">Group: Default Wage Type (HourlyRate/ShiftRate):</label>
                                                <select value={wageTypeColumn} onChange={(e) => setWageTypeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                                    <option value="">-- None --</option>
                                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-sm font-bold text-gray-800 mb-2">Group: Default Rate:</label>
                                                <select value={rateColumn} onChange={(e) => setRateColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                                    <option value="">-- None --</option>
                                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-sm font-bold text-gray-800 mb-2">Group: Salary Code:</label>
                                                <select value={salaryCodeColumn} onChange={(e) => setSalaryCodeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                                    <option value="">-- None --</option>
                                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-sm font-bold text-gray-800 mb-2">Group: Salary Code 2 (needs activation):</label>
                                                <select value={salaryCode2Column} onChange={(e) => setSalaryCode2Column(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                                    <option value="">-- None --</option>
                                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                                </select>
                                            </div>
                                        </>
                                    )}
                                </div>
                            </div>

                            {/* Sections Card */}
                            <div className="bg-gray-50 p-6 rounded-2xl border border-gray-200">
                                <div className="flex justify-between items-center mb-4 border-b border-gray-200 pb-2">
                                    <h3 className="text-lg font-bold text-gray-900">Sections</h3>
                                    {appMode === 'update' && (
                                        <span className="text-xs bg-amber-100 text-amber-800 font-semibold px-2.5 py-0.5 rounded-full">
                                            Update Mode
                                        </span>
                                    )}
                                    {appMode === 'delete' && (
                                        <span className="text-xs bg-red-100 text-red-800 font-semibold px-2.5 py-0.5 rounded-full">
                                            Delete Mode
                                        </span>
                                    )}
                                </div>
                                <div className={`grid ${appMode === 'delete' ? 'grid-cols-1 md:grid-cols-2' : appMode === 'update' ? 'md:grid-cols-2 lg:grid-cols-4' : 'md:grid-cols-3'} gap-4`}>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">
                                            {appMode === 'delete' ? 'Section Name (to Delete):' : appMode === 'update' ? 'Section (Match by):' : 'Section Name:'}
                                        </label>
                                        <span className="block text-xs text-gray-500 mb-2">
                                            {appMode === 'delete' ? 'Column with section names to delete' : appMode === 'update' ? 'Existing section name in Planday' : 'Section name'}
                                        </span>
                                        <select value={sectionNameColumn} onChange={(e) => setSectionNameColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-teal-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">
                                            Section Department:
                                        </label>
                                        <span className="block text-xs text-gray-500 mb-2">
                                            Department the section belongs to
                                        </span>
                                        <select value={sectionDeptColumn} onChange={(e) => setSectionDeptColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-teal-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    {appMode !== 'delete' && (
                                        <div>
                                            <label className="block text-sm font-bold text-gray-800 mb-1">
                                                Section Color (Opt):
                                            </label>
                                            <span className="block text-xs text-gray-500 mb-2">
                                                Hex color code (e.g. #2CD08E)
                                            </span>
                                            <select value={sectionColorColumn} onChange={(e) => setSectionColorColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-teal-500 font-medium">
                                                <option value="">-- None --</option>
                                                {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                    )}
                                    {appMode === 'update' && (
                                        <div>
                                            <label className="block text-sm font-bold text-gray-800 mb-1 flex items-center gap-1.5">
                                                <span>UPDATE - Section Name:</span>
                                                <span className="text-xs font-normal text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">New Name</span>
                                            </label>
                                            <span className="block text-xs text-gray-500 mb-2">
                                                New name for section (Optional)
                                            </span>
                                            <select value={sectionUpdateNameColumn} onChange={(e) => setSectionUpdateNameColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-amber-500 font-medium">
                                                <option value="">-- None --</option>
                                                {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                    )}
                                </div>
                            </div>

                            {/* Positions Card */}
                            {appMode !== 'delete' && (
                            <div className="bg-gray-50 p-6 rounded-2xl border border-gray-200">
                                <div className="flex justify-between items-center mb-4 border-b border-gray-200 pb-2">
                                    <h3 className="text-lg font-bold text-gray-900">Positions</h3>
                                    {appMode === 'update' && (
                                        <span className="text-xs bg-amber-100 text-amber-800 font-semibold px-2.5 py-0.5 rounded-full">
                                            Update Mode
                                        </span>
                                    )}
                                </div>
                                <div className="grid md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">
                                            {appMode === 'update' ? 'Position (Match by):' : 'Position Name:'}
                                        </label>
                                        <span className="block text-xs text-gray-500 mb-2">
                                            {appMode === 'update' ? 'Existing position name in Planday' : 'Position name'}
                                        </span>
                                        <select value={posNameColumn} onChange={(e) => setPosNameColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    {appMode === 'update' && (
                                        <div>
                                            <label className="block text-sm font-bold text-gray-800 mb-1 flex items-center gap-1.5">
                                                <span>UPDATE - Position Name:</span>
                                                <span className="text-xs font-normal text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">New Name</span>
                                            </label>
                                            <span className="block text-xs text-gray-500 mb-2">
                                                New name for position (Optional)
                                            </span>
                                            <select value={posUpdateNameColumn} onChange={(e) => setPosUpdateNameColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-amber-500 font-medium">
                                                <option value="">-- None --</option>
                                                {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                    )}
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">
                                            {appMode === 'update' ? 'Position Dept (Opt):' : 'Position Dept:'}
                                        </label>
                                        <span className="block text-xs text-gray-500 mb-2">
                                            {appMode === 'update' ? 'Department filter / scope (Opt)' : 'Department position belongs to'}
                                        </span>
                                        <select value={posDeptColumn} onChange={(e) => setPosDeptColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">
                                            {appMode === 'update' ? 'Position Group (Opt):' : 'Position Group:'}
                                        </label>
                                        <span className="block text-xs text-gray-500 mb-2">
                                            {appMode === 'update' ? 'Group filter / scope (Opt)' : 'Employee group position belongs to'}
                                        </span>
                                        <select value={posGroupColumn} onChange={(e) => setPosGroupColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. Section (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">Section assignment</span>
                                        <select value={posSectionColumn} onChange={(e) => setPosSectionColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. Color (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">Hex color code (e.g. #21E56E)</span>
                                        <select value={posColorColumn} onChange={(e) => setPosColorColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. Valid From (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">Valid from date</span>
                                        <select value={posValidFromColumn} onChange={(e) => setPosValidFromColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. Valid To (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">Valid to date</span>
                                        <select value={posValidToColumn} onChange={(e) => setPosValidToColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. Skills (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">Comma-separated skills</span>
                                        <select value={posSkillsColumn} onChange={(e) => setPosSkillsColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. In Payroll / Affect Rev (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">true/false (defaults to true)</span>
                                        <select value={posPayrollColumn} onChange={(e) => setPosPayrollColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-1">Pos. Rev Unit (Opt):</label>
                                        <span className="block text-xs text-gray-500 mb-2">Revenue unit name</span>
                                        <select value={posRevenueUnitColumn} onChange={(e) => setPosRevenueUnitColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-purple-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                </div>
                            </div>
                            )}

                            {/* Skills Card */}
                            {appMode !== 'delete' && (
                            <div className="bg-gray-50 p-6 rounded-2xl border border-gray-200">
                                <h3 className="text-lg font-bold text-gray-900 mb-4 border-b border-gray-200 pb-2">Skills</h3>
                                <div className="grid md:grid-cols-2 gap-4">
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-2">Skill Name:</label>
                                        <select value={skillNameColumn} onChange={(e) => setSkillNameColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-teal-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-gray-800 mb-2">Skill Description (Opt):</label>
                                        <select value={skillDescColumn} onChange={(e) => setSkillDescColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-teal-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                </div>
                            </div>
                            )}
                        </div>

                        {/* Mapped Data Live Preview Table */}
                        <div>
                            <h3 className="text-md font-bold text-gray-900 mb-3">Preview Mapped Data (First 10 Rows)</h3>
                            <div className="border border-gray-200 rounded-xl overflow-x-auto shadow-sm">
                                <table className="w-full text-left text-sm whitespace-nowrap">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-3 w-12 text-center">#</th>
                                            <th className="p-3">{appMode === 'delete' ? 'Department to Delete' : 'Department'} (Mapped: {deptColumn || 'None'})</th>
                                            {appMode !== 'delete' && deptUpdateNameColumn && (
                                                <th className="p-3 text-amber-700 bg-amber-50">UPDATE - Dept Name (Mapped: {deptUpdateNameColumn})</th>
                                            )}
                                            {appMode !== 'delete' && deptNumberColumn && (
                                                <th className="p-3">Dept Number (Mapped: {deptNumberColumn})</th>
                                            )}
                                            <th className="p-3">{appMode === 'delete' ? 'Employee Group to Delete' : 'Group'} (Mapped: {groupColumn || 'None'})</th>
                                            {sectionNameColumn && (
                                                <th className="p-3 text-emerald-800 bg-emerald-50/50">{appMode === 'delete' ? 'Section to Delete' : 'Section'} (Mapped: {sectionNameColumn})</th>
                                            )}
                                            {appMode === 'update' && sectionUpdateNameColumn && (
                                                <th className="p-3 text-amber-700 bg-amber-50">UPDATE - Section Name (Mapped: {sectionUpdateNameColumn})</th>
                                            )}
                                            {appMode !== 'delete' && (
                                                <>
                                                    <th className="p-3">Position Name (Mapped: {posNameColumn || 'None'})</th>
                                                    <th className="p-3">Position Dept (Mapped: {posDeptColumn || 'None'})</th>
                                                    <th className="p-3">Position Group (Mapped: {posGroupColumn || 'None'})</th>
                                                    <th className="p-3">Skill Name (Mapped: {skillNameColumn || 'None'})</th>
                                                </>
                                            )}
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {parsedRows.slice(0, 10).map((row, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50/80">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-gray-800">{deptColumn && row[deptColumn] ? row[deptColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                {appMode !== 'delete' && deptUpdateNameColumn && (
                                                    <td className="p-3 font-medium text-amber-900 bg-amber-50/50">{row[deptUpdateNameColumn] ? row[deptUpdateNameColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                )}
                                                {appMode !== 'delete' && deptNumberColumn && (
                                                    <td className="p-3 font-medium text-gray-800">{row[deptNumberColumn] !== undefined && row[deptNumberColumn] !== null && String(row[deptNumberColumn]).trim() !== '' ? row[deptNumberColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                )}
                                                <td className="p-3 font-medium text-gray-800">{groupColumn && row[groupColumn] ? row[groupColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                {sectionNameColumn && (
                                                    <td className="p-3 font-medium text-emerald-900 bg-emerald-50/30">{row[sectionNameColumn] ? row[sectionNameColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                )}
                                                {appMode === 'update' && sectionUpdateNameColumn && (
                                                    <td className="p-3 font-medium text-amber-900 bg-amber-50/50">{row[sectionUpdateNameColumn] ? row[sectionUpdateNameColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                )}
                                                {appMode !== 'delete' && (
                                                    <>
                                                        <td className="p-3 font-medium text-gray-800">{posNameColumn && row[posNameColumn] ? row[posNameColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                        <td className="p-3 font-medium text-gray-800">{posDeptColumn && row[posDeptColumn] ? row[posDeptColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                        <td className="p-3 font-medium text-gray-800">{posGroupColumn && row[posGroupColumn] ? row[posGroupColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                        <td className="p-3 font-medium text-gray-800">{skillNameColumn && row[skillNameColumn] ? row[skillNameColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                    </>
                                                )}
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
                                disabled={!deptColumn && !deptUpdateNameColumn && !groupColumn && !skillNameColumn && !posNameColumn && !sectionNameColumn}
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
                                {reviewItems.filter(i => i.type === 'section' && i.action === 'update').map(item => (
                                    <div key={item.id} className="border border-emerald-100 bg-emerald-50/20 rounded-xl p-4 shadow-sm">
                                        <div className="flex justify-between items-center mb-3">
                                            <h3 className="font-bold text-emerald-900">{item.name}</h3>
                                            {item.departmentName && (
                                                <span className="text-xs bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-medium">
                                                    Dept: {item.departmentName}
                                                </span>
                                            )}
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Section Name</label>
                                                <input 
                                                    type="text" 
                                                    value={item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'name', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-emerald-500"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Color (Hex)</label>
                                                <div className="flex items-center gap-2">
                                                    <input 
                                                        type="color" 
                                                        value={item.updatePayload?.color && /^#[0-9A-Fa-f]{6}$/.test(item.updatePayload.color) ? item.updatePayload.color : '#2CD08E'}
                                                        onChange={(e) => handleUpdatePayloadChange(item.id, 'color', e.target.value)}
                                                        className="w-9 h-9 p-0 border border-gray-300 rounded-lg cursor-pointer bg-white"
                                                    />
                                                    <input 
                                                        type="text" 
                                                        value={item.updatePayload?.color !== undefined ? item.updatePayload.color : ''}
                                                        onChange={(e) => handleUpdatePayloadChange(item.id, 'color', e.target.value)}
                                                        placeholder="#2CD08E"
                                                        className="flex-1 p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-emerald-500 font-mono"
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                ))}
                                {reviewItems.filter(i => i.type === 'position' && i.action === 'update').map(item => (
                                    <div key={item.id} className="border border-teal-100 bg-teal-50/20 rounded-xl p-4 shadow-sm">
                                        <div className="flex justify-between items-center mb-3">
                                            <h3 className="font-bold text-teal-900">{item.name}</h3>
                                            {item.updatePayload?.departmentName && (
                                                <span className="text-xs bg-teal-100 text-teal-800 px-2 py-0.5 rounded font-medium">
                                                    Dept: {item.updatePayload.departmentName}
                                                </span>
                                            )}
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Position Name</label>
                                                <input 
                                                    type="text" 
                                                    value={item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'name', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-teal-500"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Color (Hex)</label>
                                                <div className="flex items-center gap-2">
                                                    <input 
                                                        type="color" 
                                                        value={item.updatePayload?.color && /^#[0-9A-Fa-f]{6}$/.test(item.updatePayload.color) ? item.updatePayload.color : '#0F766E'}
                                                        onChange={(e) => handleUpdatePayloadChange(item.id, 'color', e.target.value)}
                                                        className="w-9 h-9 p-0 border border-gray-300 rounded-lg cursor-pointer bg-white"
                                                    />
                                                    <input 
                                                        type="text" 
                                                        value={item.updatePayload?.color !== undefined ? item.updatePayload.color : ''}
                                                        onChange={(e) => handleUpdatePayloadChange(item.id, 'color', e.target.value)}
                                                        placeholder="#0F766E"
                                                        className="flex-1 p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-teal-500 font-mono"
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                ))}
                                {reviewItems.filter(i => i.type === 'skill' && i.action === 'update').map(item => (
                                    <div key={item.id} className="border border-yellow-200 bg-yellow-50/20 rounded-xl p-4 shadow-sm">
                                        <h3 className="font-bold text-yellow-900 mb-3">{item.name}</h3>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Skill Name</label>
                                                <input 
                                                    type="text" 
                                                    value={item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'name', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-yellow-500"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">Description</label>
                                                <input 
                                                    type="text" 
                                                    value={item.updatePayload?.description !== undefined ? item.updatePayload.description : ''}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'description', e.target.value)}
                                                    placeholder="Skill description..."
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-yellow-500"
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
                        {/* Validation Warnings Banner */}
                        {validationWarnings.length > 0 && (
                            <div className="bg-amber-50 border-l-4 border-amber-500 p-4 rounded-r-lg shadow-sm">
                                <div className="flex">
                                    <div className="flex-shrink-0">
                                        <AlertIcon className="h-5 w-5 text-amber-500" />
                                    </div>
                                    <div className="ml-3">
                                        <h3 className="text-sm font-medium text-amber-800">
                                            Some data was excluded during parsing
                                        </h3>
                                        <div className="mt-2 text-sm text-amber-700">
                                            <ul className="list-disc pl-5 space-y-1">
                                                {validationWarnings.map((warning, idx) => (
                                                    <li key={idx}>{warning}</li>
                                                ))}
                                            </ul>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* Overview Stats */}
                        <div className={`grid grid-cols-2 sm:grid-cols-3 ${overviewTiles.length > 6 ? 'lg:grid-cols-4 xl:grid-cols-7' : 'lg:grid-cols-6'} gap-4`}>
                            {overviewTiles.map((tile, idx) => (
                                <div key={idx} className="bg-white p-5 rounded-2xl shadow-sm border border-gray-200">
                                    <span className="block text-xs font-bold text-gray-500 uppercase tracking-wider">{tile.label}</span>
                                    <span className={`text-2xl font-extrabold ${tile.color} mt-1 block`}>{tile.count}</span>
                                </div>
                            ))}
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
                                    <option value="section">Section</option>
                                </select>
                                {newItemType === 'section' && (
                                    <>
                                        <select
                                            value={newItemDeptId}
                                            onChange={(e) => setNewItemDeptId(e.target.value ? Number(e.target.value) : '')}
                                            className="px-3.5 py-2 border border-gray-300 rounded-xl text-sm bg-white font-medium"
                                        >
                                            <option value="">-- Select Department --</option>
                                            {currentStructure.departments.map(d => (
                                                <option key={d.id} value={d.id}>{d.name}</option>
                                            ))}
                                        </select>
                                        <div className="flex items-center gap-1.5">
                                            <input 
                                                type="color" 
                                                value={newItemColor && /^#[0-9A-Fa-f]{6}$/.test(newItemColor) ? newItemColor : '#2CD08E'} 
                                                onChange={(e) => setNewItemColor(e.target.value)}
                                                className="w-9 h-9 p-0 border border-gray-300 rounded-lg cursor-pointer bg-white"
                                            />
                                            <input 
                                                type="text" 
                                                value={newItemColor} 
                                                onChange={(e) => setNewItemColor(e.target.value)}
                                                placeholder="#2CD08E"
                                                className="w-24 px-2 py-2 border border-gray-300 rounded-xl text-xs bg-white font-mono"
                                            />
                                        </div>
                                    </>
                                )}
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
                                    {reviewStats.toCreate > 0 && (
                                        <button 
                                            onClick={() => setReviewTab('create')}
                                            className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-colors ${reviewTab === 'create' ? 'bg-blue-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
                                        >
                                            To Create ({reviewStats.toCreate})
                                        </button>
                                    )}
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
                                                        <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                                                            item.type === 'department' ? 'bg-blue-100 text-blue-800' : 
                                                            item.type === 'employee_group' ? 'bg-purple-100 text-purple-800' :
                                                            item.type === 'section' ? 'bg-gray-200 text-gray-800' :
                                                            item.type === 'position' ? 'bg-teal-100 text-teal-800' :
                                                            'bg-yellow-100 text-yellow-900 border border-yellow-300'
                                                        }`}>
                                                            {item.type === 'department' ? 'Department' : 
                                                             item.type === 'section' ? 'Section' :
                                                             item.type === 'employee_group' ? 'Employee Group' :
                                                             item.type === 'skill' ? 'Skill' : 'Position'}
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
                                                        {item.type === 'department' && (
                                                            <div className="space-y-1">
                                                                {item.action === 'update' && item.updatePayload?.name && item.updatePayload.name !== item.name && (
                                                                    <div><span className="font-semibold text-amber-800">New Name:</span> <span className="font-bold text-gray-900">{item.updatePayload.name}</span></div>
                                                                )}
                                                                {item.action === 'update' && item.updatePayload?.number !== undefined && item.updatePayload.number !== (item.number || null) && (
                                                                    <div><span className="font-semibold text-amber-800">New Number:</span> <span className="font-bold text-gray-900">{item.updatePayload.number || '<None>'}</span></div>
                                                                )}
                                                                {!(item.action === 'update' && (item.updatePayload?.name !== item.name || (item.updatePayload?.number !== undefined && item.updatePayload.number !== (item.number || null)))) && (item.updatePayload?.number || item.number) && (
                                                                    <div><span className="font-semibold text-gray-800">Number:</span> {item.updatePayload?.number || item.number}</div>
                                                                )}
                                                            </div>
                                                        )}
                                                        {item.type === 'section' && (
                                                            <div className="space-y-1">
                                                                {item.updatePayload?.departmentName && <div><span className="font-semibold text-gray-800">Dept:</span> {item.updatePayload.departmentName}</div>}
                                                                {item.action === 'update' && item.updatePayload?.name && item.updatePayload.name !== item.name && (
                                                                    <div><span className="font-semibold text-amber-800">New Name:</span> <span className="font-bold text-gray-900">{item.updatePayload.name}</span></div>
                                                                )}
                                                                {item.updatePayload?.color && (
                                                                    <div className="flex items-center gap-1.5">
                                                                        <span className="font-semibold text-gray-800">Color:</span>
                                                                        <span className="w-3 h-3 rounded-full inline-block border border-gray-300" style={{ backgroundColor: item.updatePayload.color }}></span>
                                                                        <span className="font-mono text-xs">{item.updatePayload.color}</span>
                                                                    </div>
                                                                )}
                                                            </div>
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
                                                        {item.type === 'skill' && item.updatePayload?.description && (
                                                            <div><span className="font-semibold text-gray-800">Description:</span> {item.updatePayload.description}</div>
                                                        )}
                                                        {item.type === 'position' && item.updatePayload && (
                                                            <div className="space-y-1">
                                                                {item.updatePayload.departmentName && <div><span className="font-semibold text-gray-800">Department:</span> {item.updatePayload.departmentName}</div>}
                                                                {item.updatePayload.employeeGroupName && <div><span className="font-semibold text-gray-800">Employee Group:</span> {item.updatePayload.employeeGroupName}</div>}
                                                                {item.action === 'update' && item.updatePayload.name && item.updatePayload.name !== item.name && (
                                                                    <div><span className="font-semibold text-amber-800">New Name:</span> <span className="font-bold text-gray-900">{item.updatePayload.name}</span></div>
                                                                )}
                                                                {item.updatePayload.color && (
                                                                    <div className="flex items-center gap-1.5">
                                                                        <span className="font-semibold text-gray-800">Color:</span>
                                                                        <span className="w-3 h-3 rounded-full inline-block border border-gray-300" style={{ backgroundColor: item.updatePayload.color }}></span>
                                                                        <span className="font-mono text-xs">{item.updatePayload.color}</span>
                                                                    </div>
                                                                )}
                                                                {item.updatePayload.sectionName && <div><span className="font-semibold text-gray-800">Section:</span> {item.updatePayload.sectionName}</div>}
                                                                {item.updatePayload.affectRevenue !== undefined && <div><span className="font-semibold text-gray-800">Payroll %:</span> {String(item.updatePayload.affectRevenue)}</div>}
                                                                {(item.updatePayload.validFrom || item.updatePayload.validTo) && <div><span className="font-semibold text-gray-800">Validity:</span> {item.updatePayload.validFrom || 'Any'} &rarr; {item.updatePayload.validTo || 'Any'}</div>}
                                                                {item.updatePayload.skills && <div><span className="font-semibold text-gray-800">Skills:</span> {item.updatePayload.skills}</div>}
                                                                {item.updatePayload.revenueUnitName && <div><span className="font-semibold text-gray-800">Rev Unit:</span> {item.updatePayload.revenueUnitName}</div>}
                                                            </div>
                                                        )}
                                                        {(item.type === 'department' && !(item.updatePayload?.name !== item.name || item.updatePayload?.number || item.number)) || 
                                                         (item.type === 'section' && !(item.updatePayload?.departmentName || item.updatePayload?.name || item.updatePayload?.color)) ||
                                                         (item.type === 'employee_group' && !item.updatePayload) ||
                                                         (item.type === 'position' && !item.updatePayload) ||
                                                         (item.type === 'skill' && !item.updatePayload?.description) ? (
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
                                onClick={() => {
                                    if (appMode === 'delete' && deleteInputMethod === 'selection') {
                                        setStep(1);
                                    } else {
                                        setStep(2);
                                    }
                                }}
                                className="px-5 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl transition-colors text-sm"
                            >
                                {(appMode === 'delete' && deleteInputMethod === 'selection') ? 'Back to Selection' : 'Back to Mapping / Config'}
                            </button>

                            <button 
                                onClick={() => setShowExecuteConfirm(true)}
                                disabled={reviewItems.filter(i => i.action === 'create' || i.action === 'delete' || i.action === 'update').length === 0}
                                className={`px-8 py-3.5 text-white font-bold rounded-xl shadow-lg transition-all text-sm flex items-center gap-2 ${
                                    reviewItems.filter(i => i.action === 'create' || i.action === 'delete' || i.action === 'update').length === 0
                                        ? 'bg-gray-400 cursor-not-allowed shadow-none'
                                        : appMode === 'delete' 
                                            ? 'bg-red-600 hover:bg-red-700 hover:shadow-xl' 
                                            : 'bg-green-600 hover:bg-green-700 hover:shadow-xl'
                                }`}
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
                                {appMode === 'delete' ? 'Deleting selected Departments, Employee Groups, Sections, Positions, and Skills in your Planday portal...' : appMode === 'update' ? 'Updating selected items in your Planday portal...' : 'Creating new Departments, Employee Groups, Sections, Positions, and Skills in your Planday portal...'}
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
                                        {[...reviewItems].sort((a, b) => {
                                            const orderA = TYPE_ORDER[a.type] || 99;
                                            const orderB = TYPE_ORDER[b.type] || 99;
                                            if (orderA !== orderB) return orderA - orderB;
                                            return a.name.localeCompare(b.name);
                                        }).map((item, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-xs">
                                                    {item.type === 'department' ? 'Department' : 
                                                     item.type === 'employee_group' ? 'Employee Group' :
                                                     item.type === 'section' ? 'Section' :
                                                     item.type === 'position' ? 'Position' :
                                                     item.type === 'skill' ? 'Skill' : 'Item'}
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
                                    onClick={() => { 
                                        setDeleteSelectedDepts(new Set());
                                        setDeleteSelectedGroups(new Set());
                                        setDeleteSelectedSections(new Set());
                                        setDeleteSelectedPositions(new Set());
                                        setDeleteSelectedSkills(new Set());
                                        setReviewItems([]);
                                        setExecutionCompleted(false);
                                        setStep(1); 
                                        setDeleteInputMethod('file'); 
                                        handleManualRefresh();
                                    }}
                                    className="px-6 py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow-md transition-all text-sm flex items-center gap-2"
                                >
                                    <RefreshIcon className="w-4 h-4" /> {appMode === 'delete' ? 'Start New Bulk Delete' : appMode === 'update' ? 'Start New Bulk Update' : 'Start New Import'}
                                </button>
                            </div>
                        )}
                    </div>
                )}
            </main>

            {/* Footer / App Info */}
            <footer className="py-6 text-center mt-auto">
                {/* 1. Relative wrapper with 'group' to listen for hover */}
                <div className="relative group inline-flex items-center gap-1 cursor-help">
                    <span className="text-gray-500 text-sm font-medium">App Info</span>
                    
                    {/* 2. Icon changes color when hovering the group */}
                    <InfoIcon className="h-4 w-4 text-gray-400 group-hover:text-blue-500 transition-colors" />

                    {/* 3. The Tooltip Popover (appears above on hover) */}
                    <div className="absolute bottom-full left-1/2 transform -translate-x-1/2 mb-3 w-80 p-4 bg-gray-900 text-white text-xs rounded-lg shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 text-center leading-relaxed">
                        <p className="mb-2">
                            <strong>Client-Side Processing & Secure:</strong> This app is a React-based JavaScript web application that runs entirely in your browser.
                        </p>
                        <p className="mb-3">
                            Your Excel files and employee data are processed locally in your browser and are never sent to any third‑party servers; they are only transmitted directly to the official <a href="https://openapi.planday.com/" target="_blank" rel="noopener noreferrer" className="text-blue-300 hover:text-blue-200 underline">Planday Open API</a> over a secure encrypted connection (HTTPS).
                        </p>
                        <p className="text-gray-400 border-t border-gray-700 pt-2 mt-2">Version 1.2</p>
                        <p className="text-gray-400">Made with ❤️ by the Community</p>

                        {/* Downward-pointing arrow triangle */}
                        <div className="absolute top-full left-1/2 transform -translate-x-1/2 border-4 border-transparent border-t-gray-900"></div>
                    </div>
                </div>
            </footer>

            {/* Token Guide Modal */}
            <TokenGuideModal 
                isOpen={showTokenGuide} 
                onClose={() => setShowTokenGuide(false)} 
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
                            {appMode === 'delete' ? (
                                <>
                                    {reviewStats.toDeleteDept > 0 && <li className="text-red-600">Delete {reviewStats.toDeleteDept} Departments</li>}
                                    {reviewStats.toDeleteGroup > 0 && <li className="text-red-600">Delete {reviewStats.toDeleteGroup} Employee Groups</li>}
                                    {reviewStats.toDeleteSection > 0 && <li className="text-red-600">Delete {reviewStats.toDeleteSection} Sections</li>}
                                    {reviewStats.toDeletePosition > 0 && <li className="text-red-600">Delete {reviewStats.toDeletePosition} Positions</li>}
                                    {reviewStats.toDeleteSkill > 0 && <li className="text-red-600">Delete {reviewStats.toDeleteSkill} Skills</li>}
                                    {reviewStats.toDelete === 0 && <li className="text-gray-500">No items selected to delete</li>}
                                </>
                            ) : appMode === 'update' ? (
                                <>
                                    {reviewStats.toUpdateDept > 0 && <li className="text-amber-600">Update {reviewStats.toUpdateDept} Departments</li>}
                                    {reviewStats.toUpdateGroup > 0 && <li className="text-amber-600">Update {reviewStats.toUpdateGroup} Employee Groups</li>}
                                    {reviewStats.toUpdateSection > 0 && <li className="text-amber-600">Update {reviewStats.toUpdateSection} Sections</li>}
                                    {reviewStats.toUpdatePosition > 0 && <li className="text-amber-600">Update {reviewStats.toUpdatePosition} Positions</li>}
                                    {reviewStats.toUpdateSkill > 0 && <li className="text-amber-600">Update {reviewStats.toUpdateSkill} Skills</li>}
                                    {reviewStats.toUpdate === 0 && <li className="text-gray-500">No items scheduled to update</li>}
                                </>
                            ) : (
                                <>
                                    {reviewStats.toCreateDept > 0 && <li>Create {reviewStats.toCreateDept} new Departments</li>}
                                    {reviewStats.toCreateGroup > 0 && <li>Create {reviewStats.toCreateGroup} new Employee Groups</li>}
                                    {reviewStats.toCreateSection > 0 && <li>Create {reviewStats.toCreateSection} new Sections</li>}
                                    {reviewStats.toCreatePosition > 0 && <li>Create {reviewStats.toCreatePosition} new Positions</li>}
                                    {reviewStats.toCreateSkill > 0 && <li>Create {reviewStats.toCreateSkill} new Skills</li>}
                                    {reviewStats.toUpdate > 0 && <li className="text-amber-600">Update {reviewStats.toUpdate} existing items</li>}
                                    {reviewStats.toDelete > 0 && <li className="text-red-600">Delete {reviewStats.toDelete} items</li>}
                                </>
                            )}
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
