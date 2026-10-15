import { PlandayApiCredentials, Department, EmployeeGroup, Skill, RevenueUnit, Position, Section } from '../types';

const AUTH_URL = 'https://id.planday.com/connect/token';
const REVOCATION_URL = 'https://id.planday.com/connect/revocation';
const API_BASE_URL = 'https://openapi.planday.com';

let credentials_internal: PlandayApiCredentials | null = null;
let accessToken: string | null = null;
let tokenExpiry: number | null = null;
let globalDelayUntil = 0;

export function initializeService(credentials: PlandayApiCredentials) {
    if (credentials_internal?.clientId !== credentials.clientId || credentials_internal?.refreshToken !== credentials.refreshToken) {
        accessToken = null;
        tokenExpiry = null;
    }
    credentials_internal = { ...credentials };
}

export function resetService() {
    credentials_internal = null;
    accessToken = null;
    tokenExpiry = null;
}

export async function revokeToken(): Promise<void> {
    try {
        let clientId = credentials_internal?.clientId;
        let refreshToken = credentials_internal?.refreshToken;

        if (!refreshToken) {
            const saved = sessionStorage.getItem('plandayCredentials');
            if (saved) {
                try {
                    const parsed = JSON.parse(saved);
                    clientId = clientId || parsed.clientId;
                    refreshToken = parsed.refreshToken;
                } catch { }
            }
        }

        if (clientId && refreshToken) {
            const payload = new URLSearchParams({
                'client_id': clientId,
                'token': refreshToken,
                'token_type_hint': 'refresh_token',
            });

            await fetch(REVOCATION_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: payload.toString(),
            });
        }
    } catch (err) {
        console.warn('Failed to revoke token with Planday:', err);
    } finally {
        resetService();
    }
}

async function wait(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function getAccessToken(): Promise<string> {
    if (accessToken && tokenExpiry && Date.now() < tokenExpiry) {
        return accessToken;
    }

    if (!credentials_internal) {
        throw new Error("Planday service not initialized with credentials.");
    }

    const payload = new URLSearchParams({
        'client_id': credentials_internal.clientId,
        'grant_type': 'refresh_token',
        'refresh_token': credentials_internal.refreshToken,
    });

    const response = await fetch(AUTH_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: payload.toString(),
    });

    if (!response.ok) {
        const errorText = await response.text();
        sessionStorage.removeItem('plandayCredentials');
        credentials_internal = null;
        accessToken = null;
        tokenExpiry = null;
        throw new Error(`Failed to refresh access token: ${response.status} ${errorText}. Your credentials may be invalid or expired. Please re-enter them.`);
    }

    const data = await response.json();
    accessToken = data.access_token;
    tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;

    if (data.refresh_token && data.refresh_token !== credentials_internal.refreshToken) {
        credentials_internal.refreshToken = data.refresh_token;
        sessionStorage.setItem('plandayCredentials', JSON.stringify(credentials_internal));
    }
    
    return accessToken;
}

// Optimized retry logic with NetworkError handling and exponential backoff
async function fetchWithAuth(url: string, options: RequestInit = { method: 'GET' }, retries = 5, backoffDelay = 2000): Promise<Response> {
    if (!credentials_internal) throw new Error("Service not initialized");

    const now = Date.now();
    if (globalDelayUntil > now) {
        await wait(globalDelayUntil - now);
    }

    const token = await getAccessToken();
    const headers = {
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0',
        ...options.headers,
        'Authorization': `Bearer ${token}`,
        'X-ClientId': credentials_internal.clientId,
    };

    try {
        const response = await fetch(url, { ...options, headers, cache: 'no-store' });

        // Handle Rate Limiting (429)
        if (response.status === 429) {
            if (retries > 0) {
                const retryAfterHeader = response.headers.get('Retry-After');
                const xRateLimitReset = response.headers.get('x-ratelimit-reset');
                
                let waitTime = 1000; // Default 1s
                
                if (retryAfterHeader) {
                    waitTime = parseInt(retryAfterHeader, 10) * 1000;
                } else if (xRateLimitReset) {
                    waitTime = (parseInt(xRateLimitReset, 10) + 1) * 1000;
                }

                globalDelayUntil = Math.max(globalDelayUntil, Date.now() + waitTime);

                console.warn(`Rate limited (429). Retrying in ${waitTime}ms...`);
                await wait(waitTime);
                return fetchWithAuth(url, options, retries - 1, backoffDelay);
            }
        }

        // Proactive Rate Limiting checks
        if (response.ok) {
            const remaining = response.headers.get('x-ratelimit-remaining');
            const reset = response.headers.get('x-ratelimit-reset');
            if (remaining !== null && reset !== null) {
                if (parseInt(remaining, 10) === 0) {
                    const waitTime = (parseInt(reset, 10) + 1) * 1000;
                    globalDelayUntil = Math.max(globalDelayUntil, Date.now() + waitTime);
                }
            }
        }

        // Handle Server Errors (5xx)
        if (response.status >= 500 && retries > 0) {
             console.warn(`Server error ${response.status}. Retrying...`);
             await wait(1000); 
             return fetchWithAuth(url, options, retries - 1, backoffDelay);
        }

        return response;
    } catch (error: any) {
        const isNetworkError = error.name === 'TypeError' || error.message.includes('NetworkError') || error.message.includes('fetch');
        
        if (isNetworkError && retries > 0) {
            console.warn(`Network error detected: ${error.message}. Pausing for ${backoffDelay}ms and retrying...`);
            await wait(backoffDelay);
            return fetchWithAuth(url, options, retries - 1, backoffDelay * 1.5);
        }
        throw error;
    }
}


function extractIdFromLocation(response: Response): number {
    const location = response.headers.get('Location') || response.headers.get('location');
    if (location) {
        const parts = location.split('/');
        const last = parts[parts.length - 1] || parts[parts.length - 2];
        if (last && !isNaN(parseInt(last, 10))) {
            return parseInt(last, 10);
        }
    }
    return 0;
}

export async function fetchPaginatedData(endpoint: string): Promise<any[]> {
    let allData: any[] = [];
    let offset = 0;
    const limit = 50; 

    while (true) {
        const url = `${API_BASE_URL}${endpoint}${endpoint.includes('?') ? '&' : '?'}limit=${limit}&offset=${offset}`;
        
        const response = await fetchWithAuth(url);
        if (!response.ok) throw new Error(`Failed to fetch ${endpoint}: ${await response.text()}`);
        const text = await response.text();
        const result = safeJsonParse(text);

        let items = [];
        if (Array.isArray(result)) {
            items = result;
            allData = allData.concat(items);
            break; 
        } else if (result && result.data && Array.isArray(result.data)) {
            items = result.data;
            allData = allData.concat(items);
            if (items.length < limit || (result.paging && result.paging.total <= allData.length)) {
                break;
            }
            offset += items.length;
        } else {
            break;
        }
    }
    return allData;
}

function safeJsonParse(text: string): any {
    if (!text || !text.trim()) return null;
    try {
        return JSON.parse(text);
    } catch {
        return null;
    }
}

function parseError(text: string): string {
    try {
        const json = safeJsonParse(text);
        if (!json) {
            if (text.toLowerCase().includes('<!doctype html>')) return 'Server returned HTML response.';
            return text.substring(0, 300);
        }
        const parts: string[] = [];

        if (json.modelState) {
             const ms = Object.entries(json.modelState)
                .map(([key, msgs]) => `${key.replace('model.', '')}: ${(Array.isArray(msgs) ? msgs.join(', ') : msgs)}`)
                .join('; ');
             if(ms) parts.push(ms);
        }

        if (json.errors) {
            if (Array.isArray(json.errors)) {
                const errs = json.errors.map((e: any) => `${e.key || ''}: ${e.message || ''}`).join(', ');
                if(errs) parts.push(errs);
            } else if (typeof json.errors === 'object') {
                const errs = Object.entries(json.errors)
                    .map(([key, msgs]) => `${key}: ${(Array.isArray(msgs) ? msgs.join(', ') : msgs)}`)
                    .join('; ');
                if(errs) parts.push(errs);
            }
        }

        if (json.error && typeof json.error === 'object') {
            const err = json.error;
            if (err.validation_errors && Array.isArray(err.validation_errors)) {
                const msgs = err.validation_errors.map((ve: any) => `${ve.property_name}: ${ve.error_message || ve.error_message_key}`);
                if(msgs.length > 0) parts.push(msgs.join('; '));
            }
            if (err.code) parts.push(`${err.code}${err.message ? ': ' + err.message : ''}`);
            else if (err.message) parts.push(err.message);
        }
        
        if (json.message) parts.push(json.message);
        if (json.error_description) parts.push(json.error_description);
        if (json.title && !json.message) parts.push(json.title);
        
        if (parts.length > 0) {
            return [...new Set(parts)].join(' | ');
        }

        if (Object.keys(json).length > 0) return JSON.stringify(json);
    } catch { }
    
    if (text.toLowerCase().includes('<!doctype html>')) return 'Server returned HTML response.';
    return text.substring(0, 300);
}

// --- Planday Departments API ---

export async function fetchDepartments(): Promise<Department[]> {
    return fetchPaginatedData('/hr/v1.0/departments');
}

export async function createDepartment(name: string, number: string | null = null): Promise<Department> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/departments`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, number: number || null })
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const extractedId = extractIdFromLocation(response);
    const parsed = safeJsonParse(text);
    const result = parsed || { id: extractedId };
    // Planday skills API uses skillId in responses, let's normalise it
    if (result.skillId && !result.id) result.id = result.skillId;
    if (result.data && result.data.skillId && !result.data.id) result.data.id = result.data.skillId;
    return result.data || result;
}

export async function updateDepartment(id: number, name: string, number: string | null = null): Promise<Department> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/departments/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, number: number || null })
    });
    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
    const text = await response.text();
    const parsed = safeJsonParse(text);
    if (parsed) {
        return parsed.data || parsed;
    }
    return { id, name, number: number || null };
}

export async function deleteDepartment(id: number): Promise<void> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/departments/${id}`, {
        method: 'DELETE'
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
}

// --- Planday Employee Groups API ---

export async function fetchEmployeeGroups(): Promise<EmployeeGroup[]> {
    return fetchPaginatedData('/hr/v1.0/employeegroups');
}

export async function fetchEmployeeGroupDefaultRates(): Promise<any[]> {
    const response = await fetchWithAuth(`${API_BASE_URL}/pay/v1.0/payrates/employeeGroups/default`);
    if (!response.ok) {
        if (response.status === 404) return []; // In case it's not available
        throw new Error(`Failed to fetch default rates: ${await response.text()}`);
    }
    const text = await response.text();
    const result = safeJsonParse(text) || [];
    return Array.isArray(result) ? result : (result.data || []);
}

export async function createEmployeeGroup(name: string): Promise<EmployeeGroup> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/employeegroups`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name })
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const extractedId = extractIdFromLocation(response);
    const parsed = safeJsonParse(text);
    const result = parsed || { id: extractedId };
    // Planday skills API uses skillId in responses, let's normalise it
    if (result.skillId && !result.id) result.id = result.skillId;
    if (result.data && result.data.skillId && !result.data.id) result.data.id = result.data.skillId;
    return result.data || result;
}

export async function updateEmployeeGroupDefaultRate(
    id: number,
    payload: {
        defaultWageType?: 'HourlyRate' | 'ShiftRate' | '';
        hourlyRate?: number;
        shiftRate?: number;
        salaryCode?: string;
        salaryCode2?: string;
    }
): Promise<void> {    const body: any = {
        employeeGroupId: id,
        defaultWageType: payload.defaultWageType || 'HourlyRate',
        hourlyRate: payload.hourlyRate || 0,
        shiftRate: payload.shiftRate || 0,
        salaryCode: payload.salaryCode || "",
        salaryCode2: payload.salaryCode2 || ""
    };

    const response = await fetchWithAuth(`${API_BASE_URL}/pay/v1.0/payrates/employeeGroups/${id}/default`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to set default rate for Employee Group: ${response.status} ${parseError(errText)}`);
    }
}

export async function deleteEmployeeGroup(id: number): Promise<void> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/employeegroups/${id}`, {
        method: 'DELETE'
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
}

// --- Planday Skills API ---

export async function fetchSkills(): Promise<Skill[]> {
    try {
        const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/skills`);
        if (response.ok) {
            const text = await response.text();
            const result = safeJsonParse(text);
            const raw = Array.isArray(result) ? result : (result?.data && Array.isArray(result.data) ? result.data : null);
            if (raw) {
                return raw
                    .filter((item: any) => !item.isDeleted && item.status !== 'Deleted' && item.isActive !== false)
                    .map((item: any) => ({
                        ...item,
                        id: item.id || item.skillId || item.SkillId,
                        name: item.name || item.title || ''
                    }));
            }
        }
    } catch (e) {
        console.warn("Direct fetch of /hr/v1.0/skills failed, trying paginated fallback:", e);
    }

    const data = await fetchPaginatedData('/hr/v1.0/skills');
    return data
        .filter((item: any) => !item.isDeleted && item.status !== 'Deleted' && item.isActive !== false)
        .map((item: any) => ({
            ...item,
            id: item.id || item.skillId || item.SkillId,
            name: item.name || item.title || ''
        }));
}

export async function createSkill(name: string, description: string = ''): Promise<Skill> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/skills`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, description, isTimeLimited: false })
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const extractedId = extractIdFromLocation(response);
    const parsed = safeJsonParse(text);
    const result = parsed || { id: extractedId };
    // Planday skills API uses skillId in responses, let's normalise it
    if (result.skillId && !result.id) result.id = result.skillId;
    if (result.data && result.data.skillId && !result.data.id) result.data.id = result.data.skillId;
    return result.data || result;
}

// --- Planday Revenue Units API ---

export async function fetchRevenueUnits(): Promise<RevenueUnit[]> {
    return fetchPaginatedData('/revenue/v1.0/revenueunits');
}

// --- Planday Positions API ---

export async function fetchPositionById(positionId: number): Promise<Position> {
    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/positions/${positionId}`);
    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
    const text = await response.text();
    const parsed = safeJsonParse(text);
    const data = parsed?.data || parsed;
    return {
        ...data,
        id: data.id !== undefined ? data.id : (data.positionId !== undefined ? data.positionId : positionId),
        name: data.name || data.title || '',
        startDate: data.startDate || data.StartDate || undefined,
        endDate: data.endDate || data.EndDate || undefined,
        validFrom: data.validFrom || data.startDate || data.StartDate || undefined,
        validTo: data.validTo || data.endDate || data.EndDate || undefined,
        isActive: data.isActive !== undefined ? data.isActive : (data.IsActive !== undefined ? data.IsActive : undefined),
        departmentId: data.departmentId !== undefined ? data.departmentId : (data.DepartmentId !== undefined ? data.DepartmentId : undefined),
        sectionId: data.sectionId !== undefined ? data.sectionId : (data.SectionId !== undefined ? data.SectionId : undefined),
        employeeGroupId: data.employeeGroupId !== undefined ? data.employeeGroupId : (data.EmployeeGroupId !== undefined ? data.EmployeeGroupId : undefined),
        affectRevenue: data.affectRevenue !== undefined ? data.affectRevenue : (data.AffectRevenue !== undefined ? data.AffectRevenue : undefined),
        revenueUnitId: data.revenueUnitId !== undefined ? data.revenueUnitId : (data.RevenueUnitId !== undefined ? data.RevenueUnitId : (data.revenueUnit?.id || undefined)),
        skillIds: Array.isArray(data.skillIds) ? data.skillIds : (Array.isArray(data.SkillIds) ? data.SkillIds : [])
    };
}

export async function fetchPositions(skillsOrHasSkills?: boolean | Skill[] | number): Promise<Position[]> {
    const isItemActive = (item: any): boolean => {
        if (!item) return false;
        if (item.isActive === false || item.IsActive === false || item.isActive === 'false' || item.IsActive === 'false') return false;
        if (item.isDeleted || item.status === 'Deleted') return false;
        return true;
    };

    const rawPositions = await fetchPaginatedData('/scheduling/v1.0/positions');
    const basePositions: Position[] = rawPositions
        .filter(isItemActive)
        .map((p: any) => ({
            ...p,
            id: p.id !== undefined ? p.id : (p.positionId !== undefined ? p.positionId : 0),
            name: p.name || p.title || '',
            departmentId: p.departmentId !== undefined ? p.departmentId : (p.DepartmentId !== undefined ? p.DepartmentId : undefined),
            sectionId: p.sectionId !== undefined ? p.sectionId : (p.SectionId !== undefined ? p.SectionId : undefined),
            employeeGroupId: p.employeeGroupId !== undefined ? p.employeeGroupId : (p.EmployeeGroupId !== undefined ? p.EmployeeGroupId : undefined),
            startDate: p.startDate || p.StartDate || undefined,
            endDate: p.endDate || p.EndDate || undefined,
            validFrom: p.validFrom || p.startDate || p.StartDate || undefined,
            validTo: p.validTo || p.endDate || p.EndDate || undefined,
            isActive: p.isActive !== undefined ? p.isActive : (p.IsActive !== undefined ? p.IsActive : undefined),
            affectRevenue: p.affectRevenue !== undefined ? p.affectRevenue : (p.AffectRevenue !== undefined ? p.AffectRevenue : undefined),
            revenueUnitId: p.revenueUnitId !== undefined ? p.revenueUnitId : (p.RevenueUnitId !== undefined ? p.RevenueUnitId : (p.revenueUnit?.id || undefined)),
            skillIds: Array.isArray(p.skillIds) ? p.skillIds : (Array.isArray(p.SkillIds) ? p.SkillIds : [])
        }));

    // If there are any skills created in the portal (minimum 1), the app will have to
    // use GET /scheduling/v1.0/positions/{positionId} when looking up positions in order
    // to check if the position has any skills assigned ("skillIds").
    let hasSkills = false;
    if (typeof skillsOrHasSkills === 'boolean') {
        hasSkills = skillsOrHasSkills;
    } else if (Array.isArray(skillsOrHasSkills)) {
        hasSkills = skillsOrHasSkills.length > 0;
    } else if (typeof skillsOrHasSkills === 'number') {
        hasSkills = skillsOrHasSkills > 0;
    } else {
        try {
            const portalSkills = await fetchSkills();
            hasSkills = Boolean(portalSkills && portalSkills.length > 0);
        } catch {
            hasSkills = false;
        }
    }

    if (hasSkills && basePositions.length > 0) {
        const batchSize = 6;
        const detailedPositions: Position[] = [];
        for (let i = 0; i < basePositions.length; i += batchSize) {
            const chunk = basePositions.slice(i, i + batchSize);
            const chunkResults = await Promise.all(
                chunk.map(async (pos) => {
                    if (!pos.id) return pos;
                    try {
                        const detail = await fetchPositionById(pos.id);
                        if (!isItemActive(detail)) {
                            return null;
                        }
                        return {
                            ...pos,
                            ...detail,
                            affectRevenue: detail.affectRevenue !== undefined ? detail.affectRevenue : pos.affectRevenue,
                            revenueUnitId: detail.revenueUnitId !== undefined ? detail.revenueUnitId : pos.revenueUnitId,
                            skillIds: Array.isArray(detail.skillIds) ? detail.skillIds : []
                        };
                    } catch (err) {
                        console.warn(`Could not lookup position ${pos.id} for skills:`, err);
                        return pos;
                    }
                })
            );
            detailedPositions.push(...chunkResults.filter((p): p is Position => p !== null));
        }
        return detailedPositions.filter(isItemActive);
    }

    return basePositions.filter(isItemActive);
}

export async function createPosition(payload: any): Promise<Position> {
    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/positions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const extractedId = extractIdFromLocation(response);
    const parsed = safeJsonParse(text);
    const result = parsed || { id: extractedId };
    // Planday skills API uses skillId in responses, let's normalise it
    if (result.skillId && !result.id) result.id = result.skillId;
    if (result.data && result.data.skillId && !result.data.id) result.data.id = result.data.skillId;
    return result.data || result;
}

export async function fetchPortalInfo(): Promise<any> {
    const response = await fetchWithAuth(`${API_BASE_URL}/portal/v1.0/info`);
    if (!response.ok) return null;
    const text = await response.text();
    const parsed = safeJsonParse(text);
    return parsed ? (parsed.data || parsed) : null;
}

// --- Planday Sections API ---

export async function fetchSections(departmentId?: number): Promise<Section[]> {
    const isSectionActive = (item: any): boolean => {
        if (!item) return false;
        if (item.isActive === false || item.IsActive === false || item.isActive === 'false' || item.IsActive === 'false') return false;
        if (item.isDeleted || item.status === 'Deleted') return false;
        return true;
    };

    const endpoint = departmentId !== undefined 
        ? `/scheduling/v1.0/sections?departmentId=${departmentId}` 
        : `/scheduling/v1.0/sections`;
    const data = await fetchPaginatedData(endpoint);
    return data
        .filter(isSectionActive)
        .map((item: any) => ({
            id: item.id !== undefined ? item.id : (item.sectionId !== undefined ? item.sectionId : 0),
            name: item.name || '',
            departmentId: item.departmentId !== undefined ? item.departmentId : (departmentId || 0),
            color: item.color || null,
            isActive: item.isActive !== undefined ? item.isActive : (item.IsActive !== undefined ? item.IsActive : undefined)
        }));
}

export async function fetchAllSections(departments?: Department[]): Promise<Section[]> {
    const isSectionActive = (s: Section) => s.isActive !== false && (s as any).IsActive !== false && (s as any).isActive !== 'false' && (s as any).IsActive !== 'false';

    // If departments provided, try fetching per department to be guaranteed complete data
    if (departments && departments.length > 0) {
        try {
            // First attempt: try general endpoint if it returns everything
            const general = await fetchSections().catch(() => null);
            if (general && general.length > 0) {
                return general.filter(isSectionActive);
            }
        } catch { }

        // Fallback or full coverage: fetch for each department in parallel
        const results = await Promise.all(
            departments.map(async (d) => {
                try {
                    return await fetchSections(d.id);
                } catch (err) {
                    console.warn(`Could not fetch sections for department ${d.id}:`, err);
                    return [];
                }
            })
        );
        const combined = results.flat();
        const seen = new Set<number>();
        return combined
            .filter(isSectionActive)
            .filter(s => {
                if (s.id && seen.has(s.id)) return false;
                if (s.id) seen.add(s.id);
                return true;
            });
    }

    // Otherwise try general endpoint
    try {
        const sections = await fetchSections();
        return sections.filter(isSectionActive);
    } catch {
        return [];
    }
}

export async function createSection(departmentId: number, name: string, color?: string | null): Promise<Section> {
    const body: any = { 
        name: name.trim(),
        color: color ? color.trim() : null 
    };

    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/sections/${departmentId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const extractedId = extractIdFromLocation(response);
    const parsed = safeJsonParse(text);
    const result = parsed || { id: extractedId };
    const id = result.id || result.sectionId || (result.data && (result.data.id || result.data.sectionId)) || extractedId;
    return {
        id,
        name: name.trim(),
        departmentId,
        color: color || null
    };
}

export async function updateSection(departmentId: number, sectionId: number, name: string, color?: string | null): Promise<Section> {
    const body: any = { 
        name: name.trim(),
        color: color !== undefined ? (color ? color.trim() : null) : null
    };

    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/sections/${departmentId}/${sectionId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const parsed = safeJsonParse(text);
    if (parsed && (parsed.data || parsed.id || parsed.sectionId)) {
        const res = parsed.data || parsed;
        return {
            id: res.id || res.sectionId || sectionId,
            name: res.name || name,
            departmentId: res.departmentId || departmentId,
            color: res.color !== undefined ? res.color : (color || null)
        };
    }
    return { id: sectionId, name: name.trim(), departmentId, color: color || null };
}

export async function deleteSection(departmentId: number, sectionId: number): Promise<void> {
    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/sections/${departmentId}/${sectionId}`, {
        method: 'DELETE'
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
}

export async function updatePosition(id: number, payload: any): Promise<Position> {
    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/positions/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const parsed = safeJsonParse(text);
    return parsed?.data || parsed || { id, name: payload.name || '' };
}

export async function deletePosition(id: number): Promise<void> {
    const response = await fetchWithAuth(`${API_BASE_URL}/scheduling/v1.0/positions/${id}`, {
        method: 'DELETE'
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
}

export async function updateSkill(id: number, name: string, description: string = ''): Promise<Skill> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/skills/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: name.trim(), description: description.trim(), isTimeLimited: false })
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }

    const text = await response.text();
    const parsed = safeJsonParse(text);
    return parsed?.data || parsed || { id, name: name.trim(), description };
}

export async function deleteSkill(id: number): Promise<void> {
    const response = await fetchWithAuth(`${API_BASE_URL}/hr/v1.0/skills?skillId=${id}`, {
        method: 'DELETE'
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`${response.status}: ${parseError(errText)}`);
    }
}

