import { PlandayApiCredentials, Department, EmployeeGroup } from '../types';

const AUTH_URL = 'https://id.planday.com/connect/token';
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
        ...options.headers,
        'Authorization': `Bearer ${token}`,
        'X-ClientId': credentials_internal.clientId,
    };

    try {
        const response = await fetch(url, { ...options, headers });

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

export async function fetchPaginatedData(endpoint: string): Promise<any[]> {
    let allData: any[] = [];
    let offset = 0;
    const limit = 50; 

    while (true) {
        const url = `${API_BASE_URL}${endpoint}${endpoint.includes('?') ? '&' : '?'}limit=${limit}&offset=${offset}`;
        
        const response = await fetchWithAuth(url);
        if (!response.ok) throw new Error(`Failed to fetch ${endpoint}: ${await response.text()}`);
        const result = await response.json();

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

function parseError(text: string): string {
    try {
        const json = JSON.parse(text);
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

    const result = await response.json();
    return result.data || result;
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

    const result = await response.json();
    return result.data || result;
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

export async function fetchPortalInfo(): Promise<any> {
    const response = await fetchWithAuth(`${API_BASE_URL}/portal/v1.0/info`);
    if (!response.ok) return null;
    const json = await response.json();
    return json.data;
}
