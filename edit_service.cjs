const fs = require('fs');
let code = fs.readFileSync('services/plandayService.ts', 'utf8');

const fetchGroupsTarget = `export async function fetchEmployeeGroups(): Promise<EmployeeGroup[]> {
    return fetchPaginatedData('/hr/v1.0/employeegroups');
}`;

const fetchGroupsReplace = `export async function fetchEmployeeGroups(): Promise<EmployeeGroup[]> {
    return fetchPaginatedData('/hr/v1.0/employeegroups');
}

export async function fetchEmployeeGroupDefaultRates(): Promise<any[]> {
    const response = await fetchWithAuth(\`\${API_BASE_URL}/pay/v1.0/payrates/employeeGroups/default\`);
    if (!response.ok) {
        if (response.status === 404) return []; // In case it's not available
        throw new Error(\`Failed to fetch default rates: \${await response.text()}\`);
    }
    const result = await response.json();
    return Array.isArray(result) ? result : (result.data || []);
}`;

code = code.replace(fetchGroupsTarget, fetchGroupsReplace);
fs.writeFileSync('services/plandayService.ts', code);
console.log("Updated service");
