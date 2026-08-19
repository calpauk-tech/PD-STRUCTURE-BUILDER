const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetBlank = `    // Step 2: Download Blank Excel Template
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
        ];`;

const replaceBlank = `    // Step 2: Download Blank Excel Template
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
        ];`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetBlank);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetBlank.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceBlank);
    fs.writeFileSync('App.tsx', code);
    console.log("Success blank template replacement");
} else {
    console.log("Failed blank template replacement");
}
