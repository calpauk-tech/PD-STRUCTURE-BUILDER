const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetCurrent = `        // Build Header
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
        }`;

const replaceCurrent = `        for (let i = 0; i < Math.max(1, maxLen); i++) {
            const row: any = {};
            if (includeDepartments) {
                row["Departments"] = currentStructure.departments[i]?.name || "";
                row["Department Number (optional)"] = currentStructure.departments[i]?.number || "";
            }
            if (includeEmployeeGroups) {
                row["Employee Groups"] = currentStructure.employeeGroups[i]?.name || "";
                if (includeWageColumns) {
                    row["Default Wage Type (HourlyRate/ShiftRate)"] = "";
                    row["Rate"] = "";
                    row["Salary Code"] = "";
                    row["Salary Code 2 (Rare)"] = "";
                }
            }
            data.push(row);
        }`;

if (code.includes(targetCurrent)) {
    code = code.replace(targetCurrent, replaceCurrent);
    fs.writeFileSync('App.tsx', code);
    console.log("Success current");
} else {
    console.log("Failed to find targetCurrent");
}

code = fs.readFileSync('App.tsx', 'utf8');
const targetBlank = `        const headers: string[] = [];
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

const replaceBlank = `        const headers: string[] = [];
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

if (code.includes(targetBlank)) {
    code = code.replace(targetBlank, replaceBlank);
    fs.writeFileSync('App.tsx', code);
    console.log("Success blank");
} else {
    console.log("Failed to find targetBlank");
}
