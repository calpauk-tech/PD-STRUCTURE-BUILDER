const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetDownload = `    // Step 2: Download Current Structure Excel
    const handleDownloadCurrentStructure = () => {
        const wb = XLSX.utils.book_new();
        
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

const replaceDownload = `    // Step 2: Download Current Structure Excel
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
        }`;

// remove all whitespace to match
function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetDownload);

if (codeNoSpace.includes(targetNoSpace)) {
    // replace using a regex that allows any whitespace
    const regexStr = targetDownload.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceDownload);
    fs.writeFileSync('App.tsx', code);
    console.log("Success download");
} else {
    console.log("Failed download");
}
