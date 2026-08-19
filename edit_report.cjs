const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetReport = `    // Export execution report as Excel
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
        XLSX.writeFile(wb, \`Planday_Structure_Import_Report_\${new Date().toISOString().slice(0, 10)}.xlsx\`);
    };`;

const replaceReport = `    // Export execution report as Excel
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
        XLSX.writeFile(wb, \`Planday_Structure_Report_\${new Date().toISOString().slice(0, 10)}.xlsx\`);
    };`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetReport);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetReport.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceReport);
    fs.writeFileSync('App.tsx', code);
    console.log("Success report replacement");
} else {
    console.log("Failed report replacement");
}
