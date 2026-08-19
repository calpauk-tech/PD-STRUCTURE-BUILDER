const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetMethod = `    // Helper: auto-detect column matching
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
    };`;

const replaceMethod = `    // Helper: auto-detect column matching
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
    };`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetMethod);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetMethod.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceMethod);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing auto detect");
} else {
    console.log("Failed replacing auto detect");
}
