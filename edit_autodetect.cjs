const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetAuto = `    const autoDetectColumns = (headers: string[]) => {
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

const replaceAuto = `    const autoDetectColumns = (headers: string[]) => {
        let detectedDept = '';
        let detectedDeptNum = '';
        let detectedGroup = '';
        let detectedWageType = '';
        let detectedRate = '';
        let detectedSalaryCode = '';
        let detectedSalaryCode2 = '';

        for (const h of headers) {
            const clean = h.trim().toLowerCase();
            
            if (!detectedDeptNum && (clean.includes('department number') || clean.includes('dept number') || clean === 'number')) {
                detectedDeptNum = h;
            } else if (!detectedDept && (clean.includes('department') || clean.includes('afdeling') || clean.includes('avdelning') || clean.includes('abteilung') || clean === 'dept')) {
                detectedDept = h;
            }
            
            if (!detectedGroup && (clean.includes('group') || clean.includes('rolle') || clean.includes('stilling') || clean.includes('position') || clean.includes('employee group') || clean === 'role' || clean === 'groups')) {
                detectedGroup = h;
            }

            if (!detectedWageType && (clean.includes('wage type') || clean.includes('default wage type'))) {
                detectedWageType = h;
            }
            if (!detectedRate && (clean === 'rate' || clean.includes('hourly') || clean.includes('shift'))) {
                detectedRate = h;
            }
            if (!detectedSalaryCode2 && (clean.includes('salary code 2') || clean.includes('rare'))) {
                detectedSalaryCode2 = h;
            } else if (!detectedSalaryCode && (clean.includes('salary code'))) {
                detectedSalaryCode = h;
            }
        }
        
        // Default fallbacks if standard 2-column template
        if (!detectedDept && headers.length > 0 && includeDepartments) {
            detectedDept = headers[0];
        }
        if (!detectedGroup && headers.length > 1 && includeEmployeeGroups && headers[1] !== detectedDeptNum) {
            detectedGroup = headers[1];
        } else if (!detectedGroup && headers.length === 1 && !includeDepartments && includeEmployeeGroups) {
            detectedGroup = headers[0];
        }
        
        setDeptColumn(detectedDept);
        setDeptNumberColumn(detectedDeptNum);
        setGroupColumn(detectedGroup);
        setWageTypeColumn(detectedWageType);
        setRateColumn(detectedRate);
        setSalaryCodeColumn(detectedSalaryCode);
        setSalaryCode2Column(detectedSalaryCode2);
    };`;

code = code.replace(targetAuto, replaceAuto);
fs.writeFileSync('App.tsx', code);
console.log("Replaced auto");
