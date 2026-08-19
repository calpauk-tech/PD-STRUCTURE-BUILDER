const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const target1 = `        const newDeptsSet = new Set<string>();
        const newGroupsSet = new Set<string>();
        
        parsedRows.forEach(row => {
            if (deptColumn && row[deptColumn]) {
                const val = String(row[deptColumn]).trim();
                if (val) newDeptsSet.add(val);
            }
            if (groupColumn && row[groupColumn]) {
                const val = String(row[groupColumn]).trim();
                if (val) newGroupsSet.add(val);
            }
        });`;

const replace1 = `        const newDeptsMap = new Map<string, string | null>();
        const newGroupsMap = new Map<string, any>();
        
        parsedRows.forEach(row => {
            if (deptColumn && row[deptColumn]) {
                const val = String(row[deptColumn]).trim();
                let deptNum = null;
                if (deptNumberColumn && row[deptNumberColumn]) {
                    deptNum = String(row[deptNumberColumn]).trim();
                }
                if (val) {
                    if (!newDeptsMap.has(val) || deptNum) {
                        newDeptsMap.set(val, deptNum);
                    }
                }
            }
            if (groupColumn && row[groupColumn]) {
                const val = String(row[groupColumn]).trim();
                if (val) {
                    const groupData: any = newGroupsMap.get(val) || {};
                    if (wageTypeColumn && row[wageTypeColumn]) groupData.wageType = String(row[wageTypeColumn]).trim();
                    if (rateColumn && row[rateColumn]) groupData.rate = String(row[rateColumn]).trim();
                    if (salaryCodeColumn && row[salaryCodeColumn]) groupData.salaryCode = String(row[salaryCodeColumn]).trim();
                    if (salaryCode2Column && row[salaryCode2Column]) groupData.salaryCode2 = String(row[salaryCode2Column]).trim();
                    newGroupsMap.set(val, groupData);
                }
            }
        });`;

if(code.includes(target1)){
    code = code.replace(target1, replace1);
    fs.writeFileSync('App.tsx', code);
    console.log("Success target1");
}

const target2 = `        // Departments from file
        newDeptsSet.forEach(name => {
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            if (isDuplicate && overwriteMode === 'keep') {
                items.push({
                    type: 'department',
                    name,
                    action: 'skip_duplicate',
                    status: 'skipped',
                    resultMessage: 'Already exists in Planday'
                });
            } else {
                items.push({
                    type: 'department',
                    name,
                    action: 'create',
                    status: 'pending'
                });
            }
        });

        // Employee Groups from file
        newGroupsSet.forEach(name => {
            const isDuplicate = existingGroupNames.has(name.toLowerCase());
            if (isDuplicate && overwriteMode === 'keep') {
                items.push({
                    type: 'employee_group',
                    name,
                    action: 'skip_duplicate',
                    status: 'skipped',
                    resultMessage: 'Already exists in Planday'
                });
            } else {
                items.push({
                    type: 'employee_group',
                    name,
                    action: 'create',
                    status: 'pending'
                });
            }
        });`;

const replace2 = `        // Departments from file
        newDeptsMap.forEach((deptNum, name) => {
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            
            if (isDuplicate && overwriteMode === 'keep') {
                const existingDept = currentStructure.departments.find(d => d.name.trim().toLowerCase() === name.toLowerCase());
                if (deptNum && existingDept && existingDept.number !== deptNum) {
                    items.push({
                        id: existingDept.id,
                        type: 'department',
                        name,
                        number: deptNum,
                        action: 'update',
                        status: 'pending',
                        resultMessage: 'Updating number from file'
                    });
                } else {
                    items.push({
                        type: 'department',
                        name,
                        number: deptNum,
                        action: 'skip_duplicate',
                        status: 'skipped',
                        resultMessage: 'Already exists in Planday'
                    });
                }
            } else {
                items.push({
                    type: 'department',
                    name,
                    number: deptNum,
                    action: 'create',
                    status: 'pending'
                });
            }
        });

        // Employee Groups from file
        newGroupsMap.forEach((groupData, name) => {
            const isDuplicate = existingGroupNames.has(name.toLowerCase());
            
            let updatePayload: any = undefined;
            if (Object.keys(groupData).length > 0) {
                updatePayload = {};
                if (groupData.wageType) updatePayload.defaultWageType = groupData.wageType;
                if (groupData.rate && groupData.wageType === 'HourlyRate') updatePayload.hourlyRate = groupData.rate;
                if (groupData.rate && groupData.wageType === 'ShiftRate') updatePayload.shiftRate = groupData.rate;
                if (groupData.salaryCode) updatePayload.salaryCode = groupData.salaryCode;
                if (groupData.salaryCode2) updatePayload.salaryCode2 = groupData.salaryCode2;
            }

            if (isDuplicate && overwriteMode === 'keep') {
                const existingGroup = currentStructure.employeeGroups.find(g => g.name.trim().toLowerCase() === name.toLowerCase());
                if (updatePayload && existingGroup) {
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
        });`;

if(code.includes(target2)){
    code = code.replace(target2, replace2);
    fs.writeFileSync('App.tsx', code);
    console.log("Success target2");
}

