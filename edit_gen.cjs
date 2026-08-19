const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetGen = `        // Departments from file
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

const replaceGen = `        // Departments from file
        newDeptsMap.forEach((deptNum, name) => {
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            const existingDept = currentStructure.departments.find(d => d.name.trim().toLowerCase() === name.toLowerCase());
            
            if (appMode === 'update') {
                if (existingDept) {
                    items.push({
                        id: existingDept.id,
                        type: 'department',
                        name,
                        number: deptNum,
                        action: 'update',
                        status: 'pending',
                        updatePayload: { name, number: deptNum },
                        resultMessage: 'Update scheduled'
                    });
                } else {
                    items.push({
                        type: 'department',
                        name,
                        action: 'skip_duplicate',
                        status: 'error',
                        resultMessage: 'Not found in Planday'
                    });
                }
            } else {
                if (isDuplicate && overwriteMode === 'keep') {
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
            }
        });

        // Employee Groups from file
        newGroupsMap.forEach((groupData, name) => {
            const isDuplicate = existingGroupNames.has(name.toLowerCase());
            const existingGroup = currentStructure.employeeGroups.find(g => g.name.trim().toLowerCase() === name.toLowerCase());
            
            let updatePayload: any = undefined;
            if (Object.keys(groupData).length > 0 || appMode === 'update') {
                updatePayload = {};
                if (groupData.wageType) updatePayload.defaultWageType = groupData.wageType;
                if (groupData.rate && groupData.wageType === 'HourlyRate') updatePayload.hourlyRate = groupData.rate;
                if (groupData.rate && groupData.wageType === 'ShiftRate') updatePayload.shiftRate = groupData.rate;
                if (groupData.salaryCode) updatePayload.salaryCode = groupData.salaryCode;
                if (groupData.salaryCode2) updatePayload.salaryCode2 = groupData.salaryCode2;
            }

            if (appMode === 'update') {
                if (existingGroup) {
                    items.push({
                        id: existingGroup.id,
                        type: 'employee_group',
                        name,
                        action: 'update',
                        status: 'pending',
                        updatePayload,
                        resultMessage: 'Update scheduled'
                    });
                } else {
                    items.push({
                        type: 'employee_group',
                        name,
                        action: 'skip_duplicate',
                        status: 'error',
                        resultMessage: 'Not found in Planday'
                    });
                }
            } else {
                if (isDuplicate && overwriteMode === 'keep') {
                    if (updatePayload && existingGroup && Object.keys(updatePayload).length > 0) {
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
            }
        });`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetGen);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetGen.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceGen);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing generate review logic");
} else {
    console.log("Failed replacing generate review logic");
}
