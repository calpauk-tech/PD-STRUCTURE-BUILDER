const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetReviewList = `    // Step 3 -> 4: Process File Rows and Generate Review Items (Import)
    const generateReviewList = () => {
        const existingDeptNames = new Set(currentStructure.departments.map(d => d.name.trim().toLowerCase()));
        const existingGroupNames = new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));

        const newDeptsSet = new Set<string>();
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
        });

        const items: StructureReviewItem[] = [];

        // If Overwrite mode: mark all existing Planday structure for deletion
        if (overwriteMode === 'overwrite') {
            currentStructure.departments.forEach(d => {
                items.push({
                    id: d.id,
                    type: 'department',
                    name: d.name,
                    action: appMode === 'update' ? 'update' : 'delete',
                    status: 'pending'
                });
            });
            currentStructure.employeeGroups.forEach(g => {
                items.push({
                    id: g.id,
                    type: 'employee_group',
                    name: g.name,
                    action: appMode === 'update' ? 'update' : 'delete',
                    status: 'pending'
                });
            });
        }

        // Departments from file
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
        });

        setReviewItems(items);
        setStep(appMode === 'update' ? 2 : 3); // Config or Review
    };`;

const replaceReviewList = `    // Step 3 -> 4: Process File Rows and Generate Review Items (Import)
    const generateReviewList = () => {
        const existingDeptNames = new Set(currentStructure.departments.map(d => d.name.trim().toLowerCase()));
        const existingGroupNames = new Set(currentStructure.employeeGroups.map(g => g.name.trim().toLowerCase()));

        const newDeptsMap = new Map<string, string | null>();
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
        });

        const items: StructureReviewItem[] = [];

        // If Overwrite mode: mark all existing Planday structure for deletion
        if (overwriteMode === 'overwrite') {
            currentStructure.departments.forEach(d => {
                items.push({
                    id: d.id,
                    type: 'department',
                    name: d.name,
                    number: d.number,
                    action: appMode === 'update' ? 'update' : 'delete',
                    status: 'pending'
                });
            });
            currentStructure.employeeGroups.forEach(g => {
                items.push({
                    id: g.id,
                    type: 'employee_group',
                    name: g.name,
                    action: appMode === 'update' ? 'update' : 'delete',
                    status: 'pending'
                });
            });
        }

        // Departments from file
        newDeptsMap.forEach((deptNum, name) => {
            const isDuplicate = existingDeptNames.has(name.toLowerCase());
            
            if (isDuplicate && overwriteMode === 'keep') {
                const existingDept = currentStructure.departments.find(d => d.name.trim().toLowerCase() === name.toLowerCase());
                if (deptNum && existingDept && existingDept.number !== deptNum) {
                    // Update department number if it's different and provided in file
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
        });

        setReviewItems(items);
        setStep(appMode === 'update' ? 2 : 3); // Config or Review
    };`;

if (code.includes(targetReviewList)) {
    code = code.replace(targetReviewList, replaceReviewList);
    fs.writeFileSync('App.tsx', code);
    console.log("Replaced Review List");
} else {
    console.log("Failed to find Review List");
}
