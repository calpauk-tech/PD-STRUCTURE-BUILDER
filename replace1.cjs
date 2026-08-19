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

// remove all whitespace to match
function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(target1);

if (codeNoSpace.includes(targetNoSpace)) {
    // replace using a regex that allows any whitespace
    const regexStr = target1.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replace1);
    fs.writeFileSync('App.tsx', code);
    console.log("Success target1");
} else {
    console.log("Failed target1");
}

