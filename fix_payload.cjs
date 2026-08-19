const fs = require('fs');
let code = fs.readFileSync('services/plandayService.ts', 'utf8');

const targetMethod = `    const body: any = {};
    if (payload.defaultWageType) {
        body.defaultWageType = payload.defaultWageType;
        if (payload.defaultWageType === 'HourlyRate' && payload.hourlyRate !== undefined) {
            body.hourlyRate = payload.hourlyRate;
        } else if (payload.defaultWageType === 'ShiftRate' && payload.shiftRate !== undefined) {
            body.shiftRate = payload.shiftRate;
        }
    }
    
    if (payload.salaryCode) body.salaryCode = payload.salaryCode;
    if (payload.salaryCode2) body.salaryCode2 = payload.salaryCode2;

    if (Object.keys(body).length === 0) return; // Nothing to update

    const response = await fetchWithAuth(\`\${API_BASE_URL}/pay/v1.0/payrates/employeeGroups/\${id}/default\`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });`;

const replacementMethod = `    const body: any = {
        employeeGroupId: id,
        defaultWageType: payload.defaultWageType || 'HourlyRate',
        hourlyRate: payload.hourlyRate || 0,
        shiftRate: payload.shiftRate || 0,
        salaryCode: payload.salaryCode || "",
        salaryCode2: payload.salaryCode2 || ""
    };

    const response = await fetchWithAuth(\`\${API_BASE_URL}/pay/v1.0/payrates/employeeGroups/\${id}/default\`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetMethod);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetMethod.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replacementMethod);
    fs.writeFileSync('services/plandayService.ts', code);
    console.log("Success replacing payload structure");
} else {
    console.log("Failed replacing payload structure");
}
