const fs = require('fs');
let code = fs.readFileSync('services/plandayService.ts', 'utf8');

const targetMethod = `    // Only send the fields that are required by the backend API based on the logic
    // Fetch current rates to perform a safe merge since PUT expects the full object
    let currentData: any = {};
    try {
        const getRes = await fetchWithAuth(\`\${API_BASE_URL}/pay/v1.0/payrates/employeeGroups/\${id}/default\`);
        if (getRes.ok) {
            const data = await getRes.json();
            currentData = data.data || data;
        }
    } catch (e) {
        // Ignore 404s if it doesn't exist yet
    }

    const body: any = {
        employeeGroupId: id,
        defaultWageType: payload.defaultWageType || currentData.defaultWageType || 'HourlyRate',
        hourlyRate: payload.hourlyRate !== undefined ? payload.hourlyRate : (currentData.hourlyRate || 0),
        shiftRate: payload.shiftRate !== undefined ? payload.shiftRate : (currentData.shiftRate || 0),
        salaryCode: payload.salaryCode !== undefined ? payload.salaryCode : (currentData.salaryCode || ""),
        salaryCode2: payload.salaryCode2 !== undefined ? payload.salaryCode2 : (currentData.salaryCode2 || "")
    };`;

const replaceMethod = `    const body: any = {
        employeeGroupId: id,
        defaultWageType: payload.defaultWageType || 'HourlyRate',
        hourlyRate: payload.hourlyRate || 0,
        shiftRate: payload.shiftRate || 0,
        salaryCode: payload.salaryCode || "",
        salaryCode2: payload.salaryCode2 || ""
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
    fs.writeFileSync('services/plandayService.ts', code);
    console.log("Success replacing payload structure for overwrite");
} else {
    console.log("Failed replacing payload structure");
}
