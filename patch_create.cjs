const fs = require('fs');
let content = fs.readFileSync('services/plandayService.ts', 'utf8');

const extractorFunc = `
function extractIdFromLocation(response: Response): number {
    const location = response.headers.get('Location') || response.headers.get('location');
    if (location) {
        const parts = location.split('/');
        const last = parts[parts.length - 1];
        if (last && !isNaN(parseInt(last, 10))) {
            return parseInt(last, 10);
        }
    }
    return 0;
}
`;

if (!content.includes('extractIdFromLocation')) {
    content = content.replace('export async function fetchPaginatedData', extractorFunc + '\nexport async function fetchPaginatedData');
}

function fixCreateFunction(funcName) {
    const regex = new RegExp(`export async function ${funcName}\\([\\s\\S]*?\\{\\s*const response = await fetchWithAuth\\([\\s\\S]*?\\}\\);\\s*if \\(!response\\.ok\\) \\{[\\s\\S]*?\\}\\s*const text = await response\\.text\\(\\);\\s*const result = text \\? JSON\\.parse\\(text\\) : \\{ id: 0 \\};\\s*return result\\.data \\|\\| result;\\s*\\}`, 'm');
    
    content = content.replace(regex, (match) => {
        return match.replace(`const result = text ? JSON.parse(text) : { id: 0 };`, 
                             `const extractedId = extractIdFromLocation(response);\n    const result = text ? JSON.parse(text) : { id: extractedId };\n    // Planday skills API uses skillId in responses, let's normalise it\n    if (result.skillId && !result.id) result.id = result.skillId;\n    if (result.data && result.data.skillId && !result.data.id) result.data.id = result.data.skillId;`);
    });
}

fixCreateFunction('createSkill');
fixCreateFunction('createDepartment');
fixCreateFunction('createEmployeeGroup');
fixCreateFunction('createPosition');

fs.writeFileSync('services/plandayService.ts', content);
