const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetImport = `    fetchPortalInfo,
    updateEmployeeGroupDefaultRate,
    updateDepartment
} from './services/plandayService';`;

const replaceImport = `    fetchPortalInfo,
    updateEmployeeGroupDefaultRate,
    updateDepartment,
    fetchEmployeeGroupDefaultRates
} from './services/plandayService';`;

code = code.replace(targetImport, replaceImport);
fs.writeFileSync('App.tsx', code);
console.log("Updated imports");
