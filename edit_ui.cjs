const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetUI = `                                            <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800 ml-7">
                                                <input 
                                                    type="checkbox" 
                                                    checked={includeWageColumns} 
                                                    onChange={(e) => setIncludeWageColumns(e.target.checked)}
                                                    className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                                                />
                                                <span>Include Default Wage Data columns for updates</span>
                                            </label>`;

const replaceUI = `                                            <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-gray-800 ml-7">
                                                <input 
                                                    type="checkbox" 
                                                    checked={includeWageColumns} 
                                                    onChange={(e) => setIncludeWageColumns(e.target.checked)}
                                                    className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                                                />
                                                <span>Include Default Wage Data columns</span>
                                            </label>`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetUI);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetUI.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceUI);
    fs.writeFileSync('App.tsx', code);
    console.log("Success UI replace");
} else {
    console.log("Failed UI replace");
}

