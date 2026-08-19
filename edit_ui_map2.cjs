const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetUIMap = `                )}

                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && appMode === 'import' && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <div className="flex justify-between items-center mb-2">
                                <h2 className="text-2xl font-bold text-gray-900">Map File Columns</h2>`;

const replaceUIMap = `                )}

                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && (appMode === 'import' || (appMode === 'update' && deleteInputMethod === 'file')) && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">
                        <div>
                            <div className="flex justify-between items-center mb-2">
                                <h2 className="text-2xl font-bold text-gray-900">Map File Columns</h2>`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetUIMap);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetUIMap.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceUIMap);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing UI Map condition");
} else {
    console.log("Failed replacing UI Map condition");
}
