const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetMappingUI = `                        {/* Selectors */}
                        <div className="grid md:grid-cols-2 gap-6 bg-gray-50 p-6 rounded-2xl border border-gray-200">
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">
                                    Departments Column:
                                </label>
                                <select 
                                    value={deptColumn} 
                                    onChange={(e) => setDeptColumn(e.target.value)}
                                    className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent font-medium"
                                >
                                    <option value="">-- None / Do Not Import --</option>
                                    {parsedHeaders.map(h => (
                                        <option key={h} value={h}>{h}</option>
                                    ))}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">
                                    Employee Groups Column:
                                </label>
                                <select 
                                    value={groupColumn} 
                                    onChange={(e) => setGroupColumn(e.target.value)}
                                    className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent font-medium"
                                >
                                    <option value="">-- None / Do Not Import --</option>
                                    {parsedHeaders.map(h => (
                                        <option key={h} value={h}>{h}</option>
                                    ))}
                                </select>
                            </div>
                        </div>`;

const replaceMappingUI = `                        {/* Selectors */}
                        <div className="grid md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4 bg-gray-50 p-6 rounded-2xl border border-gray-200">
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Departments:</label>
                                <select value={deptColumn} onChange={(e) => setDeptColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Dept Number (Opt):</label>
                                <select value={deptNumberColumn} onChange={(e) => setDeptNumberColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Employee Groups:</label>
                                <select value={groupColumn} onChange={(e) => setGroupColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Default Wage Type:</label>
                                <select value={wageTypeColumn} onChange={(e) => setWageTypeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Rate:</label>
                                <select value={rateColumn} onChange={(e) => setRateColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Salary Code:</label>
                                <select value={salaryCodeColumn} onChange={(e) => setSalaryCodeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div>
                                <label className="block text-sm font-bold text-gray-800 mb-2">Salary Code 2 (Rare):</label>
                                <select value={salaryCode2Column} onChange={(e) => setSalaryCode2Column(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                    <option value="">-- None --</option>
                                    {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                        </div>`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetMappingUI);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetMappingUI.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceMappingUI);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing mapping UI");
} else {
    console.log("Failed replacing mapping UI");
}
