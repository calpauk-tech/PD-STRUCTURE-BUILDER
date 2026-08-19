const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetMapping = `                        {/* Selectors */}
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

const replaceMapping = `                        {/* Selectors */}
                        <div className="grid md:grid-cols-2 gap-6 bg-gray-50 p-6 rounded-2xl border border-gray-200">
                            <div className="space-y-4">
                                <div>
                                    <label className="block text-sm font-bold text-gray-800 mb-2">Departments Column:</label>
                                    <select value={deptColumn} onChange={(e) => setDeptColumn(e.target.value)} className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                        <option value="">-- None / Do Not Import --</option>
                                        {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <label className="block text-sm font-bold text-gray-800 mb-2">Department Number Column:</label>
                                    <select value={deptNumberColumn} onChange={(e) => setDeptNumberColumn(e.target.value)} className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                        <option value="">-- None (Optional) --</option>
                                        {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                    </select>
                                </div>
                            </div>
                            <div className="space-y-4">
                                <div>
                                    <label className="block text-sm font-bold text-gray-800 mb-2">Employee Groups Column:</label>
                                    <select value={groupColumn} onChange={(e) => setGroupColumn(e.target.value)} className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                        <option value="">-- None / Do Not Import --</option>
                                        {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <label className="block text-sm font-bold text-gray-800 mb-2">Default Wage Type (HourlyRate/ShiftRate):</label>
                                    <select value={wageTypeColumn} onChange={(e) => setWageTypeColumn(e.target.value)} className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                        <option value="">-- None (Optional) --</option>
                                        {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <label className="block text-sm font-bold text-gray-800 mb-2">Rate Column:</label>
                                    <select value={rateColumn} onChange={(e) => setRateColumn(e.target.value)} className="w-full p-2.5 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                        <option value="">-- None (Optional) --</option>
                                        {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                    </select>
                                </div>
                                <div className="grid grid-cols-2 gap-2">
                                    <div>
                                        <label className="block text-xs font-bold text-gray-800 mb-2">Salary Code Column:</label>
                                        <select value={salaryCodeColumn} onChange={(e) => setSalaryCodeColumn(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-gray-800 mb-2">Salary Code 2 (Rare):</label>
                                        <select value={salaryCode2Column} onChange={(e) => setSalaryCode2Column(e.target.value)} className="w-full p-2 bg-white border border-gray-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 font-medium">
                                            <option value="">-- None --</option>
                                            {parsedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                        </select>
                                    </div>
                                </div>
                            </div>
                        </div>`;

code = code.replace(targetMapping, replaceMapping);
fs.writeFileSync('App.tsx', code);
console.log("Replaced mapping UI");
