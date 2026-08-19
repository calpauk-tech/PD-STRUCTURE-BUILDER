const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetTableMap = `                        {/* Mapped Data Live Preview Table */}
                        <div>
                            <h3 className="text-md font-bold text-gray-900 mb-3">Preview Mapped Data (First 10 Rows)</h3>
                            <div className="border border-gray-200 rounded-xl overflow-x-auto shadow-sm">
                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-3 w-12 text-center">#</th>
                                            <th className="p-3">Department (Mapped: {deptColumn || 'None'})</th>
                                            <th className="p-3">Employee Group (Mapped: {groupColumn || 'None'})</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {parsedRows.slice(0, 10).map((row, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50/80">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-gray-800">
                                                    {deptColumn && row[deptColumn] ? row[deptColumn] : <span className="text-gray-300 italic">-</span>}
                                                </td>
                                                <td className="p-3 font-medium text-gray-800">
                                                    {groupColumn && row[groupColumn] ? row[groupColumn] : <span className="text-gray-300 italic">-</span>}
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>`;

const replaceTableMap = `                        {/* Mapped Data Live Preview Table */}
                        <div>
                            <h3 className="text-md font-bold text-gray-900 mb-3">Preview Mapped Data (First 10 Rows)</h3>
                            <div className="border border-gray-200 rounded-xl overflow-x-auto shadow-sm">
                                <table className="w-full text-left text-sm whitespace-nowrap">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-3 w-12 text-center">#</th>
                                            <th className="p-3">Department (Mapped: {deptColumn || 'None'})</th>
                                            <th className="p-3">Dept # (Mapped: {deptNumberColumn || 'None'})</th>
                                            <th className="p-3">Group (Mapped: {groupColumn || 'None'})</th>
                                            <th className="p-3">Wage Type (Mapped: {wageTypeColumn || 'None'})</th>
                                            <th className="p-3">Rate (Mapped: {rateColumn || 'None'})</th>
                                            <th className="p-3">Sal. Code (Mapped: {salaryCodeColumn || 'None'})</th>
                                            <th className="p-3">Sal. Code 2 (Mapped: {salaryCode2Column || 'None'})</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {parsedRows.slice(0, 10).map((row, idx) => (
                                            <tr key={idx} className="hover:bg-gray-50/80">
                                                <td className="p-3 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                <td className="p-3 font-medium text-gray-800">{deptColumn && row[deptColumn] ? row[deptColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{deptNumberColumn && row[deptNumberColumn] ? row[deptNumberColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{groupColumn && row[groupColumn] ? row[groupColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{wageTypeColumn && row[wageTypeColumn] ? row[wageTypeColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{rateColumn && row[rateColumn] ? row[rateColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{salaryCodeColumn && row[salaryCodeColumn] ? row[salaryCodeColumn] : <span className="text-gray-300 italic">-</span>}</td>
                                                <td className="p-3 font-medium text-gray-800">{salaryCode2Column && row[salaryCode2Column] ? row[salaryCode2Column] : <span className="text-gray-300 italic">-</span>}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetTableMap);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetTableMap.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceTableMap);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing mapping table");
} else {
    console.log("Failed replacing mapping table");
}
