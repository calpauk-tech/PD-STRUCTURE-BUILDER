const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetTable = `                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-4 w-12 text-center">#</th>
                                            <th className="p-4">Type</th>
                                            <th className="p-4">Name</th>
                                            <th className="p-4">Action</th>
                                            <th className="p-4">Status / Details</th>
                                            <th className="p-4 text-center w-16">Remove</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {filteredReviewItems.length > 0 ? (
                                            filteredReviewItems.map((item, idx) => (
                                                <tr key={idx} className="hover:bg-gray-50/80 transition-colors">
                                                    <td className="p-4 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                    <td className="p-4">
                                                        <span className={\`px-2.5 py-1 rounded-full text-xs font-bold \${item.type === 'department' ? 'bg-blue-100 text-blue-800' : 'bg-indigo-100 text-indigo-800'}\`}>
                                                            {item.type === 'department' ? 'Department' : 'Employee Group'}
                                                        </span>
                                                    </td>
                                                    <td className="p-4 font-bold text-gray-900">{item.name}</td>
                                                    <td className="p-4">
                                                        {item.action === 'create' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-green-100 text-green-800">
                                                                Create
                                                            </span>
                                                        )}
                                                        {item.action === 'skip_duplicate' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">
                                                                Skip Duplicate
                                                            </span>
                                                        )}
                                                        {item.action === 'delete' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-red-100 text-red-800">
                                                                Delete
                                                            </span>
                                                        )}
                                                        {item.action === 'update' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">
                                                                Update
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td className="p-4 text-xs text-gray-600">
                                                        {item.resultMessage || (item.action === 'create' ? 'Ready to create' : item.action === 'delete' ? 'Scheduled for deletion' : item.action === 'update' ? 'Scheduled for update' : 'Existing item')}
                                                    </td>
                                                    <td className="p-4 text-center">
                                                        <button 
                                                            onClick={() => handleRemoveReviewItem(idx)}
                                                            className="text-gray-400 hover:text-red-600 transition-colors p-1"
                                                            title="Remove from queue"
                                                        >
                                                            <TrashIcon className="w-4 h-4" />
                                                        </button>
                                                    </td>
                                                </tr>
                                            ))
                                        ) : (
                                            <tr>
                                                <td colSpan={6} className="p-8 text-center text-gray-400 text-sm italic">`;

const replaceTable = `                                <table className="w-full text-left text-sm">
                                    <thead className="bg-gray-100 text-gray-700 uppercase text-xs font-bold border-b border-gray-200">
                                        <tr>
                                            <th className="p-4 w-12 text-center">#</th>
                                            <th className="p-4">Type</th>
                                            <th className="p-4">Name</th>
                                            <th className="p-4">Action</th>
                                            <th className="p-4">Payload Details</th>
                                            <th className="p-4">Status / Details</th>
                                            <th className="p-4 text-center w-16">Remove</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100">
                                        {filteredReviewItems.length > 0 ? (
                                            filteredReviewItems.map((item, idx) => (
                                                <tr key={idx} className="hover:bg-gray-50/80 transition-colors">
                                                    <td className="p-4 text-center text-gray-400 font-mono text-xs">{idx + 1}</td>
                                                    <td className="p-4">
                                                        <span className={\`px-2.5 py-1 rounded-full text-xs font-bold \${item.type === 'department' ? 'bg-blue-100 text-blue-800' : 'bg-indigo-100 text-indigo-800'}\`}>
                                                            {item.type === 'department' ? 'Department' : 'Employee Group'}
                                                        </span>
                                                    </td>
                                                    <td className="p-4 font-bold text-gray-900">{item.name}</td>
                                                    <td className="p-4">
                                                        {item.action === 'create' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-green-100 text-green-800">
                                                                Create
                                                            </span>
                                                        )}
                                                        {item.action === 'skip_duplicate' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">
                                                                Skip Duplicate
                                                            </span>
                                                        )}
                                                        {item.action === 'delete' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-red-100 text-red-800">
                                                                Delete
                                                            </span>
                                                        )}
                                                        {item.action === 'update' && (
                                                            <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">
                                                                Update
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td className="p-4 text-xs text-gray-600">
                                                        {item.type === 'department' && (item.updatePayload?.number || item.number) && (
                                                            <div><span className="font-semibold text-gray-800">Number:</span> {item.updatePayload?.number || item.number}</div>
                                                        )}
                                                        {item.type === 'employee_group' && item.updatePayload ? (
                                                            <div className="space-y-1">
                                                                {item.updatePayload.defaultWageType && <div><span className="font-semibold text-gray-800">Wage Type:</span> {item.updatePayload.defaultWageType}</div>}
                                                                {item.updatePayload.hourlyRate !== undefined && <div><span className="font-semibold text-gray-800">Hourly Rate:</span> {item.updatePayload.hourlyRate}</div>}
                                                                {item.updatePayload.shiftRate !== undefined && <div><span className="font-semibold text-gray-800">Shift Rate:</span> {item.updatePayload.shiftRate}</div>}
                                                                {item.updatePayload.salaryCode && <div><span className="font-semibold text-gray-800">Sal Code:</span> {item.updatePayload.salaryCode}</div>}
                                                                {item.updatePayload.salaryCode2 && <div><span className="font-semibold text-gray-800">Sal Code 2:</span> {item.updatePayload.salaryCode2}</div>}
                                                            </div>
                                                        ) : null}
                                                        {(item.type === 'department' && !(item.updatePayload?.number || item.number)) || (item.type === 'employee_group' && !item.updatePayload) ? (
                                                            <span className="italic text-gray-400">None</span>
                                                        ) : null}
                                                    </td>
                                                    <td className="p-4 text-xs text-gray-600">
                                                        {item.resultMessage || (item.action === 'create' ? 'Ready to create' : item.action === 'delete' ? 'Scheduled for deletion' : item.action === 'update' ? 'Scheduled for update' : 'Existing item')}
                                                    </td>
                                                    <td className="p-4 text-center">
                                                        <button 
                                                            onClick={() => handleRemoveReviewItem(idx)}
                                                            className="text-gray-400 hover:text-red-600 transition-colors p-1"
                                                            title="Remove from queue"
                                                        >
                                                            <TrashIcon className="w-4 h-4" />
                                                        </button>
                                                    </td>
                                                </tr>
                                            ))
                                        ) : (
                                            <tr>
                                                <td colSpan={7} className="p-8 text-center text-gray-400 text-sm italic">`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetTable);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetTable.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceTable);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing review table");
} else {
    console.log("Failed replacing review table");
}
