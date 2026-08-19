const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetUI = `{reviewItems.filter(i => i.type === 'employee_group' && i.action === 'update').length === 0 ? (
                            <div className="p-4 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-sm font-medium">
                                No Employee Groups selected for update. Departments cannot have rate settings applied.
                            </div>
                        ) : (
                            <div className="space-y-4 max-h-[60vh] overflow-y-auto pr-2">
                                {reviewItems.filter(i => i.type === 'employee_group' && i.action === 'update').map(item => (`;

const replaceUI = `{reviewItems.filter(i => i.action === 'update').length === 0 ? (
                            <div className="p-4 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-sm font-medium">
                                No items selected for update.
                            </div>
                        ) : (
                            <div className="space-y-4 max-h-[60vh] overflow-y-auto pr-2">
                                {reviewItems.filter(i => i.type === 'department' && i.action === 'update').map(item => (
                                    <div key={item.id} className="border border-green-100 bg-green-50/20 rounded-xl p-4 shadow-sm">
                                        <h3 className="font-bold text-green-900 mb-3">{item.name}</h3>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Name</label>
                                                <input 
                                                    type="text"
                                                    value={item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'name', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-500"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-gray-700 mb-1">New Number (Optional)</label>
                                                <input 
                                                    type="text"
                                                    value={item.updatePayload?.number !== undefined ? item.updatePayload.number : (item.number || '')}
                                                    onChange={(e) => handleUpdatePayloadChange(item.id, 'number', e.target.value)}
                                                    className="w-full p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-500"
                                                />
                                            </div>
                                        </div>
                                    </div>
                                ))}
                                {reviewItems.filter(i => i.type === 'employee_group' && i.action === 'update').map(item => (`;

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

