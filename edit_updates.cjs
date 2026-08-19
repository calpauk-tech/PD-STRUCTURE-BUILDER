const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const target1 = `                    if (item.type === 'department') {
                        // Note: Only rate updates for Employee Groups are implemented currently
                        item.status = 'skipped';
                        item.resultMessage = 'No updates defined for departments';
                    } else {`;

const replace1 = `                    if (item.type === 'department') {
                        const newName = item.updatePayload?.name !== undefined ? item.updatePayload.name : item.name;
                        const newNum = item.updatePayload?.number !== undefined ? item.updatePayload.number : (item.number || null);
                        if (newName !== item.name || newNum !== (item.number || null) || appMode === 'import') {
                            await updateDepartment(item.id, newName, newNum);
                            item.status = 'success';
                            item.resultMessage = 'Updated successfully';
                        } else {
                            item.status = 'skipped';
                            item.resultMessage = 'No updates configured';
                        }
                    } else {`;

const target2 = `                    if (item.type === 'department') {
                        const created = await createDepartment(item.name);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = \`Created (ID: \${created.id})\`;
                    } else {`;

const replace2 = `                    if (item.type === 'department') {
                        const created = await createDepartment(item.name, item.number || null);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = \`Created (ID: \${created.id})\`;
                    } else {`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);

if (codeNoSpace.includes(removeWhitespace(target1))) {
    const regex1 = new RegExp(target1.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'));
    code = code.replace(regex1, replace1);
    console.log("Success target1");
}

if (codeNoSpace.includes(removeWhitespace(target2))) {
    const regex2 = new RegExp(target2.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'));
    code = code.replace(regex2, replace2);
    console.log("Success target2");
}

fs.writeFileSync('App.tsx', code);
